"""
Multi-market IMEI export automation for t-mobiledealerordering.com

v3 changes (parallel execution):
  - Markets now run concurrently, each in its own Playwright instance /
    browser / page, via a ThreadPoolExecutor. Nothing is shared across
    threads (no shared Browser, Page, or Playwright object) — that's the
    hard requirement for Playwright's sync API to be safe under threads.
  - MAX_WORKERS controls how many markets run at once. Change the constant
    below, or pass max_workers=N to run_imei_export(). Start low (2-3) and
    watch the logs before pushing higher — concurrent load against the
    same portal/credential set is untested; nothing here confirms the
    portal handles it the same way as serial requests.
  - Logging is now per-market: each market gets its own log file
    (logs/imei_export_<run_id>_<market>.log) so a concurrent run doesn't
    interleave into one unreadable file. Console output is still combined
    but every line is tagged with the market's thread name so you can
    follow one market's story visually.
  - The run-level summary at the end still goes to a single combined
    summary log + console, same as before.
  - Everything from v2 (filter drift re-verification, too-many-documents
    detection, retry logic, screenshot-on-failure) is unchanged — only the
    execution model and logging changed.
"""

from playwright.sync_api import sync_playwright
import pandas as pd
import re
import time
import logging
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from datetime import datetime, timedelta
import asyncio
import sys

if sys.platform.startswith("win"):
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

BASE_DIR = Path(__file__).resolve().parent.parent
LOGINS_FILE = BASE_DIR / "marketlogins.xlsx"
DOWNLOAD_DIR = BASE_DIR / "data" / "imei_exports"
MERGED_OUTPUT = BASE_DIR / "data" / "imei_merged.xlsx"
LOG_DIR = BASE_DIR / "logs"
SCREENSHOT_DIR = BASE_DIR / "logs" / "screenshots"

RUN_ID = datetime.now().strftime("%Y%m%d_%H%M%S")

# ─────────────────────────────────────────
# 🔹 CONCURRENCY — CHANGE THIS TO CONTROL PARALLELISM
# ─────────────────────────────────────────
# Number of markets processed simultaneously. Each worker opens its own
# browser, so 4 workers = 4 headless Chromium instances running at once.
# Bump this up only after you've confirmed the portal doesn't rate-limit,
# CAPTCHA, or kick sessions under concurrent load from the same
# credentials file / IP — that hasn't been tested yet.
MAX_WORKERS = 5

# ─────────────────────────────────────────
# 🔹 LOGGING
# ─────────────────────────────────────────
LOG_DIR.mkdir(parents=True, exist_ok=True)
SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)

# Run-level summary logger — one shared file + console, used only for the
# final cross-market summary block, not per-date chatter.
summary_logger = logging.getLogger("imei_export.summary")
summary_logger.setLevel(logging.DEBUG)
summary_logger.handlers.clear()
summary_logger.propagate = False

_summary_file_handler = logging.FileHandler(LOG_DIR / f"imei_export_{RUN_ID}_summary.log", encoding="utf-8")
_summary_file_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
_summary_console_handler = logging.StreamHandler()
_summary_console_handler.setFormatter(logging.Formatter("%(message)s"))
summary_logger.addHandler(_summary_file_handler)
summary_logger.addHandler(_summary_console_handler)

# Per-market loggers — created lazily, one per market, each with its own
# log file. Cached in a dict guarded by a lock since multiple threads may
# ask for a logger, but each thread only ever touches its own market's
# entry after creation, so there's no contention on the logger object
# itself once it exists.
_market_loggers = {}
_market_logger_lock = threading.Lock()


def get_market_logger(market):
    with _market_logger_lock:
        if market in _market_loggers:
            return _market_loggers[market]

        mlogger = logging.getLogger(f"imei_export.market.{market}")
        mlogger.setLevel(logging.DEBUG)
        mlogger.propagate = False  # don't also bubble up to root/console twice

        file_handler = logging.FileHandler(
            LOG_DIR / f"imei_export_{RUN_ID}_{market}.log", encoding="utf-8"
        )
        file_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))

        console_handler = logging.StreamHandler()
        # threadName lets you visually separate interleaved console output
        # from multiple markets running at once
        console_handler.setFormatter(logging.Formatter(f"[%(threadName)s] %(message)s"))

        mlogger.addHandler(file_handler)
        mlogger.addHandler(console_handler)
        _market_loggers[market] = mlogger
        return mlogger


def log(market, msg, level="info"):
    mlogger = get_market_logger(market)
    getattr(mlogger, level)(f"[{market}] {msg}")


# ─────────────────────────────────────────
# 🔹 BROWSER FACTORY (unchanged)
# ─────────────────────────────────────────
def new_browser(p):
    return p.chromium.launch(
        headless=True,
        args=[
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-blink-features=AutomationControlled",
        ],
    )


def new_page(browser):
    context = browser.new_context(
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/115.0.0.0 Safari/537.36"
        ),
        viewport={"width": 1920, "height": 1080},
        java_script_enabled=True,
        accept_downloads=True,
    )
    page = context.new_page()
    page.add_init_script(
        "Object.defineProperty(navigator, 'webdriver', { get: () => undefined })"
    )
    return page


# ─────────────────────────────────────────
# 🔹 LOGIN
# ─────────────────────────────────────────
def login(page, market, username, password):
    try:
        page.goto("https://www.t-mobiledealerordering.com/")
        page.fill("#userid", username)
        page.fill("#password", password)
        page.click("input[name='AgreeTerms']")
        page.click("a[name='login']")
        page.wait_for_load_state("load")
        page.wait_for_timeout(5000)
    except Exception as e:
        log(market, f"LOGIN FAILED — {e}", "error")
        return False

    if "login.do" in page.url:
        log(market, "LOGIN FAILED — check credentials in marketlogins.xlsx", "error")
        return False

    log(market, "Login success")
    return True


# ─────────────────────────────────────────
# 🔹 FRAME HELPERS — visibility-aware, duplicate-safe
# ─────────────────────────────────────────
def _visible_matches(page, locator_str):
    """Returns list of (frame, locator_handle) for every VISIBLE match across all frames."""
    matches = []
    for frame in page.frames:
        try:
            loc = frame.locator(locator_str)
            n = loc.count()
            for i in range(n):
                el = loc.nth(i)
                if el.is_visible():
                    matches.append((frame, el))
        except Exception:
            pass
    return matches


def find_in_frames(page, locator_str, market="?", timeout=15, warn_on_dup=True):
    for _ in range(timeout):
        matches = _visible_matches(page, locator_str)
        if matches:
            if len(matches) > 1 and warn_on_dup:
                log(market, f"⚠️ {len(matches)} visible matches for '{locator_str}' — "
                             f"panel may be in a duplicated/corrupted state, using the last one", "warning")
            return matches[-1]
        time.sleep(1)
    return None, None


def click_in_frames(page, locator_str, market="?", timeout=15):
    frame, el = find_in_frames(page, locator_str, market, timeout)
    if el is None:
        return False
    el.click()
    return True


def select_in_frames(page, locator_str, value, market="?", timeout=15):
    frame, el = find_in_frames(page, locator_str, market, timeout)
    if el is None:
        return False
    el.select_option(value)
    return True


def fill_in_frames(page, locator_str, value, market="?", timeout=15):
    frame, el = find_in_frames(page, locator_str, market, timeout)
    if el is None:
        return False
    el.fill(value)
    el.dispatch_event("change")
    return True


def read_value_in_frames(page, locator_str, market="?", timeout=10):
    """Returns the current value of a visible select/input, or None if not found."""
    frame, el = find_in_frames(page, locator_str, market, timeout, warn_on_dup=False)
    if el is None:
        return None
    try:
        return el.input_value()
    except Exception:
        return None


def get_results_status(page, market, timeout=15):
    """
    Returns one of:
      ("count", N)     - N documents found (N may be 0)
      ("too_many", None) - "More than 100 documents found" state
      (None, None)      - header never appeared within timeout
    """
    count_pattern = re.compile(r"(\d+)\s+Documents?\s+Found", re.I)
    too_many_pattern = re.compile(r"More than \d+ documents found", re.I)

    for _ in range(timeout):
        for frame in page.frames:
            try:
                for el in frame.locator(".left-panel-search-results-header").all():
                    text = el.inner_text()
                    if too_many_pattern.search(text):
                        return "too_many", None
                    m = count_pattern.search(text)
                    if m:
                        return "count", int(m.group(1))
            except Exception:
                pass
        time.sleep(1)
    return None, None


# ─────────────────────────────────────────
# 🔹 SEARCH PANEL SETUP / STATE VERIFICATION
# ─────────────────────────────────────────
def open_search_panel(page, market, timeout=30):
    frame, el = find_in_frames(page, "select[name='rc_status_head1']", market, timeout, warn_on_dup=False)
    if el is None:
        raise RuntimeError(
            f"Could not find the Status dropdown within {timeout}s after login. "
            f"The search panel may need an extra click to reach — check what's on screen."
        )
    log(market, "Search panel found")


class SearchPanelLostError(RuntimeError):
    """Raised when the search panel itself is gone (not just drifted values) —
    e.g. session died mid-run. No point retrying per-date fixes against a
    panel that isn't there; the caller should abort the rest of that
    market's dates instead of retrying each one 3x for nothing."""
    pass


def wait_for_value(page, locator_str, expected_value, market, timeout=10):
    """
    Polls a select/input's actual value until it matches what we set, or
    times out. select_option() returning doesn't mean the portal's onchange
    handler has finished processing it — confirming the value stuck is the
    only way to know the correction actually registered before we proceed.
    """
    for _ in range(timeout * 2):
        val = read_value_in_frames(page, locator_str, market, timeout=1)
        if val == expected_value:
            return True
        time.sleep(0.5)
    return False


def verify_and_fix_filters(page, market):
    """
    Re-checks Status and Creation Date every time it's called (not just once).
    The portal has been observed silently resetting these mid-run — this
    re-applies them whenever they've drifted instead of assuming they held.

    A value of None (element not found at all) is treated differently from
    a wrong-but-present value: it means the panel itself is gone, not that
    a field drifted. One re-navigation attempt is made; if that fails too,
    this raises SearchPanelLostError so the caller can stop wasting retries
    on a session that isn't coming back.
    """
    status_val = read_value_in_frames(page, "select[name='rc_status_head1']", market)
    if status_val is None:
        log(market, "Status dropdown missing entirely — search panel may be lost, attempting recovery", "warning")
        try:
            open_search_panel(page, market, timeout=15)
        except Exception:
            raise SearchPanelLostError(
                "Search panel disappeared and could not be recovered — session likely dead"
            )
        status_val = read_value_in_frames(page, "select[name='rc_status_head1']", market)
        if status_val is None:
            raise SearchPanelLostError(
                "Search panel still missing after recovery attempt — session likely dead"
            )

    if status_val != "E0004":
        log(market, f"Status drifted to '{status_val}' — re-applying Completed", "warning")
        if not select_in_frames(page, "select[name='rc_status_head1']", "E0004", market):
            raise RuntimeError("Could not re-apply Status=Completed")
        if not wait_for_value(page, "select[name='rc_status_head1']", "E0004", market):
            raise RuntimeError(
                "Status dropdown never confirmed 'Completed' after correction — "
                "would have silently queried under the wrong filter"
            )

    date_val = read_value_in_frames(page, "select[name='rc_dateattributes_select']", market)
    if date_val is None:
        raise SearchPanelLostError("Creation Date dropdown missing entirely — search panel lost")

    if date_val != "att_in_period":
        log(market, f"Creation Date drifted to '{date_val}' — re-applying In Period", "warning")
        if not select_in_frames(page, "select[name='rc_dateattributes_select']", "att_in_period", market):
            raise RuntimeError("Could not re-apply Creation Date=In Period")
        if not wait_for_value(page, "select[name='rc_dateattributes_select']", "att_in_period", market):
            raise RuntimeError(
                "Creation Date dropdown never confirmed 'In Period' after correction — "
                "would have silently queried under the wrong filter"
            )
        frame, el = find_in_frames(page, "input[name='rc_daterange_low']", market, timeout=10, warn_on_dup=False)
        if el is None:
            raise RuntimeError("From/To date inputs never appeared after re-applying In Period")


def set_date_range(page, market, date_str):
    ok_low = fill_in_frames(page, "input[name='rc_daterange_low']", date_str, market)
    ok_high = fill_in_frames(page, "input[name='rc_daterange_high']", date_str, market)
    if not (ok_low and ok_high):
        raise RuntimeError(f"Could not set date range to {date_str}")


def wait_for_go_button_ready(page, market, timeout=30):
    """
    The Go button's own onclick sets class='button-disabled' the instant
    it's clicked, and (per observed behavior) resets once the portal's
    request actually finishes. Polling this is the real completion signal —
    a fixed sleep either reads stale results (too short) or clicks Go again
    while a request is still "in progress" server-side (too long-a-gap
    between checks, or clicking before ready).
    """
    for _ in range(timeout * 2):  # check every 0.5s
        frame, el = find_in_frames(page, "#gsbuttonstart", market, timeout=1, warn_on_dup=False)
        if el is not None:
            try:
                cls = el.get_attribute("class") or ""
                if "button-disabled" not in cls:
                    return True
            except Exception:
                pass
        time.sleep(0.5)
    log(market, "Go button never returned to ready state within timeout", "warning")
    return False


def click_go(page, market):
    # If a previous request left the button disabled, wait it out before
    # clicking — clicking a disabled button is what produces "Request
    # already started" and hangs for 30s doing nothing useful.
    wait_for_go_button_ready(page, market, timeout=15)

    if not click_in_frames(page, "#gsbuttonstart", market):
        raise RuntimeError("Go button not found")

    # Give the click a moment to actually register as "started" before we
    # start polling for "finished" — otherwise we might check before the
    # class has even flipped to button-disabled and think it's already done.
    page.wait_for_timeout(300)

    if not wait_for_go_button_ready(page, market, timeout=30):
        raise RuntimeError("Go button stayed disabled after click — request may be stuck")


def select_all_results(page, market):
    if not click_in_frames(page, "#orderChkAll", market, timeout=10):
        raise RuntimeError("Select-all checkbox not found despite results being present")


def export_and_download(page, market, date_str, out_dir, timeout=60):
    frame, el = find_in_frames(page, "a.search-excel-icon", market, timeout=10)
    if el is None:
        raise RuntimeError("Export button not found despite results being present")

    with page.expect_download(timeout=timeout * 1000) as download_info:
        el.click()
    download = download_info.value

    suffix = Path(download.suggested_filename).suffix or ".xls"
    safe_date = date_str.replace("/", "-")
    out_path = out_dir / f"{market}_{safe_date}{suffix}"
    download.save_as(out_path)
    log(market, f"Downloaded {out_path.name}")
    return out_path


def save_failure_screenshot(page, market, date_str):
    safe_date = date_str.replace("/", "-")
    path = SCREENSHOT_DIR / f"{market}_{safe_date}_{RUN_ID}.png"
    try:
        page.screenshot(path=str(path), full_page=True)
        log(market, f"Saved failure screenshot: {path}", "error")
    except Exception as e:
        log(market, f"Could not save screenshot: {e}", "error")


# ─────────────────────────────────────────
# 🔹 DATE RANGE HELPER
# ─────────────────────────────────────────
def daterange(start_str, end_str):
    start = datetime.strptime(start_str, "%m/%d/%Y")
    end = datetime.strptime(end_str, "%m/%d/%Y")
    if end < start:
        raise ValueError("End date is before start date")
    current = start
    while current <= end:
        yield current.strftime("%m/%d/%Y")
        current += timedelta(days=1)


# ─────────────────────────────────────────
# 🔹 PER-DATE PROCESSING (with retry + state recovery)
# ─────────────────────────────────────────
def process_date(page, market, date_str, out_dir, max_attempts=3):
    """
    Returns (status, detail):
      status in {"success", "no_orders", "failed"}
      detail = file path (success), None (no_orders), or error string (failed)
    """
    last_error = None
    for attempt in range(1, max_attempts + 1):
        try:
            verify_and_fix_filters(page, market)
            set_date_range(page, market, date_str)
            click_go(page, market)

            kind, count = get_results_status(page, market, timeout=10)

            if kind is None:
                # click_go already confirmed the request completed, so a
                # missing header here means the portal renders nothing for
                # a genuine zero-result day — not a broken click. Don't
                # burn retries on it.
                log(market, f"{date_str}: no results header (no orders)")
                return "no_orders", None

            if kind == "too_many":
                log(market, f"{date_str}: 'more than 100 documents' — filters likely drifted, "
                             f"retrying (attempt {attempt}/{max_attempts})", "warning")
                last_error = "more than 100 documents found (filter state drifted)"
                continue

            if count == 0:
                log(market, f"{date_str}: no orders")
                return "no_orders", None

            log(market, f"{date_str}: {count} documents found")
            select_all_results(page, market)
            path = export_and_download(page, market, date_str, out_dir)
            return "success", path

        except SearchPanelLostError:
            # Not worth retrying — the panel is gone, not just wrong. Let
            # this propagate so export_market can abort the rest of this
            # market's dates instead of repeating the same failed check.
            save_failure_screenshot(page, market, date_str)
            raise

        except Exception as e:
            last_error = str(e)
            log(market, f"{date_str}: attempt {attempt}/{max_attempts} failed — {e}", "error")
            if attempt == max_attempts:
                save_failure_screenshot(page, market, date_str)
            else:
                page.wait_for_timeout(2000)

    return "failed", last_error


# ─────────────────────────────────────────
# 🔹 PER-MARKET EXPORT
# ─────────────────────────────────────────
def export_market(page, row, start_date, end_date, out_dir):
    market = row["Market"]
    username = row["Username"]
    password = row["Password"]

    result = {"market": market, "success": [], "no_orders": [], "failed": {}, "login_failed": False}

    if not login(page, market, username, password):
        result["login_failed"] = True
        return result

    try:
        open_search_panel(page, market)
    except Exception as e:
        log(market, f"Could not open search panel — aborting market: {e}", "error")
        save_failure_screenshot(page, market, "search_panel")
        result["failed"]["search_panel"] = str(e)
        return result

    for date_str in daterange(start_date, end_date):
        try:
            status, detail = process_date(page, market, date_str, out_dir)
        except SearchPanelLostError as e:
            log(market, f"Aborting remaining dates for this market — {e}", "error")
            remaining = [d for d in daterange(date_str, end_date)]
            for d in remaining:
                result["failed"][d] = "search panel lost — market session aborted"
            break

        if status == "success":
            result["success"].append(detail)
        elif status == "no_orders":
            result["no_orders"].append(date_str)
        else:
            result["failed"][date_str] = detail

    return result


# ─────────────────────────────────────────
# 🔹 PARALLEL WORKER — one per market, fully isolated Playwright instance
# ─────────────────────────────────────────
def run_market_isolated(row, start_date, end_date, out_dir):
    """
    Runs a single market end-to-end inside its own thread, with its own
    sync_playwright() instance, Browser, and Page. Nothing here is shared
    with any other thread — that isolation is what makes the sync
    Playwright API safe to use under a ThreadPoolExecutor. Never pass a
    Browser/Page/Playwright object created in one thread into another.
    """
    market = row.get("Market", "UNKNOWN")
    threading.current_thread().name = market  # shows up in console log lines

    with sync_playwright() as p:
        browser = new_browser(p)
        page = new_page(browser)
        try:
            result = export_market(page, row, start_date, end_date, out_dir)
        except Exception as e:
            log(market, f"Market aborted by unexpected error — {e}", "error")
            save_failure_screenshot(page, market, "market_level_crash")
            result = {
                "market": market, "success": [], "no_orders": [],
                "failed": {"MARKET_LEVEL": str(e)}, "login_failed": False,
            }
        finally:
            page.close()
            browser.close()
        return result


# ─────────────────────────────────────────
# 🔹 MERGE
# ─────────────────────────────────────────
def merge_exports(file_paths, output_path, interactive=True):
    if not file_paths:
        summary_logger.warning("No files to merge")
        return None

    frames = []
    reference_columns = None

    for path in file_paths:
        df = pd.read_excel(path)
        if reference_columns is None:
            reference_columns = list(df.columns)
        elif list(df.columns) != reference_columns:
            raise RuntimeError(
                f"Schema mismatch: {path.name} has columns {list(df.columns)}, "
                f"expected {reference_columns}. Refusing to merge blindly."
            )
        market = path.stem.split("_")[0]
        df.insert(0, "Market", market)
        frames.append(df)

    merged = pd.concat(frames, ignore_index=True)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    final_path = write_excel_with_retry(merged, output_path, interactive=interactive)
    summary_logger.info(f"Merged {len(file_paths)} files -> {final_path} ({len(merged)} rows)")
    return final_path


def write_excel_with_retry(df, output_path, interactive=True):
    """
    All the scraping already succeeded by the time we get here — a locked
    output file shouldn't throw that work away.

    interactive=True  (console use): waits on input() for the user to
        close the file, then retries the SAME path indefinitely.
    interactive=False (Streamlit/web use): input() has no console to read
        from and would hang the session with no visible error, so instead
        this falls back once to a timestamped filename and returns that
        path — the caller (UI) is responsible for telling the user the
        name changed.
    """
    if interactive:
        while True:
            try:
                df.to_excel(output_path, index=False)
                return output_path
            except PermissionError:
                summary_logger.error(
                    f"Cannot write {output_path} — it looks like it's open "
                    f"(e.g. in Excel). Close the file, then press Enter to retry."
                )
                input()
    else:
        try:
            df.to_excel(output_path, index=False)
            return output_path
        except PermissionError:
            fallback = output_path.with_name(
                f"{output_path.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{output_path.suffix}"
            )
            summary_logger.warning(
                f"{output_path} is locked (likely open elsewhere) — "
                f"saved to {fallback} instead"
            )
            df.to_excel(fallback, index=False)
            return fallback


# ─────────────────────────────────────────
# 🔹 RUNNER
# ─────────────────────────────────────────
def run_imei_export(start_date, end_date, selected_markets=None, interactive=True, max_workers=MAX_WORKERS):
    df = pd.read_excel(LOGINS_FILE)
    df.columns = df.columns.str.strip()

    if selected_markets:
        df = df[df["Market"].str.lower().isin([m.lower() for m in selected_markets])]

    rows = [row for _, row in df.iterrows()]
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

    all_results = []
    all_files = []

    # Each market gets its own thread, its own Playwright instance, its own
    # browser — max_workers caps how many run at the same time. Threads
    # that finish before others just pick up the next queued market.
    with ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="mkt") as executor:
        futures = {
            executor.submit(run_market_isolated, row, start_date, end_date, DOWNLOAD_DIR): row.get("Market", "UNKNOWN")
            for row in rows
        }
        for future in as_completed(futures):
            market = futures[future]
            try:
                result = future.result()
            except Exception as e:
                # Should be rare — run_market_isolated already catches its
                # own exceptions — but don't let one bad market silently
                # drop out of the summary if something slips through.
                summary_logger.error(f"[{market}] Unhandled exception escaped worker thread — {e}")
                result = {
                    "market": market, "success": [], "no_orders": [],
                    "failed": {"WORKER_LEVEL": str(e)}, "login_failed": False,
                }
            all_results.append(result)
            all_files.extend(result["success"])

    # ── Summary ──
    summary_logger.info("=" * 60)
    summary_logger.info("RUN SUMMARY")
    summary_logger.info("=" * 60)
    for r in all_results:
        if r["login_failed"]:
            summary_logger.info(f"{r['market']}: LOGIN FAILED")
            continue
        summary_logger.info(
            f"{r['market']}: {len(r['success'])} downloaded, "
            f"{len(r['no_orders'])} no-orders, {len(r['failed'])} failed"
        )
        if r["failed"]:
            for date_str, reason in r["failed"].items():
                summary_logger.info(f"    FAILED {date_str}: {reason}")
    summary_logger.info("=" * 60)

    merged_path = merge_exports(all_files, MERGED_OUTPUT, interactive=interactive)
    return {"path": merged_path, "results": all_results}


if __name__ == "__main__":
    # max_workers here overrides MAX_WORKERS for this run only, if you want
    # a quick one-off change without editing the constant above.
    run_imei_export("08/31/2026", "09/06/2026", ["CO", "IN","MA","NY","VA","MDDC"], max_workers=MAX_WORKERS)