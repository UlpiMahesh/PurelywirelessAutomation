from __future__ import annotations

import logging
import re
import time
from pathlib import Path
from urllib.parse import unquote

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
LOG_DIR = BASE_DIR / "logs"
SCREENSHOT_DIR = BASE_DIR / "screenshots"
CREDENTIALS_FILE = BASE_DIR.parent / "marketlogins.xlsx"

# Put your existing credentials workbook here:
# data/marketlogins.xlsx
#
# Required columns:
# Market | Username | Password

TMOBILE_URL = "https://www.t-mobiledealerordering.com/"

STATUS_VALUE = "EALL"       # Any
DATE_VALUE = "last_year"    # Last 12 Months
ID_TYPE_VALUE = "OBJECT_ID" # Order Number

INPUT_MARKET_COL = "market"
INPUT_ORDER_COL = "order id"

logger = logging.getLogger("tmobile_tracking")
logger.setLevel(logging.INFO)
logger.handlers.clear()

LOG_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)

log_file = LOG_DIR / f"tracking_{time.strftime('%Y%m%d_%H%M%S')}.log"
fh = logging.FileHandler(log_file, encoding="utf-8")
ch = logging.StreamHandler()
fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
fh.setFormatter(fmt)
ch.setFormatter(fmt)
logger.addHandler(fh)
logger.addHandler(ch)


def log(message: str, level: str = "info") -> None:
    getattr(logger, level)(message)


def new_context(browser):
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
    context.add_init_script(
        "Object.defineProperty(navigator, 'webdriver', { get: () => undefined })"
    )
    return context


def visible_matches(page, selector: str):
    matches = []
    for frame in page.frames:
        try:
            loc = frame.locator(selector)
            for i in range(loc.count()):
                el = loc.nth(i)
                if el.is_visible():
                    matches.append((frame, el))
        except Exception:
            continue
    return matches


def find_in_frames(page, selector: str, timeout: int = 15, last: bool = True):
    for _ in range(timeout * 2):
        matches = visible_matches(page, selector)
        if matches:
            return matches[-1] if last else matches[0]
        time.sleep(0.5)
    return None, None


def read_value(page, selector: str, timeout: int = 10):
    _, el = find_in_frames(page, selector, timeout)
    if el is None:
        return None
    try:
        return el.input_value()
    except Exception:
        return None


def select_value(page, selector: str, value: str, timeout: int = 15):
    _, el = find_in_frames(page, selector, timeout)
    if el is None:
        raise RuntimeError(f"Could not find visible selector: {selector}")
    el.select_option(value)
    return el


def fill_value(page, selector: str, value: str, timeout: int = 15):
    _, el = find_in_frames(page, selector, timeout)
    if el is None:
        raise RuntimeError(f"Could not find visible selector: {selector}")
    el.fill(str(value))
    el.dispatch_event("change")
    return el


def click_element(page, selector: str, timeout: int = 15):
    _, el = find_in_frames(page, selector, timeout)
    if el is None:
        raise RuntimeError(f"Could not find visible selector: {selector}")
    el.click()
    return el


def wait_for_go_ready(page, timeout: int = 30):
    for _ in range(timeout * 2):
        _, el = find_in_frames(page, "#gsbuttonstart", timeout=1)
        if el is not None:
            try:
                cls = el.get_attribute("class") or ""
                if "button-disabled" not in cls:
                    return True
            except Exception:
                pass
        time.sleep(0.5)
    return False


def click_go(page):
    if not wait_for_go_ready(page, 15):
        log("Go button was not ready before search; continuing to click.", "warning")

    _, el = find_in_frames(page, "#gsbuttonstart", timeout=15)
    if el is None:
        raise RuntimeError("Go button not found")

    el.click()
    page.wait_for_timeout(300)

    if not wait_for_go_ready(page, 45):
        raise RuntimeError("Go button did not return to ready state after search")


def classify_login(page):
    text_parts = []
    for frame in page.frames:
        try:
            text_parts.append(frame.locator("body").inner_text(timeout=1500))
        except Exception:
            pass

    text = "\n".join(text_parts).lower()
    url = (page.url or "").lower()

    if "password expired" in text or "password has expired" in text:
        return "PASSWORD EXPIRED"
    if "account locked" in text or "account is locked" in text:
        return "ACCOUNT LOCKED"
    if "invalid password" in text or "invalid user" in text:
        return "BAD CREDENTIALS"
    if "login.do" in url:
        return "LOGIN FAILED"
    return None


def login(page, market: str, username: str, password: str):
    log(f"[{market}] Logging in")

    page.goto(TMOBILE_URL, timeout=60000, wait_until="domcontentloaded")

    page.locator("#userid").fill(str(username))
    page.locator("#password").fill(str(password))

    agree = page.locator("input[name='AgreeTerms']")
    if agree.count() and agree.first.is_visible():
        if not agree.first.is_checked():
            agree.first.check()

    page.locator("a[name='login']").click()
    page.wait_for_timeout(5000)

    status = classify_login(page)
    if status:
        raise RuntimeError(status)

    log(f"[{market}] Login successful")


def prepare_search_panel(page, market: str):
    # These are name-based selectors because the generated numeric IDs change.
    select_value(page, "select[name='rc_status_head1']", STATUS_VALUE)
    select_value(page, "select[name='rc_dateattributes_select']", DATE_VALUE)
    select_value(page, "select[name='rc_attsubcharUI']", ID_TYPE_VALUE)

    if read_value(page, "select[name='rc_status_head1']") != STATUS_VALUE:
        raise RuntimeError("Status did not remain Any (EALL)")

    if read_value(page, "select[name='rc_dateattributes_select']") != DATE_VALUE:
        raise RuntimeError("Creation Date did not remain Last 12 Months")

    log(f"[{market}] Search panel prepared: Status=Any, Creation Date=Last 12 Months")


def search_order(page, market: str, order_id: str):
    prepare_search_panel(page, market)

    fill_value(page, "input[name='rc_object_id']", order_id)
    click_go(page)

    # The portal renders the order number as a link with documentstatusdetailprepare.do.
    deadline = time.time() + 30
    while time.time() < deadline:
        matches = []
        for frame in page.frames:
            try:
                loc = frame.locator(
                    "a[href*='documentstatusdetailprepare.do'][href*='objecttype=order']"
                )
                for i in range(loc.count()):
                    el = loc.nth(i)
                    if not el.is_visible():
                        continue
                    txt = el.inner_text().strip()
                    if txt == str(order_id).strip():
                        matches.append((frame, el))
            except Exception:
                continue

        if matches:
            return matches[-1]

        # A no-result search normally leaves the result area without the order link.
        time.sleep(0.5)

    return None, None


def open_order(page, order_id: str, market: str, retries: int = 3):
    for attempt in range(1, retries + 1):
        log(
            f"[ORDER DETAILS] [{market}] "
            f"Searching order {order_id} "
            f"(attempt {attempt}/{retries})"
        )

        frame, link = search_order(page, market, order_id)

        if link is not None:
            log(
                f"[ORDER DETAILS] [{market}] "
                f"Order {order_id} found on attempt {attempt}"
            )

            link.click()

            deadline = time.time() + 30
            pattern = re.compile(
                rf"Order:\s*{re.escape(str(order_id))}\b",
                re.I,
            )

            while time.time() < deadline:
                for f in page.frames:
                    try:
                        text = f.locator(
                            "body"
                        ).inner_text(timeout=1000)

                        if pattern.search(text):
                            return True

                    except Exception:
                        continue

                time.sleep(0.5)

            log(
                f"[ORDER DETAILS] [{market}] "
                f"Order link found but detail page did not load "
                f"for {order_id}",
                "warning",
            )

        else:
            log(
                f"[ORDER DETAILS] [{market}] "
                f"Order {order_id} not detected on attempt "
                f"{attempt}/{retries}",
                "warning",
            )

        if attempt < retries:
            time.sleep(2)

    return False


TRACKNUM_RE = re.compile(r"tracknum=([^&'\"\\)]+)", re.I)


def extract_tracking_ids(page, market: str, order_id: str):
    """
    Collect every UPS tracking number shown on the order detail page.
    The portal can repeat the same tracking number on several item rows,
    so we deduplicate while preserving first-seen order.
    """
    found = []
    seen = set()

    for frame in page.frames:
        try:
            links = frame.locator("a[onclick*='tracknum=']")
            for i in range(links.count()):
                link = links.nth(i)
                if not link.is_visible():
                    continue

                onclick = link.get_attribute("onclick") or ""
                text = (link.inner_text() or "").strip()

                candidates = TRACKNUM_RE.findall(unquote(onclick))
                if not candidates and text:
                    # Fallback: tracking IDs normally begin with 1Z.
                    candidates = re.findall(r"\b1Z[0-9A-Z]+\b", text, re.I)

                for candidate in candidates:
                    candidate = candidate.strip()
                    if candidate and candidate.upper() not in seen:
                        seen.add(candidate.upper())
                        found.append(candidate)
        except Exception:
            continue

    log(f"[{market}] {order_id}: {len(found)} unique tracking ID(s)")
    return found


def screenshot_failure(page, market: str, order_id: str):
    safe_market = re.sub(r"[^A-Za-z0-9_-]+", "_", market)
    safe_order = re.sub(r"[^A-Za-z0-9_-]+", "_", str(order_id))
    path = SCREENSHOT_DIR / f"{safe_market}_{safe_order}.png"
    try:
        page.screenshot(path=str(path), full_page=True)
        log(f"[{market}] Saved screenshot: {path.name}", "error")
    except Exception:
        pass


def load_orders(input_file: Path):
    df = pd.read_excel(input_file)
    df.columns = [str(c).strip().lower() for c in df.columns]

    required = {INPUT_MARKET_COL, INPUT_ORDER_COL}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            f"Input file is missing required columns: {', '.join(sorted(missing))}. "
            f"Expected columns: Market and Order ID."
        )

    df = df[[INPUT_MARKET_COL, INPUT_ORDER_COL]].copy()
    df[INPUT_MARKET_COL] = df[INPUT_MARKET_COL].fillna("").astype(str).str.strip()
    df[INPUT_ORDER_COL] = (
        df[INPUT_ORDER_COL]
        .fillna("")
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.strip()
    )

    df = df[(df[INPUT_MARKET_COL] != "") & (df[INPUT_ORDER_COL] != "")]
    return df


def load_credentials(credentials_file=None):
    credentials_file = Path(credentials_file) if credentials_file else CREDENTIALS_FILE

    if not credentials_file.exists():
        raise FileNotFoundError(
            f"Credentials file not found: {credentials_file}\n"
            "Expected columns: Market, Username, Password."
        )

    creds = pd.read_excel(credentials_file)
    creds.columns = [str(c).strip().lower() for c in creds.columns]

    required = {"market", "username", "password"}
    missing = required - set(creds.columns)
    if missing:
        raise ValueError(
            f"Credentials file is missing columns: {', '.join(sorted(missing))}. "
            "Expected: Market, Username, Password."
        )

    creds["market_key"] = creds["market"].fillna("").astype(str).str.strip().str.lower()
    return creds.set_index("market_key")


def build_output(rows):
    max_tracking = max((len(r["tracking_ids"]) for r in rows), default=0)

    columns = ["Market", "Order ID"] + [
        f"Tracking ID {i}" for i in range(1, max_tracking + 1)
    ] + ["Status", "Tracking IDs"]

    output_rows = []
    for r in rows:
        tracking_ids = r["tracking_ids"]
        values = [r["market"], r["order_id"]]
        values.extend(tracking_ids)
        values.extend([""] * (max_tracking - len(tracking_ids)))
        values.append(r.get("status", ""))
        values.append(", ".join(tracking_ids))
        output_rows.append(values)

    return pd.DataFrame(output_rows, columns=columns)


def run_tracking(input_file, credentials_file=None, selected_markets=None):
    """
    Run the T-Mobile tracking automation for an uploaded orders workbook.

    Returns the absolute path to the generated Excel report.
    """
    input_file = Path(input_file)
    if not input_file.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    orders = load_orders(input_file)

    if selected_markets:
        selected_keys = {str(m).strip().lower() for m in selected_markets}
        orders = orders[
            orders[INPUT_MARKET_COL].str.lower().isin(selected_keys)
        ].copy()

    if orders.empty:
        raise ValueError("No orders remain after applying the selected markets.")

    credentials = load_credentials(credentials_file)

    log(f"Loaded {len(orders)} order(s) from {input_file.name}")

    results = []

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=[
                "--disable-dev-shm-usage",
                "--disable-blink-features=AutomationControlled",
            ],
        )

        try:
            # One browser session per market. This avoids logging in again for every order.
            for market, market_orders in orders.groupby(INPUT_MARKET_COL, sort=False):
                key = market.lower()

                if key not in credentials.index:
                    for _, row in market_orders.iterrows():
                        results.append({
                            "market": market,
                            "order_id": row[INPUT_ORDER_COL],
                            "tracking_ids": [],
                            "status": "MARKET CREDENTIALS NOT FOUND",
                        })
                    log(f"[{market}] Credentials not found", "error")
                    continue

                cred = credentials.loc[key]
                username = str(cred["username"]).strip()
                password = str(cred["password"])

                context = new_context(browser)
                page = context.new_page()

                try:
                    login(page, market, username, password)

                    _, search_el = find_in_frames(
                        page, "select[name='rc_status_head1']", timeout=30
                    )
                    if search_el is None:
                        raise RuntimeError(
                            "Search panel not found after login. "
                            "The portal may require navigation to Generic Search."
                        )

                    for _, row in market_orders.iterrows():
                        order_id = row[INPUT_ORDER_COL]

                        try:
                            log(f"[{market}] Processing order {order_id}")

                            opened = open_order(page, order_id, market)
                            if not opened:
                                log(f"[{market}] {order_id}: order not found", "warning")
                                results.append({
                                    "market": market,
                                    "order_id": order_id,
                                    "tracking_ids": [],
                                    "status": "ORDER NOT FOUND",
                                })
                                continue

                            tracking_ids = extract_tracking_ids(page, market, order_id)

                            results.append({
                                "market": market,
                                "order_id": order_id,
                                "tracking_ids": tracking_ids,
                                "status": "OK" if tracking_ids else "NO TRACKING ID",
                            })

                        except Exception as exc:
                            log(
                                f"[{market}] {order_id}: {type(exc).__name__}: {exc}",
                                "error",
                            )
                            screenshot_failure(page, market, order_id)

                            results.append({
                                "market": market,
                                "order_id": order_id,
                                "tracking_ids": [],
                                "status": f"ERROR: {type(exc).__name__}",
                            })

                except Exception as exc:
                    log(f"[{market}] Market session failed: {exc}", "error")

                    for _, row in market_orders.iterrows():
                        results.append({
                            "market": market,
                            "order_id": row[INPUT_ORDER_COL],
                            "tracking_ids": [],
                            "status": f"MARKET ERROR: {type(exc).__name__}",
                        })

                finally:
                    context.close()
        finally:
            browser.close()

    output_df = build_output(results)

    output_file = OUTPUT_DIR / f"tracking_results_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
    output_df.to_excel(output_file, index=False)

    log(f"Finished. Output: {output_file}")
    return str(output_file)


def main():
    """Standalone entry point using input/orders.xlsx."""
    input_file = INPUT_DIR / "orders.xlsx"
    return run_tracking(input_file=input_file, credentials_file=CREDENTIALS_FILE)


if __name__ == "__main__":
    main()
