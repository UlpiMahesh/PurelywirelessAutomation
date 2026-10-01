#!/usr/bin/env python3
"""
Download the "Sales Report By Manufacturer" raw Excel export from RT POS
(https://www.myrtpos.com/newbdi/) for a given date range.

Usage examples
--------------
    # credentials from env vars RTPOS_USER / RTPOS_PASS, or prompted if unset
    python rtpos_sales_download.py --start 22-Sep-2026 --end 29-Sep-2026

    # watch the browser, custom output folder
    python rtpos_sales_download.py --start 22-Sep-2026 --end 29-Sep-2026 \
        --headed --out-dir D:\\reports

    # extra filters (name = the form field's `name` attribute on the page)
    python rtpos_sales_download.py --start 2026-09-22 --end 2026-09-29 \
        --filter frmMarketID=HOUSTON

Setup
-----
    pip install playwright openpyxl
    playwright install chromium

Exit codes
----------
    0 success | 1 unexpected error | 2 bad input | 3 login failed
    4 gave up after retries | 130 interrupted
"""
from __future__ import annotations

import argparse
import functools
import getpass
import logging
import logging.handlers
import os
import sys
import time
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Callable, TypeVar

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import Page
from playwright.sync_api import TimeoutError as PlaywrightTimeout
from playwright.sync_api import sync_playwright

# --------------------------------------------------------------------------- #
# Constants
# --------------------------------------------------------------------------- #
BASE_URL = "https://www.myrtpos.com/newbdi/"
REPORT_URL = BASE_URL + "SalesReport_ByManufacturer.fwx"
REPORT_NAME = "sales_by_manufacturer"

# Selectors taken from the site's HTML.
SEL_USER = "#secUserID"
SEL_PASS = "#secPassword"
SEL_SAVE_COOKIE = "#secSaveCookie"
SEL_LOGIN_SUBMIT = "input[type=submit][value=Login]"
SEL_DATE_START = "#frmStart"
SEL_DATE_END = "#frmEnd"

DATE_FORMATS = ("%d-%b-%Y", "%Y-%m-%d", "%m/%d/%Y")
SITE_DATE_FORMAT = "%m/%d/%Y"  # format the site's hidden inputs expect
MAX_RANGE_DAYS_DEFAULT = 93
MIN_FILE_BYTES = 200

log = logging.getLogger("rtpos")


# --------------------------------------------------------------------------- #
# Exceptions. Only TransientError is retried.
# --------------------------------------------------------------------------- #
class RtposError(Exception):
    exit_code = 1


class ValidationError(RtposError):
    """Bad user input or an unexpected page structure. Retrying won't help."""
    exit_code = 2


class AuthError(RtposError):
    """Login rejected. Never retried, to avoid locking the account."""
    exit_code = 3


class TransientError(RtposError):
    """Timeouts, network errors, expired sessions, 5xx. Safe to retry."""
    exit_code = 4


# --------------------------------------------------------------------------- #
# Config and input validation
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class Config:
    start: date
    end: date
    out_dir: Path
    log_dir: Path
    headless: bool
    reuse_session: bool
    state_file: Path
    filters: dict
    retries: int
    retry_delay: float
    timeout_ms: int
    download_timeout_ms: int


def parse_date(text: str) -> date:
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text.strip(), fmt).date()
        except ValueError:
            continue
    raise ValidationError(
        f"Unrecognised date {text!r}. Use DD-Mon-YYYY (22-Sep-2026), "
        "YYYY-MM-DD or MM/DD/YYYY."
    )


def parse_filters(items: list[str] | None) -> dict:
    filters = {}
    for item in items or []:
        name, sep, value = item.partition("=")
        if not sep or not name.strip():
            raise ValidationError(f"Bad --filter {item!r}; expected NAME=VALUE.")
        filters[name.strip()] = value.strip()
    return filters


def build_config(args: argparse.Namespace) -> Config:
    start, end = parse_date(args.start), parse_date(args.end)
    if start > end:
        raise ValidationError(f"Start date {start} is after end date {end}.")
    span = (end - start).days + 1
    if span > args.max_days:
        raise ValidationError(
            f"Range is {span} days; the limit is {args.max_days} "
            "(raise it with --max-days if intended)."
        )
    if end > date.today():
        log.warning("End date %s is in the future; data may be incomplete.", end)

    out_dir = Path(args.out_dir).expanduser().resolve()
    log_dir = Path(args.log_dir).expanduser().resolve()
    for d in (out_dir, log_dir):
        try:
            d.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            raise ValidationError(f"Cannot create folder {d}: {exc}") from exc
        if not os.access(d, os.W_OK):
            raise ValidationError(f"Folder is not writable: {d}")

    if args.retries < 1:
        raise ValidationError("--retries must be at least 1.")

    return Config(
        start=start,
        end=end,
        out_dir=out_dir,
        log_dir=log_dir,
        headless=not args.headed,
        reuse_session=args.reuse_session,
        state_file=log_dir / ".rtpos_session.json",
        filters=parse_filters(args.filter),
        retries=args.retries,
        retry_delay=args.retry_delay,
        timeout_ms=args.timeout * 1000,
        download_timeout_ms=args.download_timeout * 1000,
    )


def get_credentials() -> tuple[str, str]:
    """Env vars first; otherwise prompt. (Swap this out when you pick a
    proper secret store, e.g. Windows Credential Manager / keyring.)"""
    user = os.environ.get("RTPOS_USER", "").strip()
    pwd = os.environ.get("RTPOS_PASS", "")
    if not user or not pwd:
        if not sys.stdin.isatty():
            raise ValidationError(
                "No credentials: set RTPOS_USER and RTPOS_PASS, or run interactively."
            )
        user = user or input("RT POS username: ").strip()
        pwd = pwd or getpass.getpass("RT POS password: ")
    if not user or not pwd:
        raise ValidationError("Username and password must not be empty.")
    return user, pwd


# --------------------------------------------------------------------------- #
# Retry helper
# --------------------------------------------------------------------------- #
T = TypeVar("T")


def with_retries(attempts: int, base_delay: float) -> Callable:
    """Retry only TransientError, with exponential backoff."""

    def decorator(fn: Callable[..., T]) -> Callable[..., T]:
        @functools.wraps(fn)
        def wrapper(*args, **kwargs) -> T:
            for attempt in range(1, attempts + 1):
                try:
                    return fn(*args, attempt=attempt, **kwargs)
                except TransientError as exc:
                    if attempt == attempts:
                        raise TransientError(
                            f"Failed after {attempts} attempts: {exc}"
                        ) from exc
                    delay = base_delay * (2 ** (attempt - 1))
                    log.warning(
                        "Attempt %d/%d failed: %s. Retrying in %.0fs.",
                        attempt, attempts, exc, delay,
                    )
                    time.sleep(delay)
            raise AssertionError("unreachable")

        return wrapper

    return decorator


# --------------------------------------------------------------------------- #
# Page interactions
# --------------------------------------------------------------------------- #
def is_login_page(page: Page) -> bool:
    return page.locator(SEL_PASS).count() > 0


def goto_checked(page: Page, url: str) -> None:
    resp = page.goto(url, wait_until="domcontentloaded")
    if resp is None:
        return
    if resp.status >= 500:
        raise TransientError(f"Server error {resp.status} for {url}")
    if resp.status >= 400:
        raise ValidationError(f"HTTP {resp.status} for {url} (wrong URL or no access?)")


def login(page: Page, creds: tuple[str, str], keep_cookie: bool) -> None:
    user, pwd = creds
    log.info("Logging in.")
    page.locator(SEL_USER).fill(user)
    page.locator(SEL_PASS).fill(pwd)
    page.locator(SEL_SAVE_COOKIE).set_checked(keep_cookie)
    with page.expect_navigation(wait_until="domcontentloaded"):
        page.locator(SEL_LOGIN_SUBMIT).click()
    if is_login_page(page):
        raise AuthError("Login rejected (still on the login page). Check credentials.")


def open_report(page: Page, cfg: Config, creds: tuple[str, str]) -> None:
    goto_checked(page, REPORT_URL)
    if is_login_page(page):
        login(page, creds, keep_cookie=cfg.reuse_session)
        goto_checked(page, REPORT_URL)
        if is_login_page(page):
            raise AuthError("Logged in but the report redirected back to login.")
    try:
        page.locator(SEL_DATE_START).wait_for(state="attached")
        page.locator(SEL_DATE_END).wait_for(state="attached")
    except PlaywrightTimeout as exc:
        raise ValidationError(
            "Date fields (#frmStart/#frmEnd) not found; the report page layout "
            f"may have changed. URL: {page.url}"
        ) from exc
    log.info("Report page loaded.")


_JS_SET_DATES = """
([s, e]) => {
  const a = document.querySelector('#frmStart');
  const b = document.querySelector('#frmEnd');
  a.value = s; b.value = e;
  try {  // keep the visible picker in sync; purely cosmetic
    if (window.jQuery && window.moment) {
      const p = jQuery('#reportrange').data('daterangepicker');
      if (p) { p.setStartDate(moment(s, 'MM/DD/YYYY')); p.setEndDate(moment(e, 'MM/DD/YYYY')); }
      const span = document.querySelector('#reportrange span');
      if (span) span.textContent = moment(s, 'MM/DD/YYYY').format('MMMM D, YYYY') +
                                   ' - ' + moment(e, 'MM/DD/YYYY').format('MMMM D, YYYY');
    }
  } catch (err) {}
  return [a.value, b.value];
}
"""


def set_date_range(page: Page, start: date, end: date) -> None:
    s, e = start.strftime(SITE_DATE_FORMAT), end.strftime(SITE_DATE_FORMAT)
    got = page.evaluate(_JS_SET_DATES, [s, e])
    if list(got) != [s, e]:
        raise ValidationError(f"Date fields did not take the value: wanted {[s, e]}, got {got}")
    log.info("Date range set: %s to %s (inclusive).", s, e)


def apply_filters(page: Page, filters: dict) -> None:
    for name, value in filters.items():
        loc = page.locator(f'[name="{name}"]')
        if loc.count() != 1:
            raise ValidationError(f"Filter field {name!r} not found exactly once on the page.")
        tag = loc.evaluate("el => el.tagName.toLowerCase()")
        if tag == "select":
            options = loc.evaluate("el => Array.from(el.options).map(o => o.value)")
            match = next((o for o in options if o.strip() == value), None)
            if match is None:
                sample = ", ".join(o.strip() for o in options[:15])
                raise ValidationError(
                    f"{value!r} is not an option of {name!r}. Examples: {sample}"
                )
            # force: the real <select> is visually hidden behind a Select2 widget
            loc.select_option(value=match, force=True)
        else:
            loc.fill(value)
        log.info("Filter set: %s = %s", name, value)


def sniff_format(path: Path) -> str:
    """Return 'xlsx', 'xls', 'html' or raise. Also catches login pages."""
    with path.open("rb") as fh:
        head = fh.read(4096)
    if head[:2] == b"PK":
        return "xlsx"
    if head[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "xls"
    text = head.decode("utf-8", errors="ignore").lower()
    if "secpassword" in text or "secuserid" in text:
        raise TransientError("Downloaded file is the login page (session expired).")
    if text.lstrip().startswith("<"):
        return "html"  # some legacy systems ship HTML tables as .xls
    raise ValidationError("Downloaded file is not a recognised Excel format.")


def validate_download(path: Path) -> str:
    size = path.stat().st_size
    if size < MIN_FILE_BYTES:
        raise TransientError(f"Downloaded file is suspiciously small ({size} bytes).")
    fmt = sniff_format(path)
    if fmt == "xlsx":
        try:
            import openpyxl  # optional deep check

            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            try:
                rows = sum(ws.max_row or 0 for ws in wb.worksheets)
            finally:
                wb.close()
            if rows <= 1:
                log.warning("Workbook has no data rows (header only?). Empty range?")
            else:
                log.info("Workbook OK: ~%d rows across %d sheet(s).", rows, len(wb.worksheets))
        except ImportError:
            log.info("openpyxl not installed; skipped deep workbook check.")
        except Exception as exc:  # corrupt zip etc.
            raise TransientError(f"Workbook failed to open: {exc}") from exc
    else:
        log.warning("Export is %s, not .xlsx. Saved as-is.", fmt.upper())
    return fmt


def download_excel(page: Page, cfg: Config) -> Path:
    button = page.locator("span.button-inner", has_text="Raw Excel")
    if button.count() != 1:
        raise ValidationError("'Raw Excel' button not found exactly once; layout changed?")

    page.on("dialog", lambda d: (log.warning("Page dialog: %s", d.message), d.accept()))

    log.info("Requesting export (timeout %ds).", cfg.download_timeout_ms // 1000)
    with page.expect_download(timeout=cfg.download_timeout_ms) as info:
        button.click()
    download = info.value
    if download.failure():
        raise TransientError(f"Download failed: {download.failure()}")

    temp = cfg.out_dir / f".{REPORT_NAME}.{os.getpid()}.part"
    download.save_as(temp)
    try:
        fmt = validate_download(temp)
    except Exception:
        temp.unlink(missing_ok=True)
        raise

    ext = ".xlsx" if fmt == "xlsx" else ".xls"
    final = cfg.out_dir / f"{REPORT_NAME}_{cfg.start:%Y-%m-%d}_to_{cfg.end:%Y-%m-%d}{ext}"
    if final.exists():
        log.warning("Overwriting existing file %s", final.name)
    os.replace(temp, final)  # atomic: no half-written final file
    return final


def save_failure_screenshot(page: Page, cfg: Config, attempt: int) -> None:
    try:
        shot = cfg.log_dir / f"failure_{datetime.now():%Y%m%d_%H%M%S}_a{attempt}.png"
        page.screenshot(path=str(shot), full_page=True)
        log.error("Saved failure screenshot: %s", shot)
    except Exception:  # never let diagnostics mask the real error
        pass


# --------------------------------------------------------------------------- #
# One full attempt (fresh browser each time, so retries start clean)
# --------------------------------------------------------------------------- #
def make_runner(cfg: Config, creds: tuple[str, str]) -> Callable[[], Path]:
    @with_retries(cfg.retries, cfg.retry_delay)
    def run(attempt: int) -> Path:
        log.info("=== Attempt %d/%d ===", attempt, cfg.retries)
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=cfg.headless)
            try:
                ctx_args = {"accept_downloads": True}
                if cfg.reuse_session and cfg.state_file.exists():
                    ctx_args["storage_state"] = str(cfg.state_file)
                context = browser.new_context(**ctx_args)
                context.set_default_timeout(cfg.timeout_ms)
                page = context.new_page()
                try:
                    open_report(page, cfg, creds)
                    if cfg.reuse_session:
                        context.storage_state(path=str(cfg.state_file))
                        try:
                            os.chmod(cfg.state_file, 0o600)
                        except OSError:
                            pass
                    set_date_range(page, cfg.start, cfg.end)
                    apply_filters(page, cfg.filters)
                    return download_excel(page, cfg)
                except PlaywrightTimeout as exc:
                    save_failure_screenshot(page, cfg, attempt)
                    raise TransientError(f"Timed out: {str(exc).splitlines()[0]}") from exc
                except PlaywrightError as exc:
                    save_failure_screenshot(page, cfg, attempt)
                    raise TransientError(f"Browser/network error: {str(exc).splitlines()[0]}") from exc
                except RtposError:
                    save_failure_screenshot(page, cfg, attempt)
                    raise
            finally:
                browser.close()

    return run


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def setup_logging(log_dir: Path) -> None:
    fmt = logging.Formatter("%(asctime)s %(levelname)-7s %(message)s")
    root = logging.getLogger("rtpos")
    root.setLevel(logging.INFO)
    root.handlers.clear()
    console = logging.StreamHandler()
    console.setFormatter(fmt)
    root.addHandler(console)
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
        fh = logging.handlers.RotatingFileHandler(
            log_dir / "rtpos_download.log", maxBytes=1_000_000, backupCount=5, encoding="utf-8"
        )
        fh.setFormatter(fmt)
        root.addHandler(fh)
    except OSError as exc:
        root.warning("File logging disabled: %s", exc)


def parse_args(argv: list[str] | None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Download the RT POS 'Sales Report By Manufacturer' raw Excel export.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("--start", required=True, help="Start date, inclusive (e.g. 22-Sep-2026)")
    p.add_argument("--end", required=True, help="End date, inclusive (e.g. 29-Sep-2026)")
    p.add_argument("--out-dir", default="downloads", help="Where to save the export")
    p.add_argument("--log-dir", default="logs", help="Logs and failure screenshots")
    p.add_argument("--filter", action="append", metavar="NAME=VALUE",
                   help="Extra form filter, repeatable (e.g. frmMarketID=HOUSTON)")
    p.add_argument("--headed", action="store_true", help="Show the browser window")
    p.add_argument("--reuse-session", action="store_true",
                   help="Cache the login cookie on disk to skip login next run")
    p.add_argument("--retries", type=int, default=3, help="Max attempts")
    p.add_argument("--retry-delay", type=float, default=5.0, help="Base backoff seconds (doubles each retry)")
    p.add_argument("--timeout", type=int, default=30, help="Per-action timeout, seconds")
    p.add_argument("--download-timeout", type=int, default=180, help="Export wait, seconds")
    p.add_argument("--max-days", type=int, default=MAX_RANGE_DAYS_DEFAULT, help="Safety cap on range length")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    setup_logging(Path(args.log_dir).expanduser())
    try:
        cfg = build_config(args)
        creds = get_credentials()
        path = make_runner(cfg, creds)()
        log.info("SUCCESS: saved %s (%.1f KB)", path, path.stat().st_size / 1024)
        return 0
    except RtposError as exc:
        log.error("%s: %s", type(exc).__name__, exc)
        return exc.exit_code
    except KeyboardInterrupt:
        log.error("Interrupted.")
        return 130
    except Exception:
        log.exception("Unexpected error")
        return 1


if __name__ == "__main__":
    sys.exit(main())
    #Arnoldsealar3