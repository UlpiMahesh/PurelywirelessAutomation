from playwright.sync_api import sync_playwright
import pandas as pd
import re
import time
from pathlib import Path
from openpyxl import Workbook
import uuid
import asyncio
import sys

if sys.platform.startswith("win"):
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

BASE_DIR = Path(__file__).resolve().parent.parent
LOGINS_FILE = BASE_DIR / "marketlogins.xlsx"


# ─────────────────────────────────────────
# 🔹 BROWSER FACTORY
# ─────────────────────────────────────────
def new_browser(p):
    browser = p.chromium.launch(
        headless=True,
        args=[
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-blink-features=AutomationControlled",
        ],
    )
    return browser


def new_page(browser):
    context = browser.new_context(
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/115.0.0.0 Safari/537.36"
        ),
        viewport={"width": 1920, "height": 1080},
        java_script_enabled=True,
    )
    page = context.new_page()
    page.add_init_script(
        "Object.defineProperty(navigator, 'webdriver', { get: () => undefined })"
    )
    return page


# ─────────────────────────────────────────
# 🔹 LOGIN
# ─────────────────────────────────────────
def login(page, username, password):
    page.goto("https://www.t-mobiledealerordering.com/")
    page.fill("#userid", username)
    page.fill("#password", password)
    page.click("input[name='AgreeTerms']")
    page.click("a[name='login']")
    page.wait_for_load_state("load")
    page.wait_for_timeout(5000)

    print(f"[{username}] URL: {page.url}")
    print(f"[{username}] Frames: {[f.url for f in page.frames]}")

    if "login.do" in page.url:
        print(f"[{username}] ❌ LOGIN FAILED")
        return False

    print(f"[{username}] ✅ Login success")
    return True


# ─────────────────────────────────────────
# 🔹 HELPERS
# ─────────────────────────────────────────
def find_in_frames(page, locator_str, timeout=15):
    for _ in range(timeout):
        for frame in page.frames:
            try:
                loc = frame.locator(locator_str)
                if loc.count() > 0:
                    return frame, loc
            except Exception:
                pass
        time.sleep(1)
    return None, None


def click_in_frames(page, locator_str, timeout=15):
    for _ in range(timeout):
        for frame in page.frames:
            try:
                loc = frame.locator(locator_str)
                if loc.count() > 0:
                    loc.first.click()
                    return True
            except Exception:
                pass
        time.sleep(1)
    return False


def parse_price(price_text):
    """
    Parse price text into a float.

    Formats seen:
      "12,979.80 USD"   → 12979.80
      "12,979.80&nbsp;USD" (already stripped by inner_text) → 12979.80
      "USD 12,979.80"   → 12979.80
    Returns float or None if unparseable.
    """
    match = re.search(r"[\d,]+\.\d+", price_text.replace("\xa0", " "))
    if match:
        return float(match.group(0).replace(",", ""))
    return None


# ─────────────────────────────────────────
# 🔹 CATALOG PRICE SCRAPER
# ─────────────────────────────────────────
def scrape_catalog_prices(page, market, item_type, sap_code, timeout=15):
    frame, _ = find_in_frames(page, ".catalauge-item-holder", timeout=timeout)
    if not frame:
        print(f"[{market}] ❌ No catalog items found for type={item_type}")
        return []

    items = frame.locator(".catalauge-item-holder").all()
    devices = []

    for item in items:
        try:
            name = item.locator(".cat-prd-dsc").inner_text().strip()
            sku  = item.locator(".cat-prd-id").inner_text().strip()

            price_locator = item.locator(".cat-prd-prc")
            if price_locator.count() == 0:
                print(f"[{market}] ⚠️ No price element for SKU {sku} — skipping")
                continue

            price_text = price_locator.first.inner_text().strip()
            price = parse_price(price_text)

            if price is None:
                print(f"[{market}] ⚠️ Could not parse price: '{price_text}' for SKU {sku} — skipping")
                continue

            devices.append({
                "Market":   market,
                "SAP Code": sap_code,
                "SKU":      sku,
                "Name":     name,
                "Price":    price,
                "Type":     item_type,
            })

        except Exception as e:
            print(f"[{market}] ⚠️ Error parsing item: {e}")
            continue

    return devices


# ─────────────────────────────────────────
# 🔹 PER-MARKET SCRAPE
# ─────────────────────────────────────────
def scrape_prices_for_market(page, row):
    market   = row["Market"]
    username = row["Username"]
    password = row["Password"]
    sap_code = row.get("SAP Codes", "")

    if not login(page, username, password):
        return []

    frame, loc = find_in_frames(page, "#credithold-tab-msg", timeout=20)
    if frame:
        print(f"[{market}] ✅ Home page loaded")
    else:
        print(f"[{market}] ⚠️ credithold-tab-msg not found — proceeding anyway")

    if not click_in_frames(page, "//a[@onclick='show_catalog_view()']"):
        print(f"[{market}] ❌ Catalog button not found")
        return []

    time.sleep(3)
    catalog_devices = scrape_catalog_prices(page, market, "Catalog", sap_code)
    print(f"[{market}] ✅ Catalog items: {len(catalog_devices)}")

    cpo_selectors = [
        "a:has-text('CPO')",
        "a:has-text('Pre-Owned')",
        "//a[contains(text(),'CPO')]",
        "//a[contains(text(),'Pre-Owned')]",
    ]

    cpo_clicked = False
    for sel in cpo_selectors:
        if click_in_frames(page, sel, timeout=5):
            cpo_clicked = True
            print(f"[{market}] ✅ CPO tab clicked via: {sel}")
            break

    if not cpo_clicked:
        print(f"[{market}] ❌ CPO tab not found — returning catalog only")
        return catalog_devices

    cpo_loaded = False
    for _ in range(20):
        for frame in page.frames:
            try:
                frame_text = frame.inner_text("body", timeout=500)
                has_cpo_text = (
                    "cpo" in frame_text.lower()
                    or "pre-owned" in frame_text.lower()
                    or "certified" in frame_text.lower()
                )
                has_items = frame.locator(".catalauge-item-holder").count() > 0
                if has_cpo_text and has_items:
                    cpo_loaded = True
                    break
            except Exception:
                pass
        if cpo_loaded:
            break
        time.sleep(1)

    if not cpo_loaded:
        print(f"[{market}] ⚠️ CPO page did not load — returning catalog only")
        return catalog_devices

    time.sleep(1)
    cpo_devices = scrape_catalog_prices(page, market, "CPO", sap_code)
    print(f"[{market}] ✅ CPO items: {len(cpo_devices)}")

    return catalog_devices + cpo_devices


# ─────────────────────────────────────────
# 🔹 RUNNER
# ─────────────────────────────────────────
def run_prices(selected_markets=None):
    df = pd.read_excel(LOGINS_FILE)
    df.columns = df.columns.str.strip()

    if selected_markets:
        df = df[df["Market"].str.lower().isin([m.lower() for m in selected_markets])]

    rows = [row for _, row in df.iterrows()]
    results_map = {}

    with sync_playwright() as p:
        browser = new_browser(p)

        for row in rows:
            page = new_page(browser)
            try:
                devices = scrape_prices_for_market(page, row)
                results_map[row["Market"]] = devices
                print(f"[{row['Market']}] Total devices scraped: {len(devices)}")
            finally:
                page.close()

        browser.close()

    output = BASE_DIR / f"data/prices_{uuid.uuid4().hex}.xlsx"
    output.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.append(["SAP Code", "Market", "SKU", "Name", "Price (USD)", "Type"])

    for market, devices in results_map.items():
        for d in devices:
            ws.append([
                d.get("SAP Code", ""),
                d["Market"],
                d["SKU"],
                d["Name"],
                d["Price"],
                d.get("Type", ""),
            ])

    wb.save(output)
    print(f"✅ Prices saved: {output}")
    return str(output)


if __name__ == "__main__":
    run_prices(["Houston"])