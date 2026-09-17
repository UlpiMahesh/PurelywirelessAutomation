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
# 🔹 BROWSER FACTORY  (consistent anti-bot setup everywhere)
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
    # Remove webdriver flag
    page.add_init_script(
        "Object.defineProperty(navigator, 'webdriver', { get: () => undefined })"
    )
    return page


# ─────────────────────────────────────────
# 🔹 COMMON LOGIN
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
    """
    Poll all frames up to `timeout` seconds looking for a locator with count > 0.
    Returns (frame, locator) or (None, None).
    """
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
    """Click the first match of locator_str across all frames. Returns True on success."""
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


def parse_allocation(alloc_text):
    """
    Parse allocation text into (available, total).

    Formats seen in the wild:
      "Allocation : 57 of 113 EA"   → available=57,  total=113  ← has remaining stock
      "Allocation : 113 EA"         → available=0,   total=113  ← fully allocated (skip)
      "57 / 113"                    → available=57,  total=113
    Returns (available, total) or (None, None) if unparseable.
    """
    # Format: "57 of 113 EA"  — two numbers joined by "of"
    match_of = re.search(r"(\d+)\s+of\s+(\d+)", alloc_text, re.IGNORECASE)
    if match_of:
        return int(match_of.group(1)), int(match_of.group(2))

    # Format: "57 / 113"
    match_slash = re.search(r"(\d+)\s*/\s*(\d+)", alloc_text)
    if match_slash:
        return int(match_slash.group(1)), int(match_slash.group(2))

    # Single number only — means 0 available (fully allocated), total = that number
    nums = re.findall(r"\d+", alloc_text)
    if len(nums) == 1:
        return 0, int(nums[0])

    if len(nums) >= 2:
        return int(nums[0]), int(nums[1])

    return None, None


def scrape_catalog_items(page, market, item_type, sap_code, timeout=15):
    """
    Scrape .catalauge-item-holder items from whichever frame has them.
    Returns a list of device dicts.
    """
    frame, _ = find_in_frames(page, ".catalauge-item-holder", timeout=timeout)
    if not frame:
        print(f"[{market}] ❌ No catalog items found for type={item_type}")
        return []

    items = frame.locator(".catalauge-item-holder").all()
    devices = []

    for item in items:
        try:
            name = item.locator(".cat-prd-dsc").inner_text().strip()
            sku = item.locator(".cat-prd-id").inner_text().strip()
            alloc_text = item.locator(".cat-prd-qty").inner_text().strip()

            available, total = parse_allocation(alloc_text)
            if available is None:
                print(f"[{market}] ⚠️ Could not parse: '{alloc_text}' — skipping")
                continue

            if available > 0:
                devices.append({
                    "Market": market,
                    "SAP Code": sap_code,
                    "SKU": sku,
                    "Name": name,
                    "Available": available,
                    "Total": total,
                    "Type": item_type,
                })
        except Exception as e:
            print(f"[{market}] ⚠️ Error parsing item: {e}")
            continue

    return devices


# ─────────────────────────────────────────
# 🔹 ALLOCATION
# ─────────────────────────────────────────
def scrape_allocation_page(page, row):
    market = row["Market"]
    username = row["Username"]
    password = row["Password"]
    sap_code = row.get("SAP Codes", "")

    if not login(page, username, password):
        return []
    # Wait for home page capacity element to appear before proceeding
    frame, loc = find_in_frames(page, "#credithold-tab-msg", timeout=20)
    if frame:
        print(f"[{market}] ✅ Home page loaded")
    else:
        print(f"[{market}] ⚠️ credithold-tab-msg not found — proceeding anyway")

    # ── 1. Click Catalog tab ──
    if not click_in_frames(page, "//a[@onclick='show_catalog_view()']"):
        print(f"[{market}] ❌ Catalog button not found")
        return []

    time.sleep(3)
    catalog_devices = scrape_catalog_items(page, market, "Catalog", sap_code)
    print(f"[{market}] ✅ Catalog items: {len(catalog_devices)}")

    # ── 2. Click CPO tab ──
    # Try common selectors for the CPO link
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

    # ── 3. Wait for CPO content to be DIFFERENT from catalog ──
    # Strategy: wait until a frame contains "CPO" or "Pre-Owned" text
    # AND catalog item holders are present (page has re-rendered)
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

    time.sleep(1)  # let DOM settle
    cpo_devices = scrape_catalog_items(page, market, "CPO", sap_code)
    print(f"[{market}] ✅ CPO items: {len(cpo_devices)}")

    return catalog_devices + cpo_devices


# ─────────────────────────────────────────
# 🔹 AMOUNTS
# ─────────────────────────────────────────
def scrape_amount_page(page, row):
    market = row["Market"]
    username = row["Username"]
    password = row["Password"]

    if not login(page, username, password):
        return {"Market": market, "Capacity": "LOGIN FAILED"}

    content = page.content().lower()
    if "runtime error" in content:
        return {"Market": market, "Capacity": "ERROR"}
    if "password expired" in content:
        return {"Market": market, "Capacity": "PASSWORD EXPIRED"}

    print(f"[{market}] URL: {page.url}")

    # The Ordering Capacity amount lives inside #credithold-tab-msg
    # which is inside the "Ordering Capacity" section of the start workarea frame.
    # Strategy: search all frames for the element #credithold-tab-msg and
    # extract the dollar figure from its text.  We avoid the credit card limits
    # table (which also contains $ values) by targeting this specific element.

    for attempt in range(20):
        for frame in page.frames:
            try:
                # Try the specific container first
                locator = frame.locator("#credithold-tab-msg")
                if locator.count() > 0:
                    text = locator.first.inner_text(timeout=2000)
                    print(f"[{market}] credithold-tab-msg text: {text[:300]}")

                    # The capacity figure appears after phrases like
                    match = re.search(
                        r"available ordering capacity.*?\bis\s*(-?\$?[\d,]+(?:\.\d+)?)",
                        text,
                        re.IGNORECASE,
                    )
                    if match:
                        raw = match.group(1).strip()
                        capacity = "(" + raw if raw.endswith(")") else raw
                        print(f"[{market}] ✅ Capacity found: {capacity}")
                        return {"Market": market, "Capacity": capacity}

                    # Fallback: first $ amount in this block
                    # (still safer than searching the whole page)
                    fallback = re.search(r"\$\s?[\d,]+(?:\.\d+)?\)?", text)
                    if fallback:
                        raw = fallback.group(0).strip()
                        capacity = "(" + raw if raw.endswith(")") else raw
                        print(f"[{market}] ✅ Capacity (fallback): {capacity}")
                        return {"Market": market, "Capacity": capacity}
                    number_only = re.search(r"(\d[\d,]*\.\d+)", text)

                    if number_only:
                        capacity = "$" + number_only.group(1)
                        print(f"[{market}] ✅ Capacity (number-only): {capacity}")
                        return {"Market": market, "Capacity": capacity}

            except Exception:
                pass

        print(f"[{market}] ⏳ Attempt {attempt+1} — capacity not parsed yet")
        time.sleep(1)

    print(f"[{market}] ❌ Ordering capacity not found")
    return {"Market": market, "Capacity": "NOT FOUND"}


# ─────────────────────────────────────────
# 🔹 RUNNERS
# ─────────────────────────────────────────
def run_allocation(selected_markets=None):
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
                devices = scrape_allocation_page(page, row)
                results_map[row["Market"]] = devices
                print(f"[{row['Market']}] Total devices with allocation: {len(devices)}")
            finally:
                page.close()

        browser.close()

    # Save Excel — now includes Type column
    output = BASE_DIR / f"data/allocation_{uuid.uuid4().hex}.xlsx"
    output.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.append(["SAP Code","Market", "SKU", "Name", "Available", "Total", "Type"])

    for market, devices in results_map.items():
        for d in devices:
            ws.append([d.get("SAP Code", ""), d["Market"], d["SKU"], d["Name"], d["Available"], d["Total"], d.get("Type", "")])


    wb.save(output)
    print(f"✅ Allocation saved: {output}")
    return str(output)


def run_amounts(selected_markets=None):
    df = pd.read_excel(LOGINS_FILE)
    df.columns = df.columns.str.strip()

    if selected_markets:
        df = df[df["Market"].str.lower().isin([m.lower() for m in selected_markets])]

    rows = [row for _, row in df.iterrows()]
    results = []

    with sync_playwright() as p:
        browser = new_browser(p)

        for row in rows:
            page = new_page(browser)
            try:
                result = scrape_amount_page(page, row)
                results.append(result)
                print(f"[{row['Market']}] Capacity: {result['Capacity']}")
            finally:
                page.close()

        browser.close()

    output = BASE_DIR / f"data/amounts_{uuid.uuid4().hex}.xlsx"
    output.parent.mkdir(parents=True, exist_ok=True)

    pd.DataFrame(results).to_excel(output, index=False)
    print(f"✅ Amounts saved: {output}")
    return str(output)