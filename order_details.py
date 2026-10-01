from __future__ import annotations

import re
import time
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright

from tmobile_tracking_automation.main import (
    INPUT_MARKET_COL,
    INPUT_ORDER_COL,
    OUTPUT_DIR,
    SCREENSHOT_DIR,
    log,
    load_credentials,
    load_orders,
    login,
    new_context,
    find_in_frames,
    open_order,
    screenshot_failure,
)

ORDER_DETAILS_OUTPUT_DIR = OUTPUT_DIR / "order_details"
ORDER_DETAILS_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# The order-detail page uses label/value table rows. We match the label text
# rather than relying on generated element IDs or table position.
def _normalize_text(value: str) -> str:
    return " ".join((value or "").replace("\xa0", " ").split()).strip()


def _get_labeled_value(page, label: str, timeout: int = 10) -> str:
    """Return the value from the table row whose td.identifier equals label."""
    target = _normalize_text(label).rstrip(":").lower()
    deadline = time.time() + timeout

    while time.time() < deadline:
        for frame in page.frames:
            try:
                labels = frame.locator("td.identifier")
                count = labels.count()
                for i in range(count):
                    label_el = labels.nth(i)
                    if not label_el.is_visible():
                        continue

                    actual = _normalize_text(label_el.inner_text()).rstrip(":").lower()
                    if actual != target:
                        continue

                    row = label_el.locator("xpath=ancestor::tr[1]")
                    value = row.locator("td.value").first
                    if value.count() and value.is_visible():
                        return _normalize_text(value.inner_text())
            except Exception:
                continue

        time.sleep(0.25)

    return ""


def _get_order_id_from_page(page, fallback: str, timeout: int = 10) -> str:
    """Read the displayed 'Order: <id>' value from the order-detail page."""
    pattern = re.compile(r"\bOrder:\s*([0-9A-Za-z_-]+)\b", re.I)
    deadline = time.time() + timeout

    while time.time() < deadline:
        for frame in page.frames:
            try:
                text = frame.locator("body").inner_text(timeout=1000)
                match = pattern.search(text)
                if match:
                    return match.group(1).strip()
            except Exception:
                continue
        time.sleep(0.25)

    return str(fallback).strip()


def extract_order_details(page, market: str, order_id: str) -> dict:
    """Extract only the fields required for the Order Details report."""
    result = {
        "market": market,
        "order_id": _get_order_id_from_page(page, order_id),
        "your_reference": _get_labeled_value(page, "Your Reference:"),
        "shipping_conditions": _get_labeled_value(page, "Shipping Conditions:"),
        "items": _get_labeled_value(page, "Items:"),
        "shipping": _get_labeled_value(page, "Shipping:"),
        "order_total": _get_labeled_value(page, "Order Total:"),
        "delivery_address": _get_labeled_value(page, "Delivery Address:"),
    }

    log(
        f"[{market}] {order_id}: "
        f"Reference='{result['your_reference']}', "
        f"Shipping Conditions='{result['shipping_conditions']}', "
        f"Items='{result['items']}', "
        f"Shipping='{result['shipping']}', "
        f"Order Total='{result['order_total']}', "
        f"Delivery Address='{result['delivery_address']}'"
    )
    return result

def clean_amount(value):
    if not value:
        return None

    value = str(value).replace("USD", "").replace(",", "").strip()

    try:
        return float(value)
    except ValueError:
        return None


def build_order_details_output(rows) -> pd.DataFrame:
    output_rows = []

    for r in rows:
        output_rows.append({
            "Market": r.get("market", ""),
            "Order ID": r.get("order_id", ""),
            "Your Reference": r.get("your_reference", ""),
            "Shipping Conditions": r.get("shipping_conditions", ""),
            "Items": clean_amount(r.get("items")),
            "Shipping": clean_amount(r.get("shipping")),
            "Order Total": clean_amount(r.get("order_total")),
            "Delivery Address": r.get("delivery_address", ""),
        })

    return pd.DataFrame(
        output_rows,
        columns=[
            "Market",
            "Order ID",
            "Your Reference",
            "Shipping Conditions",
            "Items",
            "Shipping",
            "Order Total",
            "Delivery Address",
        ],
    )



def run_order_details(input_file, credentials_file=None, selected_markets=None):
    """
    Run the separate T-Mobile Order Details automation.

    Input workbook must contain:
        Market | Order ID

    Output contains only:
        Market | Order ID | Your Reference | Items | Shipping | Order Total
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
    log(f"[ORDER DETAILS] Loaded {len(orders)} order(s) from {input_file.name}")

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
            # Reuse one login/session per market, just like the tracking automation.
            for market, market_orders in orders.groupby(INPUT_MARKET_COL, sort=False):
                key = market.lower()

                if key not in credentials.index:
                    for _, row in market_orders.iterrows():
                        results.append({
                            "market": market,
                            "order_id": str(row[INPUT_ORDER_COL]),
                            "your_reference": "",
                            "shipping_conditions": "",
                            "items": "",
                            "shipping": "",
                            "order_total": "",
                            "delivery_address": "",
                        })
                    log(f"[ORDER DETAILS] [{market}] Credentials not found", "error")
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
                        order_id = str(row[INPUT_ORDER_COL]).strip()

                        try:
                            log(f"[ORDER DETAILS] [{market}] Processing order {order_id}")

                            opened = open_order(page, order_id, market)
                            if not opened:
                                log(
                                    f"[ORDER DETAILS] [{market}] {order_id}: order not found",
                                    "warning",
                                )
                                results.append({
                                    "market": market,
                                    "order_id": order_id,
                                    "your_reference": "",
                            "shipping_conditions": "",
                                    "items": "",
                                    "shipping": "",
                                    "order_total": "",
                            "delivery_address": "",
                                })
                                continue

                            details = extract_order_details(page, market, order_id)
                            results.append(details)

                        except Exception as exc:
                            log(
                                f"[ORDER DETAILS] [{market}] {order_id}: "
                                f"{type(exc).__name__}: {exc}",
                                "error",
                            )
                            screenshot_failure(page, f"order_details_{market}", order_id)
                            results.append({
                                "market": market,
                                "order_id": order_id,
                                "your_reference": "",
                            "shipping_conditions": "",
                                "items": "",
                                "shipping": "",
                                "order_total": "",
                            "delivery_address": "",
                            })

                except Exception as exc:
                    log(
                        f"[ORDER DETAILS] [{market}] Market session failed: {exc}",
                        "error",
                    )
                    for _, row in market_orders.iterrows():
                        results.append({
                            "market": market,
                            "order_id": str(row[INPUT_ORDER_COL]),
                            "your_reference": "",
                            "shipping_conditions": "",
                            "items": "",
                            "shipping": "",
                            "order_total": "",
                            "delivery_address": "",
                        })
                finally:
                    context.close()

        finally:
            browser.close()

    output_df = build_order_details_output(results)
    output_file = ORDER_DETAILS_OUTPUT_DIR / (
        f"order_details_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
    )
    output_df.to_excel(output_file, index=False)

    log(f"[ORDER DETAILS] Finished. Output: {output_file}")
    return str(output_file)


if __name__ == "__main__":
    print("Use run_order_details() from app.py.")

