import streamlit as st
from playwright_service import run_allocation, run_amounts
from fetch_price import run_prices
from export_imei_reports_v2 import run_imei_export
import pandas as pd
from pathlib import Path
from datetime import date, timedelta

BASE_DIR = Path(__file__).resolve().parent
LOGINS_FILE = BASE_DIR / "marketlogins.xlsx"

DIRECT  = ["Houston", "Tulsa", "RGV", "SanAntonio", "Waco", "Corpus", "Tampa", "Dallas", "Arizona"]
OLD_SUB = ["Houston", "SanAntonio", "Atlanta", "Arizona", "CALIFORNIA", "SanFrancisco"]
NEW_SUB = ["CT", "CO", "IN", "MA", "NY", "VA", "MD/DC"]

df = pd.read_excel(LOGINS_FILE)
ALL_MARKETS = df["Market"].dropna().unique().tolist()

st.title("📊 Purely Wireless Automation")

st.markdown("#### Quick Select")
qcol1, qcol2, qcol3, qcol4 = st.columns(4)

if qcol1.button("🏢 Direct", use_container_width=True):
    st.session_state["market_multiselect"] = [m for m in DIRECT if m in ALL_MARKETS]

if qcol2.button("📋 Old Sub Dealers", use_container_width=True):
    st.session_state["market_multiselect"] = [m for m in OLD_SUB if m in ALL_MARKETS]

if qcol3.button("🆕 New Sub Dealers", use_container_width=True):
    st.session_state["market_multiselect"] = [m for m in NEW_SUB if m in ALL_MARKETS]

if qcol4.button("✅ All Markets", use_container_width=True):
    st.session_state["market_multiselect"] = ALL_MARKETS

selected = st.multiselect(
    "Selected Markets",
    options=ALL_MARKETS,
    key="market_multiselect",
)

st.markdown("---")

col1, col2, col3, col4 = st.columns(4)

if col1.button("📦 Get Allocation", use_container_width=True):
    if not selected:
        st.warning("Select at least one market.")
    else:
        with st.spinner(f"Running allocation for {len(selected)} markets..."):
            file = run_allocation(selected)
        with open(file, "rb") as f:
            st.download_button("⬇️ Download Allocation", f, file_name="allocation.xlsx")

if col2.button("💰 Get Amounts", use_container_width=True):
    if not selected:
        st.warning("Select at least one market.")
    else:
        with st.spinner(f"Running amounts for {len(selected)} markets..."):
            file = run_amounts(selected)
        with open(file, "rb") as f:
            st.download_button("⬇️ Download Amounts", f, file_name="amounts.xlsx")

if col3.button("🏷️ Get Prices", use_container_width=True):
    if not selected:
        st.warning("Select at least one market.")
    else:
        with st.spinner(f"Running prices for {len(selected)} markets..."):
            file = run_prices(selected)
        with open(file, "rb") as f:
            st.download_button("⬇️ Download Prices", f, file_name="prices.xlsx")

with col4:
    imei_clicked = st.button("📱 Export IMEIs", use_container_width=True)

st.markdown("#### IMEI Export Date Range")
dcol1, dcol2 = st.columns(2)
imei_start = dcol1.date_input("From", value=date.today() - timedelta(days=6))
imei_end = dcol2.date_input("To", value=date.today())

if imei_clicked:
    if not selected:
        st.warning("Select at least one market.")
    elif imei_end < imei_start:
        st.error("'To' date is before 'From' date.")
    else:
        start_str = imei_start.strftime("%m/%d/%Y")
        end_str = imei_end.strftime("%m/%d/%Y")

        with st.spinner(f"Running IMEI export for {len(selected)} markets, {imei_start} to {imei_end}..."):
            outcome = run_imei_export(start_str, end_str, selected, interactive=False)

        merged_path = outcome["path"]
        results = outcome["results"]

        # Surface anything that failed instead of hiding it behind a
        # download button that implies everything worked.
        failures = []
        for r in results:
            if r["login_failed"]:
                failures.append(f"{r['market']}: login failed")
            for d, reason in r.get("failed", {}).items():
                failures.append(f"{r['market']} {d}: {reason}")

        if failures:
            with st.expander(f"⚠️ {len(failures)} issue(s) during this run", expanded=True):
                for f in failures:
                    st.warning(f)

        if merged_path:
            with open(merged_path, "rb") as f:
                st.download_button(
                    "⬇️ Download IMEI Export",
                    f,
                    file_name=Path(merged_path).name,
                )
        else:
            st.error("No files were downloaded — nothing to export. Check the issues above.")