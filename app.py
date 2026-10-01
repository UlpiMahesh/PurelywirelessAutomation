import streamlit as st
from pathlib import Path
from datetime import date, timedelta
import tempfile
import pandas as pd

from playwright_service import run_allocation, run_amounts
from fetch_price import run_prices
from export_imei_reports_v2 import run_imei_export
from tmobile_tracking_automation.main import run_tracking
from order_details import run_order_details

st.set_page_config(page_title="Purely Wireless Automation", page_icon="📡", layout="wide", initial_sidebar_state="collapsed")

BASE_DIR = Path(__file__).resolve().parent
LOGINS_FILE = BASE_DIR / "marketlogins.xlsx"
DIRECT = ["Houston", "Tulsa", "RGV", "SanAntonio", "Waco", "Corpus", "Tampa", "Dallas", "Arizona"]
OLD_SUB = ["Houston", "SanAntonio", "Atlanta", "Arizona", "CALIFORNIA", "SanFrancisco"]
NEW_SUB = ["CT", "CO", "IN", "MA", "NY", "VA", "MD/DC"]


@st.cache_data
def load_markets(path):
    df = pd.read_excel(path)
    return df["Market"].dropna().astype(str).unique().tolist()


ALL_MARKETS = load_markets(LOGINS_FILE)
GROUPS = {"Direct": DIRECT, "Old sub dealers": OLD_SUB, "New sub dealers": NEW_SUB, "All markets": ALL_MARKETS}


def set_markets(names):
    st.session_state["market_multiselect"] = [m for m in names if m in ALL_MARKETS]


def save_uploaded_file(uploaded_file, prefix):
    temp_dir = Path(tempfile.gettempdir()) / "purely_wireless_automation"
    temp_dir.mkdir(parents=True, exist_ok=True)
    output_path = temp_dir / f"{prefix}_{uploaded_file.name}"
    output_path.write_bytes(uploaded_file.getvalue())
    return output_path


def download_file(path, label, key, filename=None):
    if not path:
        return
    path = Path(path)
    if not path.exists():
        st.error(f"Output file was not found: {path}")
        return
    with open(path, "rb") as f:
        st.download_button(label, f, file_name=filename or path.name, width="stretch", key=key, icon=":material/download:")


def card_head(icon, title, desc, tag=""):
    tag_html = f'<span class="tag">{tag}</span>' if tag else ""
    return (f'<div class="ch"><span class="ms tile">{icon}</span>{tag_html}</div>'
            f'<div class="ct">{title}</div><div class="cd">{desc}</div>')


def section(title, note=""):
    st.markdown(f'<div class="sec"><h2>{title}</h2><span>{note}</span></div>', unsafe_allow_html=True)


st.markdown("<style>@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@500;600;700&family=Geist:wght@400;500;600&family=JetBrains+Mono:wght@400;500&family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@24,400,0,0&display=swap');</style>", unsafe_allow_html=True)
st.markdown("""
<style>
:root{--bg:#f4f5fb;--high:#e9ecf6;--lowest:#f1f3fa;--tx:#171a2e;--tx2:#454b66;--mut:#7b819d;--line:rgba(70,80,130,.14);--p:#4f56e8;--s:#0c9a8a;--t:#c98a1b;--i:#7c4dea}
.stApp{color-scheme:light;background:var(--bg);color:var(--tx);font-family:'Geist',system-ui,sans-serif}
.stApp:before{content:"";position:fixed;inset:0;pointer-events:none;background:radial-gradient(520px 380px at 25% -5%,rgba(79,86,232,.10),transparent 70%),radial-gradient(460px 380px at 100% 40%,rgba(12,154,138,.08),transparent 70%)}
[data-testid="stMarkdownContainer"] p{color:var(--tx)}
[data-testid="stHeader"],[data-testid="stToolbar"],[data-testid="stDecoration"],#MainMenu,footer{display:none!important}
.block-container{max-width:1400px;padding:1.1rem 2rem 3rem!important}
[data-testid="stVerticalBlock"]{gap:1rem}
h1,h2,h3,.ct,.brand b,.big{font-family:'Plus Jakarta Sans',sans-serif}
.ms{font-family:'Material Symbols Rounded';font-weight:400;font-style:normal;line-height:1;letter-spacing:normal;text-transform:none;display:inline-block;white-space:nowrap;-webkit-font-smoothing:antialiased}
.st-key-topbar{background:rgba(255,255,255,.88);box-shadow:0 6px 24px -16px rgba(40,50,110,.3);backdrop-filter:blur(16px);border:1px solid var(--line);border-radius:20px;padding:10px 20px}
.brand{display:flex;align-items:center;gap:12px}.mark{width:36px;height:36px;border-radius:11px;background:linear-gradient(135deg,var(--s),var(--i));display:grid;place-items:center;color:#fff;font:700 14px 'Plus Jakarta Sans'}
.brand b{font-size:17px;font-weight:600;letter-spacing:-.01em}
.state{display:flex;justify-content:flex-end}.state span{display:inline-flex;align-items:center;gap:8px;font:500 12px 'JetBrains Mono',monospace;color:var(--s);background:rgba(12,154,138,.12);border-radius:999px;padding:6px 14px}
.state i{width:7px;height:7px;border-radius:50%;background:var(--s);animation:pl 2s infinite}@keyframes pl{50%{opacity:.3}}
.st-key-topbar [data-testid="stRadio"]{display:flex;justify-content:center}
[data-testid="stRadio"] [role="radiogroup"]{background:var(--lowest);border-radius:999px;padding:4px;gap:2px;flex-wrap:nowrap}
[data-testid="stRadio"] label{border-radius:999px;padding:6px 18px;margin:0;cursor:pointer}
[data-testid="stRadio"] label>div:first-child{display:none}
[data-testid="stRadio"] label p{color:var(--tx2);font-size:14px}
[data-testid="stRadio"] label:has(input:checked){background:var(--high);box-shadow:inset 0 1px 1px rgba(255,255,255,.08)}
[data-testid="stRadio"] label:has(input:checked) p{color:var(--tx);font-weight:500}
.hero{display:flex;align-items:flex-end;justify-content:space-between;gap:16px;flex-wrap:wrap;padding:14px 0 2px}
.hero h1{font-size:34px;font-weight:600;letter-spacing:-.025em;margin:0;padding:0;line-height:1.15}.hero p{color:var(--tx2);margin:6px 0 0;font-size:15px}
.stats{display:flex;gap:10px}.stat{background:#fff;border:1px solid var(--line);border-radius:999px;padding:8px 18px}
.stat small{display:block;font:500 11px 'JetBrains Mono',monospace;color:var(--mut)}.stat b{font:600 18px 'Plus Jakarta Sans';line-height:1.2}
.sec{display:flex;align-items:baseline;gap:12px;margin-top:10px}.sec h2{font-size:18px;font-weight:600;margin:0;padding:0}.sec span{color:var(--mut);font-size:13px}
[class*="st-key-card"]{--a:var(--p);position:relative;background:linear-gradient(160deg,color-mix(in srgb,var(--a) 9%,transparent),transparent 50%),#fff;box-shadow:0 8px 26px -16px rgba(40,50,110,.3);border:1px solid var(--line);border-radius:24px;padding:22px;transition:transform .25s,border-color .25s}
[class*="st-key-card_"]:hover{transform:translateY(-3px);border-color:color-mix(in srgb,var(--a) 45%,transparent)}
.st-key-card_alloc{--a:var(--p)}.st-key-card_amt{--a:var(--s)}.st-key-card_price{--a:var(--t)}.st-key-card_imei{--a:var(--i)}.st-key-card_track{--a:var(--s)}.st-key-card_details{--a:var(--p)}
.ch{display:flex;justify-content:space-between;align-items:center;margin-bottom:14px}
.tile{width:42px;height:42px;border-radius:14px;background:color-mix(in srgb,var(--a) 16%,var(--high));color:var(--a);display:grid;place-items:center;font-size:22px}
.tag{font:500 11px 'JetBrains Mono',monospace;color:var(--tx2);background:var(--high);border-radius:999px;padding:3px 10px}
.ct{font-size:18px;font-weight:600;letter-spacing:-.01em}.cd{color:var(--tx2);font-size:13px;margin:4px 0 6px;line-height:1.5}
[data-testid="stBaseButton-secondary"],button[kind="secondary"],[data-testid="stDownloadButton"] button{background:var(--high);color:var(--tx2);border:1px solid transparent;border-radius:999px;min-height:42px;font-weight:500;transition:all .2s}
[data-testid="stBaseButton-secondary"]:hover,button[kind="secondary"]:hover,[data-testid="stDownloadButton"] button:hover{border-color:var(--p);color:var(--tx);background:var(--high)}
[data-testid="stBaseButton-primary"],button[kind="primary"]{background:var(--a,var(--p));color:#fff;border:0;border-radius:999px;min-height:44px;font-weight:600;box-shadow:0 8px 22px -10px var(--a,var(--p))}
[data-testid="stBaseButton-primary"]:hover,button[kind="primary"]:hover{filter:brightness(1.1);color:#fff;border:0}
[data-testid="stBaseButton-primary"] p,button[kind="primary"] p{color:#fff}
[data-baseweb="select"]>div,[data-baseweb="input"]>div,[data-testid="stDateInput"] [data-baseweb="input"]{background:var(--lowest)!important;border:1px solid var(--line)!important;border-radius:14px!important;color:var(--tx)!important}
[data-baseweb="tag"]{background:var(--high)!important;color:var(--tx)!important;border-radius:999px!important}
[data-testid="stWidgetLabel"] p,label p{color:var(--mut)!important;font-size:12px}
[data-testid="stFileUploaderDropzone"]{background:var(--lowest);border:1.5px dashed rgba(12,154,138,.45);border-radius:18px}
[data-testid="stFileUploaderDropzone"] button{border-radius:999px}
[data-testid="stAlert"]{border-radius:14px;border:1px solid var(--line)}
@media (max-width:900px){.block-container{padding:1rem!important}.hero h1{font-size:26px}}
</style>
""", unsafe_allow_html=True)


with st.container(key="topbar"):
    c_brand, c_nav, c_state = st.columns([2.2, 3, 2.2], vertical_alignment="center")
    c_brand.markdown('<div class="brand"><div class="mark">PW</div><b>Purely Wireless</b></div>', unsafe_allow_html=True)
    with c_nav:
        page = st.radio("Navigation", ["Dashboard", "Reports", "Market Details"], horizontal=True, label_visibility="collapsed", key="navigation")
    c_state.markdown('<div class="state"><span><i></i>Ready</span></div>', unsafe_allow_html=True)


def dashboard():
    selected_now = st.session_state.get("market_multiselect", [])
    st.markdown(
        f'<div class="hero"><div><h1>Purely Wireless Automation</h1><p>Market automations and order details processing.</p></div>'
        f'<div class="stats"><div class="stat"><small>Markets</small><b>{len(ALL_MARKETS)}</b></div>'
        f'<div class="stat"><small>Selected</small><b>{len(selected_now)}</b></div></div></div>', unsafe_allow_html=True)

    section("Market selection", "Pick a group or choose markets one by one.")
    with st.container(key="card_scope"):
        cols = st.columns(5)
        for col, (label, names) in zip(cols[:4], GROUPS.items()):
            n = len([m for m in names if m in ALL_MARKETS])
            col.button(f"{label} ({n})", key=f"grp_{label}", on_click=set_markets, args=(names,), width="stretch")
        cols[4].button("Clear", key="grp_clear", on_click=set_markets, args=([],), width="stretch", icon=":material/close:")
        selected = st.multiselect("Selected markets", ALL_MARKETS, key="market_multiselect", placeholder="Search or select markets...")

    section("Market automations", "Runs against T-Mobile with the saved market logins.")
    c1, c2, c3, c4 = st.columns(4)

    with c1, st.container(key="card_alloc"):
        st.markdown(card_head("layers", "Get Allocation", "Fetch allocation data from the portal."), unsafe_allow_html=True)
        if st.button("Run allocation", key="allocation", type="primary", width="stretch", icon=":material/play_arrow:"):
            if not selected: st.warning("Select at least one market.")
            else:
                try:
                    with st.spinner(f"Running allocation for {len(selected)} market(s)..."): file = run_allocation(selected)
                    if file: st.success("Allocation completed."); download_file(file, "Download allocation", "allocation_download", "allocation.xlsx")
                except Exception as exc: st.error(f"Allocation failed: {exc}")

    with c2, st.container(key="card_amt"):
        st.markdown(card_head("account_balance_wallet", "Get Amounts", "Fetch amount information from the portal."), unsafe_allow_html=True)
        if st.button("Run amounts", key="amounts", type="primary", width="stretch", icon=":material/play_arrow:"):
            if not selected: st.warning("Select at least one market.")
            else:
                try:
                    with st.spinner(f"Running amounts for {len(selected)} market(s)..."): file = run_amounts(selected)
                    if file: st.success("Amounts completed."); download_file(file, "Download amounts", "amounts_download", "amounts.xlsx")
                except Exception as exc: st.error(f"Amounts failed: {exc}")

    with c3, st.container(key="card_price"):
        st.markdown(card_head("sell", "Get Prices", "Fetch device pricing from the portal."), unsafe_allow_html=True)
        if st.button("Run prices", key="prices", type="primary", width="stretch", icon=":material/play_arrow:"):
            if not selected: st.warning("Select at least one market.")
            else:
                try:
                    with st.spinner(f"Running prices for {len(selected)} market(s)..."): file = run_prices(selected)
                    if file: st.success("Prices completed."); download_file(file, "Download prices", "prices_download", "prices.xlsx")
                except Exception as exc: st.error(f"Prices failed: {exc}")

    with c4, st.container(key="card_imei"):
        st.markdown(card_head("sim_card_download", "Export IMEIs", "Export IMEIs for a date range."), unsafe_allow_html=True)
        d1, d2 = st.columns(2)
        start = d1.date_input("From", date.today() - timedelta(days=6), key="imei_from")
        end = d2.date_input("To", date.today(), key="imei_to")
        if st.button("Run IMEI export", key="imei", type="primary", width="stretch", icon=":material/play_arrow:"):
            if not selected: st.warning("Select at least one market.")
            elif end < start: st.error("'To' date is before 'From' date.")
            else:
                try:
                    with st.spinner(f"Running IMEI export for {len(selected)} market(s)..."):
                        outcome = run_imei_export(start.strftime("%m/%d/%Y"), end.strftime("%m/%d/%Y"), selected, interactive=False)
                    failures = []
                    for result in outcome["results"]:
                        if result["login_failed"]: failures.append(f"{result['market']}: login failed")
                        for device, reason in result.get("failed", {}).items(): failures.append(f"{result['market']} {device}: {reason}")
                    if failures:
                        with st.expander(f"{len(failures)} issue(s)", expanded=True):
                            for failure in failures: st.warning(failure)
                    if outcome["path"]:
                        st.success("IMEI export completed."); download_file(outcome["path"], "Download IMEI export", "imei_download")
                    else: st.error("No files were downloaded, nothing to export.")
                except Exception as exc: st.error(f"IMEI export failed: {exc}")

    section("Order details processing", "Upload an order Excel file. Markets are read from its Market column.")
    t_col, d_col = st.columns(2)

    with t_col, st.container(key="card_track"):
        st.markdown(card_head("local_shipping", "Fetch Tracking", "Get tracking information for each order.", ".xlsx / .xls"), unsafe_allow_html=True)
        tracking_file = st.file_uploader("Order Excel file", type=["xlsx", "xls"], key="tracking_upload")
        if st.button("Start tracking", key="start_tracking", type="primary", width="stretch", icon=":material/play_arrow:"):
            if not tracking_file: st.warning("Upload an order Excel file first.")
            else:
                try:
                    input_path = save_uploaded_file(tracking_file, "tracking")
                    with st.spinner("Fetching tracking IDs..."):
                        output_file = run_tracking(input_file=input_path, credentials_file=LOGINS_FILE, selected_markets=None)
                    if output_file:
                        st.success("Tracking report completed."); download_file(output_file, "Download tracking report", "tracking_download")
                    else: st.warning("No tracking report was returned.")
                except Exception as exc: st.error(f"Tracking failed: {exc}")

    with d_col, st.container(key="card_details"):
        st.markdown(card_head("receipt_long", "Fetch Order Details", "Get detailed information for each order.", ".xlsx / .xls"), unsafe_allow_html=True)
        details_file = st.file_uploader("Order Excel file", type=["xlsx", "xls"], key="details_upload")
        if st.button("Start extraction", key="start_details", type="primary", width="stretch", icon=":material/play_arrow:"):
            if not details_file: st.warning("Upload an order Excel file first.")
            else:
                try:
                    input_path = save_uploaded_file(details_file, "order_details")
                    with st.spinner("Fetching order details..."):
                        output_file = run_order_details(input_file=input_path, credentials_file=LOGINS_FILE, selected_markets=None)
                    if output_file:
                        st.success("Order details report completed."); download_file(output_file, "Download order details", "details_download")
                    else: st.warning("No order-details report was returned.")
                except Exception as exc: st.error(f"Order details failed: {exc}")


def placeholder(title, text):
    st.markdown(f'<div class="hero"><div><h1>{title}</h1><p>{text}</p></div></div>', unsafe_allow_html=True)
    with st.container(key="card_empty"):
        st.markdown(card_head("construction", "Coming later", "Nothing has been added here yet."), unsafe_allow_html=True)


if page == "Dashboard":
    dashboard()
elif page == "Reports":
    placeholder("Reports", "Reports will live here.")
else:
    placeholder("Market Details", "Market information and configuration will live here.")