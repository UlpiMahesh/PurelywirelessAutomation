from pathlib import Path
import re
import sys
import pandas as pd
import numpy as np
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
INPUT = ROOT / "input"
OUTPUT = ROOT / "output"


def clean_id(series):
    """Clean Excel IDs such as 12345, 12345.0, or ' 12345 '."""
    x = series.astype("string").str.strip()
    x = x.str.replace(r"\.0$", "", regex=True)
    return x.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "<NA>": pd.NA})


def clean_text(series):
    return series.astype("string").str.strip()


def find_week_sheet(path):
    """Return the latest sheet named WEEK OF DDMMYY."""
    candidates = []
    for sheet in pd.ExcelFile(path).sheet_names:
        match = re.match(r"^WEEK OF\s+(\d{6})$", str(sheet).strip(), re.I)
        if match:
            try:
                dt = pd.to_datetime(match.group(1), format="%d%m%y")
                candidates.append((dt, sheet))
            except Exception:
                pass

    if not candidates:
        raise ValueError(
            f"{path.name}: No valid 'WEEK OF DDMMYY' sheet was found."
        )

    return max(candidates, key=lambda x: x[0])[1]


def require_columns(df, required, description):
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(
            f"{description} is missing required column(s): {missing}"
        )



def load_master():
    path = INPUT / "master.xlsx"
    df = pd.read_excel(path, sheet_name="sub dealer cust codes")

    sap_col = next(
        (c for c in df.columns if str(c).strip().lower() == "sapcode"),
        None,
    )
    if sap_col is None:
        raise ValueError(
            "master.xlsx: 'sub dealer cust codes' has no SAPCODE column."
        )

    required = [sap_col, "Cust/Store", "Market", "Sub-Dealer Corporation"]
    require_columns(df, required, "master.xlsx / sub dealer cust codes")

    df = df.rename(
        columns={
            sap_col: "SAPCODE",
            "Cust/Store": "CustCode",
            "Sub-Dealer Corporation": "MasterDealer",
        }
    )

    df["SAPCODE"] = clean_id(df["SAPCODE"])
    df["CustCode"] = clean_id(df["CustCode"])
    df["Market"] = clean_text(df["Market"])
    df["MasterDealer"] = clean_text(df["MasterDealer"])

    # DealerGroup is optional in the master. If it exists, use it.
    # If it does not exist, create a stable group from SAPCODE + dealer name.
    if "DealerGroup" in df.columns:
        df["DealerGroup"] = clean_text(df["DealerGroup"])
        missing_group = df["DealerGroup"].isna() | df["DealerGroup"].eq("")
        df.loc[missing_group, "DealerGroup"] = (
            df.loc[missing_group, "SAPCODE"].fillna("").astype("string")
            + "|"
            + df.loc[missing_group, "MasterDealer"]
            .fillna("")
            .astype("string")
            .str.upper()
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )
    else:
        dealer_key = (
            df["MasterDealer"]
            .fillna("")
            .astype("string")
            .str.upper()
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )
        df["DealerGroup"] = (
            df["SAPCODE"].fillna("").astype("string") + "|" + dealer_key
        )

    # One SAPCODE must represent one market.
    sap = (
        df[["SAPCODE", "Market"]]
        .dropna(subset=["SAPCODE", "Market"])
        .drop_duplicates()
    )
    sap["MarketKey"] = (
        sap["Market"].astype("string")
        .str.casefold()
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    market_conflicts = sap.groupby("SAPCODE")["MarketKey"].nunique()
    bad = market_conflicts[market_conflicts > 1]
    if len(bad):
        raise ValueError(
            f"master.xlsx maps one SAPCODE to multiple markets: {bad.index.tolist()}"
        )
    sap = sap.drop_duplicates("SAPCODE")[["SAPCODE", "Market"]]

    # A CustCode must belong to one SAPCODE. Duplicate rows are allowed
    # because the same account can appear under several dealer-name aliases.
    # If a CustCode has multiple DealerGroups, resolve it using the unique
    # group that is shared by multiple CustCodes under the same SAPCODE.
    # This handles master rows such as:
    #   4040827 -> Sky Wireless_Oakland + PM Wireless
    #   4040829 -> Alpha Tel LLC       + PM Wireless
    #   4040830 -> Philip Zhang...      + PM Wireless
    # where PM Wireless is the common dealer group.
    cust_rows = (
        df[["CustCode", "SAPCODE", "MasterDealer", "DealerGroup"]]
        .dropna(subset=["CustCode"])
        .drop_duplicates()
    )

    cust_sap_conflicts = cust_rows.groupby("CustCode")["SAPCODE"].nunique()
    bad = cust_sap_conflicts[cust_sap_conflicts > 1]
    if len(bad):
        raise ValueError(
            "master.xlsx maps CustCode to multiple SAPCODEs: "
            f"{bad.index.tolist()[:20]}"
        )

    group_counts = (
        cust_rows[["SAPCODE", "CustCode", "DealerGroup"]]
        .drop_duplicates()
        .groupby(["SAPCODE", "DealerGroup"])["CustCode"]
        .nunique()
    )

    resolved_rows = []
    unresolved = []

    for (cust_code, sapcode), g in cust_rows.groupby(["CustCode", "SAPCODE"], dropna=False):
        groups = [gval for gval in g["DealerGroup"].dropna().unique()]
        if len(groups) <= 1:
            chosen = groups[0] if groups else pd.NA
        else:
            candidates = [
                grp for grp in groups
                if group_counts.get((sapcode, grp), 0) > 1
            ]
            if len(candidates) == 1:
                chosen = candidates[0]
            else:
                unresolved.append((cust_code, sapcode, groups))
                continue

        row = g.iloc[0].copy()
        row["DealerGroup"] = chosen
        # Prefer the DealerGroup name for display rather than an alias row.
        if pd.notna(chosen) and "|" in str(chosen):
            display = str(chosen).split("|", 1)[1]
            row["MasterDealer"] = display
        resolved_rows.append(row)

    if unresolved:
        raise ValueError(
            "Could not safely resolve duplicate DealerGroups in master. "
            f"Examples: {unresolved[:10]}"
        )

    cust = pd.DataFrame(resolved_rows)
    cust = cust.drop_duplicates("CustCode")

    return sap.reset_index(drop=True), cust.reset_index(drop=True)



def load_normal_wishlist(sap, cust):
    path = INPUT / "normal_wishlist.xlsx"
    sheets = pd.ExcelFile(path).sheet_names
    sheet = next(
        (s for s in sheets if "wishListRequests" in str(s)),
        sheets[-1],
    )

    df = pd.read_excel(path, sheet_name=sheet)
    df = df.rename(
        columns={
            "cust ": "CustCode",
            "cust": "CustCode",
            "QuantityRequested": "WishedQty",
            "ProductDescription": "DeviceName",
        }
    )

    require_columns(
        df,
        ["CustCode", "SKU", "WishedQty"],
        f"normal_wishlist.xlsx / {sheet}",
    )

    if "DeviceName" not in df.columns:
        df["DeviceName"] = pd.NA

    df["CustCode"] = clean_id(df["CustCode"])
    df["SKU"] = clean_id(df["SKU"])
    df["WishedQty"] = pd.to_numeric(df["WishedQty"], errors="coerce")
    df = df.dropna(subset=["CustCode", "SKU", "WishedQty"]).copy()
    df = df.drop(columns=["Market"], errors="ignore")

    df = df.merge(
        cust[["CustCode", "SAPCODE", "DealerGroup", "MasterDealer"]],
        on="CustCode",
        how="left",
        validate="many_to_one",
    )

    missing = df["SAPCODE"].isna()
    if missing.any():
        examples = df.loc[missing, "CustCode"].dropna().astype(str).unique()[:10]
        raise ValueError(
            "Normal wishlist contains CustCode(s) not found in master: "
            f"{examples.tolist()}"
        )

    df = df.merge(sap, on="SAPCODE", how="left", validate="many_to_one")
    df["Dealer"] = clean_text(df["MasterDealer"])
    df["Type"] = "NORMAL SUB-DEALER"

    # Wished quantity is now aggregated at DealerGroup + SKU.
    # Multiple CustCodes belonging to the same dealer are one dealer.
    return (
        df.groupby(
            ["Type", "SAPCODE", "Market", "DealerGroup", "Dealer", "SKU"],
            dropna=False,
        )
        .agg(
            WishedQty=("WishedQty", "sum"),
            DeviceName=("DeviceName", "first"),
        )
        .reset_index()
    )



def load_normal_orders(sap, cust):
    path = INPUT / "normal_orders.xlsx"
    sheet = find_week_sheet(path)
    df = pd.read_excel(path, sheet_name=sheet)

    df = df.rename(
        columns={
            "Custcode": "CustCode",
            "CustCode": "CustCode",
            "SKU's": "SKU",
            "QTY": "OrderedQty",
        }
    )

    require_columns(
        df,
        ["CustCode", "SKU", "OrderedQty"],
        f"normal_orders.xlsx / {sheet}",
    )

    df["CustCode"] = clean_id(df["CustCode"])
    df["SKU"] = clean_id(df["SKU"])
    df["OrderedQty"] = pd.to_numeric(df["OrderedQty"], errors="coerce")
    df = df.dropna(subset=["CustCode", "SKU", "OrderedQty"]).copy()

    df = df.merge(
        cust[["CustCode", "SAPCODE", "DealerGroup"]],
        on="CustCode",
        how="left",
        validate="many_to_one",
    )

    missing = df["SAPCODE"].isna()
    if missing.any():
        examples = df.loc[missing, "CustCode"].dropna().astype(str).unique()[:10]
        raise ValueError(
            "Normal orders contains CustCode(s) not found in master: "
            f"{examples.tolist()}"
        )

    df["Type"] = "NORMAL SUB-DEALER"

    # CRITICAL: orders from all CustCodes in the same DealerGroup are combined.
    orders = (
        df.groupby(
            ["Type", "SAPCODE", "DealerGroup", "SKU"],
            dropna=False,
        )
        .agg(OrderedQty=("OrderedQty", "sum"))
        .reset_index()
    )

    return orders, sheet



def load_bot_wishlist(sap, cust):
    path = INPUT / "bot_wishlist.xlsx"
    df = pd.read_excel(path, sheet_name="By_Dealer")

    df = df.rename(
        columns={
            "SAPCode": "SAPCODE",
            "ProductName": "DeviceName",
            "DealerName": "Dealer",
            "WishListQty": "WishedQty",
            "cust": "CustCode",
            "Cust": "CustCode",
        }
    )

    require_columns(
        df,
        ["SAPCODE", "SKU", "CustCode", "WishedQty", "DeviceName", "Dealer"],
        "bot_wishlist.xlsx / By_Dealer",
    )

    df["SAPCODE"] = clean_id(df["SAPCODE"])
    df["CustCode"] = clean_id(df["CustCode"])
    df["SKU"] = clean_id(df["SKU"])
    df["WishedQty"] = pd.to_numeric(df["WishedQty"], errors="coerce")
    df = df.dropna(subset=["SAPCODE", "CustCode", "SKU", "WishedQty"]).copy()

    # Resolve BOT CustCode to the same DealerGroup used by normal data.
    df = df.merge(
        cust[["CustCode", "SAPCODE", "DealerGroup", "MasterDealer"]],
        on="CustCode",
        how="left",
        validate="many_to_one",
        suffixes=("", "_MASTER"),
    )

    missing = df["DealerGroup"].isna()
    if missing.any():
        examples = df.loc[missing, "CustCode"].dropna().astype(str).unique()[:10]
        raise ValueError(
            "BOT wishlist contains CustCode(s) not found in master: "
            f"{examples.tolist()}"
        )

    # Validate the SAPCODE in BOT wishlist agrees with the master CustCode mapping.
    mismatch = df["SAPCODE"].ne(df["SAPCODE_MASTER"])
    if mismatch.any():
        examples = (
            df.loc[mismatch, ["CustCode", "SAPCODE", "SAPCODE_MASTER"]]
            .head(10)
            .to_dict("records")
        )
        raise ValueError(
            "BOT wishlist has SAPCODE/CustCode combinations that disagree with "
            f"master: {examples}"
        )

    df = df.drop(columns=["SAPCODE_MASTER"])

    df = df.merge(sap, on="SAPCODE", how="left", validate="many_to_one")
    missing = df["Market"].isna()
    if missing.any():
        examples = df.loc[missing, "SAPCODE"].dropna().astype(str).unique()[:10]
        raise ValueError(
            "BOT wishlist contains SAPCODE(s) not found in master: "
            f"{examples.tolist()}"
        )

    # Use master dealer name for consistent grouping/display.
    df["Dealer"] = clean_text(df["MasterDealer"])
    df["Type"] = "BOT SUB-DEALER"

    return (
        df.groupby(
            ["Type", "SAPCODE", "Market", "DealerGroup", "Dealer", "SKU"],
            dropna=False,
        )
        .agg(
            WishedQty=("WishedQty", "sum"),
            DeviceName=("DeviceName", "first"),
        )
        .reset_index()
    )



def load_bot_orders(sap, cust):
    path = INPUT / "bot_orders.xlsx"
    df = pd.read_excel(path, sheet_name="ORDERTEMP")

    df = df.rename(
        columns={
            "Account": "SAPCODE",
            "Ship_To_ID": "CustCode",
            "Quantity": "OrderedQty",
        }
    )

    require_columns(
        df,
        ["SAPCODE", "CustCode", "SKU", "OrderedQty"],
        "bot_orders.xlsx / ORDERTEMP",
    )

    df["SAPCODE"] = clean_id(df["SAPCODE"])
    df["CustCode"] = clean_id(df["CustCode"])
    df["SKU"] = clean_id(df["SKU"])
    df["OrderedQty"] = pd.to_numeric(df["OrderedQty"], errors="coerce")
    df = df.dropna(subset=["SAPCODE", "CustCode", "SKU", "OrderedQty"]).copy()

    # Master mapping is preferred. However, BOT order files can contain a
    # valid CustCode that is not yet present in the master.
    #
    # In that case DO NOT fail the entire report:
    #   DealerGroup = SAPCODE|CustCode
    #   Dealer      = CustCode
    #
    # This keeps the order visible without guessing a dealer name/group.
    df = df.merge(
        cust[["CustCode", "SAPCODE", "DealerGroup", "MasterDealer"]],
        on="CustCode",
        how="left",
        validate="many_to_one",
        suffixes=("", "_MASTER"),
    )

    unknown = df["DealerGroup"].isna()

    if unknown.any():
        unknown_codes = (
            df.loc[unknown, "CustCode"]
            .dropna()
            .astype(str)
            .drop_duplicates()
            .tolist()
        )

        print(
            "      BOT order CustCode(s) not found in master; "
            "using CustCode as dealer name: "
            + ", ".join(unknown_codes)
        )

        # Use the BOT order's SAPCODE because that is the authoritative
        # market/account identifier supplied by the client order file.
        df.loc[unknown, "DealerGroup"] = (
            df.loc[unknown, "SAPCODE"].astype("string")
            + "|"
            + df.loc[unknown, "CustCode"].astype("string")
        )
        df.loc[unknown, "MasterDealer"] = (
            df.loc[unknown, "CustCode"].astype("string")
        )

    # For CustCodes that ARE in the master, make sure the order SAPCODE
    # agrees with the master mapping.
    known = ~unknown
    mismatch = known & df["SAPCODE"].ne(df["SAPCODE_MASTER"])

    if mismatch.any():
        examples = (
            df.loc[mismatch, ["CustCode", "SAPCODE", "SAPCODE_MASTER"]]
            .head(10)
            .to_dict("records")
        )
        raise ValueError(
            "BOT orders has SAPCODE/CustCode combinations that disagree with "
            f"master: {examples}"
        )

    df = df.drop(columns=["SAPCODE_MASTER"], errors="ignore")

    # The BOT order file already has a Market column. It is descriptive data
    # from the source file, not the authoritative market mapping for this
    # automation. Remove it before merging the master SAPCODE -> Market map
    # so pandas cannot create Market_x / Market_y and break downstream code.
    df = df.drop(columns=["Market"], errors="ignore")

    df = df.merge(
        sap,
        on="SAPCODE",
        how="left",
        validate="many_to_one",
    )

    missing_market = df["Market"].isna()

    if missing_market.any():
        examples = (
            df.loc[missing_market, "SAPCODE"]
            .dropna()
            .astype(str)
            .unique()[:10]
        )
        raise ValueError(
            "BOT orders contains SAPCODE(s) not found in master market mapping: "
            f"{examples.tolist()}"
        )

    df["Type"] = "BOT SUB-DEALER"
    df["Dealer"] = clean_text(df["MasterDealer"])

    return (
        df.groupby(
            ["Type", "SAPCODE", "DealerGroup", "SKU"],
            dropna=False,
        )
        .agg(OrderedQty=("OrderedQty", "sum"))
        .reset_index()
    )

def format_qty(value):
    """Format positive shortage quantities for the dealer summary cell."""
    if pd.isna(value):
        return "0"
    value = float(value)
    if value.is_integer():
        return str(int(value))
    return f"{value:g}"


def build_dealer_difference_text(group):
    """
    Show ONLY dealers who did not fully satisfy their wishlist.

    Positive difference:
        Wished > Ordered -> include dealer and shortage.

    Zero:
        Wished == Ordered -> do not include.

    Negative:
        Wished < Ordered -> do not include.

    Example:
        Alyzia Wireless - 15, Shine Wireless - 30

    An over-order such as:
        ORA Investments - -4
    is intentionally NOT shown.
    """
    rows = []

    for _, row in group.iterrows():
        difference = row["DealerDifferenceQty"]

        # Only dealers who ordered less or did not order.
        if pd.isna(difference) or float(difference) <= 0:
            continue

        dealer = row["Dealer"]
        if pd.isna(dealer) or str(dealer).strip() == "":
            dealer = "UNKNOWN DEALER"
        else:
            dealer = str(dealer).strip()

        rows.append((dealer, difference))

    rows.sort(key=lambda x: x[0].casefold())

    return ", ".join(
        f"{dealer} - {format_qty(difference)}"
        for dealer, difference in rows
    )


def build_dealer_difference_text(group):
    """
    Show EVERY dealer for the market/device, not only dealers with a shortage.

    Example:
        Alyzia Wireless - 15, Shine Wireless - 0, ABC Wireless - -5

    Number = WishedQty - OrderedQty for that dealer.
    Positive = ordered less / didn't order.
    Zero = fully ordered.
    Negative = ordered more.
    """
    rows = []

    for _, row in group.iterrows():
        dealer = row["Dealer"]
        if pd.isna(dealer) or str(dealer).strip() == "":
            dealer = "UNKNOWN DEALER"
        else:
            dealer = str(dealer).strip()

        difference = row["DealerDifferenceQty"]
        rows.append((dealer, difference))

    # Stable alphabetical dealer order makes the report easier to compare week to week.
    rows.sort(key=lambda x: x[0].casefold())
    return ", ".join(
        f"{dealer} - {format_qty(difference)}"
        for dealer, difference in rows
    )


def main():
    print("[1/6] Loading master...")
    sap, cust = load_master()
    print(f"      SAP codes: {sap.SAPCODE.nunique():,}")
    print(f"      Cust codes: {cust.CustCode.nunique():,}")
    print(f"      Dealer groups: {cust.DealerGroup.nunique():,}")

    print("[2/6] Loading normal sub-dealer wishlist...")
    normal_wishlist = load_normal_wishlist(sap, cust)
    print(f"      Rows: {len(normal_wishlist):,}")

    print("[3/6] Loading normal sub-dealer orders...")
    normal_orders, normal_sheet = load_normal_orders(sap, cust)
    print(f"      Sheet: {normal_sheet} | Rows: {len(normal_orders):,}")

    print("[4/6] Loading BOT wishlist and orders...")
    bot_wishlist = load_bot_wishlist(sap, cust)
    bot_orders = load_bot_orders(sap, cust)
    print(f"      BOT wishlist rows: {len(bot_wishlist):,}")
    print(f"      BOT order rows: {len(bot_orders):,}")

    print("[5/6] Comparing wishlist against orders...")

    wishes = pd.concat([normal_wishlist, bot_wishlist], ignore_index=True)
    orders = pd.concat([normal_orders, bot_orders], ignore_index=True)

    # IMPORTANT:
    # CustCode is NOT the dealer-level matching key anymore.
    # All CustCodes belonging to one DealerGroup are treated as one dealer.
    match_keys = ["Type", "SAPCODE", "DealerGroup", "SKU"]

    result = wishes.merge(
        orders[match_keys + ["OrderedQty"]],
        on=match_keys,
        how="left",
        validate="one_to_one",
    )

    result["OrderedQty"] = result["OrderedQty"].fillna(0)

    # Core business rule:
    # DealerGroup wished quantity - all orders across that group's CustCodes.
    # Keep the signed value. Negative means the dealer ordered more.
    result["DifferenceQty"] = (
        result["WishedQty"] - result["OrderedQty"]
    )

    result["Status"] = np.select(
        [
            result["OrderedQty"].eq(0),
            result["DifferenceQty"].gt(0) & result["OrderedQty"].gt(0),
            result["DifferenceQty"].eq(0),
            result["DifferenceQty"].lt(0),
        ],
        [
            "DIDN'T ORDER",
            "ORDERED LESS",
            "FULLY ORDERED",
            "ORDERED MORE",
        ],
        default="CHECK",
    )

    # Detailed report. CustCode remains visible so you can see the
    # CustCode that was present in the wishlist, while DealerGroup is
    # the actual dealer-level comparison key.
    # Wishlist quantities are already aggregated across all CustCodes in a
    # DealerGroup, so there is no single wishlist CustCode to display here.
    # Keep the dealer-level identity as DealerGroup + Dealer.
    result = result[
        [
            "Type",
            "SAPCODE",
            "Market",
            "DealerGroup",
            "Dealer",
            "SKU",
            "DeviceName",
            "WishedQty",
            "OrderedQty",
            "DifferenceQty",
            "Status",
        ]
    ].sort_values(
        ["Market", "Type", "Dealer", "SKU"]
    ).reset_index(drop=True)

    # ---------------------------------------------------------
    # MARKET + DEVICE SUMMARY
    # ---------------------------------------------------------
    summary = (
        result.groupby(
            ["SAPCODE", "Market", "SKU", "DeviceName"],
            dropna=False,
            as_index=False,
        )
        .agg(
            WishedQty=("WishedQty", "sum"),
            OrderedQty=("OrderedQty", "sum"),
        )
    )

    # Keep the signed market/device difference.
    summary["DifferenceQty"] = (
        summary["WishedQty"] - summary["OrderedQty"]
    )

    # ---------------------------------------------------------
    # DEALER DETAILS
    # ---------------------------------------------------------
    # DealerGroup is the dealer identity. This combines all CustCodes
    # belonging to the same dealer before calculating the shortage.
    dealer_summary = (
        result.groupby(
            ["SAPCODE", "Market", "SKU", "DealerGroup", "Dealer"],
            dropna=False,
            as_index=False,
        )
        .agg(
            DealerWishedQty=("WishedQty", "sum"),
            DealerOrderedQty=("OrderedQty", "sum"),
        )
    )

    dealer_summary["DealerDifferenceQty"] = (
        dealer_summary["DealerWishedQty"]
        - dealer_summary["DealerOrderedQty"]
    )

    dealer_text = (
        dealer_summary.groupby(
            ["SAPCODE", "Market", "SKU"],
            dropna=False,
        )
        .apply(
            build_dealer_difference_text,
            include_groups=False,
        )
        .reset_index(
            name="Dealers Who Didn't Order / Ordered Less"
        )
    )

    summary = summary.merge(
        dealer_text,
        on=["SAPCODE", "Market", "SKU"],
        how="left",
        validate="one_to_one",
    )

    summary["Dealers Who Didn't Order / Ordered Less"] = (
        summary["Dealers Who Didn't Order / Ordered Less"].fillna("")
    )

    summary = summary[
        [
            "SAPCODE",
            "Market",
            "SKU",
            "DeviceName",
            "WishedQty",
            "OrderedQty",
            "DifferenceQty",
            "Dealers Who Didn't Order / Ordered Less",
        ]
    ].sort_values(
        ["Market", "SKU"]
    ).reset_index(drop=True)

    print("[6/6] Writing report...")
    OUTPUT.mkdir(exist_ok=True)
    outfile = OUTPUT / "Wishlist_Order_Comparison.xlsx"

    with pd.ExcelWriter(outfile, engine="openpyxl") as writer:
        result.to_excel(
            writer,
            sheet_name="Wishlist vs Orders",
            index=False,
        )
        summary.to_excel(
            writer,
            sheet_name="Market Device Summary",
            index=False,
        )

        for sheet_name in ["Wishlist vs Orders", "Market Device Summary"]:
            ws = writer.book[sheet_name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill("solid", fgColor="D9EAF7")
                cell.alignment = Alignment(horizontal="center")

            for i, col in enumerate(ws.columns, 1):
                max_length = max(
                    len(str(c.value)) if c.value is not None else 0
                    for c in col
                )
                ws.column_dimensions[get_column_letter(i)].width = min(
                    max(max_length + 2, 10),
                    60,
                )

    # Console validation totals.
    diff_total = result["DifferenceQty"].sum()
    didnt_order = (result["Status"] == "DIDN'T ORDER").sum()
    ordered_less = (result["Status"] == "ORDERED LESS").sum()
    fully_ordered = (result["Status"] == "FULLY ORDERED").sum()
    ordered_more = (result["Status"] == "ORDERED MORE").sum()

    print(f"      Detailed rows: {len(result):,}")
    print(f"      Market/device rows: {len(summary):,}")
    print(f"      Total Wished:    {result['WishedQty'].sum():,.0f}")
    print(f"      Total Ordered:   {result['OrderedQty'].sum():,.0f}")
    print(f"      Difference:      {diff_total:,.0f}")
    print(f"      Didn't order:    {didnt_order:,}")
    print(f"      Ordered less:    {ordered_less:,}")
    print(f"      Fully ordered:   {fully_ordered:,}")
    print(f"      Ordered more:    {ordered_more:,}")
    print(f"      Report: {outfile}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nERROR: {exc}")
        sys.exit(1)
