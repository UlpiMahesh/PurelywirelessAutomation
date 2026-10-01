# Wishlist Order Automation

## Business rule

For each Market + DealerGroup + SKU:

    Wished Qty - Ordered Qty = Difference

All CustCodes assigned to the same DealerGroup are combined. Therefore a wishlist under one CustCode can be fulfilled by an order made under another CustCode belonging to the same DealerGroup.

### Dealer shortage list

Only dealers who ordered less than wished or did not order are shown:

- Wished 20 / Ordered 0 -> Dealer - 20
- Wished 20 / Ordered 15 -> Dealer - 5
- Wished 20 / Ordered 20 -> not shown
- Wished 20 / Ordered 25 -> not shown

Negative differences are retained in the numeric DifferenceQty but are never placed in the dealer shortage list.

## Master duplicate handling

The `sub dealer cust codes` sheet may contain multiple rows for the same CustCode.

The automation:
1. Requires a CustCode to belong to one SAPCODE.
2. Allows duplicate dealer-name aliases for a CustCode.
3. If conflicting DealerGroups exist, it resolves the CustCode to the unique DealerGroup shared by multiple CustCodes under the same SAPCODE.
4. If the ambiguity cannot be safely resolved, it stops instead of inventing a mapping.

The supplied master was resolved for the three ambiguous CustCodes:
- 4040827 -> PM Wireless
- 4040829 -> PM Wireless
- 4040830 -> PM Wireless

An `DealerGroup Audit` sheet documents the final mapping.

## Input files

Place these in `input`:
- master.xlsx
- normal_wishlist.xlsx
- normal_orders.xlsx
- bot_wishlist.xlsx
- bot_orders.xlsx

## Output

`output/Wishlist_Order_Comparison.xlsx`

Sheets:
1. Wishlist vs Orders
2. Market Device Summary

Run with `run.bat`.

### Unknown BOT order CustCodes

If a CustCode appears in `bot_orders.xlsx` but is not yet present in
`master.xlsx`, the automation does NOT stop.

It uses:
- Dealer name = CustCode
- DealerGroup = SAPCODE|CustCode

Example:
`4041804` -> dealer `4041804`

This keeps the order in the report without guessing which dealer the
CustCode belongs to. Add it to the master later if you want it grouped
with other CustCodes belonging to the same dealer.
