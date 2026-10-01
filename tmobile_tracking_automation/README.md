# T-Mobile Tracking ID Automation

## What this automation does

1. Reads `input/orders.xlsx`.
2. Expects two columns:
   - `Market`
   - `Order ID`
3. Uses `data/marketlogins.xlsx` to find the T-Mobile credentials for each Market.
4. Logs into the T-Mobile Dealer Ordering portal.
5. Sets:
   - Status = `Any` (`EALL`)
   - Creation Date = `Last 12 Months` (`last_year`)
   - ID Type = `Order Number` (`OBJECT_ID`)
6. Enters each Order ID into the Order Number field.
7. Clicks Go.
8. Opens the returned order.
9. Finds all tracking IDs on the order detail page.
10. Removes duplicate tracking IDs within the same order.
11. Writes:
    - Market
    - Order ID
    - Tracking ID 1
    - Tracking ID 2
    - Tracking ID 3
    - ...
    - Status

## Input file

Put the order file here:

`input/orders.xlsx`

Required columns:

| Market | Order ID |
|---|---|
| Dallas | 172078356 |
| Dallas | 172078357 |
| Houston | 172078425 |

Column capitalization does not matter. The script normalizes the column names.

## Credentials file

Create:

`data/marketlogins.xlsx`

Required columns:

| Market | Username | Password |
|---|---|---|
| Dallas | your_username | your_password |
| Houston | your_username | your_password |

The Market in the order file must match a Market in the credentials file (case-insensitive).

Do NOT commit this credentials file to GitHub.

## Install

Open PowerShell in the project folder:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m playwright install chromium
```

If PowerShell blocks activation, you can run:

```powershell
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m playwright install chromium
```

Then run:

```powershell
python main.py
```

Or double-click `run.bat`.

## Output

The result is saved in:

`output\tracking_results_YYYYMMDD_HHMMSS.xlsx`

Example:

| Market | Order ID | Tracking ID 1 | Tracking ID 2 | Tracking ID 3 | Status |
|---|---:|---|---|---|---|
| Dallas | 172078356 | 1ZAAA | 1ZBBB | | OK |
| Dallas | 172078357 | 1ZCCC | | | OK |
| Houston | 172078425 | | | | NO TRACKING ID |

If the same tracking ID appears on multiple item rows in one order, it is written only once.

## Important

The portal uses generated numeric HTML IDs. This project intentionally uses stable `name=` attributes and visible-frame discovery instead of hard-coding IDs such as `ID595810956`.

The supplied HTML confirms:
- Status uses `rc_status_head1` and `EALL` for Any.
- Creation Date uses `rc_dateattributes_select` and `last_year` for Last 12 Months.
- Order Number uses `rc_attsubcharUI` with `OBJECT_ID`.
- Order ID uses `rc_object_id`.
- Go uses `#gsbuttonstart`.

The order-detail page contains tracking links in the order item status area. The same tracking ID can occur on several item rows, so the script deduplicates them per order.
