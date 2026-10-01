"""
Configuration for the IMEI loader.

This is the only file you should normally need to edit.
"""
from pathlib import Path

# ---------------------------------------------------------------------------
# Folders
# ---------------------------------------------------------------------------
# BASE_DIR = the imei_loader package folder itself.
# PROJECT_ROOT = its parent (your existing automation project root).
BASE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = BASE_DIR.parent

IMEI_ROOT = PROJECT_ROOT / "imei"
INCOMING_DIR = IMEI_ROOT / "incoming"
PROCESSED_DIR = IMEI_ROOT / "processed"
FAILED_DIR = IMEI_ROOT / "failed"
LOG_DIR = PROJECT_ROOT / "logs"

# ---------------------------------------------------------------------------
# SQL Server
# ---------------------------------------------------------------------------
SQL_SERVER = "localhost"
SQL_DATABASE = "PurelyWirelessDB"
SQL_TABLE = "dbo.IMEI_INV"

# Windows authentication. Set to False and fill in the username/password
# below if you switch to a SQL login.
USE_TRUSTED_CONNECTION = True
SQL_USERNAME = ""
SQL_PASSWORD = ""

# Leave as None to auto-detect the newest installed ODBC driver.
SQL_DRIVER = None

# ---------------------------------------------------------------------------
# Markets
# ---------------------------------------------------------------------------
# Canonical market names -- these are the EXACT strings written to IMEI_INV.
# Spellings here match what is already in the table, so existing rows and new
# rows stay consistent.
CANONICAL_MARKETS = [
    "HOUSTON",
    "DALLAS",
    "AZ",
    "Tulsa",
    "Waco",
    "CORPUS",
    "Tampa",
    "RGV",
    "SA",
]

# Maps what the portal export writes -> canonical name above.
# Keys are matched case-insensitively with surrounding whitespace stripped.
# A market in the file that is neither a canonical name nor a key here is an
# ERROR: the file is rejected and you get a notification.
MARKET_ALIASES = {
    "PHOENIX": "AZ",
    "ARIZONA": "AZ",
    "SAN ANTONIO": "SA",
    "SAN ANTONIO/AUSTIN": "SA",
    "CORPUS CHRISTI": "CORPUS",
    "TULSA": "Tulsa",
    "WACO": "Waco",
    "TAMPA": "Tampa",
    "HOUSTON": "HOUSTON",
    "DALLAS": "DALLAS",
    "RGV": "RGV",
}

# If True, the canonical list above is refreshed from
# SELECT DISTINCT Market FROM IMEI_INV at run time and merged with the list.
# Off by default: the table cannot validate markets it has never seen.
REFRESH_MARKETS_FROM_DB = False

# ---------------------------------------------------------------------------
# File handling
# ---------------------------------------------------------------------------
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".csv"}

# Expected header, exactly as the portal export writes it.
EXPECTED_COLUMNS = [
    "Market",
    "Order No",
    "Cust/Store",
    "Delivery(Pack List)",
    "Issue Date",
    "Item",
    "SKU",
    "SKUDescr",
    "IMEI/Serial Number",
    "PO Number",
]

# Excel header -> IMEI_INV column
COLUMN_MAP = {
    "Market": "Market",
    "Order No": "orderNo",
    "Cust/Store": "Cust_Store_id",
    "Delivery(Pack List)": "delivery_pack_list_no",
    "Issue Date": "IssueDate",
    "Item": "Item",
    "SKU": "SKU",
    "SKUDescr": "SKUDescr",
    "IMEI/Serial Number": "IMEI",
    "PO Number": "PONumber",
}

# A file is only picked up once its size has been unchanged for this long.
# Stops the job from reading a file you are still copying in.
FILE_STABLE_SECONDS = 30

# IMEIs must be exactly this many digits. Set to None to skip the check.
IMEI_EXPECTED_LENGTH = 15

# Rows are inserted in batches of this size.
INSERT_BATCH_SIZE = 5000

# ---------------------------------------------------------------------------
# Notifications
# ---------------------------------------------------------------------------
NOTIFY_APP_NAME = "IMEI Loader"
NOTIFY_TIMEOUT_SECONDS = 20
