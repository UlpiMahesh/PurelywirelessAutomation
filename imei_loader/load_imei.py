"""
IMEI weekly loader.

Scans imei/incoming for CSV/Excel IMEI exports, validates them, checks for
duplicate IMEIs both inside the file and against IMEI_INV, and loads clean
files into SQL Server in a single transaction.

A file either loads completely or not at all. Nothing is ever half-loaded.

Run:  python -m imei_loader.load_imei
"""
from __future__ import annotations

import logging
import re
import shutil
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd

from . import config, db
from .notify import notify

log = logging.getLogger("imei_loader")

MONTHS = "jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec"


class FileRejected(Exception):
    """A file failed validation. Message is shown in the notification."""

    def __init__(self, message: str, detail: str = "", report: pd.DataFrame | None = None):
        super().__init__(message)
        self.message = message
        self.detail = detail
        self.report = report


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
def setup_logging() -> Path:
    config.LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_path = config.LOG_DIR / "imei_loader.log"

    log.setLevel(logging.DEBUG)
    log.handlers.clear()

    from logging.handlers import RotatingFileHandler

    fh = RotatingFileHandler(log_path, maxBytes=5_000_000, backupCount=10, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(
        logging.Formatter("%(asctime)s | %(levelname)-8s | %(name)s | %(message)s")
    )

    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(logging.Formatter("%(asctime)s | %(levelname)-8s | %(message)s"))

    log.addHandler(fh)
    log.addHandler(ch)
    logging.getLogger("imei_loader.db").parent = log
    return log_path


# ---------------------------------------------------------------------------
# Folder / file discovery
# ---------------------------------------------------------------------------
def ensure_folders() -> None:
    for d in (config.INCOMING_DIR, config.PROCESSED_DIR, config.FAILED_DIR, config.LOG_DIR):
        d.mkdir(parents=True, exist_ok=True)


def is_stable(path: Path) -> bool:
    """True if the file size has not changed over FILE_STABLE_SECONDS."""
    try:
        first = path.stat().st_size
    except OSError:
        return False
    time.sleep(min(config.FILE_STABLE_SECONDS, 30))
    try:
        return path.stat().st_size == first
    except OSError:
        return False


def discover_files() -> list[Path]:
    files = [
        p
        for p in sorted(config.INCOMING_DIR.iterdir())
        if p.is_file()
        and p.suffix.lower() in config.ALLOWED_EXTENSIONS
        and not p.name.startswith("~$")  # Excel lock files
    ]
    return files


def parse_period_label(filename: str) -> str:
    """
    Best-effort date range out of the filename, used only for logging and the
    notification text. Never fatal.

    Handles: IDOO_IMEI_7-12_SEP, IMEI_14-sep_to_19-sep, IMEI(14 sep - 19 sep)
    """
    name = Path(filename).stem
    lowered = name.lower()

    # 7-12_SEP  /  7 - 12 sep
    m = re.search(rf"(\d{{1,2}})\s*[-_to ]+\s*(\d{{1,2}})[\s_\-]*({MONTHS})", lowered)
    if m:
        return f"{m.group(1)}-{m.group(2)} {m.group(3).upper()}"

    # 14-sep to 19-sep  /  14 sep - 19 sep
    pairs = re.findall(rf"(\d{{1,2}})\s*[-_ ]?\s*({MONTHS})", lowered)
    if len(pairs) >= 2:
        return f"{pairs[0][0]} {pairs[0][1].upper()} - {pairs[-1][0]} {pairs[-1][1].upper()}"
    if len(pairs) == 1:
        return f"{pairs[0][0]} {pairs[0][1].upper()}"

    log.warning("Could not parse a date range from filename '%s'", filename)
    return name


# ---------------------------------------------------------------------------
# Reading & validation
# ---------------------------------------------------------------------------
def read_file(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    try:
        if suffix == ".csv":
            df = pd.read_csv(path, dtype=str, keep_default_na=False, na_values=[""])
        elif suffix == ".xls":
            df = pd.read_excel(path, dtype=str, engine="xlrd")
        else:
            df = pd.read_excel(path, dtype=str, engine="openpyxl")
    except Exception as exc:  # noqa: BLE001
        raise FileRejected(
            "File could not be opened",
            f"{type(exc).__name__}: {exc}",
        ) from exc

    df.columns = [str(c).strip() for c in df.columns]
    return df


def validate_header(df: pd.DataFrame) -> None:
    expected = set(config.EXPECTED_COLUMNS)
    actual = set(df.columns)
    missing = expected - actual
    if missing:
        raise FileRejected(
            "File header does not match the expected layout",
            "Missing columns: " + ", ".join(sorted(missing))
            + "\nColumns found: " + ", ".join(df.columns),
        )
    extra = actual - expected
    if extra:
        log.warning("Ignoring unexpected columns: %s", ", ".join(sorted(extra)))


def clean_frame(df: pd.DataFrame) -> pd.DataFrame:
    df = df[config.EXPECTED_COLUMNS].copy()
    for col in df.columns:
        df[col] = df[col].astype("string").str.strip()
    # Drop rows that are entirely blank (trailing junk rows in Excel).
    df = df.dropna(how="all")
    df = df[~(df.fillna("") == "").all(axis=1)]
    return df.reset_index(drop=True)


def build_market_lookup(conn) -> dict[str, str]:
    """Maps an upper-cased/trimmed market token -> canonical DB spelling."""
    canonical = list(config.CANONICAL_MARKETS)
    if config.REFRESH_MARKETS_FROM_DB:
        from_db = db.distinct_markets(conn)
        log.info("Markets currently in IMEI_INV: %s", ", ".join(sorted(from_db)))
        for m in from_db:
            if m not in canonical:
                canonical.append(m)

    lookup = {m.strip().upper(): m for m in canonical}
    for alias, target in config.MARKET_ALIASES.items():
        if target not in canonical:
            log.warning(
                "Alias '%s' points at '%s', which is not in CANONICAL_MARKETS. Ignoring.",
                alias,
                target,
            )
            continue
        lookup[alias.strip().upper()] = target
    return lookup


def validate_markets(df: pd.DataFrame, lookup: dict[str, str]) -> pd.DataFrame:
    keys = df["Market"].fillna("").str.strip().str.upper()
    unknown = sorted({k for k in keys.unique() if k not in lookup})
    if unknown:
        counts = keys[keys.isin(unknown)].value_counts().to_dict()
        detail_lines = [f"  {name}  ({n} rows)" for name, n in counts.items()]
        raise FileRejected(
            "Unknown market name(s): " + ", ".join(unknown),
            "These market values are not in CANONICAL_MARKETS and have no alias:\n"
            + "\n".join(detail_lines)
            + "\n\nFix: either correct the file, or add an entry to MARKET_ALIASES "
            "in config.py (e.g. \"PHOENIX\": \"AZ\").",
        )
    df = df.copy()
    df["Market"] = keys.map(lookup)
    return df


def validate_imeis(df: pd.DataFrame) -> None:
    imei = df["IMEI"].fillna("")

    blank = df[imei == ""]
    if len(blank):
        raise FileRejected(
            f"{len(blank)} row(s) have a blank IMEI",
            "First offending rows (file row numbers, header = row 1):\n"
            + ", ".join(str(i + 2) for i in blank.index[:25]),
        )

    if config.IMEI_EXPECTED_LENGTH:
        bad = df[
            ~imei.str.fullmatch(rf"\d{{{config.IMEI_EXPECTED_LENGTH}}}").fillna(False)
        ]
        if len(bad):
            sample = bad[["Market", "orderNo", "IMEI"]].head(25)
            raise FileRejected(
                f"{len(bad)} IMEI(s) are not {config.IMEI_EXPECTED_LENGTH} digits",
                "Sample:\n" + sample.to_string(index=False),
                report=bad,
            )


def parse_issue_dates(df: pd.DataFrame) -> pd.DataFrame:
    parsed = pd.to_datetime(df["IssueDate"], errors="coerce", format="mixed", dayfirst=False)
    bad = df[parsed.isna()]
    if len(bad):
        raise FileRejected(
            f"{len(bad)} row(s) have an unreadable Issue Date",
            "Sample values: " + ", ".join(bad["IssueDate"].fillna("(blank)").head(10)),
            report=bad,
        )
    df = df.copy()
    df["IssueDate"] = parsed.dt.date
    return df


def check_duplicates(conn, df: pd.DataFrame) -> None:
    """
    Two-stage duplicate check, both run BEFORE anything is inserted.
      1. duplicates within the file itself
      2. IMEIs already present in IMEI_INV
    Either one rejects the whole file.
    """
    dup_mask = df["IMEI"].duplicated(keep=False)
    if dup_mask.any():
        dups = df[dup_mask].sort_values("IMEI")
        n_unique = dups["IMEI"].nunique()
        raise FileRejected(
            f"{n_unique} IMEI(s) appear more than once inside the file",
            "Sample:\n"
            + dups[["Market", "orderNo", "IssueDate", "IMEI"]].head(25).to_string(index=False),
            report=dups,
        )

    imeis = df["IMEI"].tolist()
    log.info("Checking %s IMEIs against %s ...", len(imeis), config.SQL_TABLE)
    existing = db.find_existing_imeis(conn, imeis)
    if existing:
        clash = df[df["IMEI"].isin(existing)].sort_values("IMEI")
        raise FileRejected(
            f"{len(existing)} IMEI(s) already exist in the database",
            "Nothing was inserted. Sample:\n"
            + clash[["Market", "orderNo", "IssueDate", "IMEI"]].head(25).to_string(index=False),
            report=clash,
        )
    log.info("Duplicate check passed: all %s IMEIs are new", len(imeis))


# ---------------------------------------------------------------------------
# Load
# ---------------------------------------------------------------------------
def load_dataframe(conn, df: pd.DataFrame) -> int:
    include_id = not db.imei_id_is_identity(conn)
    loaded_at = datetime.now()

    values = []
    if include_id:
        next_id = db.max_imei_id(conn) + 1
        log.info("IMEI_ID is not an IDENTITY column; assigning ids from %s", next_id)
    else:
        next_id = None
        log.debug("IMEI_ID is an IDENTITY column; letting SQL Server assign ids")

    for offset, row in enumerate(df.itertuples(index=False)):
        record = (
            row.Market,
            row.orderNo,
            row.Cust_Store_id,
            row.delivery_pack_list_no,
            row.IssueDate,
            row.Item,
            row.SKU,
            row.SKUDescr,
            row.IMEI,
            row.PONumber,
            loaded_at,
        )
        if include_id:
            record = (next_id + offset,) + record
        values.append(record)

    inserted = db.insert_rows(conn, values, include_imei_id=include_id)
    return inserted


# ---------------------------------------------------------------------------
# Per-file orchestration
# ---------------------------------------------------------------------------
def move_file(path: Path, dest_dir: Path) -> Path:
    dest_dir.mkdir(parents=True, exist_ok=True)
    target = dest_dir / path.name
    if target.exists():
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target = dest_dir / f"{path.stem}__{stamp}{path.suffix}"
    shutil.move(str(path), str(target))
    log.info("Moved %s -> %s", path.name, target)
    return target


def write_error_sidecar(moved: Path, message: str, detail: str, report: pd.DataFrame | None) -> None:
    txt = moved.with_suffix(moved.suffix + ".error.txt")
    body = [
        f"File:    {moved.name}",
        f"Failed:  {datetime.now():%Y-%m-%d %H:%M:%S}",
        f"Reason:  {message}",
        "",
        detail or "(no further detail)",
        "",
        "Nothing from this file was written to the database.",
        "Fix the file, then move it back into imei/incoming to retry.",
    ]
    txt.write_text("\n".join(body), encoding="utf-8")
    log.info("Wrote error detail to %s", txt.name)

    if report is not None and len(report):
        csv_path = moved.with_suffix(moved.suffix + ".rejected_rows.csv")
        try:
            report.to_csv(csv_path, index=False)
            log.info("Wrote %s offending rows to %s", len(report), csv_path.name)
        except Exception:  # noqa: BLE001
            log.exception("Could not write rejected-rows report")


def process_file(path: Path, conn) -> bool:
    """Returns True on success. Handles its own move + notification."""
    label = parse_period_label(path.name)
    log.info("=" * 70)
    log.info("Processing '%s' (period: %s)", path.name, label)

    try:
        df = read_file(path)
        log.info("Read %s rows", len(df))
        validate_header(df)
        df = clean_frame(df)
        df = df.rename(columns=config.COLUMN_MAP)
        log.info("%s rows after cleaning", len(df))
        if df.empty:
            raise FileRejected("File contains no data rows", "")

        lookup = build_market_lookup(conn)
        df = validate_markets(df, lookup)
        log.info("Markets in file: %s", ", ".join(sorted(df["Market"].unique())))

        validate_imeis(df)
        df = parse_issue_dates(df)
        check_duplicates(conn, df)

        inserted = load_dataframe(conn, df)
        conn.commit()
        log.info("COMMITTED %s rows from '%s'", inserted, path.name)

    except FileRejected as rej:
        conn.rollback()
        log.error("REJECTED '%s': %s", path.name, rej.message)
        if rej.detail:
            log.error("Detail:\n%s", rej.detail)
        moved = move_file(path, config.FAILED_DIR)
        write_error_sidecar(moved, rej.message, rej.detail, rej.report)
        notify(
            "IMEI upload FAILED",
            f"{path.name}\n{rej.message}\n\nMoved to imei\\failed. Nothing was inserted.",
        )
        return False

    except Exception as exc:  # noqa: BLE001
        conn.rollback()
        log.error("UNEXPECTED failure on '%s': %s", path.name, exc)
        log.debug("Traceback:\n%s", traceback.format_exc())
        moved = move_file(path, config.FAILED_DIR)
        write_error_sidecar(
            moved,
            f"Unexpected error: {type(exc).__name__}",
            f"{exc}\n\n{traceback.format_exc()}",
            None,
        )
        notify(
            "IMEI upload FAILED",
            f"{path.name}\nUnexpected error: {type(exc).__name__}\n"
            "See logs\\imei_loader.log. Nothing was inserted.",
        )
        return False

    move_file(path, config.PROCESSED_DIR)
    notify(
        "IMEIs uploaded to DB",
        f"{inserted:,} IMEIs loaded from {path.name}\nPeriod: {label}\n"
        f"Loaded: {datetime.now():%Y-%m-%d %H:%M}",
    )
    return True


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main() -> int:
    log_path = setup_logging()
    log.info("#" * 70)
    log.info("IMEI loader started | log: %s", log_path)

    ensure_folders()
    files = discover_files()
    if not files:
        log.info("No files in %s. Nothing to do.", config.INCOMING_DIR)
        return 0

    log.info("Found %s candidate file(s): %s", len(files), ", ".join(f.name for f in files))

    ready = []
    for f in files:
        if is_stable(f):
            ready.append(f)
        else:
            log.warning("'%s' is still being written; skipping until next run", f.name)
    if not ready:
        return 0

    try:
        conn = db.connect()
    except Exception as exc:  # noqa: BLE001
        log.error("Could not connect to SQL Server: %s", exc)
        log.debug("Traceback:\n%s", traceback.format_exc())
        notify(
            "IMEI upload FAILED",
            f"Could not connect to {config.SQL_SERVER}/{config.SQL_DATABASE}.\n"
            "Files left in imei\\incoming. See logs\\imei_loader.log.",
        )
        return 2

    ok = failed = 0
    try:
        for f in ready:
            if process_file(f, conn):
                ok += 1
            else:
                failed += 1
    finally:
        conn.close()

    log.info("Run finished: %s loaded, %s failed", ok, failed)
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
