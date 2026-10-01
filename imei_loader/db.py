"""SQL Server access for the IMEI loader."""
from __future__ import annotations

import logging
from typing import Iterable, Sequence

import pyodbc

from . import config

log = logging.getLogger(__name__)

# Order matters: these are the columns we insert, in this order.
INSERT_COLUMNS = [
    "Market",
    "orderNo",
    "Cust_Store_id",
    "delivery_pack_list_no",
    "IssueDate",
    "Item",
    "SKU",
    "SKUDescr",
    "IMEI",
    "PONumber",
    "LoadedDate",
]


def _pick_driver() -> str:
    if config.SQL_DRIVER:
        return config.SQL_DRIVER
    installed = [d for d in pyodbc.drivers() if "SQL Server" in d]
    if not installed:
        raise RuntimeError(
            "No SQL Server ODBC driver found. Install 'ODBC Driver 17 for SQL Server' "
            "or set SQL_DRIVER in config.py."
        )
    preferred = [
        "ODBC Driver 18 for SQL Server",
        "ODBC Driver 17 for SQL Server",
        "SQL Server Native Client 11.0",
        "SQL Server",
    ]
    for name in preferred:
        if name in installed:
            return name
    return installed[-1]


def connect() -> pyodbc.Connection:
    driver = _pick_driver()
    parts = [
        f"DRIVER={{{driver}}}",
        f"SERVER={config.SQL_SERVER}",
        f"DATABASE={config.SQL_DATABASE}",
    ]
    if config.USE_TRUSTED_CONNECTION:
        parts.append("Trusted_Connection=yes")
    else:
        parts.append(f"UID={config.SQL_USERNAME}")
        parts.append(f"PWD={config.SQL_PASSWORD}")
    if "18" in driver:
        # Driver 18 defaults to Encrypt=yes, which breaks a plain local instance.
        parts.append("Encrypt=no")
        parts.append("TrustServerCertificate=yes")

    conn_str = ";".join(parts) + ";"
    log.debug("Connecting with driver %s to %s/%s", driver, config.SQL_SERVER, config.SQL_DATABASE)
    conn = pyodbc.connect(conn_str, autocommit=False)
    return conn


def imei_id_is_identity(conn: pyodbc.Connection) -> bool:
    """True if IMEI_ID is an IDENTITY column (so we must not insert it)."""
    sql = """
        SELECT c.is_identity
        FROM sys.columns c
        WHERE c.object_id = OBJECT_ID(?) AND c.name = 'IMEI_ID'
    """
    row = conn.cursor().execute(sql, config.SQL_TABLE).fetchone()
    if row is None:
        raise RuntimeError(
            f"Could not find column IMEI_ID on {config.SQL_TABLE}. "
            "Check SQL_TABLE in config.py."
        )
    return bool(row[0])


def max_imei_id(conn: pyodbc.Connection) -> int:
    sql = f"SELECT ISNULL(MAX(IMEI_ID), 0) FROM {config.SQL_TABLE}"
    return int(conn.cursor().execute(sql).fetchone()[0])


def distinct_markets(conn: pyodbc.Connection) -> list[str]:
    sql = f"SELECT DISTINCT Market FROM {config.SQL_TABLE} WHERE Market IS NOT NULL"
    return [r[0] for r in conn.cursor().execute(sql).fetchall()]


def find_existing_imeis(conn: pyodbc.Connection, imeis: Sequence[str]) -> set[str]:
    """
    Return the subset of `imeis` that already exist in IMEI_INV.

    Uses a temp table + join rather than a giant IN (...) so it scales to
    millions of rows and stays on one index seek per row.
    """
    if not imeis:
        return set()

    cur = conn.cursor()
    cur.execute("IF OBJECT_ID('tempdb..#imei_check') IS NOT NULL DROP TABLE #imei_check;")
    cur.execute("CREATE TABLE #imei_check (IMEI varchar(50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL);")
    cur.execute("CREATE CLUSTERED INDEX ix_imei_check ON #imei_check(IMEI);")

    cur.fast_executemany = True
    batch = config.INSERT_BATCH_SIZE
    for i in range(0, len(imeis), batch):
        chunk = [(v,) for v in imeis[i : i + batch]]
        cur.executemany("INSERT INTO #imei_check (IMEI) VALUES (?)", chunk)
    log.debug("Staged %s IMEIs into #imei_check", len(imeis))

    rows = cur.execute(
        f"""
        SELECT DISTINCT t.IMEI
        FROM #imei_check t
        INNER JOIN {config.SQL_TABLE} i ON i.IMEI = t.IMEI
        """
    ).fetchall()
    cur.execute("DROP TABLE #imei_check;")
    return {r[0] for r in rows}


def insert_rows(conn: pyodbc.Connection, rows: Iterable[tuple], include_imei_id: bool) -> int:
    """
    Insert rows inside the caller's transaction. Caller commits or rolls back.
    `rows` must match INSERT_COLUMNS order, with IMEI_ID prepended when
    include_imei_id is True.
    """
    cols = (["IMEI_ID"] if include_imei_id else []) + INSERT_COLUMNS
    placeholders = ", ".join("?" for _ in cols)
    sql = f"INSERT INTO {config.SQL_TABLE} ({', '.join(cols)}) VALUES ({placeholders})"

    cur = conn.cursor()
    cur.fast_executemany = True

    rows = list(rows)
    total = 0
    batch = config.INSERT_BATCH_SIZE
    for i in range(0, len(rows), batch):
        chunk = rows[i : i + batch]
        cur.executemany(sql, chunk)
        total += len(chunk)
        log.debug("Inserted %s / %s rows", total, len(rows))
    return total
