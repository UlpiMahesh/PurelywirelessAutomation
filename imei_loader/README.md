# IMEI Loader

Drops into your existing automation project as a package. Scans `imei/incoming`,
validates, checks duplicates, loads clean files into `PurelyWirelessDB.dbo.IMEI_INV`.

## Layout

```
<your project root>\
├── imei_loader\          <- this package
│   ├── __init__.py
│   ├── config.py         <- the only file you normally edit
│   ├── db.py
│   ├── load_imei.py
│   ├── notify.py
│   ├── run.bat
│   └── requirements.txt
├── imei\                 <- created automatically on first run
│   ├── incoming\         <- drop weekly files here
│   ├── processed\        <- successful loads land here
│   └── failed\           <- rejects land here, with .error.txt next to them
└── logs\
    └── imei_loader.log   <- rotating, 5 MB x 10
```

## Install

```bat
cd <your project root>
pip install -r imei_loader\requirements.txt
```

You also need an ODBC driver: **ODBC Driver 17 for SQL Server** (or 18).
The loader auto-detects whichever is installed.

## Run manually

```bat
cd <your project root>
python -m imei_loader.load_imei
```

Exit codes: `0` nothing to do / all loaded, `1` at least one file rejected,
`2` could not connect to SQL Server.

## Task Scheduler

Create a Basic Task:

- **Trigger:** Daily, whatever time suits. An empty `incoming` folder is a
  no-op that finishes in under a second, so there is no reason to skip days.
- **Action:** Start a program
  - Program: `C:\path\to\project\imei_loader\run.bat`
  - Start in: `C:\path\to\project`  ← **must be set, or the import fails**
- **General tab:** select **"Run only when user is logged on"**.

That last one is not optional. With *"Run whether user is logged on or not"*
the task runs in session 0, where Windows silently discards every toast
notification. You would get no success and no failure alerts, with no error
anywhere. If the machine is not logged in at that hour, notifications are not a
workable alert channel — use the log file or ask me to add email alerts.

## How a file is judged

A file is loaded **completely or not at all**. Everything below happens before
a single row is inserted; on any failure the transaction is rolled back, the
file moves to `failed\`, and you get a notification.

1. **Stability** — the file's size must be unchanged for 30s, so a file you are
   still copying in is skipped until the next run.
2. **Header** — must contain all ten portal export columns.
3. **Markets** — each value is upper-cased, trimmed, then resolved against
   `CANONICAL_MARKETS` + `MARKET_ALIASES`. Unknown → reject.
4. **IMEIs** — no blanks, exactly 15 digits.
5. **Issue Date** — every row must parse.
6. **Duplicates inside the file** — any IMEI appearing twice → reject.
7. **Duplicates against the DB** — any IMEI already in `IMEI_INV` → reject.
8. **Insert** — one transaction, batched 5,000 rows at a time.

## Market aliases

Your portal export and your database disagree on two names:

| In the file   | In IMEI_INV |
|---------------|-------------|
| `PHOENIX`     | `AZ`        |
| `SAN ANTONIO` | `SA`        |

Both are already in `MARKET_ALIASES`. When a new market appears, add one line:

```python
MARKET_ALIASES = {
    ...
    "OKLAHOMA CITY": "OKC",
}
```

and add `"OKC"` to `CANONICAL_MARKETS`. Canonical spellings are written to the
table verbatim, so keep them matching what is already there (`Tulsa`, not
`TULSA`).

## When a file is rejected

`imei\failed\` gets three things:

- the original file, untouched
- `<file>.error.txt` — the reason, plus the exact offending values
- `<file>.rejected_rows.csv` — the specific rows at fault (duplicates, bad
  IMEIs, bad dates), so you can fix them without hunting

Fix the file, move it back into `imei\incoming`, and it will be picked up on
the next run.

## IMEI_ID

The loader checks `sys.columns.is_identity` at run time. If `IMEI_ID` is an
IDENTITY column it lets SQL Server assign ids; if not, it assigns
`MAX(IMEI_ID) + 1` upward. Either way it works, but if it is **not** an
IDENTITY column you should make it one — the `MAX + 1` path is not safe if
anything else writes to the table at the same time.

## Recommended index

The duplicate check joins on `IMEI`. Without an index it will table-scan, which
gets slow as the table grows:

```sql
CREATE UNIQUE INDEX IX_IMEI_INV_IMEI ON dbo.IMEI_INV (IMEI);
```

Make it `UNIQUE` if you want the database itself to enforce what the loader
already checks — a second line of defence against anything inserting outside
this pipeline. It will fail to create if you already have duplicates, which is
a useful thing to find out.
