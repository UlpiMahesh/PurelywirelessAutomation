"""
Windows desktop notifications, with graceful degradation.

Order of attempts:
  1. win11toast  (best looking, clickable)
  2. plyer       (works on Win10/11)
  3. PowerShell balloon tip
  4. A .txt file dropped next to the logs, so a missed toast is never a
     silently missed failure.

IMPORTANT: Windows toasts only appear when the script runs in an interactive
desktop session. In Task Scheduler you MUST use "Run only when user is logged
on". With "Run whether user is logged on or not" the task runs in session 0
and every toast is swallowed with no error.
"""
from __future__ import annotations

import logging
import subprocess
from datetime import datetime
from pathlib import Path

from . import config

log = logging.getLogger(__name__)


def _try_win11toast(title: str, message: str) -> bool:
    try:
        from win11toast import toast  # type: ignore

        toast(title, message, app_id=config.NOTIFY_APP_NAME, duration="long")
        return True
    except Exception as exc:  # noqa: BLE001
        log.debug("win11toast unavailable/failed: %s", exc)
        return False


def _try_plyer(title: str, message: str) -> bool:
    try:
        from plyer import notification  # type: ignore

        notification.notify(
            title=title,
            message=message,
            app_name=config.NOTIFY_APP_NAME,
            timeout=config.NOTIFY_TIMEOUT_SECONDS,
        )
        return True
    except Exception as exc:  # noqa: BLE001
        log.debug("plyer unavailable/failed: %s", exc)
        return False


def _try_powershell(title: str, message: str) -> bool:
    """Classic tray balloon. No extra packages needed."""
    safe_title = title.replace("'", "''")
    safe_msg = message.replace("'", "''")
    script = (
        "[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms');"
        "$b = New-Object System.Windows.Forms.NotifyIcon;"
        "$b.Icon = [System.Drawing.SystemIcons]::Information;"
        "$b.BalloonTipIcon = 'Info';"
        f"$b.BalloonTipTitle = '{safe_title}';"
        f"$b.BalloonTipText = '{safe_msg}';"
        "$b.Visible = $true;"
        f"$b.ShowBalloonTip({config.NOTIFY_TIMEOUT_SECONDS * 1000});"
        f"Start-Sleep -Seconds {min(config.NOTIFY_TIMEOUT_SECONDS, 10)};"
        "$b.Dispose();"
    )
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
            check=True,
            capture_output=True,
            timeout=60,
        )
        return True
    except Exception as exc:  # noqa: BLE001
        log.debug("powershell balloon failed: %s", exc)
        return False


def _write_fallback_file(title: str, message: str) -> None:
    try:
        config.LOG_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path: Path = config.LOG_DIR / f"NOTIFY_{stamp}.txt"
        path.write_text(f"{title}\n\n{message}\n", encoding="utf-8")
        log.warning("Desktop notification could not be shown; wrote %s", path)
    except Exception:  # noqa: BLE001
        log.exception("Could not write notification fallback file")


def notify(title: str, message: str) -> None:
    """Show a desktop notification. Never raises."""
    log.info("NOTIFY | %s | %s", title, message.replace("\n", " / "))
    for attempt in (_try_win11toast, _try_plyer, _try_powershell):
        try:
            if attempt(title, message):
                return
        except Exception:  # noqa: BLE001
            log.exception("Notification backend raised")
    _write_fallback_file(title, message)
