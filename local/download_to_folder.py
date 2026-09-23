"""
Local step 1: download the Delivery Order Detail report from Accurate straight into
the local Accurate folder, as "<MM>. BBL Aura WhiteInc Sales.xlsx" - the files the
Power Query in "Accurate 2026.xlsx" reads. The previous copy of each file is moved
to a backup folder first.

    python local/download_to_folder.py                    # prev + current month
    python local/download_to_folder.py --month 2026-09
"""

from __future__ import annotations

import argparse
import calendar
import datetime as dt
import os
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))

import run_accurate_pipeline as P  # noqa: E402  (REPORT_ID, PLAN_ID, load_env, parse_month)
from accurate_client import AccurateClient  # noqa: E402

LOCAL_DIR = Path(r"D:\SCARLETT_512\SCARLETT-329\SOM\Data\Report\Sales\Accurate Report 2025")
# Outside the Accurate folder on purpose: Power Query reads every "*Sales*" file
# in that folder (recursively), so a backup copy inside it would be counted twice.
BACKUP_DIR = HERE.parent / "work" / "local_backups"


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", action="append", default=[])
    ap.add_argument("--dir", type=Path, default=LOCAL_DIR)
    args = ap.parse_args(argv)
    P.load_env(HERE.parent / ".env")

    today = dt.date.today()
    months = [P.parse_month(m, today) for m in (args.month or ["prev", "current"])]
    months = [m for m in months if m[0] > P.FROZEN_YEAR]

    client = AccurateClient(os.environ["ACCURATE_EMAIL"], os.environ["ACCURATE_PASSWORD"],
                            os.environ.get("ACCURATE_DEVICE_ID", "A-som-github-accurate-pipeline"))
    client.login()
    db = client.open_database(os.environ.get("ACCURATE_DB_NAME", "Bintang"))
    print(f"opened database {db.get('alias') or db.get('name')!r}")

    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    for year, month in months:
        start = dt.date(year, month, 1)
        end = min(dt.date(year, month, calendar.monthrange(year, month)[1]), today)
        data = client.run_report_xlsx(P.REPORT_ID, P.PLAN_ID, start.strftime("%d/%m/%Y"), end.strftime("%d/%m/%Y"))
        target = args.dir / f"{month:02d}{P.FILE_SUFFIX}"
        BACKUP_DIR.mkdir(parents=True, exist_ok=True)
        if target.exists():
            shutil.copy2(target, BACKUP_DIR / f"{target.stem} {stamp}{target.suffix}")
        # Staged outside the Accurate folder (same drive), then swapped in whole.
        tmp = BACKUP_DIR / f"incoming {target.name}"
        tmp.write_bytes(data)
        os.replace(tmp, target)
        print(f"{year}-{month:02d} ({start:%d/%m/%Y}..{end:%d/%m/%Y}): {len(data):,} bytes -> {target}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
