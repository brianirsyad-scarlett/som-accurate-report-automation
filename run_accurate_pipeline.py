"""
The local Accurate routine, on GCS:

  1. download the Delivery Order Detail report from Accurate for each month
     (default: previous + current, since deliveries get back-dated) and save it as
     "<MM>. BBL Aura WhiteInc Sales.xlsx" - the name used in the local folder
  2. rebuild "Accurate <year>.xlsx" (table Accurate_<year>) from that year's monthly
     files - Accurate 2025 is frozen and never rebuilt
  3. rebuild "0. 2025 Accurate.xlsx" (table Accurate_Report) = Accurate 2025 + later years

GCS layout (bucket_som):
  sales_parquet/raw/primary/accurate/source/<year>/<MM>. BBL Aura WhiteInc Sales.xlsx
  sales_parquet/raw/primary/accurate/Accurate 2025.xlsx        (frozen, uploaded once)
  sales_parquet/raw/primary/accurate/Accurate <year>.xlsx      (rebuilt)
  sales_parquet/raw/primary/accurate/0. 2025 Accurate.xlsx     (rebuilt; Sell In reads this)

    python run_accurate_pipeline.py                      # prev + current month
    python run_accurate_pipeline.py --month 2026-07 --month 2026-08
    python run_accurate_pipeline.py --skip-download      # rebuild only
"""

from __future__ import annotations

import argparse
import calendar
import datetime as dt
import os
import sys
from pathlib import Path

from google.cloud import storage

import build_accurate_report as B

BUCKET = "bucket_som"
PREFIX = "sales_parquet/raw/primary/accurate"
SOURCE_PREFIX = f"{PREFIX}/source"
REPORT_ID = "10600"
PLAN_ID = "DeliveryOrderDetailReport"
FILE_SUFFIX = ". BBL Aura WhiteInc Sales.xlsx"
FROZEN_YEAR = 2025
HERE = Path(__file__).resolve().parent
WORK = HERE / "work"


def load_env(path: Path) -> None:
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def parse_month(token: str, today: dt.date) -> tuple[int, int]:
    token = token.strip().lower()
    if token in ("current", ""):
        return today.year, today.month
    if token in ("prev", "previous"):
        p = today.replace(day=1) - dt.timedelta(days=1)
        return p.year, p.month
    y, m = token.replace("/", "-").split("-")
    return int(y), int(m)


def download(months: list[tuple[int, int]], bucket, today: dt.date) -> None:
    from accurate_client import AccurateClient
    client = AccurateClient(os.environ["ACCURATE_EMAIL"], os.environ["ACCURATE_PASSWORD"],
                            os.environ.get("ACCURATE_DEVICE_ID", "A-som-github-accurate-pipeline"))
    client.login()
    db = client.open_database(os.environ.get("ACCURATE_DB_NAME", "Bintang"))
    print(f"opened database {db.get('alias') or db.get('name')!r} on {client.host}")
    for year, month in months:
        start = dt.date(year, month, 1)
        end = min(dt.date(year, month, calendar.monthrange(year, month)[1]), today)
        data = client.run_report_xlsx(REPORT_ID, PLAN_ID, start.strftime("%d/%m/%Y"), end.strftime("%d/%m/%Y"))
        name = f"{month:02d}{FILE_SUFFIX}"
        local = WORK / "source" / str(year) / name
        local.parent.mkdir(parents=True, exist_ok=True)
        local.write_bytes(data)
        key = f"{SOURCE_PREFIX}/{year}/{name}"
        bucket.blob(key).upload_from_filename(str(local), timeout=600)
        print(f"{year}-{month:02d} ({start:%d/%m/%Y}..{end:%d/%m/%Y}): {len(data):,} bytes -> gs://{BUCKET}/{key}")


def rebuild(bucket) -> None:
    # Every monthly file of every non-frozen year, fresh from GCS.
    years: dict[int, Path] = {}
    for blob in bucket.list_blobs(prefix=f"{SOURCE_PREFIX}/"):
        parts = blob.name[len(SOURCE_PREFIX) + 1:].split("/")
        if len(parts) != 2 or not parts[1].endswith(".xlsx") or parts[1].startswith("~$"):
            continue
        year = int(parts[0])
        local = WORK / "source" / str(year) / parts[1]
        local.parent.mkdir(parents=True, exist_ok=True)
        blob.download_to_filename(str(local), timeout=600)
        years[year] = local.parent

    frozen = WORK / f"Accurate {FROZEN_YEAR}.xlsx"
    bucket.blob(f"{PREFIX}/Accurate {FROZEN_YEAR}.xlsx").download_to_filename(str(frozen), timeout=600)
    y2025 = B.accurate_2025(frozen)
    print(f"Accurate_{FROZEN_YEAR} (frozen): {y2025.height:,} rows")

    year_tables = []
    for year in sorted(y for y in years if y > FROZEN_YEAR):
        t = B.accurate_year(years[year])
        out = WORK / f"Accurate {year}.xlsx"
        B.write_table_xlsx(t, out, f"Accurate_{year}", str(year))
        bucket.blob(f"{PREFIX}/{out.name}").upload_from_filename(str(out), timeout=600)
        print(f"Accurate_{year}: {t.height:,} rows -> gs://{BUCKET}/{PREFIX}/{out.name}")
        year_tables.append(t)

    report = B.accurate_report(y2025, year_tables)
    out = WORK / "0. 2025 Accurate.xlsx"
    B.write_table_xlsx(report, out, "Accurate_Report", "Accurate_Report")
    bucket.blob(f"{PREFIX}/{out.name}").upload_from_filename(str(out), timeout=600)
    print(f"Accurate_Report: {report.height:,} rows -> gs://{BUCKET}/{PREFIX}/{out.name}")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--month", action="append", default=[])
    ap.add_argument("--skip-download", action="store_true")
    args = ap.parse_args(argv)
    load_env(HERE / ".env")

    today = dt.date.today()
    months = [parse_month(m, today) for m in (args.month or ["prev", "current"])]
    months = [m for m in months if m[0] > FROZEN_YEAR]
    bucket = storage.Client().bucket(BUCKET)
    if not args.skip_download:
        download(months, bucket, today)
    rebuild(bucket)
    return 0


if __name__ == "__main__":
    sys.exit(main())
