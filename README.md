# som-accurate-report-automation

The local Accurate routine (Data\Report\Sales\Accurate Report 2025) on GitHub Actions:

1. Log in to Accurate Online and download the *Delivery Order Detail* report
   (report 10600) for the previous and current month, saved as
   `<MM>. BBL Aura WhiteInc Sales.xlsx` in
   `gs://bucket_som/sales_parquet/raw/primary/accurate/source/<year>/`.
2. Rebuild `Accurate <year>.xlsx` (table `Accurate_<year>`) from that year's
   monthly files - the Power Query of `Accurate 2026.xlsx`, in Python.
   `Accurate 2025.xlsx` is frozen and only read.
3. Rebuild `0. 2025 Accurate.xlsx` (table `Accurate_Report`) = 2025 + later years,
   in `gs://bucket_som/sales_parquet/raw/primary/accurate/`. Sell In reads this.

Steps 2-3 were checked against the local workbooks: identical row for row.

Schedule: 00:45 WIB daily, plus manual runs (Actions -> Run workflow).

Secrets: `GCP_SA_KEY`, `ACCURATE_EMAIL`, `ACCURATE_PASSWORD`. Optional variable
`ACCURATE_DB_NAME` (part of the database name, default `Bintang`).

Local run: copy `.env.example` to `.env`, then `python run_accurate_pipeline.py`.
