# SAS_DB — data pipeline + free shareable dashboard

Replaces the manual Excel-merge + Power BI workflow with two automated steps.

## Setup (once)

```powershell
pip install -r requirements.txt
```

## One-time setup: the permanent shareable link (GitHub Pages)

You do this **once**. After it, everyone uses the same URL forever and it
updates itself whenever you run the pipeline.

1. Create a **free, public** repo on GitHub, e.g. `sas-dashboard`.
2. In this folder run:

   ```powershell
   git init
   git add .
   git commit -m "init"
   git branch -M main
   git remote add origin https://github.com/<you>/sas-dashboard.git
   git push -u origin main
   ```

3. On GitHub: **Settings → Pages → Source: Deploy from branch → `main` / root → Save**.
4. Your permanent link is:

   ```
   https://<you>.github.io/sas-dashboard/
   ```

   Share this once. It never changes.

> The raw files (`input.csv`, `main.xlsx`) are git-ignored, so only the
> dashboard is published. Note that the dashboard's **aggregated** numbers are
> embedded in the public page — anyone with the link can see them. If that data
> is sensitive, use email/Drive sharing (below) instead of a public link.

## Every day: download a fresh `input.csv`, then run

```powershell
python pipeline.py
```

This does exactly what you did by hand, then updates the live link:

1. Reads `main.xlsx` (creates it the first time from `input.csv`).
2. Deletes from `main.xlsx` the dates present in `input.csv` (so the 3–5 day
   overlap never duplicates), then appends the input rows and saves.
3. Rebuilds `index.html` from `main.xlsx`.
4. Commits + pushes it — the GitHub Pages link refreshes in about a minute.

Run steps individually if needed:

```powershell
python pipeline.py --update       # merge only
python pipeline.py --dashboard    # rebuild dashboard only
python pipeline.py --deploy       # publish only
python pipeline.py --no-deploy    # update + rebuild, but don't publish
```

## The dashboard (replaces Power BI)

`index.html` is a **single self-contained file**: KPIs (Grand Total, Console,
Positive, Positive %), a daily trend line, and breakdowns by Load Type /
Cluster DMH Zone / FWD-RTN / Asset, with date-range and per-dimension filters.

Prefer not to use a public link? Just email `index.html` or drop it in Google
Drive / SharePoint — it opens in any browser (needs internet for the chart
library).

## Notes

- Column A must be `Date` (day-first `DD-MM-YYYY` is handled).
  Numeric metrics: `Console`, `Positive`, `Grand Total`.
- The dashboard aggregates to date + key dimensions (dropping the very
  high-cardinality `SMH-DMH Lanes` / `bag_src_hub_name`) so the file stays small
  and fast even after months of daily data. Adjust `DASHBOARD_DIMS` in
  `pipeline.py` if you need different filters.
