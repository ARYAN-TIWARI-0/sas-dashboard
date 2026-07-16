"""
SAS_DB pipeline
===============

Replaces the manual Excel + Power BI workflow with three steps:

  1. update    -> merge input.csv into main.xlsx
                  (drops any dates from main that are present in input, then
                   appends the input rows, so re-downloads never duplicate)
  2. dashboard -> build a single self-contained index.html from main.xlsx
                  (interactive, works offline)
  3. deploy    -> commit + push index.html so a fixed GitHub Pages URL updates
                  itself (one-time repo setup required; see README)

Usage
-----
    python pipeline.py                # update + rebuild + publish (daily run)
    python pipeline.py --update       # only merge input.csv into main.xlsx
    python pipeline.py --dashboard    # only rebuild index.html
    python pipeline.py --deploy       # only publish to GitHub Pages
    python pipeline.py --no-deploy    # update + rebuild, skip publishing

Defaults assume input.csv / main.xlsx / index.html live next to this file.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

# --- configuration -----------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent

DATE_COL = "Date"                         # column A
METRICS = ["Console", "Positive", "Grand Total"]

# Computed classification column (was an Excel formula in main.xlsx):
#   =IFERROR(IF(F=H,"Local",IF(G=I,"Zonal","National")),"Local")
# F=bag_src_hub_name  H=Cluster DMH  G=bag_src_hub_zone  I=Cluster DMH Zone
CLASS_COL = "Column13"
CLASS_SRC = {
    "name": "bag_src_hub_name",   # F
    "dmh": "Cluster DMH",         # H
    "src_zone": "bag_src_hub_zone",   # G
    "dmh_zone": "Cluster DMH Zone",   # I
}

# Dimensions kept for the dashboard. The two highest-cardinality raw columns
# ("SMH-DMH Lanes", "bag_src_hub_name") are collapsed by aggregation so the
# embedded dataset stays small and the page stays fast, even after months of
# daily appends.
DASHBOARD_DIMS = [
    "Load Type",
    "Initials",
    "Asset",
    "bag_src_hub_zone",
    "Cluster DMH",
    "Cluster DMH Zone",
    "FWD/RTN",
    CLASS_COL,
]


# --- step 1: merge input.csv into main.xlsx ----------------------------------

def add_classification(df: pd.DataFrame) -> pd.DataFrame:
    """Populate the Column13 Local/Zonal/National scope as real values.
    Replaces the live Excel formula so appended rows are classified too.
    Text comparison is case-insensitive, matching Excel's '=' behaviour."""
    needed = list(CLASS_SRC.values())
    missing = [c for c in needed if c not in df.columns]
    if missing:
        print(f"[update] warning: cannot compute {CLASS_COL}, missing columns: {missing}")
        return df

    def norm(col):
        return df[col].fillna("").astype(str).str.strip().str.casefold()

    name, dmh = norm(CLASS_SRC["name"]), norm(CLASS_SRC["dmh"])
    szone, dzone = norm(CLASS_SRC["src_zone"]), norm(CLASS_SRC["dmh_zone"])
    df[CLASS_COL] = np.where(name == dmh, "Local",
                             np.where(szone == dzone, "Zonal", "National"))
    return df


def _normalize_dates(series: pd.Series) -> pd.Series:
    """Coerce a column to plain calendar dates so main/input compare cleanly.
    Input uses day-first (DD-MM-YYYY); Excel may store datetimes. dayfirst=True
    is safe for both."""
    return pd.to_datetime(series, errors="coerce", dayfirst=True).dt.normalize()


def update_main(main_path: Path, input_path: Path) -> pd.DataFrame:
    if not input_path.exists():
        sys.exit(f"[update] input file not found: {input_path}")

    new_df = pd.read_csv(input_path)
    if DATE_COL not in new_df.columns:
        sys.exit(f"[update] '{DATE_COL}' column (col A) missing from {input_path.name}")

    new_df[DATE_COL] = _normalize_dates(new_df[DATE_COL])
    new_dates = set(new_df[DATE_COL].dropna().unique())

    if main_path.exists():
        main_df = pd.read_excel(main_path)
        main_df[DATE_COL] = _normalize_dates(main_df[DATE_COL])
        # drop rows for dates that the input refreshes, then append
        kept = main_df[~main_df[DATE_COL].isin(new_dates)]
        removed = len(main_df) - len(kept)
        combined = pd.concat([kept, new_df], ignore_index=True)
        print(f"[update] main had {len(main_df)} rows; "
              f"removed {removed} rows for {len(new_dates)} refreshed date(s)")
    else:
        combined = new_df.copy()
        print(f"[update] {main_path.name} not found - creating it from {input_path.name}")

    # drop blank/formula-only rows that have no date
    combined = combined[combined[DATE_COL].notna()]
    combined = combined.sort_values(DATE_COL).reset_index(drop=True)

    # keep numeric metrics numeric
    for col in METRICS:
        if col in combined.columns:
            combined[col] = pd.to_numeric(combined[col], errors="coerce").fillna(0)

    # (re)compute the classification for all rows and keep it as the last column
    combined = add_classification(combined)
    if CLASS_COL in combined.columns:
        ordered = [c for c in combined.columns if c != CLASS_COL] + [CLASS_COL]
        combined = combined[ordered]

    combined.to_excel(main_path, index=False)
    added = ", ".join(sorted(d.strftime("%Y-%m-%d") for d in new_dates))
    print(f"[update] appended {len(new_df)} rows for date(s): {added}")
    print(f"[update] saved {len(combined)} total rows -> {main_path}")
    return combined


# --- step 2: build dashboard.html --------------------------------------------

def build_dashboard(main_path: Path, out_path: Path, df: pd.DataFrame | None = None) -> None:
    if df is None:
        if not main_path.exists():
            sys.exit(f"[dashboard] {main_path} not found - run the update step first")
        df = pd.read_excel(main_path)

    df = df.copy()
    df[DATE_COL] = _normalize_dates(df[DATE_COL])
    for col in METRICS:
        df[col] = pd.to_numeric(df.get(col, 0), errors="coerce").fillna(0)

    dims = [d for d in DASHBOARD_DIMS if d in df.columns]
    for d in dims:
        df[d] = df[d].fillna("(blank)").astype(str)

    df["_date"] = df[DATE_COL].dt.strftime("%Y-%m-%d")

    grouped = (
        df.groupby(["_date"] + dims, dropna=False)[METRICS]
        .sum()
        .reset_index()
    )

    records = [
        {
            "date": r["_date"],
            "dims": {d: r[d] for d in dims},
            "metrics": {m: float(r[m]) for m in METRICS},
        }
        for _, r in grouped.iterrows()
    ]

    payload = {
        "records": records,
        "dims": dims,
        "metrics": METRICS,
        "dates": sorted(df["_date"].dropna().unique().tolist()),
    }

    html = _HTML_TEMPLATE.replace(
        "/*__DATA__*/",
        json.dumps(payload, separators=(",", ":")),
    )
    out_path.write_text(html, encoding="utf-8")
    print(f"[dashboard] {len(records)} aggregated rows across "
          f"{len(payload['dates'])} day(s) -> {out_path}")
    print(f"[dashboard] open it in a browser, email the file, or host it free "
          f"(GitHub Pages / Netlify) to share a link")


# --- dashboard template ------------------------------------------------------
# Single file, Plotly from CDN. Data is embedded so the file is fully portable.

_HTML_TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>SAS Dashboard</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js" charset="utf-8"></script>
<style>
  :root{--bg:#0f172a;--card:#1e293b;--muted:#94a3b8;--txt:#e2e8f0;--accent:#38bdf8;--line:#334155;--panel:#0b1220;}
  *{box-sizing:border-box;}
  body{margin:0;font-family:Inter,Segoe UI,system-ui,Arial,sans-serif;background:var(--bg);color:var(--txt);}
  header{padding:16px 24px;border-bottom:1px solid var(--line);display:flex;justify-content:space-between;align-items:baseline;flex-wrap:wrap;gap:8px;}
  header h1{font-size:20px;margin:0;}
  header .sub{color:var(--muted);font-size:13px;}
  .layout{display:grid;grid-template-columns:280px 1fr;gap:16px;padding:16px 24px;align-items:start;}
  .filters{background:var(--card);border-radius:12px;padding:14px;position:sticky;top:16px;max-height:calc(100vh - 32px);overflow:auto;}
  .filters h2{font-size:12px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);margin:0 0 10px;}
  .fld{font-size:11px;color:var(--muted);margin:10px 0 4px;}
  .filters input[type=date]{width:100%;background:var(--panel);color:var(--txt);border:1px solid var(--line);border-radius:8px;padding:6px 8px;font-size:13px;}
  .quick{display:flex;gap:6px;margin-top:8px;flex-wrap:wrap;}
  .quick button{flex:1;min-width:56px;background:var(--panel);color:var(--txt);border:1px solid var(--line);border-radius:8px;padding:5px;font-size:11px;cursor:pointer;}
  .quick button:hover{border-color:var(--accent);}
  .slicer{margin-top:8px;border:1px solid var(--line);border-radius:8px;overflow:hidden;}
  .slicer-btn{width:100%;display:flex;justify-content:space-between;align-items:center;background:var(--panel);color:var(--txt);border:0;padding:8px 10px;font-size:12px;cursor:pointer;text-align:left;}
  .slicer-btn b{font-weight:600;}
  .slicer-btn .cap{color:var(--accent);font-size:11px;margin-left:8px;white-space:nowrap;}
  .slicer-panel{display:none;padding:8px;background:var(--card);border-top:1px solid var(--line);}
  .slicer-panel.open{display:block;}
  .slicer-search{width:100%;background:var(--panel);color:var(--txt);border:1px solid var(--line);border-radius:6px;padding:5px 7px;font-size:12px;margin-bottom:6px;}
  .slicer-acts{display:flex;gap:10px;margin-bottom:6px;}
  .slicer-acts a{color:var(--accent);font-size:11px;cursor:pointer;}
  .slicer-opts{max-height:180px;overflow:auto;}
  .opt{display:flex;align-items:center;gap:7px;padding:3px 2px;font-size:12px;cursor:pointer;}
  .opt input{accent-color:var(--accent);}
  .reset{width:100%;margin-top:14px;background:var(--accent);color:#04223a;border:0;border-radius:8px;padding:9px;font-weight:600;cursor:pointer;}
  .kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:16px;}
  .kpi{background:var(--card);border-radius:12px;padding:14px 16px;}
  .kpi .v{font-size:24px;font-weight:700;}
  .kpi .l{font-size:12px;color:var(--muted);margin-top:2px;}
  .toolbar{display:flex;align-items:center;gap:8px;margin-bottom:12px;font-size:12px;color:var(--muted);}
  .toolbar select{background:var(--panel);color:var(--txt);border:1px solid var(--line);border-radius:8px;padding:5px 8px;font-size:12px;}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;}
  .chart{background:var(--card);border-radius:12px;padding:8px;min-height:320px;}
  .chart.wide{grid-column:1 / -1;}
  @media(max-width:900px){.layout{grid-template-columns:1fr;}.grid{grid-template-columns:1fr;}.kpis{grid-template-columns:repeat(2,1fr);}}
</style>
</head>
<body>
<header>
  <div><h1>SAS Operations Dashboard</h1><div class="sub" id="daterange"></div></div>
  <div class="sub">Generated from main.xlsx</div>
</header>
<div class="layout">
  <aside class="filters">
    <h2>Filters</h2>
    <div class="fld">From date</div><input type="date" id="fFrom"/>
    <div class="fld">To date</div><input type="date" id="fTo"/>
    <div class="quick">
      <button data-days="7">7D</button>
      <button data-days="30">30D</button>
      <button data-days="0">All</button>
    </div>
    <div id="dimFilters"></div>
    <button class="reset" id="reset">Reset all filters</button>
  </aside>
  <main>
    <div class="kpis" id="kpis"></div>
    <div class="toolbar">
      <span>Measure for charts:</span>
      <select id="measure"></select>
    </div>
    <div class="grid">
      <div class="chart wide" id="cTrend"></div>
      <div class="chart" id="cSource"></div>
      <div class="chart" id="cZone"></div>
      <div class="chart" id="cDirection"></div>
      <div class="chart" id="cScope"></div>
      <div class="chart wide" id="cAsset"></div>
    </div>
  </main>
</div>
<script>
const PAYLOAD = /*__DATA__*/;
const {records, dims, metrics, dates} = PAYLOAD;
const PRIMARY = metrics[metrics.length-1]; // "Grand Total"
let measure = PRIMARY;
const PLOT_LAYOUT = {paper_bgcolor:"transparent",plot_bgcolor:"transparent",
  font:{color:"#e2e8f0",size:12},margin:{t:40,r:16,b:60,l:70},
  legend:{orientation:"h",y:-0.2},colorway:["#38bdf8","#f472b6","#a3e635","#fbbf24","#c084fc"]};
const CONF = {displayModeBar:false,responsive:true};

const uniq = k => [...new Set(records.map(r=>r.dims[k]))].sort();
const fmt = n => n.toLocaleString(undefined,{maximumFractionDigits:0});

// selections[dim] = Set of chosen values; empty/absent = all
const selections = {};

// ---- build slicers (clickable checkbox dropdowns) ----
const dimBox = document.getElementById("dimFilters");
dims.forEach(dim=>{
  const values = uniq(dim);
  const wrap=document.createElement("div"); wrap.className="slicer";
  const btn=document.createElement("button"); btn.className="slicer-btn";
  btn.innerHTML=`<b>${dim}</b><span class="cap" data-cap>All</span>`;
  const panel=document.createElement("div"); panel.className="slicer-panel";
  const search=document.createElement("input"); search.className="slicer-search"; search.placeholder="Search...";
  const acts=document.createElement("div"); acts.className="slicer-acts";
  acts.innerHTML=`<a data-act="all">Select all</a><a data-act="none">Clear</a>`;
  const opts=document.createElement("div"); opts.className="slicer-opts";
  values.forEach(v=>{
    const row=document.createElement("label"); row.className="opt";
    const cb=document.createElement("input"); cb.type="checkbox"; cb.value=v;
    const span=document.createElement("span"); span.textContent=v;
    row.appendChild(cb); row.appendChild(span); opts.appendChild(row);
    cb.addEventListener("change",()=>{
      selections[dim]=new Set([...opts.querySelectorAll("input:checked")].map(c=>c.value));
      updateCap(btn,dim,values.length); render();
    });
  });
  search.addEventListener("input",()=>{
    const q=search.value.toLowerCase();
    opts.querySelectorAll(".opt").forEach(o=>{
      o.style.display=o.textContent.toLowerCase().includes(q)?"flex":"none";
    });
  });
  acts.querySelector('[data-act="all"]').addEventListener("click",()=>{
    opts.querySelectorAll("input").forEach(c=>c.checked=true);
    selections[dim]=new Set(values); updateCap(btn,dim,values.length); render();
  });
  acts.querySelector('[data-act="none"]').addEventListener("click",()=>{
    opts.querySelectorAll("input").forEach(c=>c.checked=false);
    selections[dim]=new Set(); updateCap(btn,dim,values.length); render();
  });
  btn.addEventListener("click",()=>panel.classList.toggle("open"));
  panel.appendChild(search); panel.appendChild(acts); panel.appendChild(opts);
  wrap.appendChild(btn); wrap.appendChild(panel); dimBox.appendChild(wrap);
});
function updateCap(btn,dim,total){
  const s=selections[dim];
  const cap=btn.querySelector("[data-cap]");
  cap.textContent=(!s||s.size===0||s.size===total)?"All":`${s.size} selected`;
}

// ---- date filters ----
const fFrom=document.getElementById("fFrom"), fTo=document.getElementById("fTo");
fFrom.min=fTo.min=dates[0]; fFrom.max=fTo.max=dates[dates.length-1];
fFrom.value=dates[0]; fTo.value=dates[dates.length-1];
fFrom.addEventListener("change",render); fTo.addEventListener("change",render);
document.querySelectorAll(".quick button").forEach(b=>b.addEventListener("click",()=>{
  const days=+b.dataset.days, last=dates[dates.length-1];
  if(days===0){fFrom.value=dates[0];}
  else{const d=new Date(last);d.setDate(d.getDate()-(days-1));fFrom.value=d.toISOString().slice(0,10);}
  fTo.value=last; render();
}));

// ---- measure selector ----
const measureSel=document.getElementById("measure");
metrics.forEach(m=>{const o=document.createElement("option");o.value=m;o.textContent=m;measureSel.appendChild(o);});
measureSel.value=measure;
measureSel.addEventListener("change",()=>{measure=measureSel.value;render();});

document.getElementById("reset").addEventListener("click",()=>{
  fFrom.value=dates[0]; fTo.value=dates[dates.length-1];
  dims.forEach(d=>selections[d]=new Set());
  dimBox.querySelectorAll("input[type=checkbox]").forEach(c=>c.checked=false);
  dimBox.querySelectorAll(".slicer-btn").forEach(b=>b.querySelector("[data-cap]").textContent="All");
  render();
});
document.getElementById("daterange").textContent = dates.length
  ? `${dates[0]} to ${dates[dates.length-1]} - ${dates.length} day(s)` : "no data";

// ---- filtering + charts ----
function filtered(){
  const from=fFrom.value, to=fTo.value;
  return records.filter(r=>{
    if(r.date<from||r.date>to) return false;
    for(const d in selections){const s=selections[d]; if(s&&s.size&&!s.has(r.dims[d])) return false;}
    return true;
  });
}
function sumBy(rows, keyFn, metric){
  const m={}; rows.forEach(r=>{const k=keyFn(r); m[k]=(m[k]||0)+r.metrics[metric];}); return m;
}
function topBar(el, rows, dim, title, horizontal){
  const m=sumBy(rows,r=>r.dims[dim],measure);
  let entries=Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,12);
  const labels=entries.map(e=>e[0]), vals=entries.map(e=>e[1]);
  const trace = horizontal
    ? {type:"bar",orientation:"h",y:labels.slice().reverse(),x:vals.slice().reverse(),marker:{color:"#38bdf8"}}
    : {type:"bar",x:labels,y:vals,marker:{color:"#38bdf8"}};
  Plotly.react(el,[trace],{...PLOT_LAYOUT,title},CONF);
}
function render(){
  const rows=filtered();
  const totals={}; metrics.forEach(mt=>totals[mt]=rows.reduce((s,r)=>s+r.metrics[mt],0));
  const posRate = totals[PRIMARY] ? (100*totals["Positive"]/totals[PRIMARY]) : 0;
  const kpis=[
    ["Grand Total",fmt(totals["Grand Total"])],
    ["Console",fmt(totals["Console"])],
    ["Positive",fmt(totals["Positive"])],
    ["Positive %",posRate.toFixed(2)+"%"],
  ];
  document.getElementById("kpis").innerHTML = kpis.map(k=>
    `<div class="kpi"><div class="v">${k[1]}</div><div class="l">${k[0]}</div></div>`).join("");
  const byDate=metrics.map(mt=>{
    const m=sumBy(rows,r=>r.date,mt);
    const xs=Object.keys(m).sort();
    return {type:"scatter",mode:"lines+markers",name:mt,x:xs,y:xs.map(d=>m[d])};
  });
  Plotly.react("cTrend",byDate,{...PLOT_LAYOUT,title:"Daily volume trend (all measures)"},CONF);
  if(dims.includes("Load Type")) topBar("cSource",rows,"Load Type","By Load Type",false);
  if(dims.includes("Cluster DMH Zone")) topBar("cZone",rows,"Cluster DMH Zone","By Cluster DMH Zone",false);
  if(dims.includes("FWD/RTN")) topBar("cDirection",rows,"FWD/RTN","By FWD / RTN",false);
  if(dims.includes("Column13")) topBar("cScope",rows,"Column13","By Scope (Local / Zonal / National)",false);
  if(dims.includes("Asset")) topBar("cAsset",rows,"Asset","By Asset (top 12)",true);
}
render();
</script>
</body>
</html>
"""


# --- step 3: publish to the permanent GitHub Pages link ----------------------

def _git(*args: str, cwd: Path) -> subprocess.CompletedProcess:
    return subprocess.run(["git", *args], cwd=cwd, capture_output=True, text=True)


def deploy_dashboard(out_path: Path) -> None:
    """Commit + push the dashboard so the fixed GitHub Pages URL updates itself.

    One-time setup is required (see README): create a public GitHub repo, run
    git init / remote add / push, and enable Pages. After that, this pushes the
    latest dashboard on every run.
    """
    repo = out_path.resolve().parent

    if _git("rev-parse", "--is-inside-work-tree", cwd=repo).returncode != 0:
        print("[deploy] skipped: this folder is not a git repo yet.")
        print("[deploy] one-time setup (see README) then re-run:")
        print("         git init && git add . && git commit -m init")
        print("         git remote add origin <your-repo-url> && git push -u origin main")
        return

    if not _git("remote", "get-url", "origin", cwd=repo).stdout.strip():
        print("[deploy] skipped: no 'origin' remote. Add one:")
        print("         git remote add origin <your-repo-url>")
        return

    _git("add", out_path.name, cwd=repo)
    if not _git("diff", "--cached", "--quiet", cwd=repo).returncode:
        print("[deploy] no dashboard changes to publish.")
        return

    msg = f"dashboard update {datetime.now():%Y-%m-%d %H:%M}"
    if _git("commit", "-m", msg, cwd=repo).returncode != 0:
        print("[deploy] commit failed - check `git status`.")
        return

    push = _git("push", cwd=repo)
    if push.returncode != 0:
        print(f"[deploy] push failed:\n{push.stderr.strip()}")
        return

    url = _git("remote", "get-url", "origin", cwd=repo).stdout.strip()
    print(f"[deploy] published. Live link updates in ~1 min "
          f"(GitHub Pages URL for {url}).")


# --- cli ---------------------------------------------------------------------

def main() -> None:
    ap = argparse.ArgumentParser(description="Update main.xlsx and build a shareable dashboard.")
    ap.add_argument("--input", default=str(BASE_DIR / "input.csv"))
    ap.add_argument("--main", default=str(BASE_DIR / "main.xlsx"))
    ap.add_argument("--out", default=str(BASE_DIR / "index.html"),
                    help="dashboard file (index.html gives a clean Pages URL)")
    ap.add_argument("--update", action="store_true", help="only merge input.csv into main.xlsx")
    ap.add_argument("--dashboard", action="store_true", help="only rebuild the dashboard")
    ap.add_argument("--deploy", action="store_true", help="only push the dashboard to GitHub Pages")
    ap.add_argument("--no-deploy", action="store_true", help="skip publishing on a full run")
    args = ap.parse_args()

    main_path, input_path, out_path = Path(args.main), Path(args.input), Path(args.out)
    run_all = not (args.update or args.dashboard or args.deploy)

    df = None
    if args.update or run_all:
        df = update_main(main_path, input_path)
    if args.dashboard or run_all:
        build_dashboard(main_path, out_path, df)
    if args.deploy or (run_all and not args.no_deploy):
        deploy_dashboard(out_path)


if __name__ == "__main__":
    main()
