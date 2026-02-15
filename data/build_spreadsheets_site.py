#!/usr/bin/env python3
"""
Generate static HTML pages for .xlsx workbooks (no JS, no CDN).
- Creates: sheets/<workbook>.html for each workbook
- Creates: spreadsheets.html index listing all workbooks

Usage:
  python build_spreadsheets_site.py
  python build_spreadsheets_site.py --xlsx-dir xlsx --out-dir . --sheets-dir sheets

Notes:
- Requires: pandas, openpyxl
    pip install pandas openpyxl
- This renders values into HTML tables. It does not preserve Excel formatting.
"""

import argparse
import html
from pathlib import Path

import pandas as pd

# Desired ordering (edit as you like). Files not found are skipped.
ORDERED_XLSX = [
  "actualexpectedgrades.xlsx",
  "altinstrument.xlsx",
  "basedata.xlsx",
  "bootstrap_params.xlsx",
  "bpref.xlsx",
  "bstartbundled.xlsx",
  "bundleddata.xlsx",
  "bundledse.xlsx",
  "bundled_structural_param.xlsx",
  "choicesetevolution.xlsx",
  "effortparams.xlsx",
  "effortraw.xlsx",
  "effort_data.xlsx",
  "enrl_order.xlsx",
  "estimates_joint_threetype_cap.xlsx",
  "estimates_joint_threetype_start.xlsx",
  "estimates_joint_threetype_start_wfp.xlsx",
  "estimates_joint_threetype_wfp.xlsx",
  "evaluation_class.xlsx",
  "GEgrade3results.xlsx",
  "GEgrade3results_noeff.xlsx",
  "GEgrade3results_noeffld.xlsx",
  "GEgrade3results_nu2.xlsx",
  "GEgrade3start_noeff.xlsx",
  "GEgrade3start_noeffld.xlsx",
  "GEtables_effort.xlsx",
  "GEtables_noeff.xlsx",
  "GE_counterfacts_data_sort.xlsx",
  "gridderiv.xlsx",
  "gridderivact.xlsx",
  "instructorrank.xlsx",
  "multisemevaluation.xlsx",
  "numbers.xlsx",
  "PENoGrades.xlsx",
  "PEresults.xlsx",
  "PE_counterfacts_data.xlsx",
  "profdecomp.xlsx",
  "profprefs.xlsx",
  "profreduced.xlsx",
  "profresults.xlsx",
  "robustresults.xlsx",
  "room_capacity.xlsx",
  "student_data.xlsx",
  "super_list.xlsx",
  "t12.xlsx",
  "t4.xlsx",
  "tables_1to4.xlsx",
  "tables_6to14.xlsx",
  "tables_appB.xlsx",
  "tables_appD.xlsx"
]

PAGE_CSS = """
:root { color-scheme: light dark; }
body {
  font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
  margin: 24px;
  line-height: 1.4;
}
a { text-decoration: none; }
a:hover { text-decoration: underline; }
.header {
  display: flex;
  gap: 14px;
  align-items: baseline;
  flex-wrap: wrap;
  margin-bottom: 16px;
}
.header h1 {
  font-size: 22px;
  margin: 0;
}
.meta {
  color: #666;
  font-size: 13px;
}
.sheet-nav {
  margin: 10px 0 18px;
  padding: 10px 12px;
  border: 1px solid rgba(125,125,125,.35);
  border-radius: 10px;
}
.sheet-nav a {
  margin-right: 10px;
  white-space: nowrap;
}
table {
  border-collapse: collapse;
  margin: 10px 0 24px;
  width: max-content;
  max-width: 100%;
}
td, th {
  border: 1px solid rgba(125,125,125,.35);
  padding: 4px 8px;
  vertical-align: top;
  font-size: 13px;
}
th {
  position: sticky;
  top: 0;
  background: rgba(200,200,200,.25);
  backdrop-filter: blur(6px);
}
.section {
  margin-top: 26px;
}
.section h2 {
  font-size: 18px;
  margin: 0 0 6px;
}
.small {
  font-size: 12px;
  color: #666;
}
.wrap {
  overflow-x: auto;
  border: 1px solid rgba(125,125,125,.25);
  border-radius: 12px;
  padding: 10px;
}
"""

def sanitize_filename(name: str) -> str:
  stem = Path(name).stem
  safe = "".join(ch if (ch.isalnum() or ch in "-_") else "_" for ch in stem)
  return safe + ".html"

def df_to_html_table(df: pd.DataFrame) -> str:
  df2 = df.copy()
  df2 = df2.where(pd.notnull(df2), "")
  df2 = df2.astype(str)
  # to_html() escapes by default, so no applymap() needed
  return df2.to_html(index=False, header=False, border=0)


def render_workbook(xlsx_path: Path, out_path: Path, back_href: str) -> None:
  xls = pd.ExcelFile(xlsx_path, engine="openpyxl")
  sheet_names = xls.sheet_names

  nav_links = []
  for s in sheet_names:
    anchor = "sheet-" + "".join(ch if ch.isalnum() else "-" for ch in s).strip("-").lower()
    nav_links.append(f'<a href="#{anchor}">{html.escape(s)}</a>')
  nav_html = '<div class="sheet-nav"><div class="small">Sheets:</div>' + "".join(nav_links) + "</div>"

  parts = []
  parts.append(f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{html.escape(xlsx_path.name)}</title>
  <style>{PAGE_CSS}</style>
</head>
<body>
  <div class="header">
    <h1>{html.escape(xlsx_path.name)}</h1>
    <div class="meta"><a href="{back_href}">← Back to list</a></div>
  </div>
  {nav_html}
""")

  for s in sheet_names:
    anchor = "sheet-" + "".join(ch if ch.isalnum() else "-" for ch in s).strip("-").lower()
    df = pd.read_excel(xlsx_path, sheet_name=s, header=None, engine="openpyxl")
    rows, cols = df.shape
    table_html = df_to_html_table(df)
    parts.append(f"""
  <div class="section" id="{anchor}">
    <h2>{html.escape(s)}</h2>
    <div class="small">{rows} rows × {cols} columns</div>
    <div class="wrap">{table_html}</div>
  </div>
""")

  parts.append("""
</body>
</html>
""")

  out_path.write_text("".join(parts), encoding="utf-8")

def render_index(xlsx_files, out_path: Path, sheets_dir: str) -> None:
  items = []
  for f in xlsx_files:
    page = sanitize_filename(f.name)
    items.append(f'<li><a href="{sheets_dir}/{page}" target="_blank" rel="noopener">{html.escape(f.name)}</a></li>')
  html_doc = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Spreadsheets</title>
  <style>
    body {{ font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; margin: 24px; }}
    h1 {{ margin: 0 0 10px; font-size: 22px; }}
    .note {{ color: #666; font-size: 13px; margin-bottom: 14px; }}
    ul {{ line-height: 1.8; }}
    a {{ text-decoration: none; }}
    a:hover {{ text-decoration: underline; }}
  </style>
</head>
<body>
  <h1>Spreadsheets</h1>
  <div class="note">Click a filename to open its static HTML view in a new tab.</div>
  <ul>
    {''.join(items)}
  </ul>
</body>
</html>
"""
  out_path.write_text(html_doc, encoding="utf-8")

def main():
  ap = argparse.ArgumentParser()
  ap.add_argument("--xlsx-dir", default=".", help="Directory containing .xlsx files")
  ap.add_argument("--out-dir", default=".", help="Where to write spreadsheets.html")
  ap.add_argument("--sheets-dir", default="sheets", help="Where to write sheet pages (relative to out-dir)")
  args = ap.parse_args()

  xlsx_dir = Path(args.xlsx_dir)
  out_dir = Path(args.out_dir)
  sheets_dir = out_dir / args.sheets_dir
  sheets_dir.mkdir(parents=True, exist_ok=True)

  found = []
  for name in ORDERED_XLSX:
    p = xlsx_dir / name
    if p.exists():
      found.append(p)
    else:
      print(f"[skip] not found: {p}")

  if not found:
    found = sorted(xlsx_dir.glob("*.xlsx"))

  if not found:
    raise SystemExit(f"No .xlsx files found in {xlsx_dir}")

  for p in found:
    out_page = sheets_dir / sanitize_filename(p.name)
    render_workbook(p, out_page, back_href="../spreadsheets.html")
    print(f"[ok] wrote {out_page}")

  render_index(found, out_dir / "spreadsheets.html", sheets_dir=args.sheets_dir)
  print(f"[ok] wrote {out_dir / 'spreadsheets.html'}")

if __name__ == "__main__":
  main()
