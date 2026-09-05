"""
build_data.py — Prepares dashboard_data.json for the Corporate CSR Monitor
dashboard artifact from whichever category output files currently exist.

Run this any time a category's output .xlsx changes, then re-publish the
artifact from dashboard.html with the freshly written dashboard_data.json
injected. See build_html.py for the injection step.

Categories and their expected output files (edit these paths as real runs
land — a category with no file yet, or a header-only file, renders as a
"no runs completed yet" empty state rather than being skipped):
"""

import json
import os
import re
from collections import Counter

import openpyxl

# ── Configure the current output file for each category here ──────────────
# All three read from one shared folder — drop/overwrite the matching
# filename there after a run and re-run this script. Nothing else to edit.
OUTPUT_DIR = r"C:\Users\Swati Shukla\Downloads\Corporate_CSR_Monitor\outputs"

SOURCES = {
    "companies": {
        "path": os.path.join(OUTPUT_DIR, "Companies_output.xlsx"),
        "sourceList": "CSR_Companies_2024-25.csv",
        "label": "CSR Companies",
    },
    "rnd": {
        "path": os.path.join(OUTPUT_DIR, "RnD_Companies_output.xlsx"),
        "sourceList": "RnD_Companies_2022-23.csv",
        "label": "R&D Companies",
    },
    "institutes": {
        "path": os.path.join(OUTPUT_DIR, "Institute_CSR_Donations.xlsx"),
        "sourceList": "Institutes (IIT / IISc / IISER)",
        "label": "Institutes",
    },
}

OUT_PATH = os.path.join(os.path.dirname(__file__), "dashboard_data.json")

_LEGAL_SUFFIX_RE = re.compile(
    r'\b(ltd\.?|limited|pvt\.?|private|inc\.?|incorporated|corp\.?|corporation|'
    r'plc|llp|group|industries|holdings|company|co\.?)\b', re.IGNORECASE)

_MONTHS = {m.lower(): i + 1 for i, m in enumerate(
    ["January", "February", "March", "April", "May", "June", "July", "August",
     "September", "October", "November", "December"])}


def _norm(s):
    return re.sub(r'[^a-z0-9]', '', (s or '').lower())


def _company_root(company):
    s = _LEGAL_SUFFIX_RE.sub('', company or '')
    return _norm(s) or _norm(company)


def _parse_date(s):
    s = (s or "").strip()
    m = re.match(r'(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})', s)
    if m:
        mo = _MONTHS.get(m.group(2).lower())
        if mo:
            try:
                return f"{int(m.group(3)):04d}-{mo:02d}-{int(m.group(1)):02d}"
            except Exception:
                pass
    m = re.match(r'([A-Za-z]+)\s+(\d{4})', s)
    if m:
        mo = _MONTHS.get(m.group(1).lower())
        if mo:
            return f"{int(m.group(2)):04d}-{mo:02d}-01"
    m = re.match(r'(\d{4})-(\d{2})-(\d{2})', s)
    if m:
        return s[:10]
    m = re.match(r'(\d{4})$', s)
    if m:
        return f"{s}-07-01"
    return None


def _is_institute_mode(headers):
    return "Institute Name" in headers and "Donor Company" in headers


def load_category(key, cfg):
    path = cfg["path"]
    if not path or not os.path.exists(path):
        return empty_result(cfg, status="pending")

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    try:
        headers = [str(h).strip() if h is not None else "" for h in next(rows_iter)]
    except StopIteration:
        return empty_result(cfg, status="pending")

    institute_mode = _is_institute_mode(headers)
    idx = {h: i for i, h in enumerate(headers)}

    def cell(row, name, default=""):
        i = idx.get(name)
        if i is None or i >= len(row):
            return default
        v = row[i]
        return (str(v).strip() if v is not None else default)

    out_rows = []
    n = 0
    for row in rows_iter:
        if row is None or all(v is None for v in row):
            continue
        n += 1
        if institute_mode:
            subject = cell(row, "Donor Company") or cell(row, "Institute Name")
            secondary = cell(row, "Institute Name")
        else:
            subject = cell(row, "Company Name")
            secondary = ""
        date_raw = cell(row, "Date")
        out_rows.append({
            "id": f"r{n:04d}",
            "company": subject,
            "companyGroup": _company_root(subject),
            "secondary": secondary,
            "category": cell(row, "CSR Category"),
            "theme": cell(row, "Theme"),
            "causeFocus": cell(row, "Cause Focus (This Year)"),
            "initiative": cell(row, "Initiative / Program"),
            "amount": cell(row, "Amount"),
            "partner": cell(row, "Partner Organization"),
            "headline": cell(row, "News Headline"),
            "update": cell(row, "Detailed Update"),
            "source": cell(row, "Source"),
            "date": date_raw,
            "dateSort": _parse_date(date_raw),
            "searchPeriod": cell(row, "Search Period"),
            "sentiment": cell(row, "Sentiment") or "Positive",
            "searchUrl": cell(row, "Searchable Link"),
        })

    if not out_rows:
        return empty_result(cfg, status="pending")

    cat_counts = Counter(r["category"] for r in out_rows if r["category"])
    groups = {}
    for r in out_rows:
        g = r["companyGroup"]
        if not g:
            continue
        groups.setdefault(g, {"name": r["company"], "count": 0})
        groups[g]["count"] += 1
    top_companies = sorted(groups.values(), key=lambda x: -x["count"])[:12]
    dates = sorted(d for d in (r["dateSort"] for r in out_rows) if d)

    summary = {
        "status": "ready",
        "label": cfg["label"],
        "sourceList": cfg["sourceList"],
        "totalUpdates": len(out_rows),
        "nSubjects": len(groups),
        "latestDate": dates[-1] if dates else None,
        "earliestDate": dates[0] if dates else None,
        "categoryCounts": dict(cat_counts),
        "topCompanies": top_companies,
        "instituteMode": institute_mode,
    }
    return {"summary": summary, "rows": out_rows}


def empty_result(cfg, status="pending"):
    return {
        "summary": {
            "status": status,
            "label": cfg["label"],
            "sourceList": cfg["sourceList"],
            "totalUpdates": 0,
            "nSubjects": 0,
            "latestDate": None,
            "earliestDate": None,
            "categoryCounts": {},
            "topCompanies": [],
            "instituteMode": cfg.get("instituteMode", False),
        },
        "rows": [],
    }


def main():
    data = {key: load_category(key, cfg) for key, cfg in SOURCES.items()}
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    for key, d in data.items():
        s = d["summary"]
        print(f"  {key:12s} status={s['status']:8s} rows={s['totalUpdates']:4d} subjects={s['nSubjects']:3d} latest={s['latestDate']}")
    print(f"\nWrote {OUT_PATH}")


if __name__ == "__main__":
    main()
