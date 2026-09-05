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

import csv
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
        "canonicalList": r"C:\Users\Swati Shukla\Downloads\Corporate_CSR_Monitor\dist\CSR_Companies_2024-25.csv",
    },
    "rnd": {
        "path": os.path.join(OUTPUT_DIR, "RnD_Companies_output.xlsx"),
        "sourceList": "RnD_Companies_2022-23.csv",
        "label": "R&D Companies",
        "canonicalList": r"C:\Users\Swati Shukla\Downloads\Corporate_CSR_Monitor\dist\RnD_Companies_2022-23.csv",
    },
    "institutes": {
        "path": os.path.join(OUTPUT_DIR, "Institute_CSR_Donations.xlsx"),
        "sourceList": "Institutes (IIT / IISc / IISER)",
        "label": "Institutes",
        # No fixed list here — institute mode discovers whichever donor
        # company is named, so there's no authoritative name to snap to.
        "canonicalList": None,
        "instituteMode": True,
    },
}

OUT_PATH = os.path.join(os.path.dirname(__file__), "dashboard_data.json")

_LEGAL_SUFFIX_RE = re.compile(
    r'\b(ltd\.?|limited|pvt\.?|private|inc\.?|incorporated|corp\.?|corporation|'
    r'plc|llp|group|industries|holdings|company|co\.?)\b', re.IGNORECASE)
_PAREN_RE = re.compile(r'\([^)]*\)')
_AND_WORD_RE = re.compile(r'\band\b', re.IGNORECASE)
_FOUNDATION_RE = re.compile(r'\bfoundation\b', re.IGNORECASE)

_MONTHS = {m.lower(): i + 1 for i, m in enumerate(
    ["January", "February", "March", "April", "May", "June", "July", "August",
     "September", "October", "November", "December"])}


def _norm(s):
    return re.sub(r'[^a-z0-9]', '', (s or '').lower())


def _company_root(company):
    """Collapse name variants that are the same real company — legal
    suffixes (Ltd/Limited/Corporation/...), a parenthetical aside or
    abbreviation ("(IOCL)", "(Reliance Foundation)"), the word "foundation"
    on its own, and "and" vs "&" — down to one grouping key. Deliberately
    aggressive: this is for UI grouping/filter-dropdown dedup, not for
    "is this company mentioned in this article" matching, where merging
    too much would cause false positives."""
    s = _PAREN_RE.sub(' ', company or '')
    s = _LEGAL_SUFFIX_RE.sub(' ', s)
    s = _FOUNDATION_RE.sub(' ', s)
    s = _AND_WORD_RE.sub(' ', s)
    return _norm(s) or _norm(company)


_STOPWORDS = {"and", "of", "the", "for", "&"}


def _acronym_candidates(name):
    """Generate plausible acronyms for a canonical company name, e.g.
    "Indian Oil Corporation" -> {"IOC", "IOCL"} — the bare initials, and
    the same with a trailing L, since Indian PSUs are almost always
    informally abbreviated as if "Limited" were part of the acronym
    (IOCL, NHPC, ONGC-style names already end up covered by the bare
    form). Only built from words that survive stripping legal suffixes,
    so "Limited"/"Corporation" itself never becomes part of the initials."""
    words = re.findall(r"[A-Za-z']+", _LEGAL_SUFFIX_RE.sub(' ', name))
    letters = "".join(w[0] for w in words if w.lower() not in _STOPWORDS and w)
    letters = letters.upper()
    if len(letters) < 2:
        return set()
    return {letters, letters + "L"}


def load_canonical_list(path):
    """Load a company-list CSV (one name per row, optional header) and
    build a matcher back to it: `resolve(scraped_name)` returns the
    canonical name from the list if it can confidently identify one, else
    None. Two strategies, in order: (1) normalized-root match — handles
    almost every Ltd/Limited/Pvt/Corporation/"and vs &" spelling
    difference; (2) acronym match — handles the AI response naming a
    company by its abbreviation alone (bare "IOCL", or "(IOCL)" tacked on
    as a parenthetical) with no full form for the root match to catch."""
    if not path or not os.path.exists(path):
        return None
    names = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        for row in csv.reader(f):
            if row and row[0].strip():
                names.append(row[0].strip())
    header_words = {"company", "companies", "name"}
    names = [n for n in names if n.lower() not in header_words]

    root_to_name = {}
    acronym_to_name = {}
    for name in names:
        root_to_name[_company_root(name)] = name
        for ac in _acronym_candidates(name):
            acronym_to_name.setdefault(ac, name)

    def resolve(scraped_name):
        root = _company_root(scraped_name)
        if root in root_to_name:
            return root_to_name[root]
        # Try any parenthetical aside in the scraped text ("... (IOCL)"),
        # and the bare scraped name itself if it's short enough to plausibly
        # be an all-caps acronym on its own.
        candidates = re.findall(r'\(([^)]*)\)', scraped_name or '')
        candidates.append(scraped_name or '')
        for cand in candidates:
            key = re.sub(r'[^A-Za-z]', '', cand).upper()
            if key and key in acronym_to_name:
                return acronym_to_name[key]
        return None

    return resolve


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


def _dedupe_rows(rows):
    """Defensive de-dup for data scraped before the engine's own dedup key
    stopped including category — the same real story used to get written
    twice when two different category searches both found it (e.g. once
    as "New Initiative", once as "CSR Spend"). Same fingerprint shape as
    the engine's current _dedup_key/_institute_dedup_key: company +
    headline text, or a shared source URL, regardless of category. Keeps
    the first occurrence of each fingerprint."""
    seen = set()
    out = []
    for r in rows:
        text = r["initiative"] or r["headline"] or r["update"]
        hl = _norm(text)[:40]
        keys = {f"{r['companyGroup']}|{hl}"}
        url = re.sub(r'^https?://', '', re.sub(r'^www\.', '', (r["source"] or "").strip().lower())).rstrip('/')
        if url and "/" in url:
            keys.add(f"url|{url}")
        if keys & seen:
            continue
        seen.update(keys)
        out.append(r)
    return out


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
    resolve_canonical = load_canonical_list(cfg.get("canonicalList"))

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
        canonical = resolve_canonical(subject) if resolve_canonical and not institute_mode else None
        display_company = canonical or subject
        out_rows.append({
            "id": f"r{n:04d}",
            "company": display_company,
            "scrapedAs": subject if canonical and canonical != subject else "",
            "companyGroup": _company_root(display_company),
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

    before = len(out_rows)
    out_rows = _dedupe_rows(out_rows)
    if len(out_rows) < before:
        print(f"  {key}: dropped {before - len(out_rows)} duplicate row(s) (same story, different category label)")

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
