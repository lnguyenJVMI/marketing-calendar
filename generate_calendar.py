#!/usr/bin/env python3
"""
generate_calendar.py
--------------------
Reads CommunicationsCalendar2026.xlsx (and any companion CSVs) from
.github/datafiles/ and produces index.html in the repository root by
injecting a JSON payload into template.html.

Workbook structure expected:
  - Sheet "Dated Communications": rows with a real Target Date.
      Columns used: Description, Target Date, Producer, Type, Audience,
                    BBS Job Number, Vanity Link, Premium, BBS Final Art,
                    Segment Code.
      Header-style rows (no date) are skipped.
  - Sheet "Standard Communications": recurring items without dates.
      Columns: Inserted into schedule?, Nickname / Abbreviation,
               Name of Communication, Type, How Often, Audience, Subject, Notes.
  - Sheet "Tapings": informal day-range strings like "April 7-8", "August 21-22".

Any other Excel/CSV file dropped into .github/datafiles/ that contains a
"date"-like column will be merged into the dated communications list.

The template uses the literal token __PAYLOAD_PLACEHOLDER__ for injection.
"""
from __future__ import annotations

import json
import re
import sys
from datetime import datetime, date
from pathlib import Path
from typing import Any

import pandas as pd

ROOT = Path(__file__).resolve().parent
TEMPLATE_PATH = ROOT / "template.html"
OUTPUT_PATH = ROOT / "index.html"
DATA_DIR_CANDIDATES = [
    ROOT / ".github" / "datafiles",
    ROOT / "datafiles",
    ROOT,
]
PLACEHOLDER = "__PAYLOAD_PLACEHOLDER__"

DATED_SHEET_NAMES = {"dated communications", "dated", "calendar", "schedule"}
STANDARD_SHEET_NAMES = {"standard communications", "standard", "recurring", "standing"}
TAPINGS_SHEET_NAMES = {"tapings", "taping"}

MONTHS = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #
def _is_blank(v: Any) -> bool:
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    if isinstance(v, str) and not v.strip():
        return True
    return False


def clean_str(v: Any) -> str:
    if _is_blank(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def coerce_premium(v: Any) -> str:
    """Return the literal Premium offer text, or empty string for blank/none."""
    if _is_blank(v):
        return ""
    s = clean_str(v)
    if s.lower() in {"none", "n/a", "na", "-", "—"}:
        return ""
    return s


def to_iso_date(v: Any) -> str | None:
    if _is_blank(v):
        return None
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, pd.Timestamp):
        if pd.isna(v):
            return None
        return v.strftime("%Y-%m-%d")
    s = clean_str(v)
    if not s:
        return None
    try:
        ts = pd.to_datetime(s, errors="raise")
        if pd.isna(ts):
            return None
        return ts.strftime("%Y-%m-%d")
    except Exception:
        return None


# --------------------------------------------------------------------------- #
# file discovery
# --------------------------------------------------------------------------- #
def find_data_files() -> list[Path]:
    files: list[Path] = []
    for d in DATA_DIR_CANDIDATES:
        if d.exists() and d.is_dir():
            for p in sorted(d.iterdir()):
                if p.is_file() and p.suffix.lower() in {".xlsx", ".xls", ".csv"}:
                    files.append(p)
            if files:
                break
    return files


# --------------------------------------------------------------------------- #
# parsers
# --------------------------------------------------------------------------- #
def parse_dated_sheet(df: pd.DataFrame) -> list[dict]:
    # Keep only named columns
    keep = [c for c in df.columns if not str(c).startswith("Unnamed")]
    df = df[keep].copy()

    cols = {str(c).strip().lower(): c for c in df.columns}

    def col(*names: str) -> str | None:
        for n in names:
            if n.lower() in cols:
                return cols[n.lower()]
        return None

    c_desc = col("Description", "Title", "Name")
    c_date = col("Target Date", "Date", "Air Date", "Publish Date")
    c_producer = col("Producer", "Owner")
    c_type = col("Type", "Channel", "Medium")
    c_audience = col("Audience", "Segment", "Community")
    c_job = col("BBS Job Number", "Job #", "Job Number", "Job")
    c_vanity = col("Vanity Link", "Link", "URL")
    c_premium = col("Premium")
    c_art = col("BBS Final Art", "Final Art", "Art")
    c_seg = col("Segment Code", "Segment")

    items: list[dict] = []
    for _, row in df.iterrows():
        iso = to_iso_date(row[c_date]) if c_date else None
        if not iso:
            # Section/header rows like "JANUARY" or holidays without dates
            continue
        title = clean_str(row[c_desc]) if c_desc else ""
        if not title:
            title = "(untitled)"
        items.append({
            "date": iso,
            "title": title,
            "producer": clean_str(row[c_producer]) if c_producer else "",
            "type": clean_str(row[c_type]) if c_type else "",
            "audience": clean_str(row[c_audience]) if c_audience else "",
            "job": clean_str(row[c_job]) if c_job else "",
            "vanity": clean_str(row[c_vanity]) if c_vanity else "",
            "premium": coerce_premium(row[c_premium]) if c_premium else "",
            "art": clean_str(row[c_art]) if c_art else "",
            "segment": clean_str(row[c_seg]) if c_seg else "",
        })
    items.sort(key=lambda it: it["date"])
    return items


def parse_standard_sheet(df: pd.DataFrame) -> list[dict]:
    keep = [c for c in df.columns if not str(c).startswith("Unnamed")]
    df = df[keep].copy()
    cols = {str(c).strip().lower(): c for c in df.columns}

    def col(*names: str) -> str | None:
        for n in names:
            if n.lower() in cols:
                return cols[n.lower()]
        return None

    c_inserted = col("Inserted into schedule?")
    c_nick = col("Nickname / Abbreviation", "Abbreviation", "Nickname")
    c_name = col("Name of Communication", "Name", "Title")
    c_type = col("Type")
    c_freq = col("How Often", "Frequency", "Cadence")
    c_aud = col("Audience")
    c_sub = col("Subject", "Description")
    c_notes = col("Notes")

    items: list[dict] = []
    for _, row in df.iterrows():
        name = clean_str(row[c_name]) if c_name else ""
        if not name:
            continue
        items.append({
            "name": name,
            "abbreviation": clean_str(row[c_nick]) if c_nick else "",
            "type": clean_str(row[c_type]) if c_type else "",
            "frequency": clean_str(row[c_freq]) if c_freq else "",
            "audience": clean_str(row[c_aud]) if c_aud else "",
            "subject": clean_str(row[c_sub]) if c_sub else "",
            "notes": clean_str(row[c_notes]) if c_notes else "",
            "inserted": clean_str(row[c_inserted]) if c_inserted else "",
        })
    return items


# Parse strings like "April 7-8", "August 21-22", "Sept 29-30"
TAPING_RE = re.compile(
    r"^\s*([A-Za-z]+\.?)\s+(\d{1,2})\s*(?:[-–]\s*(\d{1,2}))?\s*$"
)


def parse_taping_string(s: str, default_year: int) -> dict | None:
    s = clean_str(s)
    if not s:
        return None
    m = TAPING_RE.match(s)
    if not m:
        # Try generic date parse
        iso = to_iso_date(s)
        if iso:
            return {"label": s, "start": iso, "end": iso}
        return None
    mon_raw = m.group(1).lower().rstrip(".")
    day_a = int(m.group(2))
    day_b = int(m.group(3)) if m.group(3) else day_a
    mo = MONTHS.get(mon_raw) or MONTHS.get(mon_raw[:3])
    if not mo:
        return None
    try:
        start = date(default_year, mo, day_a)
        end = date(default_year, mo, day_b)
    except ValueError:
        return None
    return {"label": s, "start": start.strftime("%Y-%m-%d"), "end": end.strftime("%Y-%m-%d")}


def parse_tapings_sheet(df: pd.DataFrame, default_year: int) -> list[dict]:
    """The 'Tapings' sheet uses the header cell as a value too; collect both."""
    items: list[dict] = []
    seen: set[str] = set()
    # The pandas header may itself be a value (e.g. "April 7-8")
    for c in df.columns:
        s = clean_str(c)
        if not s or s.lower().startswith("unnamed"):
            continue
        if s in seen:
            continue
        seen.add(s)
        parsed = parse_taping_string(s, default_year)
        if parsed:
            items.append(parsed)
    for _, row in df.iterrows():
        for v in row.tolist():
            s = clean_str(v)
            if not s or s in seen:
                continue
            seen.add(s)
            parsed = parse_taping_string(s, default_year)
            if parsed:
                items.append(parsed)
    items.sort(key=lambda x: x["start"])
    return items


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #
def load_workbook(path: Path) -> tuple[list[dict], list[dict], list[dict]]:
    dated: list[dict] = []
    standard: list[dict] = []
    tapings: list[dict] = []

    # Default year for tapings = first year we see in the dated sheet, else current
    default_year = datetime.now().year

    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
        df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
        if not df.empty:
            dated.extend(parse_dated_sheet(df))
        return dated, standard, tapings

    xls = pd.ExcelFile(path)
    # First pass: gather dated to determine default year
    for sn in xls.sheet_names:
        if sn.strip().lower() in DATED_SHEET_NAMES:
            df = pd.read_excel(path, sheet_name=sn)
            df = df.dropna(axis=0, how="all")
            dated.extend(parse_dated_sheet(df))
    if dated:
        default_year = int(dated[0]["date"][:4])

    for sn in xls.sheet_names:
        norm = sn.strip().lower()
        if norm in DATED_SHEET_NAMES:
            continue  # already parsed
        df = pd.read_excel(path, sheet_name=sn)
        df = df.dropna(axis=0, how="all")
        if norm in STANDARD_SHEET_NAMES:
            standard.extend(parse_standard_sheet(df))
        elif norm in TAPINGS_SHEET_NAMES:
            tapings.extend(parse_tapings_sheet(df, default_year))

    return dated, standard, tapings


def main() -> int:
    if not TEMPLATE_PATH.exists():
        print(f"ERROR: template not found at {TEMPLATE_PATH}", file=sys.stderr)
        return 2

    files = find_data_files()
    if not files:
        print("WARNING: no Excel/CSV data files found; generating an empty calendar.")
        dated, standard, tapings = [], [], []
    else:
        dated_all: list[dict] = []
        standard_all: list[dict] = []
        tapings_all: list[dict] = []
        for f in files:
            print(f"Loading {f.name} ...")
            try:
                d, s, t = load_workbook(f)
            except Exception as exc:
                print(f"  Skipping {f.name}: {exc}")
                continue
            print(f"  +{len(d)} dated, +{len(s)} standard, +{len(t)} tapings")
            dated_all.extend(d)
            standard_all.extend(s)
            tapings_all.extend(t)
        dated_all.sort(key=lambda it: it["date"])
        tapings_all.sort(key=lambda it: it["start"])
        dated, standard, tapings = dated_all, standard_all, tapings_all

    payload = {
        "events": dated,
        "standing": standard,
        "tapings": tapings,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
    }

    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    if PLACEHOLDER not in template:
        print(f"ERROR: template missing {PLACEHOLDER}", file=sys.stderr)
        return 2
    html = template.replace(PLACEHOLDER, json.dumps(payload, ensure_ascii=False))
    OUTPUT_PATH.write_text(html, encoding="utf-8")
    print(f"Wrote {OUTPUT_PATH} ({len(html):,} chars). "
          f"events={len(dated)}, standing={len(standard)}, tapings={len(tapings)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
