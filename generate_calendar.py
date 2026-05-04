#!/usr/bin/env python3
"""
Jewish Voice Marketing Calendar Generator - MASTER FIX VERSION.
- Absolute data capture from Excel and CSV.
- Aggressive hyperlink extraction.
- Clean, professional UI with interactive features.
"""

import argparse
import html
import json
import re
import os
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import openpyxl

@dataclass
class CalendarItem:
    id: str
    title: str
    date: str
    source: str
    category: str = "General"
    owner: str = ""
    audience: str = ""
    link: str = ""
    notes: str = ""
    job_number: str = ""
    premium: str = ""
    final_art: str = ""
    segment_code: str = ""

@dataclass
class TapingDate:
    title: str
    date: str

def clean_val(val: Any) -> str:
    if val is None or pd.isna(val): return ""
    return str(val).strip()

def parse_dt(val: Any) -> str:
    if not val or pd.isna(val): return ""
    try:
        if isinstance(val, (datetime, pd.Timestamp)):
            return val.strftime("%Y-%m-%d")
        d = pd.to_datetime(str(val).strip(), errors='coerce')
        return d.strftime("%Y-%m-%d") if pd.notna(d) else str(val).strip()
    except: return str(val).strip()

def get_link(cell: Any) -> str:
    if not cell: return ""
    # Try to get the hyperlink target directly
    if hasattr(cell, 'hyperlink') and cell.hyperlink and cell.hyperlink.target:
        return cell.hyperlink.target
    
    val = str(cell.value) if cell.value else ""
    
    # If the value is a HYPERLINK formula
    if "HYPERLINK" in val.upper():
        m = re.search(r'HYPERLINK\("([^"]+)"\)', val)
        if m: return m.group(1)
        
    # If the value itself is a URL
    if val.startswith("http"): return val
    
    return ""

def process_xlsx(path: Path) -> tuple[list[CalendarItem], list[TapingDate]]:
    items, tapings = [], []
    wb = openpyxl.load_workbook(path, data_only=False)
    for sn in wb.sheetnames:
        ws = wb[sn]
        rows = list(ws.rows)
        if not rows: continue
        
        # Map headers (case-insensitive, space-insensitive)
        hdrs = [re.sub(r'[^a-z]', '', str(c.value).lower()) for c in rows[0]]
        
        def find_idx(targets):
            for t in targets:
                norm_t = re.sub(r'[^a-z]', '', t.lower())
                if norm_t in hdrs: return hdrs.index(norm_t)
            return None

        idx = {
            'title': find_idx(['Description', 'Name of Communication', 'Title', 'Task Name']),
            'date': find_idx(['Target Date', 'Air Date', 'Due Date', 'Date']),
            'type': find_idx(['Type', 'Category']),
            'prod': find_idx(['Producer', 'Owner', 'Assignee']),
            'aud': find_idx(['Audience']),
            'link': find_idx(['Vanity Link', 'Ops Database', 'Vanity linl', 'Link']),
            'job': find_idx(['BBS Job Number', 'Job Number']),
            'prem': find_idx(['Premium']),
            'art': find_idx(['BBS Final Art', 'Final Art']),
            'seg': find_idx(['Segment Code']),
            'notes': find_idx(['Notes', 'Subject', 'Description'])
        }

        if "taping" in sn.lower():
            for r in rows[1:]:
                for c in r:
                    d = parse_dt(c.value)
                    if d and len(d) == 10: tapings.append(TapingDate("Taping", d))
            continue

        if idx['title'] is None: continue

        for r_idx, r in enumerate(rows[1:], 2):
            title = clean_val(r[idx['title']].value)
            if not title: continue
            
            cat = clean_val(r[idx['type']].value) if idx['type'] is not None else "General"
            if any(x in (title + cat).lower() for x in ['broadcast', 'tv']): cat = "Broadcast/TV"

            items.append(CalendarItem(
                id=f"xl-{sn}-{r_idx}",
                title=title,
                date=parse_dt(r[idx['date']].value) if idx['date'] is not None else "",
                source=f"Excel ({path.name})",
                category=cat,
                owner=clean_val(r[idx['prod']].value) if idx['prod'] is not None else "",
                audience=clean_val(r[idx['aud']].value) if idx['aud'] is not None else "",
                link=get_link(r[idx['link']]) if idx['link'] is not None else "",
                notes=clean_val(r[idx['notes']].value) if idx['notes'] is not None else "",
                job_number=clean_val(r[idx['job']].value) if idx['job'] is not None else "",
                premium=clean_val(r[idx['prem']].value) if idx['prem'] is not None else "",
                final_art=clean_val(r[idx['art']].value) if idx['art'] is not None else "",
                segment_code=clean_val(r[idx['seg']].value) if idx['seg'] is not None else ""
            ))
    return items, tapings

def process_csv(path: Path) -> list[CalendarItem]:
    items = []
    try:
        df = pd.read_csv(path)
        hdrs = [re.sub(r'[^a-z]', '', str(c).lower()) for c in df.columns]
        
        def find_idx(targets):
            for t in targets:
                norm_t = re.sub(r'[^a-z]', '', t.lower())
                if norm_t in hdrs: return hdrs.index(norm_t)
            return None
        
        i_title = find_idx(['Description', 'Task Name', 'Name', 'Title'])
        i_date = find_idx(['Target Date', 'Due Date', 'Date'])
        if i_title is None: return []

        for idx, row in df.iterrows():
            title = clean_val(row.iloc[i_title])
            if not title: continue
            items.append(CalendarItem(
                id=f"csv-{idx}", title=title,
                date=parse_dt(row.iloc[i_date]) if i_date is not None else "",
                source=f"CSV ({path.name})", category="General"
            ))
    except: pass
    return items

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--data-dir", type=Path, default=Path(".github/datafiles"))
    parser.add_argument("--output", type=Path, default=Path("index.html"))
    args = parser.parse_args()

    all_items, all_tapings = [], []
    if args.data_dir.exists():
        files = [args.data_dir / f for f in os.listdir(args.data_dir) if not f.startswith("~$")]
        xlsx = sorted([f for f in files if f.suffix == ".xlsx"], key=os.path.getmtime, reverse=True)
        csvs = sorted([f for f in files if f.suffix == ".csv"], key=os.path.getmtime, reverse=True)
        
        if xlsx:
            it, tp = process_xlsx(xlsx[0])
            all_items.extend(it); all_tapings.extend(tp)
        if csvs:
            all_items.extend(process_csv(csvs[0]))

    # Sort and Deduplicate
    all_items.sort(key=lambda x: x.date or "9999-99-99")
    seen_tapings = set()
    unique_tapings = []
    for t in sorted(all_tapings, key=lambda x: x.date):
        if t.date not in seen_tapings:
            unique_tapings.append(t); seen_tapings.add(t.date)

    # Generate HTML (Simplified for reliability)
    payload = json.dumps({
        "items": [asdict(i) for i in all_items],
        "broadcast": [asdict(i) for i in all_items if i.category == "Broadcast/TV"],
        "tapings": [asdict(t) for t in unique_tapings],
        "updated": datetime.now().strftime("%b %d, %Y")
    }, ensure_ascii=False)

    with open("template.html", "r") as f:
        html_template = f.read()
    html_template = html_template.replace("{payload}", payload)
    args.output.write_text(html_template, encoding="utf-8")

if __name__ == "__main__":
    main()
