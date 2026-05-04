#!/usr/bin/env python3
"""
Jewish Voice Marketing Calendar Generator - Total Capture Version.
- Pulls ALL data from Excel and CSV.
- Extracts hidden hyperlinks from Excel.
- Merges everything into one master calendar.
"""

from __future__ import annotations

import argparse
import html
import json
import re
import os
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional, Any

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
    notes: str = ""

def normalized(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    val = str(value).strip()
    if val.lower() in ["nan", "nat", "none"]:
        return ""
    return val

def normalize_col(column: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(column).lower())

def find_col_idx(df_cols: Iterable[str], target: str) -> Optional[int]:
    target_norm = normalize_col(target)
    for i, col in enumerate(df_cols):
        if normalize_col(col) == target_norm:
            return i
    return None

def parse_date(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    if not text:
        return ""
    try:
        parsed = pd.to_datetime(text, errors="coerce")
        if pd.notna(parsed):
            return parsed.strftime("%Y-%m-%d")
    except: pass
    return text

def extract_excel_data(path: Path) -> tuple[list[CalendarItem], list[TapingDate]]:
    items = []
    tapings = []
    wb = openpyxl.load_workbook(path, data_only=False)
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows())
        if not rows: continue
        
        headers = [str(cell.value) for cell in rows[0]]
        
        # Identify columns
        c_desc = find_col_idx(headers, "Description") or find_col_idx(headers, "Name of Communication") or find_col_idx(headers, "Title")
        c_date = find_col_idx(headers, "Target Date") or find_col_idx(headers, "Air Date") or find_col_idx(headers, "Date")
        c_type = find_col_idx(headers, "Type") or find_col_idx(headers, "Category")
        c_prod = find_col_idx(headers, "Producer") or find_col_idx(headers, "Owner")
        c_aud = find_col_idx(headers, "Audience")
        c_link = find_col_idx(headers, "Vanity Link") or find_col_idx(headers, "Ops Database for Offer Codes") or find_col_idx(headers, "Vanity linl")
        c_job = find_col_idx(headers, "BBS Job Number") or find_col_idx(headers, "Job Number")
        c_prem = find_col_idx(headers, "Premium")
        c_art = find_col_idx(headers, "BBS Final Art") or find_col_idx(headers, "Final Art")
        c_seg = find_col_idx(headers, "Segment Code")
        c_notes = find_col_idx(headers, "Notes") or find_col_idx(headers, "Subject")

        # Special handling for Taping sheets
        if "taping" in sheet_name.lower():
            for r_idx, row in enumerate(rows[1:], start=2):
                for cell in row:
                    d = parse_date(cell.value)
                    if d: tapings.append(TapingDate(title="Taping Session", date=d))
            continue

        if c_desc is None: continue

        for r_idx, row in enumerate(rows[1:], start=2):
            title = normalized(row[c_desc].value)
            if not title: continue
            
            # Extract Link
            link_url = ""
            if c_link is not None:
                cell = row[c_link]
                if cell.hyperlink:
                    link_url = cell.hyperlink.target
                else:
                    val = str(cell.value) if cell.value else ""
                    if val.startswith("http"):
                        link_url = val
                    elif "HYPERLINK" in val:
                        match = re.search(r'HYPERLINK\("([^"]+)"', val)
                        if match: link_url = match.group(1)

            date = parse_date(row[c_date].value) if c_date is not None else ""
            category = normalized(row[c_type].value) if c_type is not None else "General"
            if "broadcast" in (title + " " + category).lower() or "tv" in (title + " " + category).lower():
                category = "Broadcast/TV"

            items.append(CalendarItem(
                id=f"xl-{sheet_name}-{r_idx}",
                title=title, date=date, source=f"Excel ({path.name})", category=category,
                owner=normalized(row[c_prod].value) if c_prod is not None else "",
                audience=normalized(row[c_aud].value) if c_aud is not None else "",
                link=link_url,
                notes=normalized(row[c_notes].value) if c_notes is not None else "",
                job_number=normalized(row[c_job].value) if c_job is not None else "",
                premium=normalized(row[c_prem].value) if c_prem is not None else "",
                final_art=normalized(row[c_art].value) if c_art is not None else "",
                segment_code=normalized(row[c_seg].value) if c_seg is not None else "",
            ))
    return items, tapings

def extract_csv_data(path: Path) -> list[CalendarItem]:
    items = []
    try:
        df = pd.read_csv(path)
        cols = df.columns
        c_desc = find_col_idx(cols, "Description") or find_col_idx(cols, "Task Name") or find_col_idx(cols, "Name")
        c_date = find_col_idx(cols, "Target Date") or find_col_idx(cols, "Due Date") or find_col_idx(cols, "Date")
        c_type = find_col_idx(cols, "Type") or find_col_idx(cols, "Category")
        c_prod = find_col_idx(cols, "Producer") or find_col_idx(cols, "Assignee")
        c_notes = find_col_idx(cols, "Notes") or find_col_idx(cols, "Description")

        if c_desc is None: return items

        for idx, row in df.iterrows():
            title = normalized(row.iloc[c_desc])
            if not title: continue
            items.append(CalendarItem(
                id=f"csv-{idx}",
                title=title,
                date=parse_date(row.iloc[c_date]) if c_date is not None else "",
                source=f"CSV ({path.name})",
                category=normalized(row.iloc[c_type]) if c_type is not None else "General",
                owner=normalized(row.iloc[c_prod]) if c_prod is not None else "",
                notes=normalized(row.iloc[c_notes]) if c_notes is not None else ""
            ))
    except Exception as e:
        print(f"Error reading CSV {path.name}: {e}")
    return items

def load_latest_files(folder_path: Path) -> tuple[list[CalendarItem], list[TapingDate]]:
    all_items = []
    all_tapings = []
    if not folder_path.exists(): return [], []

    xlsx_files = [folder_path / f for f in os.listdir(folder_path) if f.lower().endswith(".xlsx") and not f.startswith("~$")]
    csv_files = [folder_path / f for f in os.listdir(folder_path) if f.lower().endswith(".csv")]
    
    if xlsx_files:
        latest_xlsx = max(xlsx_files, key=os.path.getmtime)
        items, tapings = extract_excel_data(latest_xlsx)
        all_items.extend(items)
        all_tapings.extend(tapings)

    if csv_files:
        latest_csv = max(csv_files, key=os.path.getmtime)
        all_items.extend(extract_csv_data(latest_csv))
            
    return all_items, all_tapings

def generate_html(items: list[CalendarItem], tapings: list[TapingDate], title: str) -> str:
    today = datetime.now().strftime("%Y-%m-%d")
    items_sorted = sorted([i for i in items if i.date], key=lambda x: x.date)
    undated = [i for i in items if not i.date]
    broadcast = [i for i in items_sorted if i.category == "Broadcast/TV"]
    tapings_sorted = sorted(list({{t.date: t for t in tapings}}.values()), key=lambda x: x.date)

    payload = {
        "items": [asdict(i) for i in items_sorted + undated],
        "broadcast": [asdict(i) for i in broadcast],
        "tapings": [asdict(t) for t in tapings_sorted],
        "today": today,
    }
    
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(title)}</title>
<style>
:root {{ --bg:#f1f5f9; --panel:#ffffff; --ink:#1e293b; --muted:#64748b; --line:#cbd5e1; --blue:#0284c7; --jv-blue: #1e3a8a; }}
body {{ margin:0; background:var(--bg); color:var(--ink); font-family: 'Inter', sans-serif; line-height: 1.5; scroll-behavior: smooth; }}
header {{ padding: 20px 40px; background: var(--jv-blue); color:#fff; display: flex; justify-content: space-between; align-items: center; }}
header h1 {{ margin:0; font-size: 22px; font-weight: 700; }}
.last-updated {{ font-size: 12px; opacity: 0.8; }}
main {{ max-width: 1400px; margin: 0 auto; padding: 24px; }}
.tabs {{ display:flex; gap:4px; margin-bottom: 24px; background: #e2e8f0; padding: 4px; border-radius: 10px; width: fit-content; }}
.tab {{ padding:8px 24px; border-radius:8px; border:none; background:transparent; cursor:pointer; font-weight:600; color: var(--muted); }}
.tab.active {{ background:#fff; color:var(--jv-blue); box-shadow: 0 2px 4px rgba(0,0,0,0.05); }}
.panel {{ display:none; }}
.panel.active {{ display:block; }}
.card {{ background:#fff; border:1px solid var(--line); border-radius:12px; padding:24px; margin-bottom:24px; }}
.controls {{ display:grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap:16px; margin-bottom:24px; }}
input, select {{ padding:10px 14px; border:1px solid var(--line); border-radius:8px; width:100%; box-sizing:border-box; }}
table {{ width:100%; border-collapse:collapse; font-size: 14px; }}
th, td {{ padding:14px; text-align:left; border-bottom:1px solid var(--line); vertical-align: top; }}
th {{ background:#f8fafc; font-size:11px; text-transform:uppercase; color:var(--muted); font-weight: 700; }}
.chip {{ display:inline-block; padding:2px 10px; border-radius:20px; font-size:11px; font-weight:700; background:#f1f5f9; }}
.chip.broadcast {{ background:#dcfce7; color:#166534; }}
.link-btn {{ display: inline-flex; align-items: center; gap: 6px; background: var(--blue); color: #fff; padding: 6px 12px; border-radius: 6px; text-decoration: none; font-size: 12px; font-weight: 600; }}
.detail-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 8px; margin-top: 8px; }}
.detail-item {{ font-size: 11px; color: var(--muted); }}
.detail-label {{ font-weight: 700; color: var(--ink); text-transform: uppercase; font-size: 9px; display: block; }}
.taping-list {{ display:flex; flex-wrap:wrap; gap:10px; }}
.taping {{ background:#fff7ed; border:1px solid #fed7aa; color:#9a3412; padding:8px 16px; border-radius:12px; font-size:14px; font-weight:700; cursor:pointer; }}
.timeline-item {{ border-left: 3px solid var(--blue); padding: 0 0 32px 24px; position:relative; }}
.timeline-item::before {{ content:""; position:absolute; left:-9px; top:0; width:15px; height:15px; border-radius:50%; background:#fff; border: 3px solid var(--blue); }}
.timeline-date {{ font-weight:800; color:var(--blue); font-size:14px; margin-bottom:8px; }}
.timeline-content {{ background: #fff; border: 1px solid var(--line); padding: 20px; border-radius: 12px; }}
.timeline-title {{ font-size:18px; font-weight:700; margin-bottom:8px; color: var(--jv-blue); }}
.producer-text {{ font-size: 16px; font-weight: 700; color: var(--ink); }}
.audience-text {{ font-size: 13px; color: var(--muted); }}
</style>
</head>
<body>
<header>
    <h1>{html.escape(title)}</h1>
    <div class="last-updated">Last Updated: {datetime.now().strftime("%b %d, %Y")}</div>
</header>
<main>
    <div class="tabs">
        <button class="tab active" id="tab-calendar" onclick="showTab('calendar')">Marketing Calendar</button>
        <button class="tab" id="tab-broadcast" onclick="showTab('broadcast')">Broadcast Schedule</button>
    </div>
    <div id="calendar" class="panel active">
        <div class="card controls">
            <input type="text" id="search" placeholder="Search title, job #, premium..." oninput="render()">
            <select id="typeFilter" onchange="render()"><option value="">All Types</option></select>
            <select id="producerFilter" onchange="render()"><option value="">All Producers</option></select>
        </div>
        <div class="card" style="overflow-x:auto; padding: 0;">
            <table>
                <thead>
                    <tr><th style="width:100px">Date</th><th>Details</th><th style="width:120px">Type</th><th style="width:180px">Producer/Audience</th><th style="width:150px">Links</th></tr>
                </thead>
                <tbody id="tableBody"></tbody>
            </table>
        </div>
    </div>
    <div id="broadcast" class="panel">
        <div class="card">
            <h2 style="margin-top:0; font-size: 18px;">Upcoming Taping Sessions</h2>
            <div id="tapingList" class="taping-list"></div>
        </div>
        <div class="timeline" id="timeline"></div>
    </div>
</main>
<script>
const DATA = {json.dumps(payload, ensure_ascii=False)};
function showTab(id) {{
    document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.getElementById(id).classList.add('active');
    document.getElementById('tab-' + id).classList.add('active');
}}
function jumpToDate(date) {{
    showTab('broadcast');
    setTimeout(() => {{
        const el = document.querySelector(`[data-date="${{date}}"]`);
        if (el) el.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
    }}, 100);
}}
function render() {{
    const search = document.getElementById('search').value.toLowerCase();
    const type = document.getElementById('typeFilter').value;
    const prod = document.getElementById('producerFilter').value;
    const filtered = DATA.items.filter(i => {{
        const text = (i.title + i.notes + i.job_number + i.premium + i.segment_code).toLowerCase();
        return text.includes(search) && (!type || i.category === type) && (!prod || i.owner === prod);
    }});
    document.getElementById('tableBody').innerHTML = filtered.map(i => `
        <tr>
            <td style="font-weight:700; color:var(--jv-blue)">${{i.date || 'Undated'}}</td>
            <td>
                <div style="font-weight:700; font-size:15px">${{i.title}}</div>
                <div style="color:var(--muted); font-size:13px">${{i.notes}}</div>
                <div class="detail-grid">
                    ${{i.job_number ? `<div class="detail-item"><span class="detail-label">Job #</span>${{i.job_number}}</div>` : ''}}
                    ${{i.premium ? `<div class="detail-item"><span class="detail-label">Premium</span>${{i.premium}}</div>` : ''}}
                    ${{i.segment_code ? `<div class="detail-item"><span class="detail-label">Segment</span>${{i.segment_code}}</div>` : ''}}
                    ${{i.final_art ? `<div class="detail-item"><span class="detail-label">Final Art</span>${{i.final_art}}</div>` : ''}}
                </div>
            </td>
            <td><span class="chip ${{i.category === 'Broadcast/TV' ? 'broadcast' : ''}}">${{i.category}}</span></td>
            <td><div class="producer-text">${{i.owner}}</div><div class="audience-text">${{i.audience}}</div></td>
            <td>${{i.link ? `<a href="${{i.link}}" class="link-btn" target="_blank">🔗 Vanity Link</a>` : ''}}</td>
        </tr>
    `).join('');
}}
function init() {{
    const types = [...new Set(DATA.items.map(i => i.category))].filter(Boolean).sort();
    const prods = [...new Set(DATA.items.map(i => i.owner))].filter(Boolean).sort();
    document.getElementById('typeFilter').innerHTML += types.map(t => `<option value="${{t}}">${{t}}</option>`).join('');
    document.getElementById('producerFilter').innerHTML += prods.map(p => `<option value="${{p}}">${{p}}</option>`).join('');
    document.getElementById('tapingList').innerHTML = DATA.tapings.map(t => `<div class="taping" onclick="jumpToDate('${{t.date}}')">${{t.date}}</div>`).join('');
    document.getElementById('timeline').innerHTML = DATA.broadcast.map(i => `
        <div class="timeline-item" data-date="${{i.date}}">
            <div class="timeline-date">${{i.date}}</div>
            <div class="timeline-content">
                <div class="timeline-title">${{i.title}}</div>
                <div style="color:var(--muted); margin-bottom:12px">${{i.notes}}</div>
                <div class="detail-grid">
                    <div class="detail-item"><span class="detail-label">Producer</span>${{i.owner}}</div>
                    ${{i.job_number ? `<div class="detail-item"><span class="detail-label">Job #</span>${{i.job_number}}</div>` : ''}}
                    ${{i.premium ? `<div class="detail-item"><span class="detail-label">Premium</span>${{i.premium}}</div>` : ''}}
                </div>
                ${{i.link ? `<div style="margin-top:12px"><a href="${{i.link}}" class="link-btn" target="_blank">🔗 Open Vanity Link</a></div>` : ''}}
            </div>
        </div>
    `).join('');
    render();
}}
init();
</script>
</body>
</html>
"""

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--data-dir", type=Path, default=Path(".github/datafiles"))
    parser.add_argument("--output", type=Path, default=Path("index.html"))
    args = parser.parse_args()
    items, tapings = load_latest_files(args.data_dir)
    html = generate_html(items, tapings, "Jewish Voice Marketing Calendar")
    args.output.write_text(html, encoding="utf-8")
    print(f"Generated {{args.output}} with {{len(items)}} items.")

if __name__ == "__main__":
    main()
