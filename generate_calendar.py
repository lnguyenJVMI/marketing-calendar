#!/usr/bin/env python3
"""
Jewish Voice Marketing Calendar Generator - Final Robust Version.
Extracts actual URLs from Excel cells and handles auto-discovery of latest files.
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
from typing import Iterable, Optional

import pandas as pd
import openpyxl

@dataclass
class CalendarItem:
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

def normalized(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    val = str(value).strip()
    if val.lower() in ["nan", "nat", "none"]:
        return ""
    if isinstance(value, float):
        if value == int(value):
            return f"{int(value)}"
        return f"{value}"
    return val

def normalize_col(column: object) -> str:
    return re.sub(r"[^a-z0-9]", "", str(column).lower())

def find_col_idx(df_cols: Iterable[str], target: str) -> Optional[int]:
    target_norm = normalize_col(target)
    for i, col in enumerate(df_cols):
        if normalize_col(col) == target_norm:
            return i
    return None

def parse_date(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    if not text:
        return ""
    parsed = pd.to_datetime(text, errors="coerce")
    if pd.notna(parsed):
        return parsed.strftime("%Y-%m-%d")
    return text

def process_dataframe(df: pd.DataFrame, source_name: str) -> list[CalendarItem]:
    items = []
    cols = df.columns
    c_desc = find_col_idx(cols, "Description") or find_col_idx(cols, "Name of Communication") or find_col_idx(cols, "Title") or find_col_idx(cols, "Task Name")
    c_date = find_col_idx(cols, "Target Date") or find_col_idx(cols, "Air Date") or find_col_idx(cols, "Due Date") or find_col_idx(cols, "Date")
    c_type = find_col_idx(cols, "Type") or find_col_idx(cols, "Category") or find_col_idx(cols, "Section")
    c_prod = find_col_idx(cols, "Producer") or find_col_idx(cols, "Owner") or find_col_idx(cols, "Assignee")
    c_aud = find_col_idx(cols, "Audience")
    c_link = find_col_idx(cols, "Vanity Link") or find_col_idx(cols, "Ops Database for Offer Codes") or find_col_idx(cols, "Vanity linl") or find_col_idx(cols, "Link")
    c_job = find_col_idx(cols, "BBS Job Number") or find_col_idx(headers, "Job Number")
    c_prem = find_col_idx(cols, "Premium")
    c_art = find_col_idx(cols, "BBS Final Art") or find_col_idx(headers, "Final Art")
    c_seg = find_col_idx(cols, "Segment Code")
    c_notes = find_col_idx(cols, "Notes") or find_col_idx(headers, "Subject")

    if c_desc is None:
        return items

    for _, row in df.iterrows():
        title = normalized(row.iloc[c_desc])
        if not title: continue
        
        date = parse_date(row.iloc[c_date]) if c_date is not None else ""
        category = normalized(row.iloc[c_type]) if c_type is not None else "General"
        
        if "broadcast" in (title + " " + category).lower() or "tv" in (title + " " + category).lower():
            category = "Broadcast/TV"

        items.append(CalendarItem(
            title=title, date=date, source=source_name, category=category,
            owner=normalized(row.iloc[c_prod]) if c_prod is not None else "",
            audience=normalized(row.iloc[c_aud]) if c_aud is not None else "",
            link=normalized(row.iloc[c_link]) if c_link is not None else "",
            notes=normalized(row.iloc[c_notes]) if c_notes is not None else "",
            job_number=normalized(row.iloc[c_job]) if c_job is not None else "",
            premium=normalized(row.iloc[c_prem]) if c_prem is not None else "",
            final_art=normalized(row.iloc[c_art]) if c_art is not None else "",
            segment_code=normalized(row.iloc[c_seg]) if c_seg is not None else "",
        ))
    return items

def extract_excel_with_links(path: Path) -> list[CalendarItem]:
    items = []
    wb = openpyxl.load_workbook(path, data_only=False)
    
    for sheet_name in wb.sheetnames:
        if "taping" in sheet_name.lower():
            continue
            
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows: continue
        
        headers = [str(h) for h in rows[0]]
        
        # Find column indices
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

        if c_desc is None: continue

        # Process rows (skip header)
        for r_idx, row_vals in enumerate(ws.iter_rows(min_row=2), start=2):
            title = normalized(row_vals[c_desc].value)
            if not title: continue
            
            link_url = ""
            if c_link is not None:
                cell = row_vals[c_link]
                if cell.hyperlink:
                    link_url = cell.hyperlink.target
                else:
                    val = normalized(cell.value)
                    if val.startswith("http"):
                        link_url = val

            date = parse_date(row_vals[c_date].value) if c_date is not None else ""
            category = normalized(row_vals[c_type].value) if c_type is not None else "General"
            
            if "broadcast" in (title + " " + category).lower() or "tv" in (title + " " + category).lower():
                category = "Broadcast/TV"

            items.append(CalendarItem(
                title=title, date=date, source=f"Excel ({path.name})", category=category,
                owner=normalized(row_vals[c_prod].value) if c_prod is not None else "",
                audience=normalized(row_vals[c_aud].value) if c_aud is not None else "",
                link=link_url,
                notes=normalized(row_vals[c_notes].value) if c_notes is not None else "",
                job_number=normalized(row_vals[c_job].value) if c_job is not None else "",
                premium=normalized(row_vals[c_prem].value) if c_prem is not None else "",
                final_art=normalized(row_vals[c_art].value) if c_art is not None else "",
                segment_code=normalized(row_vals[c_seg].value) if c_seg is not None else "",
            ))
    return items

def load_latest_from_folder(folder_path: Path) -> tuple[list[CalendarItem], list[TapingDate]]:
    all_items: list[CalendarItem] = []
    all_tapings: list[TapingDate] = []
    
    if not folder_path.exists():
        return all_items, all_tapings

    xlsx_files = [folder_path / f for f in os.listdir(folder_path) if f.lower().endswith(".xlsx") and not f.startswith("~$")]
    csv_files = [folder_path / f for f in os.listdir(folder_path) if f.lower().endswith(".csv")]
    
    if xlsx_files:
        latest_xlsx = max(xlsx_files, key=os.path.getmtime)
        print(f"Processing Latest Excel: {latest_xlsx.name}")
        all_items.extend(extract_excel_with_links(latest_xlsx))
        
        try:
            xl = pd.ExcelFile(latest_xlsx)
            for sheet in xl.sheet_names:
                if "taping" in sheet.lower():
                    df = xl.parse(sheet)
                    for col in df.columns:
                        if any(m in str(col).lower() for m in ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]):
                            all_tapings.append(TapingDate(title="Taping Session", date=str(col)))
                        for val in df[col].dropna():
                            all_tapings.append(TapingDate(title="Taping Session", date=str(val)))
        except: pass

    if csv_files:
        latest_csv = max(csv_files, key=os.path.getmtime)
        print(f"Processing Latest CSV: {latest_csv.name}")
        try:
            df_csv = pd.read_csv(latest_csv)
            all_items.extend(process_dataframe(df_csv, f"CSV ({latest_csv.name})"))
        except Exception as e:
            print(f"Error reading CSV: {e}")
            
    return all_items, all_tapings

def generate_html(items: list[CalendarItem], tapings: list[TapingDate], title: str) -> str:
    today = datetime.now().strftime("%Y-%m-%d")
    items_sorted = sorted([i for i in items if i.date], key=lambda x: x.date)
    undated = [i for i in items if not i.date]
    broadcast = [i for i in items_sorted if i.category == "Broadcast/TV"]

    payload = {
        "items": [asdict(i) for i in items_sorted + undated],
        "broadcast": [asdict(i) for i in broadcast],
        "tapings": [asdict(t) for t in tapings],
        "today": today,
    }
    
    data_json = json.dumps(payload, ensure_ascii=False)
    
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(title)}</title>
<style>
:root {{ --bg:#f1f5f9; --panel:#ffffff; --ink:#1e293b; --muted:#64748b; --line:#cbd5e1; --blue:#0284c7; --green:#059669; --amber:#d97706; --jv-blue: #1e3a8a; }}
body {{ margin:0; background:var(--bg); color:var(--ink); font-family: 'Inter', system-ui, -apple-system, sans-serif; line-height: 1.5; }}
header {{ padding: 20px 40px; background: var(--jv-blue); color:#fff; display: flex; justify-content: space-between; align-items: center; }}
header h1 {{ margin:0; font-size: 22px; font-weight: 700; letter-spacing: -0.025em; }}
.last-updated {{ font-size: 12px; opacity: 0.8; }}
main {{ max-width: 1400px; margin: 0 auto; padding: 24px; }}
.tabs {{ display:flex; gap:4px; margin-bottom: 24px; background: #e2e8f0; padding: 4px; border-radius: 10px; width: fit-content; }}
.tab {{ padding:8px 24px; border-radius:8px; border:none; background:transparent; cursor:pointer; font-weight:600; color: var(--muted); transition: all 0.2s; }}
.tab.active {{ background:#fff; color:var(--jv-blue); box-shadow: 0 2px 4px rgba(0,0,0,0.05); }}
.panel {{ display:none; animation: fadeIn 0.3s ease-in-out; }}
.panel.active {{ display:block; }}
@keyframes fadeIn {{ from {{ opacity: 0; transform: translateY(4px); }} to {{ opacity: 1; transform: translateY(0); }} }}
.card {{ background:#fff; border:1px solid var(--line); border-radius:12px; padding:24px; margin-bottom:24px; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }}
.controls {{ display:grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap:16px; margin-bottom:24px; }}
input, select {{ padding:10px 14px; border:1px solid var(--line); border-radius:8px; width:100%; box-sizing:border-box; font-size: 14px; }}
table {{ width:100%; border-collapse:collapse; font-size: 14px; }}
th, td {{ padding:14px; text-align:left; border-bottom:1px solid var(--line); vertical-align: top; }}
th {{ background:#f8fafc; font-size:11px; text-transform:uppercase; color:var(--muted); font-weight: 700; letter-spacing: 0.05em; }}
tr:hover td {{ background: #f8fafc; }}
.chip {{ display:inline-block; padding:2px 10px; border-radius:20px; font-size:11px; font-weight:700; background:#f1f5f9; color: var(--muted); }}
.chip.broadcast {{ background:#dcfce7; color:#166534; }}
.link-btn {{ display: inline-flex; align-items: center; gap: 6px; background: var(--blue); color: #fff; padding: 6px 12px; border-radius: 6px; text-decoration: none; font-size: 12px; font-weight: 600; }}
.link-btn:hover {{ background: #0369a1; }}
.detail-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 8px; margin-top: 8px; }}
.detail-item {{ font-size: 11px; color: var(--muted); }}
.detail-label {{ font-weight: 700; color: var(--ink); text-transform: uppercase; font-size: 9px; display: block; }}
.taping-list {{ display:flex; flex-wrap:wrap; gap:10px; }}
.taping {{ background:#fff7ed; border:1px solid #fed7aa; color:#9a3412; padding:8px 16px; border-radius:12px; font-size:14px; font-weight:600; }}
.timeline-item {{ border-left: 3px solid var(--blue); padding: 0 0 32px 24px; position:relative; }}
.timeline-item::before {{ content:""; position:absolute; left:-9px; top:0; width:15px; height:15px; border-radius:50%; background:#fff; border: 3px solid var(--blue); }}
.timeline-date {{ font-weight:800; color:var(--blue); font-size:14px; margin-bottom:8px; }}
.timeline-content {{ background: #fff; border: 1px solid var(--line); padding: 20px; border-radius: 12px; }}
.timeline-title {{ font-size:18px; font-weight:700; margin-bottom:8px; color: var(--jv-blue); }}
</style>
</head>
<body>
<header>
    <h1>{html.escape(title)}</h1>
    <div class="last-updated">Last Updated: {datetime.now().strftime("%b %d, %Y")}</div>
</header>
<main>
    <div class="tabs">
        <button class="tab active" onclick="showTab('calendar')">Marketing Calendar</button>
        <button class="tab" onclick="showTab('broadcast')">Broadcast Schedule</button>
    </div>

    <div id="calendar" class="panel active">
        <div class="card controls">
            <input type="text" id="search" placeholder="Search by title, job #, or premium..." oninput="render()">
            <select id="typeFilter" onchange="render()"><option value="">All Types</option></select>
            <select id="producerFilter" onchange="render()"><option value="">All Producers</option></select>
        </div>
        <div class="card" style="overflow-x:auto; padding: 0;">
            <table id="dataTable">
                <thead>
                    <tr>
                        <th style="width: 100px">Date</th>
                        <th>Communication Details</th>
                        <th style="width: 120px">Type</th>
                        <th style="width: 150px">Producer / Audience</th>
                        <th style="width: 150px">Links & Assets</th>
                    </tr>
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
const DATA = {data_json};

function showTab(id) {{
    document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.getElementById(id).classList.add('active');
    event.currentTarget.classList.add('active');
}}

function render() {{
    const search = document.getElementById('search').value.toLowerCase();
    const type = document.getElementById('typeFilter').value;
    const producer = document.getElementById('producerFilter').value;

    const filtered = DATA.items.filter(i => {{
        const text = (i.title + i.notes + i.job_number + i.premium + i.segment_code).toLowerCase();
        const matchSearch = text.includes(search);
        const matchType = !type || i.category === type;
        const matchProducer = !producer || i.owner === producer;
        return matchSearch && matchType && matchProducer;
    }});

    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = filtered.map(i => `
        <tr>
            <td style="white-space:nowrap; font-weight: 700; color: var(--jv-blue)">${{i.date || 'Undated'}}</td>
            <td>
                <div style="font-weight: 700; font-size: 15px; margin-bottom: 4px;">${{i.title}}</div>
                <div style="color: var(--muted); font-size: 13px; margin-bottom: 8px;">${{i.notes}}</div>
                <div class="detail-grid">
                    ${{i.job_number ? `<div class="detail-item"><span class="detail-label">Job #</span>${{i.job_number}}</div>` : ''}}
                    ${{i.premium ? `<div class="detail-item"><span class="detail-label">Premium</span>${{i.premium}}</div>` : ''}}
                    ${{i.segment_code ? `<div class="detail-item"><span class="detail-label">Segment</span>${{i.segment_code}}</div>` : ''}}
                    ${{i.final_art ? `<div class="detail-item"><span class="detail-label">Final Art</span>${{i.final_art}}</div>` : ''}}
                </div>
            </td>
            <td><span class="chip ${{i.category === 'Broadcast/TV' ? 'broadcast' : ''}}">${{i.category}}</span></td>
            <td>
                <div style="font-weight: 600;">${{i.owner}}</div>
                <div style="font-size: 11px; color: var(--muted);">${{i.audience}}</div>
            </td>
            <td>
                ${{i.link ? `<a href="${{i.link}}" class="link-btn" target="_blank">🔗 Vanity Link</a>` : `<span style="color: #cbd5e1; font-size: 11px;">No link provided</span>`}}
            </td>
        </tr>
    `).join('');
}}

function init() {{
    const types = [...new Set(DATA.items.map(i => i.category))].filter(Boolean).sort();
    const producers = [...new Set(DATA.items.map(i => i.owner))].filter(Boolean).sort();
    
    document.getElementById('typeFilter').innerHTML += types.map(t => `<option value="${{t}}">${{t}}</option>`).join('');
    document.getElementById('producerFilter').innerHTML += producers.map(p => `<option value="${{p}}">${{p}}</option>`).join('');

    document.getElementById('tapingList').innerHTML = DATA.tapings.map(t => `
        <div class="taping">${{t.date}}</div>
    `).join('');

    document.getElementById('timeline').innerHTML = DATA.broadcast.map(i => `
        <div class="timeline-item">
            <div class="timeline-date">${{i.date}}</div>
            <div class="timeline-content">
                <div class="timeline-title">${{i.title}}</div>
                <div style="margin-bottom: 12px; color: var(--muted);">${{i.notes}}</div>
                <div class="detail-grid" style="margin-bottom: 16px;">
                    <div class="detail-item"><span class="detail-label">Producer</span>${{i.owner}}</div>
                    <div class="detail-item"><span class="detail-label">Audience</span>${{i.audience}}</div>
                    ${{i.job_number ? `<div class="detail-item"><span class="detail-label">Job #</span>${{i.job_number}}</div>` : ''}}
                    ${{i.premium ? `<div class="detail-item"><span class="detail-label">Premium</span>${{i.premium}}</div>` : ''}}
                    ${{i.segment_code ? `<div class="detail-item"><span class="detail-label">Segment</span>${{i.segment_code}}</div>` : ''}}
                </div>
                ${{i.link ? `<a href="${{i.link}}" class="link-btn" target="_blank">🔗 Open Vanity Link</a>` : ''}}
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

    items, tapings = load_latest_from_folder(args.data_dir)
    html_content = generate_html(items, tapings, "Jewish Voice Marketing Calendar")
    args.output.write_text(html_content, encoding="utf-8")
    print(f"Generated {args.output} with {len(items)} items from latest files in {args.data_dir}.")

if __name__ == "__main__":
    main()
