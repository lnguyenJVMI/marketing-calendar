#!/usr/bin/env python3
"""
Comprehensive Marketing Calendar Generator for Jewish Voice.
Captures all fields: Vanity Link, Premium, BBS Final Art, Segment Code, Job Number, etc.
"""

from __future__ import annotations

import argparse
import html
import json
import re
from dataclasses import dataclass, asdict, field
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional, Dict, Any

import pandas as pd

# Column detection candidates
DATE_CANDIDATES = ["target date", "air date", "date", "start date", "publish date"]
TITLE_CANDIDATES = ["description", "name of communication", "title", "task name"]
TYPE_CANDIDATES = ["type", "category", "channel"]
OWNER_CANDIDATES = ["producer", "assignee", "owner", "lead"]
LINK_CANDIDATES = ["vanity link", "link", "url"]
AUDIENCE_CANDIDATES = ["audience"]
NOTES_CANDIDATES = ["notes", "subject"]

# Specific Jewish Voice Fields
JV_FIELDS = {
    "job_number": ["bbs job number", "job number"],
    "premium": ["premium"],
    "final_art": ["bbs final art", "final art"],
    "segment_code": ["segment code"],
}

BROADCAST_KEYWORDS = ["broadcast", "tv", "television", "show", "episode", "air", "aired"]

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
    extra_details: Dict[str, str] = field(default_factory=dict)

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
    # Handle scientific notation for job numbers if they come in as floats
    if isinstance(value, float) and value > 1000:
        return f"{int(value)}"
    return val

def normalize_col(column: object) -> str:
    return re.sub(r"\s+", " ", str(column).strip().lower())

def find_column(columns: Iterable[object], candidates: list[str]) -> Optional[object]:
    lookup = {normalize_col(c): c for c in columns}
    for candidate in candidates:
        if candidate in lookup:
            return lookup[candidate]
    for col in columns:
        ncol = normalize_col(col)
        if any(candidate in ncol for candidate in candidates):
            return col
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

def load_calendar_data(path: Path) -> tuple[list[CalendarItem], list[TapingDate]]:
    items: list[CalendarItem] = []
    tapings: list[TapingDate] = []
    
    if not path.exists():
        return items, tapings

    xl = pd.ExcelFile(path)
    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        df = df.dropna(axis=1, how="all")
        
        if "taping" in sheet_name.lower():
            for col in df.columns:
                if any(m in str(col).lower() for m in ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]):
                    tapings.append(TapingDate(title="Taping Session", date=str(col)))
                for val in df[col].dropna():
                    tapings.append(TapingDate(title="Taping Session", date=str(val)))
            continue

        date_col = find_column(df.columns, DATE_CANDIDATES)
        title_col = find_column(df.columns, TITLE_CANDIDATES)
        type_col = find_column(df.columns, TYPE_CANDIDATES)
        owner_col = find_column(df.columns, OWNER_CANDIDATES)
        link_col = find_column(df.columns, LINK_CANDIDATES)
        audience_col = find_column(df.columns, AUDIENCE_CANDIDATES)
        notes_col = find_column(df.columns, NOTES_CANDIDATES)
        
        # Map JV specific fields
        jv_cols = {k: find_column(df.columns, v) for k, v in JV_FIELDS.items()}

        if not title_col:
            continue

        for _, row in df.iterrows():
            title = normalized(row.get(title_col))
            if not title:
                continue
                
            date = parse_date(row.get(date_col)) if date_col else ""
            category = normalized(row.get(type_col)) if type_col else "General"
            
            if any(k in (title + " " + category).lower() for k in BROADCAST_KEYWORDS):
                category = "Broadcast/TV"

            items.append(CalendarItem(
                title=title,
                date=date,
                source="Communications Calendar",
                category=category,
                owner=normalized(row.get(owner_col)) if owner_col else "",
                audience=normalized(row.get(audience_col)) if audience_col else "",
                link=normalized(row.get(link_col)) if link_col else "",
                notes=normalized(row.get(notes_col)) if notes_col else "",
                job_number=normalized(row.get(jv_cols["job_number"])) if jv_cols["job_number"] else "",
                premium=normalized(row.get(jv_cols["premium"])) if jv_cols["premium"] else "",
                final_art=normalized(row.get(jv_cols["final_art"])) if jv_cols["final_art"] else "",
                segment_code=normalized(row.get(jv_cols["segment_code"])) if jv_cols["segment_code"] else "",
            ))
            
    return items, tapings

def generate_html(items: list[CalendarItem], tapings: list[TapingDate], title: str) -> str:
    today = datetime.now().strftime("%Y-%m-%d")
    items_sorted = sorted([i for i in items if i.date], key=lambda x: x.date)
    undated = [i for i in items if not i.date]
    
    broadcast = [i for i in items_sorted if i.category == "Broadcast/TV"]
    upcoming_count = sum(1 for i in broadcast if i.date >= today)

    payload = {
        "items": [asdict(i) for i in items_sorted + undated],
        "broadcast": [asdict(i) for i in broadcast],
        "tapings": [asdict(t) for t in tapings],
        "today": today,
        "upcomingCount": upcoming_count,
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
input:focus {{ outline: 2px solid var(--blue); border-color: transparent; }}
table {{ width:100%; border-collapse:collapse; font-size: 14px; }}
th, td {{ padding:14px; text-align:left; border-bottom:1px solid var(--line); vertical-align: top; }}
th {{ background:#f8fafc; font-size:11px; text-transform:uppercase; color:var(--muted); font-weight: 700; letter-spacing: 0.05em; }}
tr:hover td {{ background: #f8fafc; }}
.chip {{ display:inline-block; padding:2px 10px; border-radius:20px; font-size:11px; font-weight:700; background:#f1f5f9; color: var(--muted); }}
.chip.broadcast {{ background:#dcfce7; color:#166534; }}
.chip.dm {{ background:#dbeafe; color:#1e40af; }}
.link-btn {{ display: inline-flex; align-items: center; gap: 6px; background: var(--blue); color: #fff; padding: 6px 12px; border-radius: 6px; text-decoration: none; font-size: 12px; font-weight: 600; }}
.link-btn:hover {{ background: #0369a1; }}
.detail-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 8px; margin-top: 8px; }}
.detail-item {{ font-size: 11px; color: var(--muted); }}
.detail-label {{ font-weight: 700; color: var(--ink); text-transform: uppercase; font-size: 9px; display: block; }}
.taping-list {{ display:flex; flex-wrap:wrap; gap:10px; }}
.taping {{ background:#fff7ed; border:1px solid #fed7aa; color:#9a3412; padding:8px 16px; border-radius:12px; font-size:14px; font-weight:600; }}
.timeline {{ position: relative; padding: 20px 0; }}
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
    <div class="last-updated">Updated: {datetime.now().strftime("%b %d, %Y")}</div>
</header>
<main>
    <div class="tabs">
        <button class="tab active" onclick="showTab('calendar')">Marketing Calendar</button>
        <button class="tab" onclick="showTab('broadcast')">Broadcast Schedule</button>
    </div>

    <div id="calendar" class="panel active">
        <div class="card controls">
            <input type="text" id="search" placeholder="Search by title, notes, or job #..." oninput="render()">
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
            <td><span class="chip ${{i.category === 'Broadcast/TV' ? 'broadcast' : (i.category === 'DM' ? 'dm' : '')}}">${{i.category}}</span></td>
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
    parser.add_argument("--xlsx", type=Path, required=True)
    parser.add_argument("--output", type=Path, default=Path("index.html"))
    args = parser.parse_args()

    items, tapings = load_calendar_data(args.xlsx)
    html_content = generate_html(items, tapings, "Jewish Voice Marketing Calendar")
    args.output.write_text(html_content, encoding="utf-8")
    print(f"Generated {args.output} with {len(items)} items.")

if __name__ == "__main__":
    main()
