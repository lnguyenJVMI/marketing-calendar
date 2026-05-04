#!/usr/bin/env python3
"""
Updated Marketing Calendar Generator for Jewish Voice.
Includes support for Vanity Links, Producers, Audiences, and specific sheet structures.
"""

from __future__ import annotations

import argparse
import html
import json
import re
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional

import pandas as pd

# Column detection candidates
DATE_CANDIDATES = ["target date", "air date", "date", "start date", "publish date"]
TITLE_CANDIDATES = ["description", "name of communication", "title", "task name"]
TYPE_CANDIDATES = ["type", "category", "channel"]
OWNER_CANDIDATES = ["producer", "assignee", "owner", "lead"]
LINK_CANDIDATES = ["vanity link", "link", "url"]
AUDIENCE_CANDIDATES = ["audience"]
NOTES_CANDIDATES = ["notes", "subject"]

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
    sheet: str = ""

@dataclass
class TapingDate:
    title: str
    date: str
    notes: str = ""

def normalized(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()

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
            # Handle the specific "Tapings" sheet structure seen in the sample
            # It seems to have dates as values in the first column
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

        if not title_col:
            continue

        for _, row in df.iterrows():
            title = normalized(row.get(title_col))
            if not title:
                continue
                
            date = parse_date(row.get(date_col)) if date_col else ""
            category = normalized(row.get(type_col)) if type_col else "General"
            
            # Auto-classify as Broadcast if keywords found
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
                sheet=sheet_name
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
:root {{ --bg:#f8fafc; --panel:#ffffff; --ink:#0f172a; --muted:#64748b; --line:#e2e8f0; --blue:#2563eb; --green:#10b981; --amber:#f59e0b; }}
body {{ margin:0; background:var(--bg); color:var(--ink); font-family: system-ui, -apple-system, sans-serif; }}
header {{ padding: 24px; background: #1e293b; color:#fff; }}
header h1 {{ margin:0; font-size: 24px; }}
main {{ max-width: 1200px; margin: 0 auto; padding: 20px; }}
.tabs {{ display:flex; gap:8px; margin-bottom: 20px; }}
.tab {{ padding:10px 20px; border-radius:8px; border:1px solid var(--line); background:#fff; cursor:pointer; font-weight:600; }}
.tab.active {{ background:var(--blue); color:#fff; border-color:var(--blue); }}
.panel {{ display:none; }}
.panel.active {{ display:block; }}
.card {{ background:#fff; border:1px solid var(--line); border-radius:12px; padding:20px; margin-bottom:20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
.controls {{ display:grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap:12px; margin-bottom:20px; }}
input, select {{ padding:8px 12px; border:1px solid var(--line); border-radius:6px; width:100%; box-sizing:border-box; }}
table {{ width:100%; border-collapse:collapse; }}
th, td {{ padding:12px; text-align:left; border-bottom:1px solid var(--line); }}
th {{ background:#f1f5f9; font-size:12px; text-transform:uppercase; color:var(--muted); }}
.chip {{ display:inline-block; padding:2px 8px; border-radius:12px; font-size:11px; font-weight:700; background:#f1f5f9; }}
.chip.broadcast {{ background:#fef3c7; color:#92400e; }}
.link {{ color:var(--blue); text-decoration:none; font-weight:500; }}
.link:hover {{ text-decoration:underline; }}
.taping-list {{ display:flex; flex-wrap:wrap; gap:8px; }}
.taping {{ background:#fffbeb; border:1px solid #f59e0b; color:#92400e; padding:6px 12px; border-radius:20px; font-size:13px; font-weight:600; }}
.timeline-item {{ border-left: 2px solid var(--line); padding-left: 20px; position:relative; margin-bottom:20px; }}
.timeline-item::before {{ content:""; position:absolute; left:-7px; top:0; width:12px; height:12px; border-radius:50%; background:var(--blue); }}
.date-label {{ font-weight:700; color:var(--muted); font-size:13px; margin-bottom:4px; }}
.item-title {{ font-size:16px; font-weight:600; margin-bottom:4px; }}
.item-meta {{ font-size:13px; color:var(--muted); }}
</style>
</head>
<body>
<header><h1>{html.escape(title)}</h1></header>
<main>
    <div class="tabs">
        <button class="tab active" onclick="showTab('calendar')">Marketing Calendar</button>
        <button class="tab" onclick="showTab('broadcast')">Broadcast Schedule</button>
    </div>

    <div id="calendar" class="panel active">
        <div class="card controls">
            <input type="text" id="search" placeholder="Search communications..." oninput="render()">
            <select id="typeFilter" onchange="render()"><option value="">All Types</option></select>
            <select id="producerFilter" onchange="render()"><option value="">All Producers</option></select>
        </div>
        <div class="card" style="overflow-x:auto">
            <table id="dataTable">
                <thead>
                    <tr>
                        <th>Date</th>
                        <th>Description</th>
                        <th>Type</th>
                        <th>Producer</th>
                        <th>Audience</th>
                        <th>Link / Details</th>
                    </tr>
                </thead>
                <tbody id="tableBody"></tbody>
            </table>
        </div>
    </div>

    <div id="broadcast" class="panel">
        <div class="card">
            <h2>Taping Sessions</h2>
            <div id="tapingList" class="taping-list"></div>
        </div>
        <div class="card">
            <h2>Broadcast Timeline</h2>
            <div id="timeline"></div>
        </div>
    </div>
</main>

<script>
const DATA = {data_json};

function showTab(id) {{
    document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.getElementById(id).classList.add('active');
    event.target.classList.add('active');
}}

function render() {{
    const search = document.getElementById('search').value.toLowerCase();
    const type = document.getElementById('typeFilter').value;
    const producer = document.getElementById('producerFilter').value;

    const filtered = DATA.items.filter(i => {{
        const matchSearch = i.title.toLowerCase().includes(search) || i.notes.toLowerCase().includes(search);
        const matchType = !type || i.category === type;
        const matchProducer = !producer || i.owner === producer;
        return matchSearch && matchType && matchProducer;
    }});

    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = filtered.map(i => `
        <tr>
            <td style="white-space:nowrap">${{i.date || 'Undated'}}</td>
            <td><strong>${{i.title}}</strong><br><small style="color:#64748b">${{i.notes}}</small></td>
            <td><span class="chip ${{i.category === 'Broadcast/TV' ? 'broadcast' : ''}}">${{i.category}}</span></td>
            <td>${{i.owner}}</td>
            <td>${{i.audience}}</td>
            <td>${{i.link ? `<a href="${{i.link}}" class="link" target="_blank">View Link</a>` : i.sheet}}</td>
        </tr>
    `).join('');
}}

function init() {{
    // Populate filters
    const types = [...new Set(DATA.items.map(i => i.category))].sort();
    const producers = [...new Set(DATA.items.map(i => i.owner))].sort();
    
    document.getElementById('typeFilter').innerHTML += types.map(t => `<option value="${{t}}">${{t}}</option>`).join('');
    document.getElementById('producerFilter').innerHTML += producers.map(p => `<option value="${{p}}">${{p}}</option>`).join('');

    // Render Tapings
    document.getElementById('tapingList').innerHTML = DATA.tapings.map(t => `
        <div class="taping">${{t.date}}</div>
    `).join('');

    // Render Timeline
    document.getElementById('timeline').innerHTML = DATA.broadcast.map(i => `
        <div class="timeline-item">
            <div class="date-label">${{i.date}}</div>
            <div class="item-title">${{i.title}}</div>
            <div class="item-meta">${{i.owner}} • ${{i.audience}}</div>
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
