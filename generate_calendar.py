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
    if hasattr(cell, 'hyperlink') and cell.hyperlink:
        return cell.hyperlink.target
    val = str(cell.value) if cell.value else ""
    if val.startswith("http"): return val
    if "HYPERLINK" in val:
        m = re.search(r'HYPERLINK\("([^"]+)"', val)
        if m: return m.group(1)
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
        files = [args.data_dir / f for f in os.listdir(args.data_dir) if not f.startswith('~$')]
        xlsx = sorted([f for f in files if f.suffix == '.xlsx'], key=os.path.getmtime, reverse=True)
        csvs = sorted([f for f in files if f.suffix == '.csv'], key=os.path.getmtime, reverse=True)
        
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
    payload = json.dumps({{
        "items": [asdict(i) for i in all_items],
        "broadcast": [asdict(i) for i in all_items if i.category == "Broadcast/TV"],
        "tapings": [asdict(t) for t in unique_tapings],
        "updated": datetime.now().strftime("%b %d, %Y")
    }}, ensure_ascii=False)

    html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Marketing Calendar</title>
    <style>
        :root {{ --jv-blue: #1e3a8a; --blue: #0284c7; --bg: #f1f5f9; --ink: #1e293b; --muted: #64748b; }}
        body {{ font-family: 'Inter', sans-serif; margin: 0; background: var(--bg); color: var(--ink); scroll-behavior: smooth; }}
        header {{ background: var(--jv-blue); color: white; padding: 20px 40px; display: flex; justify-content: space-between; align-items: center; }}
        main {{ max-width: 1400px; margin: 0 auto; padding: 20px; }}
        .tabs {{ display: flex; gap: 10px; margin-bottom: 20px; }}
        .tab {{ padding: 10px 20px; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; background: #e2e8f0; color: var(--muted); }}
        .tab.active {{ background: white; color: var(--jv-blue); box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        .panel {{ display: none; }} .panel.active {{ display: block; }}
        .card {{ background: white; padding: 20px; border-radius: 12px; border: 1px solid #cbd5e1; margin-bottom: 20px; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
        th, td {{ padding: 12px; text-align: left; border-bottom: 1px solid #e2e8f0; vertical-align: top; }}
        th {{ background: #f8fafc; color: var(--muted); font-size: 11px; text-transform: uppercase; }}
        .link-btn {{ background: var(--blue); color: white; padding: 6px 12px; border-radius: 6px; text-decoration: none; font-size: 12px; font-weight: bold; display: inline-block; }}
        .detail-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 10px; margin-top: 10px; }}
        .detail-item {{ font-size: 11px; color: var(--muted); }}
        .detail-label {{ font-weight: bold; color: var(--ink); display: block; text-transform: uppercase; font-size: 9px; }}
        .taping-btn {{ background: #fff7ed; border: 1px solid #fed7aa; color: #9a3412; padding: 8px 15px; border-radius: 20px; cursor: pointer; font-weight: bold; margin: 5px; }}
        .timeline-item {{ border-left: 3px solid var(--blue); padding: 0 0 30px 20px; position: relative; margin-left: 10px; }}
        .timeline-item::before {{ content: ""; position: absolute; left: -9px; top: 0; width: 15px; height: 15px; background: white; border: 3px solid var(--blue); border-radius: 50%; }}
    </style>
</head>
<body>
    <header><h1>Marketing Calendar</h1><div style="font-size: 12px;">Last Updated: <span id="update-date"></span></div></header>
    <main>
        <div class="tabs">
            <button class="tab active" onclick="showTab('cal')">Calendar</button>
            <button class="tab" onclick="showTab('bc')">Broadcast Schedule</button>
        </div>
        <div id="cal" class="panel active">
            <div class="card"><input type="text" id="search" placeholder="Search..." oninput="render()" style="width:100%; padding:10px; border-radius:8px; border:1px solid #cbd5e1;"></div>
            <div class="card" style="padding:0; overflow-x:auto;">
                <table><thead><tr><th>Date</th><th>Details</th><th>Type</th><th>Producer/Audience</th><th>Link</th></tr></thead><tbody id="tbody"></tbody></table>
            </div>
        </div>
        <div id="bc" class="panel">
            <div class="card"><h3>Taping Sessions</h3><div id="tapings"></div></div>
            <div id="timeline"></div>
        </div>
    </main>
    <script>
        const DATA = {payload};
        document.getElementById('update-date').innerText = DATA.updated;
        function showTab(id) {{
            document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
            document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
            document.getElementById(id).classList.add('active');
            event.target.classList.add('active');
        }}
        function jump(date) {{
            showTab('bc');
            setTimeout(() => {{
                const el = document.querySelector(`[data-date="${{date}}"]`);
                if(el) el.scrollIntoView({{behavior:'smooth', block:'center'}});
            }}, 100);
        }}
        function render() {{
            const s = document.getElementById('search').value.toLowerCase();
            const filtered = DATA.items.filter(i => (i.title + i.notes + i.job_number + i.premium).toLowerCase().includes(s));
            document.getElementById('tbody').innerHTML = filtered.map(i => `
                <tr>
                    <td style="font-weight:bold; color:var(--jv-blue); white-space:nowrap;">${{i.date || 'Undated'}}</td>
                    <td>
                        <div style="font-weight:bold; font-size:15px;">${{i.title}}</div>
                        <div style="color:var(--muted); font-size:13px;">${{i.notes}}</div>
                        <div class="detail-grid">
                            ${{i.job_number ? `<div class="detail-item"><span class="detail-label">Job #</span>${{i.job_number}}</div>` : ''}}
                            ${{i.premium ? `<div class="detail-item"><span class="detail-label">Premium</span>${{i.premium}}</div>` : ''}}
                            ${{i.segment_code ? `<div class="detail-item"><span class="detail-label">Segment</span>${{i.segment_code}}</div>` : ''}}
                        </div>
                    </td>
                    <td><span style="font-size:11px; font-weight:bold; padding:2px 8px; border-radius:10px; background:#f1f5f9;">${{i.category}}</span></td>
                    <td><div style="font-size:16px; font-weight:bold;">${{i.owner}}</div><div style="font-size:13px; color:var(--muted);">${{i.audience}}</div></td>
                    <td>${{i.link ? `<a href="${{i.link}}" class="link-btn" target="_blank">🔗 Link</a>` : ''}}</td>
                </tr>
            `).join('');
        }}
        document.getElementById('tapings').innerHTML = DATA.tapings.map(t => `<button class="taping-btn" onclick="jump('${{t.date}}')">${{t.date}}</button>`).join('');
        document.getElementById('timeline').innerHTML = DATA.broadcast.map(i => `
            <div class="timeline-item" data-date="${{i.date}}">
                <div style="font-weight:bold; color:var(--blue);">${{i.date}}</div>
                <div class="card">
                    <div style="font-size:18px; font-weight:bold; color:var(--jv-blue);">${{i.title}}</div>
                    <div style="color:var(--muted); margin:10px 0;">${{i.notes}}</div>
                    <div class="detail-grid">
                        <div class="detail-item"><span class="detail-label">Producer</span>${{i.owner}}</div>
                        ${{i.premium ? `<div class="detail-item"><span class="detail-label">Premium</span>${{i.premium}}</div>` : ''}}
                    </div>
                    ${{i.link ? `<div style="margin-top:15px;"><a href="${{i.link}}" class="link-btn" target="_blank">🔗 Open Link</a></div>` : ''}}
                </div>
            </div>
        `).join('');
        render();
    </script>
</body>
</html>"""
    args.output.write_text(html_template, encoding='utf-8')

if __name__ == "__main__":
    main()
