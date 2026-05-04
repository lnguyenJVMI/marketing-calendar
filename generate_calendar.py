#!/usr/bin/env python3
"""
Free / low-cost standalone marketing calendar generator.

This script produces a single HTML file with two tabs:
1. Marketing Calendar: all dated Asana/communications items, with filters and a task table.
2. Broadcast Schedule: TV/broadcast items grouped by month, plus optional taping dates.

Expected inputs are intentionally flexible:
- Asana CSV export: columns such as Name, Task Name, Due Date, Assignee, Section, Project, Notes, Completed.
- Communications Excel workbook: one or more sheets with Date/Air Date/Start Date and Title/Name/Description columns.
- Optional Tapings sheet: date ranges or start/end dates are highlighted in the Broadcast tab.

Usage:
    python3 generate_calendar.py \
      --asana-csv asana_tasks.csv \
      --communications-xlsx communications_calendar.xlsx \
      --output marketing-calendar-standalone.html
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


DATE_CANDIDATES = [
    "due date", "date", "air date", "start date", "publish date", "send date",
    "event date", "scheduled date", "deadline"
]
TITLE_CANDIDATES = [
    "name", "task name", "title", "subject", "episode", "show", "description", "item"
]
TYPE_CANDIDATES = ["type", "category", "channel", "section", "project", "medium"]
OWNER_CANDIDATES = ["assignee", "owner", "lead", "responsible", "assigned to"]
STATUS_CANDIDATES = ["status", "completed", "progress"]
NOTES_CANDIDATES = ["notes", "description", "details", "summary"]

BROADCAST_KEYWORDS = [
    "broadcast", "tv", "television", "show", "episode", "air", "aired", "taping", "tapings"
]


@dataclass
class CalendarItem:
    title: str
    date: str
    source: str
    category: str = "General"
    owner: str = ""
    status: str = ""
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


def classify_category(row: pd.Series, explicit: str, source: str, sheet: str) -> str:
    combined = " ".join(normalized(v).lower() for v in row.values)
    if explicit:
        return explicit
    if any(keyword in combined for keyword in BROADCAST_KEYWORDS) or any(keyword in sheet.lower() for keyword in BROADCAST_KEYWORDS):
        return "Broadcast/TV"
    if "email" in combined:
        return "Email"
    if "social" in combined:
        return "Social"
    if "asana" in source.lower():
        return "Asana Task"
    return "Communications"


def load_tabular_file(path: Path, source: str, sheet: str = "") -> list[CalendarItem]:
    if not path.exists():
        return []

    if path.suffix.lower() in {".xlsx", ".xls"}:
        frames = pd.read_excel(path, sheet_name=None)
    else:
        frames = {sheet or path.stem: pd.read_csv(path)}

    items: list[CalendarItem] = []
    for sheet_name, df in frames.items():
        if df.empty:
            continue
        df = df.dropna(how="all")
        date_col = find_column(df.columns, DATE_CANDIDATES)
        title_col = find_column(df.columns, TITLE_CANDIDATES)
        type_col = find_column(df.columns, TYPE_CANDIDATES)
        owner_col = find_column(df.columns, OWNER_CANDIDATES)
        status_col = find_column(df.columns, STATUS_CANDIDATES)
        notes_col = find_column(df.columns, NOTES_CANDIDATES)

        if not date_col or not title_col:
            continue

        for _, row in df.iterrows():
            date = parse_date(row.get(date_col))
            title = normalized(row.get(title_col))
            if not date or not title:
                continue
            explicit_category = normalized(row.get(type_col)) if type_col else ""
            category = classify_category(row, explicit_category, source, str(sheet_name))
            items.append(CalendarItem(
                title=title,
                date=date,
                source=source,
                category=category,
                owner=normalized(row.get(owner_col)) if owner_col else "",
                status=normalized(row.get(status_col)) if status_col else "",
                notes=normalized(row.get(notes_col)) if notes_col else "",
                sheet=str(sheet_name),
            ))
    return items


def extract_tapings(path: Path) -> list[TapingDate]:
    if not path.exists() or path.suffix.lower() not in {".xlsx", ".xls"}:
        return []
    tapings: list[TapingDate] = []
    frames = pd.read_excel(path, sheet_name=None)
    for sheet_name, df in frames.items():
        if "taping" not in str(sheet_name).lower():
            continue
        date_col = find_column(df.columns, DATE_CANDIDATES + ["range", "dates"])
        end_col = find_column(df.columns, ["end date", "finish date"])
        title_col = find_column(df.columns, TITLE_CANDIDATES)
        notes_col = find_column(df.columns, NOTES_CANDIDATES)
        if not date_col:
            continue
        for _, row in df.dropna(how="all").iterrows():
            start = parse_date(row.get(date_col))
            end = parse_date(row.get(end_col)) if end_col else ""
            if not start:
                continue
            title = normalized(row.get(title_col)) if title_col else "Taping"
            label = f"{start} – {end}" if end and end != start else start
            tapings.append(TapingDate(title=title or "Taping", date=label, notes=normalized(row.get(notes_col)) if notes_col else ""))
    return tapings


def generate_html(items: list[CalendarItem], tapings: list[TapingDate], title: str) -> str:
    today = datetime.now().strftime("%Y-%m-%d")
    items_sorted = sorted(items, key=lambda item: item.date)
    broadcast = [item for item in items_sorted if item.category.lower() in {"broadcast/tv", "broadcast", "tv"} or any(k in (item.title + " " + item.notes + " " + item.sheet).lower() for k in BROADCAST_KEYWORDS)]
    upcoming_count = sum(1 for item in broadcast if item.date >= today)

    payload = {
        "items": [asdict(item) for item in items_sorted],
        "broadcast": [asdict(item) for item in broadcast],
        "tapings": [asdict(taping) for taping in tapings],
        "today": today,
        "upcomingCount": upcoming_count,
    }
    data = json.dumps(payload, ensure_ascii=False)
    safe_title = html.escape(title)

    return f"""<!doctype html>
<html lang=\"en\">
<head>
<meta charset=\"utf-8\">
<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">
<title>{safe_title}</title>
<style>
:root {{ --bg:#f6f7fb; --panel:#ffffff; --ink:#172033; --muted:#667085; --line:#e5e7eb; --blue:#2563eb; --green:#16a34a; --amber:#d97706; --purple:#7c3aed; }}
* {{ box-sizing: border-box; }}
body {{ margin:0; background:var(--bg); color:var(--ink); font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }}
header {{ padding: 32px 28px 18px; background: linear-gradient(135deg,#0f172a,#1e40af); color:#fff; }}
header h1 {{ margin:0 0 8px; font-size: clamp(28px, 4vw, 44px); letter-spacing:-.03em; }}
header p {{ margin:0; color:#dbeafe; max-width: 900px; }}
main {{ max-width: 1180px; margin: -8px auto 48px; padding: 0 18px; }}
.tabs {{ display:flex; gap:10px; margin: 20px 0; flex-wrap:wrap; }}
.tab {{ border:1px solid var(--line); background:var(--panel); color:var(--ink); padding:12px 16px; border-radius:999px; cursor:pointer; font-weight:700; box-shadow:0 2px 8px rgba(15,23,42,.04); }}
.tab.active {{ background:var(--blue); color:#fff; border-color:var(--blue); }}
.badge {{ display:inline-block; margin-left:8px; padding:2px 8px; border-radius:999px; background:#eff6ff; color:#1d4ed8; font-size:12px; }}
.tab.active .badge {{ background:#dbeafe; color:#1e3a8a; }}
.panel {{ display:none; }}
.panel.active {{ display:block; }}
.card {{ background:var(--panel); border:1px solid var(--line); border-radius:18px; padding:18px; box-shadow: 0 10px 28px rgba(15,23,42,.06); margin-bottom:18px; }}
.controls {{ display:grid; grid-template-columns: repeat(4, minmax(160px,1fr)); gap:12px; }}
input, select {{ width:100%; padding:10px 12px; border:1px solid var(--line); border-radius:10px; background:#fff; color:var(--ink); }}
.stats {{ display:grid; grid-template-columns: repeat(4, minmax(140px,1fr)); gap:12px; margin-bottom:18px; }}
.stat {{ background:var(--panel); border:1px solid var(--line); border-radius:16px; padding:16px; }}
.stat strong {{ display:block; font-size:28px; }}
.table-wrap {{ overflow:auto; }}
table {{ width:100%; border-collapse: collapse; min-width: 840px; }}
th, td {{ padding:12px; text-align:left; border-bottom:1px solid var(--line); vertical-align:top; }}
th {{ color:var(--muted); font-size:12px; text-transform:uppercase; letter-spacing:.04em; background:#f8fafc; }}
.chip {{ display:inline-flex; align-items:center; border-radius:999px; padding:4px 9px; font-size:12px; font-weight:700; background:#eef2ff; color:#3730a3; }}
.chip.broadcast {{ background:#fef3c7; color:#92400e; }}
.chip.asana {{ background:#dcfce7; color:#166534; }}
.month {{ margin: 26px 0 12px; font-size:22px; }}
.timeline {{ position:relative; padding-left:24px; }}
.timeline:before {{ content:""; position:absolute; left:7px; top:0; bottom:0; width:2px; background:var(--line); }}
.event {{ position:relative; padding:14px 16px; background:#fff; border:1px solid var(--line); border-radius:16px; margin-bottom:12px; }}
.event:before {{ content:""; position:absolute; left:-22px; top:20px; width:12px; height:12px; border-radius:50%; background:var(--blue); box-shadow:0 0 0 4px #dbeafe; }}
.event.upcoming:before {{ background:var(--green); box-shadow:0 0 0 4px #dcfce7; }}
.event.today {{ border-color:var(--blue); box-shadow:0 0 0 3px #dbeafe; }}
.event h3 {{ margin:0 0 6px; }}
.event p {{ margin:6px 0 0; color:var(--muted); }}
.status {{ font-size:12px; padding:3px 8px; border-radius:999px; font-weight:800; background:#e5e7eb; color:#374151; }}
.status.upcoming {{ background:#dcfce7; color:#166534; }}
.status.today {{ background:#dbeafe; color:#1d4ed8; }}
.cards-grid {{ display:grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap:14px; }}
.toggle {{ display:flex; gap:8px; margin-bottom:16px; }}
.toggle button {{ border:1px solid var(--line); background:#fff; padding:9px 12px; border-radius:10px; cursor:pointer; font-weight:700; }}
.toggle button.active {{ background:#111827; color:#fff; }}
.taping-list {{ display:flex; flex-wrap:wrap; gap:10px; }}
.taping {{ border:1px solid #f59e0b; background:#fffbeb; color:#92400e; padding:10px 12px; border-radius:999px; font-weight:700; }}
.empty {{ color:var(--muted); text-align:center; padding:38px; }}
@media (max-width: 760px) {{ .controls, .stats {{ grid-template-columns:1fr; }} header {{ padding:24px 18px 14px; }} }}
</style>
</head>
<body>
<header>
  <h1>{safe_title}</h1>

</header>
<main>
  <nav class=\"tabs\" aria-label=\"Calendar tabs\">
    <button class=\"tab active\" data-tab=\"calendar\">Marketing Calendar <span class=\"badge\" id=\"allCount\">0</span></button>
    <button class=\"tab\" data-tab=\"broadcast\">Broadcast Schedule <span class=\"badge\" id=\"broadcastBadge\">0 upcoming</span></button>
  </nav>

  <section id=\"calendar\" class=\"panel active\">
    <div class=\"stats\" id=\"stats\"></div>
    <div class=\"card controls\">
      <input id=\"search\" placeholder=\"Search title, notes, owner...\" />
      <select id=\"categoryFilter\"><option value=\"\">All categories</option></select>
      <select id=\"sourceFilter\"><option value=\"\">All sources</option></select>
      <select id=\"monthFilter\"><option value=\"\">All months</option></select>
    </div>
    <div class=\"card table-wrap\">
      <table>
        <thead><tr><th>Date</th><th>Title</th><th>Category</th><th>Source</th><th>Owner</th><th>Status</th><th>Notes</th></tr></thead>
        <tbody id=\"itemRows\"></tbody>
      </table>
    </div>
  </section>

  <section id=\"broadcast\" class=\"panel\">
    <div class=\"card\">
      <div class=\"toggle\"><button id=\"timelineBtn\" class=\"active\">Timeline</button><button id=\"cardsBtn\">Cards</button></div>
      <div id=\"broadcastContent\"></div>
    </div>
    <div class=\"card\">
      <h2>Taping Dates</h2>
      <div id=\"tapings\" class=\"taping-list\"></div>
    </div>
  </section>
</main>
<script>
const DATA = {data};
const esc = value => String(value ?? '').replace(/[&<>\"]/g, c => ({{'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}}[c]));
const fmtDate = d => {{ const parsed = new Date(d + 'T00:00:00'); return isNaN(parsed) ? d : parsed.toLocaleDateString(undefined, {{month:'short', day:'numeric', year:'numeric'}}); }};
const monthLabel = d => {{ const parsed = new Date(d + 'T00:00:00'); return isNaN(parsed) ? 'Undated' : parsed.toLocaleDateString(undefined, {{month:'long', year:'numeric'}}); }};
const unique = arr => [...new Set(arr.filter(Boolean))].sort();
let broadcastMode = 'timeline';

function chipClass(value) {{
  const lower = String(value).toLowerCase();
  if (lower.includes('broadcast') || lower.includes('tv')) return 'broadcast';
  if (lower.includes('asana')) return 'asana';
  return '';
}}
function populateFilters() {{
  document.getElementById('allCount').textContent = DATA.items.length;
  document.getElementById('broadcastBadge').textContent = `${{DATA.upcomingCount}} upcoming`;
  const categories = unique(DATA.items.map(i => i.category));
  const sources = unique(DATA.items.map(i => i.source));
  const months = unique(DATA.items.map(i => monthLabel(i.date)));
  document.getElementById('categoryFilter').innerHTML = '<option value="">All categories</option>' + categories.map(v => `<option>${{esc(v)}}</option>`).join('');
  document.getElementById('sourceFilter').innerHTML = '<option value="">All sources</option>' + sources.map(v => `<option>${{esc(v)}}</option>`).join('');
  document.getElementById('monthFilter').innerHTML = '<option value="">All months</option>' + months.map(v => `<option>${{esc(v)}}</option>`).join('');
}}
function renderStats(items) {{
  const upcoming = items.filter(i => i.date >= DATA.today).length;
  const broadcast = items.filter(i => i.category.toLowerCase().includes('broadcast') || i.category.toLowerCase().includes('tv')).length;
  const sources = unique(items.map(i => i.source)).length;
  document.getElementById('stats').innerHTML = [
    ['Total items', items.length], ['Upcoming', upcoming], ['Broadcast/TV', broadcast], ['Sources', sources]
  ].map(([label, value]) => `<div class="stat"><strong>${{value}}</strong><span>${{label}}</span></div>`).join('');
}}
function filteredItems() {{
  const q = document.getElementById('search').value.toLowerCase();
  const category = document.getElementById('categoryFilter').value;
  const source = document.getElementById('sourceFilter').value;
  const month = document.getElementById('monthFilter').value;
  return DATA.items.filter(i => {{
    const haystack = `${{i.title}} ${{i.notes}} ${{i.owner}} ${{i.status}}`.toLowerCase();
    return (!q || haystack.includes(q)) && (!category || i.category === category) && (!source || i.source === source) && (!month || monthLabel(i.date) === month);
  }});
}}
function renderCalendar() {{
  const items = filteredItems();
  renderStats(items);
  document.getElementById('itemRows').innerHTML = items.length ? items.map(i => `<tr><td>${{fmtDate(i.date)}}</td><td><strong>${{esc(i.title)}}</strong></td><td><span class="chip ${{chipClass(i.category)}}">${{esc(i.category)}}</span></td><td>${{esc(i.source)}}</td><td>${{esc(i.owner)}}</td><td>${{esc(i.status)}}</td><td>${{esc(i.notes)}}</td></tr>`).join('') : `<tr><td colspan="7" class="empty">No items match the current filters.</td></tr>`;
}}
function eventCard(i) {{
  const isToday = i.date === DATA.today;
  const isUpcoming = i.date >= DATA.today;
  const status = isToday ? 'Today' : (isUpcoming ? 'Upcoming' : 'Aired');
  return `<article class="event ${{isUpcoming ? 'upcoming' : ''}} ${{isToday ? 'today' : ''}}"><h3>${{esc(i.title)}}</h3><span class="status ${{isToday ? 'today' : (isUpcoming ? 'upcoming' : '')}}">${{status}}</span> <span>${{fmtDate(i.date)}}</span><p>${{esc(i.notes || i.owner || i.sheet || '')}}</p></article>`;
}}
function renderBroadcast() {{
  const root = document.getElementById('broadcastContent');
  if (!DATA.broadcast.length) {{ root.innerHTML = '<p class="empty">No broadcast items were detected. Add words like TV, broadcast, show, episode, or air date to your source data, or set the Type/Category column to Broadcast/TV.</p>'; return; }}
  if (broadcastMode === 'cards') {{
    root.innerHTML = `<div class="cards-grid">${{DATA.broadcast.map(eventCard).join('')}}</div>`;
  }} else {{
    const grouped = DATA.broadcast.reduce((acc, item) => {{ const m = monthLabel(item.date); (acc[m] ||= []).push(item); return acc; }}, {{}});
    root.innerHTML = Object.entries(grouped).map(([month, items]) => `<h2 class="month">${{esc(month)}}</h2><div class="timeline">${{items.map(eventCard).join('')}}</div>`).join('');
  }}
  document.getElementById('tapings').innerHTML = DATA.tapings.length ? DATA.tapings.map(t => `<span class="taping">${{esc(t.title)}}: ${{esc(t.date)}}</span>`).join('') : '<span class="empty">No taping sheet/date ranges detected.</span>';
}}
document.querySelectorAll('.tab').forEach(tab => tab.addEventListener('click', () => {{
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  tab.classList.add('active'); document.getElementById(tab.dataset.tab).classList.add('active');
}}));
['search','categoryFilter','sourceFilter','monthFilter'].forEach(id => document.getElementById(id).addEventListener('input', renderCalendar));
document.getElementById('timelineBtn').addEventListener('click', () => {{ broadcastMode='timeline'; document.getElementById('timelineBtn').classList.add('active'); document.getElementById('cardsBtn').classList.remove('active'); renderBroadcast(); }});
document.getElementById('cardsBtn').addEventListener('click', () => {{ broadcastMode='cards'; document.getElementById('cardsBtn').classList.add('active'); document.getElementById('timelineBtn').classList.remove('active'); renderBroadcast(); }});
populateFilters(); renderCalendar(); renderBroadcast();
</script>
</body>
</html>"""


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate a standalone tabbed marketing calendar HTML file.")
    parser.add_argument("--asana-csv", type=Path, help="Optional Asana CSV export path.")
    parser.add_argument("--communications-xlsx", type=Path, help="Optional communications calendar Excel workbook path.")
    parser.add_argument("--output", type=Path, default=Path("marketing-calendar-standalone.html"), help="Output HTML path.")
    parser.add_argument("--title", default="Marketing Calendar", help="Calendar page title.")
    args = parser.parse_args()

    items: list[CalendarItem] = []
    if args.asana_csv:
        items.extend(load_tabular_file(args.asana_csv, "Asana", args.asana_csv.stem))
    if args.communications_xlsx:
        items.extend(load_tabular_file(args.communications_xlsx, "Communications Calendar"))
        tapings = extract_tapings(args.communications_xlsx)
    else:
        tapings = []

    if not items:
        raise SystemExit("No dated calendar items were found. Check that your files include date and title/name columns.")

    args.output.write_text(generate_html(items, tapings, args.title), encoding="utf-8")
    print(f"Wrote {args.output} with {len(items)} items and {len(tapings)} taping dates.")


if __name__ == "__main__":
    main()
