import json, re, os, openpyxl
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
import pandas as pd

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

def clean_val(val):
    return str(val).strip() if val and not pd.isna(val) else ""

def parse_dt(val):
    if not val or pd.isna(val):
        return ""
    try:
        if isinstance(val, (datetime, pd.Timestamp)):
            return val.strftime("%Y-%m-%d")
        d = pd.to_datetime(str(val).strip(), errors='coerce')
        return d.strftime("%Y-%m-%d") if pd.notna(d) else str(val).strip()
    except:
        return str(val).strip()

def get_link(cell):
    if not cell:
        return ""
    if hasattr(cell, 'hyperlink') and cell.hyperlink and cell.hyperlink.target:
        return cell.hyperlink.target
    val = str(cell.value) if cell.value else ""
    if "HYPERLINK" in val.upper():
        m = re.search(r'HYPERLINK\("([^"]+)"\)', val)
        if m:
            return m.group(1)
    return val if val.startswith("http") else ""

def process_xlsx(path):
    items = []
    tapings = []
    try:
        wb = openpyxl.load_workbook(path, data_only=False)
        for sn in wb.sheetnames:
            ws = wb[sn]
            rows = list(ws.rows)
            if not rows:
                continue
            hdrs = [re.sub(r'[^a-z]', '', str(c.value).lower()) for c in rows[0]]

            def find_idx(targets):
                for t in targets:
                    nt = re.sub(r'[^a-z]', '', t.lower())
                    if nt in hdrs:
                        return hdrs.index(nt)
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
                        if d and len(d) == 10:
                            tapings.append(TapingDate("Taping", d))
                continue

            if idx['title'] is None:
                continue

            for r_idx, r in enumerate(rows[1:], 2):
                title = clean_val(r[idx['title']].value)
                if not title:
                    continue
                cat = clean_val(r[idx['type']].value) if idx['type'] is not None else "General"
                if any(x in (title + cat).lower() for x in ['broadcast', 'tv']):
                    cat = "Broadcast/TV"
                items.append(CalendarItem(
                    id="xl-{}-{}".format(sn, r_idx),
                    title=title,
                    date=parse_dt(r[idx['date']].value) if idx['date'] is not None else "",
                    source=path.name,
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
    except Exception as e:
        print("Error processing {}: {}".format(path, e))
    return items, tapings

def process_csv(path):
    items = []
    try:
        df = pd.read_csv(path)
        hdrs = [re.sub(r'[^a-z]', '', str(c).lower()) for c in df.columns]

        def find_idx(targets):
            for t in targets:
                nt = re.sub(r'[^a-z]', '', t.lower())
                if nt in hdrs:
                    return hdrs.index(nt)
            return None

        i_title = find_idx(['Description', 'Name of Communication', 'Title', 'Task Name'])
        i_date = find_idx(['Target Date', 'Air Date', 'Date'])
        i_type = find_idx(['Type', 'Category'])
        i_prod = find_idx(['Producer', 'Owner', 'Assignee'])
        i_aud = find_idx(['Audience'])
        i_job = find_idx(['BBS Job Number', 'Job Number'])
        i_prem = find_idx(['Premium'])

        if i_title is not None:
            for idx, row in df.iterrows():
                title = clean_val(row.iloc[i_title])
                if not title:
                    continue
                cat = clean_val(row.iloc[i_type]) if i_type is not None else "General"
                if any(x in (title + cat).lower() for x in ['broadcast', 'tv']):
                    cat = "Broadcast/TV"
                items.append(CalendarItem(
                    id="csv-{}".format(idx),
                    title=title,
                    date=parse_dt(row.iloc[i_date]) if i_date is not None else "",
                    source=path.name,
                    category=cat,
                    owner=clean_val(row.iloc[i_prod]) if i_prod is not None else "",
                    audience=clean_val(row.iloc[i_aud]) if i_aud is not None else "",
                    job_number=clean_val(row.iloc[i_job]) if i_job is not None else "",
                    premium=clean_val(row.iloc[i_prem]) if i_prem is not None else ""
                ))
    except Exception as e:
        print("Error processing {}: {}".format(path, e))
    return items


# ---- MAIN LOGIC ----
all_items = []
all_tapings = []

# Look in both current directory and .github/datafiles
search_paths = [Path("."), Path(".github/datafiles")]
for search_dir in search_paths:
    if search_dir.exists():
        for f in search_dir.iterdir():
            if f.name.startswith('~$'):
                continue
            if f.suffix == '.xlsx':
                it, tp = process_xlsx(f)
                all_items.extend(it)
                all_tapings.extend(tp)
            elif f.suffix == '.csv':
                all_items.extend(process_csv(f))

print("Found {} items and {} tapings".format(len(all_items), len(all_tapings)))

payload = json.dumps({
    "items": [asdict(i) for i in all_items],
    "broadcast": [asdict(i) for i in all_items if i.category == "Broadcast/TV"],
    "tapings": [asdict(t) for t in all_tapings],
    "updated": datetime.now().strftime("%b %d, %Y")
}, ensure_ascii=False)

# Read template and inject payload
script_dir = Path(__file__).parent
template_path = script_dir / "template.html"
with open(template_path, "r") as f:
    html_template = f.read()

# Replace the placeholder with actual data
output_html = html_template.replace("__PAYLOAD_PLACEHOLDER__", payload)

with open(script_dir / "index.html", "w") as f:
    f.write(output_html)

print("Generated index.html successfully!")
