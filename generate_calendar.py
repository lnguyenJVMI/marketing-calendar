import os
import pandas as pd
import json
from datetime import datetime
from openpyxl import load_workbook

DATA_DIR = ".github/datafiles"
MAPPING_FILE = "header_mapping.json"
TEMPLATE_FILE = "template.html"
OUTPUT_FILE = "index.html"

def load_mapping():
    if os.path.exists(MAPPING_FILE):
        with open(MAPPING_FILE, 'r') as f:
            return json.load(f)
    return {}

def get_all_files():
    files = []
    if not os.path.exists(DATA_DIR):
        return files
    for f in os.listdir(DATA_DIR):
        if f.endswith(('.xlsx', '.xls', '.csv')):
            files.append(os.path.join(DATA_DIR, f))
    return files

def extract_hyperlinks(file_path, sheet_name, mapping):
    links = {}
    if not file_path.endswith(('.xlsx', '.xls')):
        return links
    
    try:
        wb = load_workbook(file_path, data_only=True)
        if sheet_name not in wb.sheetnames:
            return links
        ws = wb[sheet_name]
        
        # Find column indices for columns we want links from
        headers = [cell.value for cell in ws[1]]
        target_cols = {}
        for field in ["Vanity Link", "Premium", "BBS Final Art"]:
            col_name = mapping.get(field, field)
            if col_name in headers:
                target_cols[field] = headers.index(col_name) + 1

        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            row_links = {}
            for field, col_idx in target_cols.items():
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.hyperlink:
                    row_links[field] = cell.hyperlink.target
            if row_links:
                links[row_idx] = row_links
    except Exception as e:
        print(f"Error extracting links from {file_path} [{sheet_name}]: {e}")
    return links

def process_file(file_path, mapping):
    all_events = []
    all_standing = []
    all_tapings = []

    try:
        if file_path.endswith(('.xlsx', '.xls')):
            xl = pd.ExcelFile(file_path)
            for sheet in xl.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet)
                links = extract_hyperlinks(file_path, sheet, mapping)
                
                # Sheet-specific logic for your known workbook
                if "Standing" in sheet or "Standard" in sheet:
                    all_standing.extend(df.to_dict('records'))
                elif "Taping" in sheet:
                    all_tapings.extend(df.to_dict('records'))
                else:
                    # Treat as dated events
                    events = df.to_dict('records')
                    for i, ev in enumerate(events):
                        # Add links if found (openpyxl row is 1-indexed, data starts at row 2)
                        if (i + 2) in links:
                            ev['_links'] = links[i + 2]
                    all_events.extend(events)
        else:
            df = pd.read_csv(file_path)
            all_events.extend(df.to_dict('records'))
    except Exception as e:
        print(f"Error processing {file_path}: {e}")
    
    return all_events, all_standing, all_tapings

def normalize_events(raw_events, mapping):
    normalized = []
    
    # Map standard fields to whatever the user chose
    date_col = mapping.get("Target Date", "Target Date")
    desc_col = mapping.get("Description", "Description")
    type_col = mapping.get("Type", "Type")
    prod_col = mapping.get("Producer", "Producer")
    aud_col  = mapping.get("Audience", "Audience")
    job_col  = mapping.get("BBS Job Number", "BBS Job Number")
    van_col  = mapping.get("Vanity Link", "Vanity Link")
    prem_col = mapping.get("Premium", "Premium")
    art_col  = mapping.get("BBS Final Art", "BBS Final Art")
    seg_col  = mapping.get("Segment Code", "Segment Code")

    for ev in raw_events:
        # Get date
        d_val = ev.get(date_col)
        if pd.isna(d_val): continue
        
        try:
            if isinstance(d_val, datetime):
                dt = d_val
            else:
                dt = pd.to_datetime(str(d_val))
            date_str = dt.strftime("%Y-%m-%d")
        except:
            continue

        links = ev.get('_links', {})
        
        normalized.append({
            "date": date_str,
            "title": str(ev.get(desc_col, "")).strip(),
            "type": str(ev.get(type_col, "")).strip(),
            "producer": str(ev.get(prod_col, "")).strip(),
            "audience": str(ev.get(aud_col, "")).strip(),
            "job": str(ev.get(job_col, "")).strip() if not pd.isna(ev.get(job_col)) else "",
            "vanity": str(ev.get(van_col, "")).strip(),
            "vanity_url": links.get("Vanity Link"),
            "premium": str(ev.get(prem_col, "")).strip(),
            "premium_url": links.get("Premium"),
            "premium_on": str(ev.get(prem_col, "")).lower() not in ["", "nan", "none", "—"],
            "art": str(ev.get(art_col, "")).strip(),
            "art_url": links.get("BBS Final Art"),
            "segment": str(ev.get(seg_col, "")).strip()
        })
    
    return sorted(normalized, key=lambda x: x['date'])

def main():
    mapping = load_mapping()
    files = get_all_files()
    
    all_raw_events = []
    all_standing = []
    all_tapings = []
    
    for f in files:
        print(f"Processing {f}...")
        e, s, t = process_file(f, mapping)
        all_raw_events.extend(e)
        all_standing.extend(s)
        all_tapings.extend(t)
        
    events = normalize_events(all_raw_events, mapping)
    
    # Simple normalization for standing/tapings (keep as is but handle NaNs)
    standing = []
    for s in all_standing:
        standing.append({k: (v if not pd.isna(v) else "") for k, v in s.items()})
        
    tapings = []
    for t in all_tapings:
        # Try to parse date from whatever column looks like a date
        date_val = ""
        for k, v in t.items():
            if "date" in k.lower() and not pd.isna(v):
                date_val = str(v)
                break
        tapings.append({
            "date": date_val,
            "label": next((str(v) for k, v in t.items() if "desc" in k.lower() or "label" in k.lower()), "Taping")
        })

    payload = {
        "events": events,
        "standing": standing,
        "tapings": tapings,
        "updated": datetime.now().strftime("%Y-%m-%d %H:%M")
    }

    if not os.path.exists(TEMPLATE_FILE):
        print(f"Template {TEMPLATE_FILE} not found!")
        return

    with open(TEMPLATE_FILE, 'r') as f:
        html = f.read()
    
    html = html.replace("__PAYLOAD_PLACEHOLDER__", json.dumps(payload))
    
    with open(OUTPUT_FILE, 'w') as f:
        f.write(html)
    
    print(f"Successfully generated {OUTPUT_FILE} with {len(events)} events.")

if __name__ == "__main__":
    main()
