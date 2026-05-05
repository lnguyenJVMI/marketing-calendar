import os
import pandas as pd
import json
from openpyxl import load_workbook

DATA_DIR = ".github/datafiles"
MAPPING_FILE = "header_mapping.json"

# Standard fields we want to map to
STANDARD_FIELDS = [
    "Target Date",
    "Description",
    "Type",
    "Producer",
    "Audience",
    "BBS Job Number",
    "Vanity Link",
    "Premium",
    "BBS Final Art",
    "Segment Code"
]

def get_all_files():
    files = []
    if not os.path.exists(DATA_DIR):
        return files
    for f in os.listdir(DATA_DIR):
        if f.endswith(('.xlsx', '.xls', '.csv')):
            files.append(os.path.join(DATA_DIR, f))
    return files

def get_headers(file_path):
    headers = []
    try:
        if file_path.endswith(('.xlsx', '.xls')):
            xl = pd.ExcelFile(file_path)
            for sheet in xl.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet, nrows=0)
                headers.append({"source": f"{os.path.basename(file_path)} [{sheet}]", "columns": df.columns.tolist()})
        else:
            df = pd.read_csv(file_path, nrows=0)
            headers.append({"source": os.path.basename(file_path), "columns": df.columns.tolist()})
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
    return headers

def main():
    print("--- Calendar Header Mapping Setup ---")
    files = get_all_files()
    if not files:
        print(f"No files found in {DATA_DIR}")
        return

    all_source_headers = []
    for f in files:
        all_source_headers.extend(get_headers(f))

    mapping = {}
    
    print("\nFor each standard field, select the corresponding column from your files.")
    print("Enter the number of the column, or press Enter to skip/leave as default.")

    for field in STANDARD_FIELDS:
        print(f"\nMapping for: {field}")
        options = []
        idx = 1
        for source in all_source_headers:
            for col in source['columns']:
                options.append((source['source'], col))
                print(f"  {idx}. {col} (from {source['source']})")
                idx += 1
        
        choice = input(f"Select column for '{field}' (1-{idx-1}) or skip: ")
        if choice.isdigit() and 1 <= int(choice) < idx:
            selected = options[int(choice)-1]
            mapping[field] = selected[1]
            print(f"Mapped '{field}' to '{selected[1]}'")
        else:
            print(f"Skipped '{field}'")

    with open(MAPPING_FILE, 'w') as f:
        json.dump(mapping, f, indent=2)
    
    print(f"\nMapping saved to {MAPPING_FILE}")
    print("You can now run generate_calendar.py to use these mappings.")

if __name__ == "__main__":
    main()
