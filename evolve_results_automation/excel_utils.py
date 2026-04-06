import os
import re
import glob
import logging
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

from .config import COLUMNS, BASE_DIR, get_excel_file_for_year
from .parsing_utils import unique_row_hash

def initialize_excel(filepath: str):
    if not os.path.exists(filepath):
        logging.info(f"Excel file not found, creating {filepath}")
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(filepath, index=False)

def load_existing_results(filepath: str):
    if not os.path.exists(filepath):
        return pd.DataFrame()
    return pd.read_excel(filepath, dtype=str)

def save_results(filepath: str, df: pd.DataFrame):
    df = df.copy()
    date_cols_ddmmyyyy = ["Result Sent", "Certificate", "E-Certificate sent"]
    for col in date_cols_ddmmyyyy:
        if col in df.columns:
            df[col] = df[col].apply(format_ddmmyyyy)
    df.to_excel(filepath, index=False)
    autofilter_and_autofit(filepath)

def format_ddmmyyyy(val):
    if not val or pd.isna(val):
        return ""
    try:
        if isinstance(val, str) and len(val) == 10 and val[2] == '/' and val[5] == '/':
            return val
        dt = pd.to_datetime(val, errors='coerce')
        if pd.isna(dt):
            return val
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return val

def autofilter_and_autofit(filepath: str):
    wb = load_workbook(filepath)
    ws = wb.active
    ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                cell_len = len(str(cell.value)) if cell.value is not None else 0
                if cell_len > max_length:
                    max_length = cell_len
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = max(10, min(max_length + 2, 50))
    wb.save(filepath)
    wb.close()

def save_year_to_excel(year, rows_by_year, silent=False):
    """Save rows for a given year to Excel, merging with existing data and deduplicating."""
    if year not in rows_by_year:
        return
    year_rows = rows_by_year[year]
    excel_file = get_excel_file_for_year(year)
    initialize_excel(excel_file)

    # Load existing data (may contain rows from previous accounts)
    existing_df = load_existing_results(excel_file)

    # Create DataFrame from current rows (includes PDF links from processing)
    new_df = pd.DataFrame(year_rows)

    # Merge existing + new, then deduplicate
    if not existing_df.empty:
        combined = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        combined = new_df.copy()

    # Remove garbage rows missing core fields
    core_fields = ["Completed", "First name", "Last name"]
    for cf in core_fields:
        if cf in combined.columns:
            combined = combined[combined[cf].notna() & (combined[cf] != "")]

    # Deduplicate using unique_row_hash, keep last (in-memory rows have latest PDF links)
    combined['_hash'] = combined.apply(unique_row_hash, axis=1)
    combined = combined.drop_duplicates(subset='_hash', keep='last')
    combined = combined.drop(columns='_hash')

    # Enforce column order from config
    combined = combined.reindex(columns=COLUMNS)

    # Sort by completion date
    if "Completed" in combined.columns:
        try:
            combined["Completed_sort"] = pd.to_datetime(combined["Completed"], format="%d/%m/%Y", errors="coerce")
            combined = combined.sort_values(by="Completed_sort", ascending=True).drop(columns="Completed_sort")
        except Exception as e:
            logging.warning(f"Sorting failed for year {year}: {e}")

    save_results(excel_file, combined)
    if not silent:
        logging.info(f"Saved {len(combined)} rows to {year}/exam_results.xlsx")

def load_all_existing_data():
    """Load all existing year Excel files and return hashes + rows needing PDF download.
    
    Returns:
        tuple: (existing_hashes: set, rows_by_year: dict, pdf_resume_count: int)
    """
    existing_hashes = set()
    rows_by_year = {}
    pdf_resume_count = 0
    for excel_path in glob.glob(os.path.join(BASE_DIR, '*', 'exam_results.xlsx')):
        folder_name = os.path.basename(os.path.dirname(excel_path))
        if not re.match(r'^\d{4}$', folder_name):
            continue
        ef = load_existing_results(excel_path)
        if not ef.empty:
            for _, r in ef.iterrows():
                # Skip garbage rows missing core fields
                completed_val = str(r.get("Completed", "")).strip()
                if not completed_val or completed_val == "nan":
                    continue
                
                existing_hashes.add(unique_row_hash(r))
                # If row needs PDF, add to rows_by_year for processing
                report_dl = str(r.get("PDF report save time", "")).strip()
                if not report_dl or report_dl == "nan":
                    try:
                        yr = datetime.strptime(completed_val, "%d/%m/%Y").year
                    except (ValueError, TypeError):
                        continue  # Skip rows with unparseable dates
                    if yr not in rows_by_year:
                        rows_by_year[yr] = []
                    row_dict = {k: ("" if pd.isna(v) else v) for k, v in r.items()}
                    rows_by_year[yr].append(row_dict)
                    pdf_resume_count += 1
    logging.info(f"Loaded {len(existing_hashes)} existing results from Excel files")
    if pdf_resume_count > 0:
        logging.info(f"Found {pdf_resume_count} existing rows needing PDF download (resuming)")
    return existing_hashes, rows_by_year, pdf_resume_count