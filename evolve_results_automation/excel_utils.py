import os
import re
import glob
import logging
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

from .config import COLUMNS, BASE_DIR, get_excel_file_for_year
from .parsing_utils import unique_row_hash


def _normalize(val):
    """Normalize a cell value to a clean string (empty string for None/NaN)."""
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() == "nan" else s


def initialize_excel(filepath: str):
    if not os.path.exists(filepath):
        logging.info(f"Excel file not found, creating {filepath}")
        wb = Workbook()
        ws = wb.active
        ws.append(COLUMNS)
        wb.save(filepath)
        wb.close()


def load_existing_results(filepath: str):
    """Load rows from an Excel file as a list of dicts with string values."""
    if not os.path.exists(filepath):
        return []
    wb = load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)

    # First row is the header
    try:
        headers = [str(h) if h is not None else "" for h in next(rows_iter)]
    except StopIteration:
        wb.close()
        return []

    result = []
    for row_values in rows_iter:
        row_dict = {headers[i]: _normalize(row_values[i]) for i in range(len(headers))}
        result.append(row_dict)

    wb.close()
    return result


def save_results(filepath: str, rows: list):
    """Write rows (list of dicts) to an Excel file with formatting applied in one save."""
    date_cols_ddmmyyyy = ["Result Sent", "Certificate", "E-Certificate sent"]
    wb = Workbook()
    ws = wb.active
    ws.append(COLUMNS)
    for row in rows:
        values = []
        for col in COLUMNS:
            val = row.get(col, "")
            if col in date_cols_ddmmyyyy:
                val = format_ddmmyyyy(val)
            values.append(val)
        ws.append(values)

    # Apply formatting in-memory before saving (single I/O operation)
    ws.auto_filter.ref = ws.dimensions
    # Freeze top row so header stays visible when scrolling
    ws.freeze_panes = "A2"
    # Style header row (City & Guilds red with white bold text)
    header_font = Font(bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='E30613', end_color='E30613', fill_type='solid')
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
    # Auto-fit column widths
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


def format_ddmmyyyy(val):
    if not val:
        return ""
    s = str(val).strip()
    if not s:
        return ""
    try:
        if len(s) == 10 and s[2] == '/' and s[5] == '/':
            return s
        dt = datetime.strptime(s, "%Y-%m-%d %H:%M:%S")
        return dt.strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return s


def save_year_to_excel(year, rows_by_year, silent=False):
    """Save rows for a given year to Excel, merging with existing data and deduplicating."""
    if year not in rows_by_year:
        return
    year_rows = rows_by_year[year]
    excel_file = get_excel_file_for_year(year)
    initialize_excel(excel_file)

    # Load existing data (may contain rows from previous accounts)
    existing = load_existing_results(excel_file)

    # Merge existing + new
    combined = existing + year_rows

    # Remove garbage rows missing core fields
    core_fields = ["Completed", "First name", "Last name"]
    combined = [r for r in combined
                if all(r.get(cf, "").strip() for cf in core_fields)]

    # Deduplicate using unique_row_hash, keep last (in-memory rows have latest PDF links)
    seen = {}
    for row in combined:
        seen[unique_row_hash(row)] = row
    combined = list(seen.values())

    # Enforce column order from config (fill missing columns with empty string)
    combined = [{col: row.get(col, "") for col in COLUMNS} for row in combined]

    # Sort by completion date
    def sort_key(row):
        try:
            return datetime.strptime(row.get("Completed", ""), "%d/%m/%Y")
        except (ValueError, TypeError):
            return datetime.max
    try:
        combined.sort(key=sort_key)
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
        rows = load_existing_results(excel_path)
        for r in rows:
            # Skip garbage rows missing core fields
            completed_val = r.get("Completed", "").strip()
            if not completed_val:
                continue

            existing_hashes.add(unique_row_hash(r))
            # If row needs PDF, add to rows_by_year for processing
            report_dl = r.get("PDF report save time", "").strip()
            if not report_dl:
                try:
                    yr = datetime.strptime(completed_val, "%d/%m/%Y").year
                except (ValueError, TypeError):
                    continue  # Skip rows with unparseable dates
                if yr not in rows_by_year:
                    rows_by_year[yr] = []
                rows_by_year[yr].append(r)
                pdf_resume_count += 1
    logging.info(f"Loaded {len(existing_hashes)} existing results from Excel files")
    if pdf_resume_count > 0:
        logging.info(f"Found {pdf_resume_count} existing rows needing PDF download (resuming)")
    return existing_hashes, rows_by_year, pdf_resume_count