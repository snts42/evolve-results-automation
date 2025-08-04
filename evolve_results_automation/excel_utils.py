import os
import pandas as pd
from openpyxl import load_workbook
from .logging_utils import log

def initialize_excel(filepath: str, columns):
    if not os.path.exists(filepath):
        log(f"Excel file not found. Creating {filepath}")
        df = pd.DataFrame(columns=columns)
        df.to_excel(filepath, index=False)
        autofilter_and_autofit(filepath)

def load_existing_results(filepath: str):
    if not os.path.exists(filepath):
        return pd.DataFrame()
    return pd.read_excel(filepath, dtype=str)

def save_results(filepath: str, df: pd.DataFrame):
    # Format date columns as requested
    date_cols_ddmmyyyy = ["Result Sent", "Certificate", "E-Certificate"]
    for col in date_cols_ddmmyyyy:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: format_ddmmyyyy(x))
    df.to_excel(filepath, index=False)
    autofilter_and_autofit(filepath)

def format_ddmmyyyy(val):
    import pandas as pd
    if not val or pd.isna(val):
        return ""
    try:
        # If already in DD/MM/YYYY, return as is
        if isinstance(val, str) and len(val) == 10 and val[2] == '/' and val[5] == '/':
            return val
        # Try parsing as datetime string
        dt = pd.to_datetime(val, errors='coerce')
        if pd.isna(dt):
            return val
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return val

def autofilter_and_autofit(filepath: str):
    from openpyxl.utils import get_column_letter
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
        ws.column_dimensions[col_letter].width = max(10, min(max_length + 2, 60))
    wb.save(filepath)