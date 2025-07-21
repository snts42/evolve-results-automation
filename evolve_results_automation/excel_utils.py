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
    df.to_excel(filepath, index=False)
    autofilter_and_autofit(filepath)

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