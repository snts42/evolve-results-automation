import os
import re
import logging
from collections import Counter, defaultdict
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from .config import COLUMNS, get_excel_file_for_year, list_year_excel_files
from .parsing_utils import unique_row_hash


def _normalize(val):
    """Normalize a cell value to a clean string (empty string for None/NaN)."""
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() == "nan" else s


def initialize_excel(filepath: str):
    if not os.path.exists(filepath):
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        logging.info(f"Excel file not found, creating {filepath}")
        wb = Workbook()
        ws = wb.active
        ws.title = "Results"
        ws.append(COLUMNS)
        wb.save(filepath)
        wb.close()


def load_existing_results(filepath: str):
    """Load rows from an Excel file as a list of dicts with string values."""
    if not os.path.exists(filepath):
        return []
    wb = load_workbook(filepath, read_only=True, data_only=True)
    try:
        ws = wb["Results"] if "Results" in wb.sheetnames else wb.active
        rows_iter = ws.iter_rows(values_only=True)

        # First row is the header
        try:
            headers = [str(h) if h is not None else "" for h in next(rows_iter)]
        except StopIteration:
            return []

        result = []
        for row_values in rows_iter:
            row_dict = {h: _normalize(v) for h, v in zip(headers, row_values)}
            result.append(row_dict)

        return result
    finally:
        wb.close()


def _build_results_workbook(rows: list):
    """Build an in-memory Workbook with the Results sheet. Returns the unsaved wb."""
    date_cols_ddmmyyyy = ["Result Sent", "Certificate", "E-Certificate sent"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    ws.sheet_properties.tabColor = 'E30613'
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
    for cell in ws[1]:
        cell.font = _HDR_FONT
        cell.fill = _HDR_FILL
    # Alternating row stripes + thin grey borders (consistent with Analytics tab)
    for row_idx in range(2, ws.max_row + 1):
        for cell in ws[row_idx]:
            cell.border = _THIN_BORDER
            cell.font = _DATA_FONT
            if row_idx % 2 == 0:
                cell.fill = _STRIPE_FILL

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

    return wb


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

    wb = _build_results_workbook(combined)
    try:
        wb.save(excel_file)
    finally:
        wb.close()
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
    for _year_str, excel_path in list_year_excel_files():
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


# =========================================================================
# ANALYTICS SHEET - Compact dashboard, side-by-side charts
# =========================================================================

# ── Colour palette ────────────────────────────────────────────────────────
_RED      = 'E30613'
_WHITE    = 'FFFFFF'
_BLACK    = '1A1A1A'
_GREY_600 = '757575'
_GREY_400 = 'BDBDBD'
_RED_BG   = 'FFEBEE'
_RED_FG   = 'C62828'

# ── Fonts / fills ─────────────────────────────────────────────────────────
_TITLE_FONT      = Font(bold=True, size=14, color=_WHITE)
_TITLE_META_FONT = Font(size=10, color=_WHITE)
_SECTION_FONT    = Font(bold=True, size=12, color=_BLACK)
_HDR_FONT        = Font(bold=True, size=11, color=_WHITE)
_HDR_FILL        = PatternFill('solid', fgColor=_RED)
_DATA_FONT       = Font(size=11, color=_BLACK)
_STRIPE_FILL     = PatternFill('solid', fgColor='FFF2F2')
_TITLE_FILL      = PatternFill('solid', fgColor=_RED)
_KPI_LABEL_FONT  = Font(bold=True, size=10, color=_GREY_600)
_KPI_VALUE_FONT  = Font(bold=True, size=14, color=_RED)
_KPI_VALUE_SM    = Font(bold=True, size=12, color=_RED)
_KPI_SUB_FONT    = Font(size=8, color=_GREY_600)
_KPI_FILL        = PatternFill('solid', fgColor='FFF2F2')
_INSIGHT_FONT    = Font(size=11, color=_BLACK)
_FOOTER_FONT     = Font(italic=True, size=7, color=_GREY_600)
_THIN_BORDER = Border(
    left=Side('thin', color=_GREY_400), right=Side('thin', color=_GREY_400),
    top=Side('thin', color=_GREY_400), bottom=Side('thin', color=_GREY_400))

_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
_MONTHS_FULL = ["January", "February", "March", "April", "May", "June",
                "July", "August", "September", "October", "November", "December"]

# ── Column layout (12 cols: A-L) ─────────────────────────────────────────
_NUM_COLS = 12
_COL_W = {1: 20}
for _ci in range(2, _NUM_COLS + 1):
    _COL_W[_ci] = 13

# Chart sizing
_CHART_W = 14.5          # cm per chart (side-by-side)
_CHART_H = 7.5           # cm height
_CHART_ROWS = 16         # rows a 7.5cm chart occupies

# Wide table column layouts: list of (start_col, end_col) per field
_RESIT_LAYOUT  = [(1, 3), (4, 5), (6, 9), (10, 10), (11, 12)]
_REBOOK_LAYOUT = [(1, 3), (4, 5), (6, 9), (10, 11), (12, 12)]
_EXTRA_LAYOUT  = [(1, 3), (4, 5), (6, 10), (11, 12)]


def _apply_col_widths(ws):
    for ci in range(1, _NUM_COLS + 1):
        ws.column_dimensions[get_column_letter(ci)].width = _COL_W.get(ci, 13)


def _exam_short(name):
    """Strip numeric prefix from exam name."""
    return re.sub(r'^\d{4}-\d{3}\s*', '', name)


def _exam_chart_label(name, max_line=18):
    """Shorten exam name for chart axis and force 2-line wrap if long."""
    s = _exam_short(name)
    s = s.replace('Functional Skills ', 'FS ')
    if len(s) <= max_line:
        return s
    # Insert newline at best word boundary near midpoint
    mid = len(s) // 2
    best = s.rfind(' ', 0, mid + 5)
    if best <= 0:
        best = s.find(' ', mid)
    if best > 0:
        return s[:best] + '\n' + s[best + 1:]
    return s


# ── Section / table helpers ───────────────────────────────────────────────

def _section(ws, row, title, start_col=1, end_col=_NUM_COLS):
    """Write a section title row, merged across given columns."""
    ws.merge_cells(start_row=row, start_column=start_col,
                   end_row=row, end_column=end_col)
    ws.cell(row=row, column=start_col, value=title).font = _SECTION_FONT
    ws.row_dimensions[row].height = 22
    return row + 1


def _wide_tbl_header(ws, row, headers, layout):
    """Write header row with merged cells spanning A-L."""
    for (sc, ec), text in zip(layout, headers):
        if ec > sc:
            ws.merge_cells(start_row=row, start_column=sc,
                           end_row=row, end_column=ec)
        c = ws.cell(row=row, column=sc, value=text)
        c.font = _HDR_FONT
        c.fill = _HDR_FILL
        c.border = _THIN_BORDER
        c.alignment = Alignment(horizontal='center', vertical='center',
                                wrap_text=True)
        for ci in range(sc + 1, ec + 1):
            cell = ws.cell(row=row, column=ci)
            cell.fill = _HDR_FILL
            cell.border = _THIN_BORDER
    ws.row_dimensions[row].height = 22


def _wide_tbl_row(ws, row, values, layout, stripe=False):
    """Write data row with merged cells spanning A-L."""
    for idx, ((sc, ec), val) in enumerate(zip(layout, values)):
        if ec > sc:
            ws.merge_cells(start_row=row, start_column=sc,
                           end_row=row, end_column=ec)
        c = ws.cell(row=row, column=sc, value=val)
        c.font = _DATA_FONT
        c.border = _THIN_BORDER
        c.alignment = Alignment(
            horizontal='left' if idx == 0 else 'center',
            vertical='center', wrap_text=True)
        if stripe:
            c.fill = _STRIPE_FILL
        for ci in range(sc + 1, ec + 1):
            cell = ws.cell(row=row, column=ci)
            cell.border = _THIN_BORDER
            if stripe:
                cell.fill = _STRIPE_FILL
    ws.row_dimensions[row].height = 20


def _short_centre_name(full_name):
    """Shorten centre name for display (max 31 chars)."""
    name = full_name.strip()
    name = re.sub(r'^\d+\s*\([^)]*\)\s*', '', name)
    if ' - ' in name:
        name = name.split(' - ')[0].strip()
    if not name:
        name = full_name.strip()
    return name[:31]


# ── Chart builder ─────────────────────────────────────────────────────────

def _bar_chart(data_ref, cat_ref, fill_hex=_RED, width=15, height=7.5,
               data_labels=True, y_max=None):
    """Create a styled BarChart object."""
    ch = BarChart()
    ch.type = "col"
    ch.style = 10
    ch.title = None
    ch.y_axis.title = None
    ch.y_axis.scaling.min = 0
    if y_max is not None:
        ch.y_axis.scaling.max = y_max
    ch.y_axis.numFmt = '0'
    ch.y_axis.delete = False
    ch.x_axis.tickLblPos = 'low'
    ch.x_axis.delete = False
    ch.width = width
    ch.height = height
    ch.legend = None
    ch.gapWidth = 80
    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(cat_ref)
    s = ch.series[0]
    s.graphicalProperties.solidFill = fill_hex
    if data_labels:
        s.dLbls = DataLabelList()
        s.dLbls.showVal = True
        s.dLbls.showSerName = False
        s.dLbls.showCatName = False
        s.dLbls.showPercent = False
        s.dLbls.showLegendKey = False
        s.dLbls.numFmt = '0'
    return ch


def _exam_breakdown_chart(dws, exam_d_start, exam_d_end, width=15, height=7.5):
    """Build horizontal bar chart: exam names on Y-axis, counts as bars."""
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font as DrawingFont

    ch = BarChart()
    ch.type = "bar"          # horizontal bars
    ch.style = 10
    ch.title = None
    ch.x_axis.title = None   # value axis (bottom)
    ch.x_axis.scaling.min = 0
    ch.x_axis.numFmt = '0'
    ch.x_axis.delete = False
    ch.y_axis.title = None    # category axis (left - exam names)
    ch.y_axis.tickLblPos = 'low'
    ch.y_axis.delete = False
    ch.width = width
    ch.height = height
    ch.legend = None
    ch.gapWidth = 60

    ch.y_axis.txPr = RichText(
        p=[Paragraph(pPr=ParagraphProperties(
            defRPr=CharacterProperties(latin=DrawingFont(typeface='Calibri'), sz=700)
        ), endParaRPr=CharacterProperties(latin=DrawingFont(typeface='Calibri'), sz=700))]
    )

    data_ref = Reference(dws, min_col=2, min_row=exam_d_start,
                         max_row=exam_d_end)
    cat_ref = Reference(dws, min_col=1, min_row=exam_d_start + 1,
                        max_row=exam_d_end)
    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(cat_ref)

    s = ch.series[0]
    s.graphicalProperties.solidFill = _RED
    s.dLbls = DataLabelList()
    s.dLbls.showVal = True
    s.dLbls.showSerName = False
    s.dLbls.showCatName = False
    s.dLbls.showPercent = False
    s.dLbls.showLegendKey = False
    s.dLbls.numFmt = '0'
    return ch


# ── Compute helpers ───────────────────────────────────────────────────────

def _compute_resit_data(rows):
    """Return (resit_tracker, rebook_opps) from rows."""
    pairs = defaultdict(list)
    for rd in rows:
        enrol = rd.get("Enrolment no.", "").strip()
        test = rd.get("Test Name", "").strip()
        if enrol and test:
            pairs[(enrol, test)].append(rd)

    resit_tracker = []
    rebook_opps = []
    def _date_key(r):
        try:
            return datetime.strptime(r.get("Completed", ""), "%d/%m/%Y")
        except (ValueError, TypeError):
            return datetime.min

    for (enrol, test), attempts in pairs.items():
        attempts.sort(key=_date_key)
        latest = attempts[-1]
        name = (f"{latest.get('First name', '')} "
                f"{latest.get('Last name', '')}").strip()
        result = latest.get("Result", "").strip()
        centre = latest.get("Centre Name", "").strip()
        if len(attempts) >= 2:
            resit_tracker.append({
                "name": name, "enrolment": enrol, "exam": test,
                "attempts": len(attempts), "result": result, "centre": centre,
            })
        if result.lower() == "fail":
            fail_date = latest.get("Completed", "")
            try:
                days = (datetime.now() - datetime.strptime(
                    fail_date, "%d/%m/%Y")).days
            except (ValueError, TypeError):
                days = 0
            rebook_opps.append({
                "name": name, "enrolment": enrol, "exam": test,
                "failed": fail_date, "days_ago": days, "centre": centre,
            })

    resit_tracker.sort(key=lambda x: x["attempts"], reverse=True)
    rebook_opps.sort(key=lambda x: x["days_ago"], reverse=True)
    return resit_tracker, rebook_opps


def _compute_extra_time(rows):
    """Return list of dicts for candidates whose duration exceeds mode."""
    by_test = defaultdict(list)
    for rd in rows:
        dur_str = rd.get("Duration", "").strip()
        test = rd.get("Test Name", "").strip()
        if dur_str and test:
            try:
                dur_val = float(dur_str)
                by_test[test].append((rd, dur_val))
            except (ValueError, TypeError):
                pass

    result = []
    for test, entries in by_test.items():
        durations = [d for _, d in entries]
        mode_dur = Counter(durations).most_common(1)[0][0] if durations else 0
        if not mode_dur:
            continue
        for rd, dur in entries:
            if dur > mode_dur * 1.1:
                name = (f"{rd.get('First name', '')} "
                        f"{rd.get('Last name', '')}").strip()
                enrol = rd.get("Enrolment no.", "").strip()
                result.append({
                    "name": name, "enrolment": enrol, "exam": test,
                    "extra_pct": f"+{round((dur / mode_dur - 1) * 100)}%",
                    "centre": rd.get("Centre Name", "").strip(),
                })
    return result


def _compute_insights(rows, by_exam, monthly, centres, rebook_opps):
    """Generate auto-insight bullet strings."""
    total = len(rows)
    insights = []
    if by_exam:
        top = by_exam[0]
        pct = round(top["count"] / total * 100) if total else 0
        insights.append(
            f'{top["name"]} is the most sat exam - '
            f'{top["count"]} sittings ({pct}% of annual volume)')

    if rebook_opps:
        insights.append(
            f'{len(rebook_opps)} candidate(s) failed and have not rebooked - '
            f'potential revenue to recover')

    busiest = max(monthly.items(), key=lambda x: x[1]["count"],
                  default=(0, {"count": 0}))
    if busiest[1]["count"] > 0:
        mi = busiest[0]
        cnt = busiest[1]["count"]
        centre_count = len(centres)
        insights.append(
            f'{_MONTHS_FULL[mi - 1]} was the busiest month - {cnt} sittings' 
            + (f' across {centre_count} centres' if centre_count > 1 else ''))

    return insights


# ── Header / KPI / footer builders ────────────────────────────────────────

def _build_header(ws, title, subtitle, year):
    """Write single-row red header: bold title left, smaller meta right."""
    r = 1
    # Title in A-D
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    c = ws.cell(row=r, column=1, value=title)
    c.font = _TITLE_FONT
    c.fill = _TITLE_FILL
    c.alignment = Alignment(vertical='center')
    for ci in range(2, 5):
        ws.cell(row=r, column=ci).fill = _TITLE_FILL
    # Meta in E-L (smaller font)
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=_NUM_COLS)
    meta = (f"{subtitle}    Year: {year}    "
            f"Generated: {datetime.now().strftime('%d/%m/%Y')}")
    m = ws.cell(row=r, column=5, value=meta)
    m.font = _TITLE_META_FONT
    m.fill = _TITLE_FILL
    m.alignment = Alignment(vertical='center')
    for ci in range(6, _NUM_COLS + 1):
        ws.cell(row=r, column=ci).fill = _TITLE_FILL
    ws.row_dimensions[r].height = 30
    return r + 2  # row 2 = spacer


def _build_kpis(ws, row, kpis, layout=None):
    """Write KPI cards with grey card background and red top accent.
    layout = list of (start_col, end_col) per KPI. Returns next row."""
    if layout is None:
        layout = [(1, 2), (3, 4), (5, 6), (7, 9), (10, 11), (12, 12)]

    for i, kpi in enumerate(kpis):
        if i >= len(layout):
            break
        sc, ec = layout[i]
        span = ec - sc + 1
        label, val = kpi[0], kpi[1]
        sub = kpi[2] if len(kpi) > 2 else None
        # Use smaller font for text values that need more space
        is_text = isinstance(val, str) and len(val) > 6
        val_font = _KPI_VALUE_SM if is_text else _KPI_VALUE_FONT

        # Merge cells for 3 rows
        for r_off in range(3):
            ws.merge_cells(start_row=row + r_off, start_column=sc,
                           end_row=row + r_off, end_column=ec)

        # Apply card styling: grey fill + red top/bottom accent
        for r_off in range(3):
            for c_off in range(span):
                cell = ws.cell(row=row + r_off, column=sc + c_off)
                cell.fill = _KPI_FILL
                top_s = Side('medium', color=_RED) if r_off == 0 else Side('thin', color=_GREY_400)
                bot_s = Side('medium', color=_RED) if r_off == 2 else Side('thin', color=_GREY_400)
                cell.border = Border(
                    top=top_s, bottom=bot_s,
                    left=Side('thin', color=_GREY_400),
                    right=Side('thin', color=_GREY_400))

        # Label row
        lc = ws.cell(row=row, column=sc, value=label)
        lc.font = _KPI_LABEL_FONT
        lc.fill = _KPI_FILL
        lc.alignment = Alignment(horizontal='center', vertical='center')

        # Value row
        vc = ws.cell(row=row + 1, column=sc, value=val)
        vc.font = val_font
        vc.fill = _KPI_FILL
        vc.alignment = Alignment(horizontal='center', vertical='center')

        # Subtitle row
        if sub:
            sc2 = ws.cell(row=row + 2, column=sc, value=sub)
            sc2.font = _KPI_SUB_FONT
            sc2.fill = _KPI_FILL
            sc2.alignment = Alignment(horizontal='center', vertical='center')

    ws.row_dimensions[row].height = 18
    ws.row_dimensions[row + 1].height = 30
    ws.row_dimensions[row + 2].height = 16
    return row + 4


def _build_footer(ws, row, text):
    """Write a grey footer line."""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=_NUM_COLS)
    c = ws.cell(row=row, column=1, value=text)
    c.font = _FOOTER_FONT
    c.alignment = Alignment(horizontal='left')
    ws.row_dimensions[row].height = 14
    return row + 1


# ── Public entry-point ────────────────────────────────────────────────────
def add_analytics_sheet(filepath, rows=None):
    """Add/replace the Analytics sheet on an existing Excel file."""
    if rows is None:
        rows = load_existing_results(filepath)
    if not rows:
        return
    wb = load_workbook(filepath)
    try:
        if wb.active and wb.active.title not in ("Results", "Analytics"):
            wb.active.title = "Results"
        _add_analytics_to_wb(wb, rows)
        wb.save(filepath)
    finally:
        wb.close()


# ── Main orchestrator ─────────────────────────────────────────────────────
def _add_analytics_to_wb(wb, rows):
    """Build the Analytics tab (single compact dashboard)."""
    if not rows:
        return

    # Remove all non-Results tabs
    for name in list(wb.sheetnames):
        if name != "Results":
            del wb[name]

    # Split rows by centre
    centres = defaultdict(list)
    for rd in rows:
        cn = rd.get("Centre Name", "").strip() or "Unknown Centre"
        centres[cn].append(rd)

    # Detect year
    year = datetime.now().year
    for rd in rows:
        try:
            year = datetime.strptime(rd.get("Completed", ""), "%d/%m/%Y").year
            break
        except (ValueError, TypeError):
            pass

    # Compute data
    total = len(rows)
    by_exam = _compute_by_exam(rows)
    resit_tracker, rebook_opps = _compute_resit_data(rows)
    extra_time = _compute_extra_time(rows)

    monthly = {}
    for rd in rows:
        try:
            mi = datetime.strptime(rd.get("Completed", ""), "%d/%m/%Y").month
            if mi not in monthly:
                monthly[mi] = {"count": 0}
            monthly[mi]["count"] += 1
        except (ValueError, TypeError):
            pass

    customers = len({rd.get("Enrolment no.", "").strip() for rd in rows
                     if rd.get("Enrolment no.", "").strip()})
    total_fails = sum(1 for rd in rows
                      if rd.get("Result", "").strip().lower() == "fail")
    resits_returned = len(resit_tracker)
    resit_conv = (f"{round(resits_returned / total_fails * 100)}%"
                  if total_fails else "N/A")
    resit_sub = (f"{resits_returned} of {total_fails} returned"
                 if total_fails else "No fails")

    top_exam = _exam_short(by_exam[0]["name"]) if by_exam else "N/A"
    top_exam_pct = round(by_exam[0]["count"] / total * 100) if by_exam and total else 0
    top_exam_sub = f"{by_exam[0]['count']} sittings ({top_exam_pct}%)" if by_exam else ""

    busiest_mi = max(range(1, 13),
                     key=lambda m: monthly.get(m, {"count": 0})["count"],
                     default=1)
    busiest_cnt = monthly.get(busiest_mi, {"count": 0})["count"]
    busiest_pct = round(busiest_cnt / total * 100) if total else 0
    busiest_sub = f"{busiest_cnt} sittings ({busiest_pct}%)" if busiest_cnt else ""

    # ── Hidden chart-data sheet ───────────────────────────────────────
    dws = wb.create_sheet("_ChartData")
    dws.sheet_state = 'hidden'

    # Monthly data (rows 1-13)
    dws.cell(1, 1, "Month")
    dws.cell(1, 2, "Exams")
    for i in range(12):
        dws.cell(2 + i, 1, _MONTHS[i])
        cnt = monthly.get(i + 1, {"count": 0})["count"]
        dws.cell(2 + i, 2, cnt if cnt else None)

    # Exam data (rows 15+): one row per exam for horizontal bar chart
    exam_d_start = 15
    dws.cell(exam_d_start, 1, "Exam")
    dws.cell(exam_d_start, 2, "Count")
    for i, ex in enumerate(by_exam):
        dws.cell(exam_d_start + 1 + i, 1, _exam_chart_label(ex["name"]))
        dws.cell(exam_d_start + 1 + i, 2, ex["count"])
    exam_d_end = exam_d_start + len(by_exam)

    # ── Analytics tab ─────────────────────────────────────────────────
    ws = wb.create_sheet("Analytics", 1)
    ws.sheet_properties.tabColor = _RED
    ws.sheet_view.showGridLines = False

    # Row 1: Header
    r = _build_header(ws, "EVOLVE SECURE-ASSESS ANALYTICS",
                      f" All Centres ({len(centres)})", year)

    # Row 3: OVERVIEW + KPIs (wide layout: Most Popular gets 3 cols)
    r = _section(ws, r, "OVERVIEW")
    extra_candidates = len({et["enrolment"] for et in extra_time if et["enrolment"]})
    extra_pct = round(extra_candidates / customers * 100) if customers else 0
    extra_sub = f"{extra_candidates} of {customers} ({extra_pct}%)" if customers else ""
    kpi_layout = [(1, 2), (3, 4), (5, 6), (7, 9), (10, 11), (12, 12)]
    r = _build_kpis(ws, r, [
        ("Total Exams", total),
        ("Unique Candidates", customers),
        ("Extra Time Candidates", extra_candidates, extra_sub),
        ("Most Popular Exam", top_exam, top_exam_sub),
        ("Busiest Month", _MONTHS_FULL[busiest_mi - 1], busiest_sub),
        ("Resit Conversion", resit_conv, resit_sub),
    ], layout=kpi_layout)

    # ── Side-by-side charts ───────────────────────────────────────────
    # Left: Monthly Volume (A-F), Right: Exam Breakdown (G-L)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws.cell(row=r, column=1, value="MONTHLY VOLUME").font = _SECTION_FONT
    ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=12)
    ws.cell(row=r, column=7, value="EXAM BREAKDOWN").font = _SECTION_FONT
    ws.row_dimensions[r].height = 22
    r += 1

    chart_anchor_row = r

    # Monthly volume chart (left) - with 25% headroom for data labels
    month_max = max((monthly.get(m, {"count": 0})["count"]
                     for m in range(1, 13)), default=0)
    ch_month = _bar_chart(
        Reference(dws, min_col=2, min_row=1, max_row=13),
        Reference(dws, min_col=1, min_row=2, max_row=13),
        fill_hex=_RED, width=_CHART_W, height=_CHART_H,
        y_max=int(month_max * 1.25) + 1 if month_max else None)
    ws.add_chart(ch_month, f"A{chart_anchor_row}")

    # Exam breakdown chart (right) - horizontal bars with names on Y-axis
    if by_exam:
        ch_exam = _exam_breakdown_chart(
            dws, exam_d_start, exam_d_end,
            width=_CHART_W, height=_CHART_H)
        ws.add_chart(ch_exam, f"G{chart_anchor_row}")

    r = chart_anchor_row + _CHART_ROWS
    r += 1  # spacer

    # ── Key Insights (top of text sections for visibility) ────────────
    insights = _compute_insights(rows, by_exam, monthly, centres, rebook_opps)
    if insights:
        r = _section(ws, r, "KEY INSIGHTS")
        for ins in insights:
            ws.merge_cells(start_row=r, start_column=1,
                           end_row=r, end_column=_NUM_COLS)
            c = ws.cell(row=r, column=1, value=f"  {ins}")
            c.font = _INSIGHT_FONT
            c.alignment = Alignment(vertical='center', wrap_text=True)
            ws.row_dimensions[r].height = 20
            r += 1
        r += 1

    r = _section(ws, r, "REBOOK OPPORTUNITIES")
    if rebook_opps:
        _wide_tbl_header(ws, r, ["Candidate", "Enrolment", "Exam",
                                  "Failed", "Days Ago"],
                         _REBOOK_LAYOUT)
        r += 1
        for idx, rb in enumerate(rebook_opps):
            _wide_tbl_row(ws, r, [rb["name"], rb["enrolment"], rb["exam"],
                                   rb["failed"], rb["days_ago"]],
                          _REBOOK_LAYOUT, stripe=(idx % 2 == 1))
            r += 1
    else:
        ws.merge_cells(start_row=r, start_column=1,
                       end_row=r, end_column=_NUM_COLS)
        ws.cell(row=r, column=1,
                value="  No rebook opportunities - all fails have "
                      "rebooked or passed on resit").font = _INSIGHT_FONT
        r += 1
    r += 1

    failed_resits = [rt for rt in resit_tracker if rt["result"].lower() == "fail"]
    if failed_resits:
        r = _section(ws, r, "FAILED RESIT TRACKER")
        _wide_tbl_header(ws, r, ["Candidate", "Enrolment", "Exam",
                                  "Attempts", "Latest Result"],
                         _RESIT_LAYOUT)
        r += 1
        for idx, rt in enumerate(failed_resits):
            _wide_tbl_row(ws, r, [rt["name"], rt["enrolment"], rt["exam"],
                                   rt["attempts"], rt["result"]],
                          _RESIT_LAYOUT, stripe=(idx % 2 == 1))
            rc = ws.cell(row=r, column=11)
            rc.fill = PatternFill('solid', fgColor=_RED_BG)
            rc.font = Font(bold=True, size=9, color=_RED_FG)
            ws.cell(row=r, column=12).fill = PatternFill('solid', fgColor=_RED_BG)
            r += 1
        r += 1

    # ── Extra Time ────────────────────────────────────────────────────
    r = _section(ws, r, "EXTRA TIME CANDIDATES")
    if extra_time:
        _wide_tbl_header(ws, r, ["Candidate", "Enrolment", "Exam",
                                  "Extra Time"],
                         _EXTRA_LAYOUT)
        r += 1
        for idx, et in enumerate(extra_time):
            _wide_tbl_row(ws, r, [et["name"], et["enrolment"], et["exam"],
                                   et["extra_pct"]],
                          _EXTRA_LAYOUT, stripe=(idx % 2 == 1))
            r += 1
    else:
        ws.merge_cells(start_row=r, start_column=1,
                       end_row=r, end_column=_NUM_COLS)
        ws.cell(row=r, column=1,
                value="  No extra time candidates detected").font = _INSIGHT_FONT
        r += 1
    r += 1

    # ── Footer ────────────────────────────────────────────────────────
    centre_names = ", ".join(_short_centre_name(cn)
                             for cn in sorted(centres.keys()))
    _build_footer(ws, r, f"Data: Results tab - {total} records - "
                         f"{len(centres)} centre(s): {centre_names} - "
                         f"Internal use - City & Guilds")

    _apply_col_widths(ws)


def _compute_by_exam(rows):
    """Group rows by Test Name. Returns list of dicts sorted by count desc.
    Each: {name, count, passed, failed, rate, avg_score}."""
    groups = defaultdict(list)
    for r in rows:
        name = r.get("Test Name", "").strip()
        if name:
            groups[name].append(r)
    result = []
    for name, group in groups.items():
        count = len(group)
        passed = sum(1 for r in group if r.get("Result", "").strip().lower() == "pass")
        scores = []
        for r in group:
            try:
                scores.append(float(r.get("Percent", "0").replace("%", "")))
            except (ValueError, TypeError):
                pass
        avg_score = round(sum(scores) / len(scores), 1) if scores else 0.0
        result.append({
            "name": name, "count": count, "passed": passed,
            "failed": count - passed,
            "rate": round(passed / count * 100, 1) if count else 0.0,
            "avg_score": avg_score,
        })
    result.sort(key=lambda x: x["count"], reverse=True)
    return result
