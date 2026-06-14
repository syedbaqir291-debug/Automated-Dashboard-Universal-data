"""
dashboard_builder.py
Generic "upload any data -> native PivotTable dashboard" engine.

Builds an .xlsx with:
  - Dashboard : title (+ OMAC Developers capsule), KPI cards (formulas
                reading PivotTable cells), one chart per dimension, and
                the live native PivotTables themselves stacked below the
                charts (so the file visibly contains real, interactive
                pivots the moment it's opened).
  - Data      : cleaned source as an Excel Table (the PivotCache source)

All PivotTables share ONE PivotCache, so a single set of Slicers (added
by the user in Excel - see README) cross-filters every pivot, chart and
KPI card at once.
"""

import re
import numpy as np
import pandas as pd
from io import BytesIO

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from openpyxl.pivot.cache import (
    CacheDefinition, CacheField, SharedItems, CacheSource, WorksheetSource
)
from openpyxl.pivot.record import Record, RecordList
from openpyxl.pivot.fields import Missing, Number, Text, Index
from openpyxl.pivot.table import (
    TableDefinition, PivotField, FieldItem, Location, DataField,
    RowColField, RowColItem, PivotTableStyle
)
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.text import CharacterProperties

NAVY = '1F4E5C'
TEAL = '2E8B8B'
LIGHT = 'EAF2F2'
ACCENT = 'C0392B'
GREY = '95A5A6'
GOLD = 'D4AC0D'
PALETTE = [TEAL, NAVY, ACCENT, 'E67E22', 'F1C40F', '27AE60', '8E44AD', '34495E']

RECORD_COL = '__Records__'
BLANK = '(Blank)'

MAX_DIMENSIONS = 8
MAX_CHART_CATEGORIES = 10
DONUT_THRESHOLD = 5  # <= this many categories -> donut, else bar

BRAND = 'OMAC Developers \u00b7 S M Baqir'


# ---------------------------------------------------------------------
# Step 1 - load & profile
# ---------------------------------------------------------------------
def load_dataframe(file_obj, filename):
    if filename.lower().endswith('.csv'):
        return pd.read_csv(file_obj)
    return pd.read_excel(file_obj)


def _norm_text(v):
    if not isinstance(v, str):
        return v
    v = v.replace('\n', ' ').replace('\r', ' ')
    v = ' '.join(v.split())
    return v


def clean_dataframe(df):
    """Clean headers AND normalize whitespace/newlines inside text cells.

    This is the generic fix for categories silently splitting into
    near-duplicates (e.g. "Finance\\n" vs "Finance " vs "Finance  Dept")
    which otherwise fragments PivotTable groupings.
    """
    df = df.copy()
    cols = []
    seen = {}
    for c in df.columns:
        c2 = _norm_text(str(c))
        if not c2 or c2.lower().startswith('unnamed'):
            c2 = 'Column'
        if c2 in seen:
            seen[c2] += 1
            c2 = f'{c2} ({seen[c2]})'
        else:
            seen[c2] = 0
        cols.append(c2)
    df.columns = cols

    for col in df.columns:
        if df[col].dtype == object or pd.api.types.is_string_dtype(df[col]):
            df[col] = df[col].map(_norm_text)

    return df


def profile_columns(df):
    """Return dimension/measure candidates + a suggested default dimension set."""
    n = len(df)
    dim_candidates, num_candidates, scored = [], [], []

    id_pattern = r'\b(id|name|email|phone|contact|address|mrn|uid|url|date|time|no|number|code)\b'

    for col in df.columns:
        s = df[col]
        nunique = s.nunique(dropna=True)
        fill_rate = s.notna().sum() / n if n else 0
        looks_like_id_name = bool(re.search(id_pattern, col.lower()))
        looks_like_id_vals = nunique == n and n > 10

        if pd.api.types.is_numeric_dtype(s):
            if not looks_like_id_name and not looks_like_id_vals:
                num_candidates.append(col)

        is_long_text = False
        if s.dtype == object or pd.api.types.is_string_dtype(s):
            try:
                avg_len = s.dropna().astype(str).str.len().mean()
            except Exception:
                avg_len = 0
            is_long_text = bool(avg_len and avg_len > 30)

        is_numeric = pd.api.types.is_numeric_dtype(s)
        numeric_too_granular = is_numeric and nunique > 20

        looks_like_datetime_vals = False
        if s.dtype == object or pd.api.types.is_string_dtype(s):
            sample = s.dropna().astype(str)
            if len(sample):
                dt_re = re.compile(
                    r'\d{1,2}\s*(?:am|pm)|\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|'
                    r'\d{1,2}[-\s](?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)', re.I)
                match_frac = sample.str.contains(dt_re).mean()
                looks_like_datetime_vals = match_frac > 0.3

        if (2 <= nunique <= max(50, int(n * 0.6)) and not is_long_text
                and not looks_like_id_vals and not looks_like_id_name
                and not numeric_too_granular and not looks_like_datetime_vals
                and fill_rate >= 0.3):
            dim_candidates.append(col)
            score = -nunique  # prefer fewer categories, but include long-tail too
            scored.append((score, col))

    scored.sort(reverse=True, key=lambda t: t[0])
    suggested = [c for _, c in scored[:6]]

    return {
        'dimension_candidates': dim_candidates,
        'numeric_candidates': num_candidates,
        'suggested_dimensions': suggested,
        'row_count': n,
    }


# ---------------------------------------------------------------------
# Step 2 - pivot table helpers
# ---------------------------------------------------------------------
def _make_pivot_fields(ncols, row_field_idx, n_items):
    pfs = []
    for i in range(ncols):
        if i == row_field_idx:
            items = [FieldItem(x=k) for k in range(n_items)] + [FieldItem(t='default')]
            pfs.append(PivotField(axis='axisRow', items=items, showAll=False))
        else:
            pfs.append(PivotField(items=[FieldItem(t='default')], showAll=False))
    return pfs


def _make_pivot(name, ncols, row_field_idx, n_items, ref, cache,
                 measure_idx, subtotal, measure_label, field_caption):
    rowItems = [RowColItem(r=0, i=0, x=[Index(v=k)]) for k in range(n_items)]
    rowItems.append(RowColItem(t='grand', r=0, i=0))
    pt = TableDefinition(
        name=name, cacheId=1, dataCaption='Values',
        applyNumberFormats=False, applyBorderFormats=False, applyFontFormats=False,
        applyPatternFormats=False, applyAlignmentFormats=False, applyWidthHeightFormats=False,
        dataOnRows=False, outline=True, outlineData=True, useAutoFormatting=True,
        itemPrintTitles=True, createdVersion=8, updatedVersion=8, minRefreshableVersion=3,
        indent=0, compact=False, compactData=False,
        location=Location(ref=ref, firstHeaderRow=1, firstDataRow=1, firstDataCol=1),
        pivotFields=_make_pivot_fields(ncols, row_field_idx, n_items),
        rowFields=[RowColField(x=row_field_idx)],
        rowItems=rowItems,
        colFields=[], colItems=[RowColItem(r=0, i=0)],
        dataFields=[DataField(name=measure_label, fld=measure_idx, subtotal=subtotal)],
        pivotTableStyleInfo=PivotTableStyle(name='PivotStyleMedium9', showRowHeaders=True,
                                             showColHeaders=True, showRowStripes=True,
                                             showColStripes=False, showLastColumn=False),
        rowHeaderCaption=field_caption,
    )
    pt.cache = cache
    return pt


def _style_title(chart, text):
    chart.title = text
    try:
        chart.title.tx.rich.p[0].r[0].rPr = CharacterProperties(sz=1200, b=True, solidFill=NAVY)
    except Exception:
        pass


# ---------------------------------------------------------------------
# Step 3 - build the workbook
# ---------------------------------------------------------------------
def build_workbook(df, dims, measure_mode, measure_col=None,
                    title='Auto-Generated Dashboard', source_label=''):
    """
    df            : pandas DataFrame (already loaded)
    dims          : list of column names to use as pivot row fields (1-8)
    measure_mode  : 'count' | 'sum' | 'average'
    measure_col   : numeric column name, required if measure_mode != 'count'

    Returns: bytes of the .xlsx file
    """
    if not dims:
        raise ValueError('Select at least one dimension column.')
    dims = dims[:MAX_DIMENSIONS]

    df = clean_dataframe(df)
    df = df.reset_index(drop=True)
    n = len(df)
    df[RECORD_COL] = 1
    headers = list(df.columns)
    ncols = len(headers)

    for d in dims:
        df[d] = df[d].apply(lambda v: BLANK if pd.isna(v) else str(v))

    # -------- measure / aggregation setup --------
    if measure_mode == 'count':
        measure_idx = headers.index(RECORD_COL)
        subtotal = 'count'
        measure_label = 'Records'
        numfmt = '#,##0'
        sort_fn = lambda d: df[d].value_counts()
    elif measure_mode == 'sum':
        if not measure_col:
            raise ValueError('measure_col required for sum')
        measure_idx = headers.index(measure_col)
        subtotal = 'sum'
        measure_label = f'Sum of {measure_col}'
        numfmt = '#,##0.00'
        sort_fn = lambda d: df.groupby(d)[measure_col].sum(min_count=1)
    else:
        if not measure_col:
            raise ValueError('measure_col required for average')
        measure_idx = headers.index(measure_col)
        subtotal = 'average'
        measure_label = f'Average {measure_col}'
        numfmt = '#,##0.00'
        sort_fn = lambda d: df.groupby(d)[measure_col].mean()

    # -------- categorical ordering (Pareto: highest value first) --------
    CAT_FIELDS = {}
    for d in dims:
        idx = headers.index(d)
        ser = sort_fn(d).fillna(0).sort_values(ascending=False)
        CAT_FIELDS[idx] = [str(v) for v in ser.index.tolist()]
    CAT_INDEX = {idx: {v: i for i, v in enumerate(order)} for idx, order in CAT_FIELDS.items()}

    # -------- cache fields --------
    cache_fields = []
    for i, h in enumerate(headers):
        if i in CAT_FIELDS:
            items = [Text(v=v) for v in CAT_FIELDS[i]]
            shared = SharedItems(_fields=items, containsString=True, containsBlank=False)
        else:
            col = df[h]
            has_nan = bool(col.isna().any())
            if pd.api.types.is_numeric_dtype(col):
                shared = SharedItems(
                    containsString=False, containsNumber=True,
                    containsInteger=bool((col.dropna() % 1 == 0).all()) if col.notna().any() else False,
                    containsBlank=has_nan)
            else:
                shared = SharedItems(containsString=True, containsBlank=has_nan)
        cache_fields.append(CacheField(name=h, sharedItems=shared))

    # -------- records --------
    def field_value(i, h, val):
        if i in CAT_FIELDS:
            return Index(v=CAT_INDEX[i][val])
        if pd.isna(val):
            return Missing()
        if isinstance(val, (int, np.integer, float, np.floating)):
            return Number(v=float(val))
        return Text(v=str(val))

    records = []
    for _, row in df.iterrows():
        fields = tuple(field_value(i, h, row[h]) for i, h in enumerate(headers))
        records.append(Record(_fields=fields))

    last_col_letter = get_column_letter(ncols)
    cache = CacheDefinition(
        refreshOnLoad=True, saveData=True, recordCount=n,
        createdVersion=8, refreshedVersion=8, minRefreshableVersion=3,
        cacheSource=CacheSource(type='worksheet',
                                 worksheetSource=WorksheetSource(ref=f'A1:{last_col_letter}{n + 1}', sheet='Data')),
        cacheFields=cache_fields,
    )
    cache.records = RecordList(r=records)

    # -------- layout: figure out where pivot tables live on Dashboard --------
    n_dims = len(dims)
    n_chart_rows = max(1, -(-n_dims // 3))  # ceil
    chart_anchor_rows = [10, 28, 46, 64, 82, 100]
    pivot_start_row = chart_anchor_rows[min(n_chart_rows, len(chart_anchor_rows)) - 1] + 24

    pivots = {}
    pivot_layout = {}  # dim -> (header_row, grand_row, n_items, fld_idx, name, caption_row)
    running_row = pivot_start_row
    for d in dims:
        idx = headers.index(d)
        n_items = len(CAT_FIELDS[idx])
        name = f'PT_{idx}_{d}'.replace(' ', '_')
        name = ''.join(ch for ch in name if ch.isalnum() or ch == '_')[:31]
        header_row = running_row + 1
        grand_row = header_row + 1 + n_items
        ref = f'A{header_row}:B{grand_row}'
        pivots[name] = _make_pivot(name, ncols, idx, n_items, ref, cache,
                                    measure_idx, subtotal, measure_label, d)
        pivot_layout[d] = (header_row, grand_row, n_items, idx, name, running_row)
        running_row = grand_row + 2

    # -------- workbook --------
    wb = Workbook()
    ws_data = wb.active
    ws_data.title = 'Data'

    header_fill = PatternFill('solid', start_color=NAVY)
    header_font = Font(bold=True, color='FFFFFF', name='Calibri')
    for j, h in enumerate(headers, start=1):
        c = ws_data.cell(row=1, column=j, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(wrap_text=True, vertical='center')

    for i, (_, row) in enumerate(df.iterrows(), start=2):
        for j, h in enumerate(headers, start=1):
            val = row[h]
            if pd.isna(val):
                val = None
            elif isinstance(val, (np.integer,)):
                val = int(val)
            elif isinstance(val, (np.floating,)):
                val = float(val)
            cell = ws_data.cell(row=i, column=j, value=val)
            cell.font = Font(name='Calibri', size=10)
            cell.alignment = Alignment(vertical='top')

    last_col = get_column_letter(ncols)
    tbl = Table(displayName='SourceData', ref=f'A1:{last_col}{n + 1}')
    tbl.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showFirstColumn=False,
                                         showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    ws_data.add_table(tbl)
    ws_data.freeze_panes = 'A2'
    for j, h in enumerate(headers, start=1):
        ws_data.column_dimensions[get_column_letter(j)].width = 22 if len(h) > 14 else 14
    rc_idx = headers.index(RECORD_COL) + 1
    ws_data.column_dimensions[get_column_letter(rc_idx)].hidden = True

    # ---- Dashboard sheet ----
    ws_dash = wb.create_sheet('Dashboard', 0)
    _build_dashboard(ws_dash, dims, headers, CAT_FIELDS, pivot_layout, pivots,
                      measure_mode, measure_col, measure_label, numfmt, title, source_label, n,
                      pivot_start_row)

    wb._sheets = [wb['Dashboard'], wb['Data']]
    wb.active = 0

    out = BytesIO()
    wb.save(out)
    out.seek(0)

    # Round-trip through openpyxl once: this consolidates the pivot cache
    # into a single shared definition (avoids duplicate pivotCache parts).
    wb2 = load_workbook(out)
    out2 = BytesIO()
    wb2.save(out2)
    out2.seek(0)
    return out2.read()


# ---------------------------------------------------------------------
# Step 4 - Dashboard sheet (branding + KPI cards + charts + live pivots)
# ---------------------------------------------------------------------
def _build_dashboard(ws, dims, headers, CAT_FIELDS, pivot_layout, pivots,
                       measure_mode, measure_col, measure_label, numfmt, title, source_label, n,
                       pivot_start_row):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    n_dims = len(dims)
    n_cols_layout = max(24, (min(n_dims, 6) * 3) or 18)
    last_col_letter = get_column_letter(n_cols_layout)

    # ---------------- Title band ----------------
    ws.merge_cells(f'A1:{get_column_letter(n_cols_layout - 6)}2')
    c = ws['A1']
    c.value = title
    c.font = Font(name='Calibri', size=22, bold=True, color='FFFFFF')
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    for row in ws[f'A1:{last_col_letter}2']:
        for cell in row:
            cell.fill = PatternFill('solid', start_color=NAVY)

    # OMAC branding capsule, top-right of the title band
    badge_c0 = n_cols_layout - 5
    ws.merge_cells(start_row=1, start_column=badge_c0, end_row=1, end_column=n_cols_layout)
    b = ws.cell(row=1, column=badge_c0)
    b.value = BRAND
    b.font = Font(name='Calibri', size=10, bold=True, color=NAVY)
    b.fill = PatternFill('solid', start_color=GOLD)
    b.alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(start_row=2, start_column=badge_c0, end_row=2, end_column=n_cols_layout)
    b2 = ws.cell(row=2, column=badge_c0)
    b2.value = 'Auto Pivot Dashboard'
    b2.font = Font(name='Calibri', size=8, italic=True, color='FFFFFF')
    b2.alignment = Alignment(horizontal='center', vertical='center')

    ws.merge_cells(f'A3:{last_col_letter}3')
    c = ws['A3']
    subtitle = f'{n:,} records'
    if source_label:
        subtitle = f'{source_label}  |  {subtitle}'
    c.value = subtitle
    c.font = Font(name='Calibri', size=11, italic=True, color=GREY)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)

    # ---------------- KPI cards ----------------
    if measure_mode == 'count':
        kpi1_label = 'Total Records'
    elif measure_mode == 'sum':
        kpi1_label = f'Total {measure_col}'
    else:
        kpi1_label = f'Overall Average {measure_col}'

    d0 = dims[0]
    h0, g0, n0, idx0, name0, _ = pivot_layout[d0]
    grand_addr0 = f'B{g0}'
    top_label_addr0 = f'A{h0 + 1}'
    top_val_addr0 = f'B{h0 + 1}'
    second_label_addr0 = f'A{h0 + 2}' if n0 >= 2 else top_label_addr0
    second_val_addr0 = f'B{h0 + 2}' if n0 >= 2 else top_val_addr0
    cat_count_formula = f'COUNTA(A{h0 + 1}:A{h0 + n0})'

    kpis = [
        (kpi1_label, f'={grand_addr0}', numfmt, NAVY),
        (f'Top {d0}', f'={top_label_addr0}', '@', TEAL, True),
        (f'{measure_label} ({d0} top)', f'={top_val_addr0}', numfmt, TEAL),
    ]

    if len(dims) >= 2:
        d1 = dims[1]
        h1, g1, n1, idx1, name1, _ = pivot_layout[d1]
        kpis.append((f'Top {d1}', f'=A{h1 + 1}', '@', ACCENT, True))
        kpis.append((f'{measure_label} ({d1} top)', f'=B{h1 + 1}', numfmt, ACCENT))
    else:
        kpis.append((f'2nd Highest {d0}', f'={second_label_addr0}', '@', ACCENT, True))
        kpis.append((f'{measure_label} (2nd)', f'={second_val_addr0}', numfmt, ACCENT))

    kpis.append((f'{d0} Categories', f'={cat_count_formula}', '#,##0', GREY))

    card_width = 3
    start_row = 5
    end_row = 8
    thin = Side(style='thin', color='D0D0D0')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for idx, kpi in enumerate(kpis[:6]):
        label, formula, fmt, color = kpi[0], kpi[1], kpi[2], kpi[3]
        is_text = len(kpi) > 4 and kpi[4]
        c0 = idx * card_width + 1
        c1 = c0 + card_width - 1
        rng_label = f'{get_column_letter(c0)}{start_row}:{get_column_letter(c1)}{start_row}'
        rng_value = f'{get_column_letter(c0)}{start_row + 1}:{get_column_letter(c1)}{end_row - 1}'
        ws.merge_cells(rng_label)
        ws.merge_cells(rng_value)
        lab = ws.cell(row=start_row, column=c0, value=label.upper())
        lab.font = Font(name='Calibri', size=9, bold=True, color=GREY)
        lab.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        val = ws.cell(row=start_row + 1, column=c0, value=formula)
        val.font = Font(name='Calibri', size=20 if is_text else 24, bold=True, color=color)
        if fmt != '@':
            val.number_format = fmt
        val.alignment = Alignment(horizontal='center', vertical='center', wrap_text=is_text)
        for r in range(start_row, end_row + 1):
            for col in range(c0, c1 + 1):
                cell = ws.cell(row=r, column=col)
                cell.border = border
                cell.fill = PatternFill('solid', start_color=LIGHT if r == start_row else 'FFFFFF')

    # ---------------- Charts ----------------
    anchor_cols = ['A', 'J', 'S']
    anchor_rows = [10, 28, 46, 64, 82, 100]
    for i, d in enumerate(dims):
        h, g, n_items, idx, name, _ = pivot_layout[d]
        col_idx = i % 3
        row_idx = i // 3
        anchor = f'{anchor_cols[col_idx]}{anchor_rows[row_idx]}'

        if n_items <= DONUT_THRESHOLD:
            chart = DoughnutChart()
            chart.add_data(Reference(ws, min_col=2, min_row=h, max_row=h + n_items), titles_from_data=True)
            chart.set_categories(Reference(ws, min_col=1, min_row=h + 1, max_row=h + n_items))
            chart.height = 8
            chart.width = 9.5
            chart.dataLabels = DataLabelList(showVal=True, showCatName=False, showSerName=False,
                                              showPercent=False, showLegendKey=False)
            chart.holeSize = 55
            _style_title(chart, d)
            chart.series[0].dPt = [
                DataPoint(idx=k, spPr=GraphicalProperties(solidFill=PALETTE[k % len(PALETTE)]))
                for k in range(n_items)
            ]
        else:
            top_n = min(n_items, MAX_CHART_CATEGORIES)
            chart = BarChart()
            chart.type = 'bar'
            chart.add_data(Reference(ws, min_col=2, min_row=h + 1, max_row=h + top_n), titles_from_data=False)
            chart.set_categories(Reference(ws, min_col=1, min_row=h + 1, max_row=h + top_n))
            chart.height = 9.5
            chart.width = 12.5
            chart.dataLabels = DataLabelList(showVal=True, showCatName=False, showSerName=False,
                                              showPercent=False, showLegendKey=False)
            chart.series[0].graphicalProperties = GraphicalProperties(solidFill=PALETTE[i % len(PALETTE)])
            chart.y_axis.majorGridlines = None
            chart.legend = None
            title_text = f'{d}  (top {top_n})' if n_items > top_n else d
            _style_title(chart, title_text)

        ws.add_chart(chart, anchor)

    # ---------------- Live PivotTables section ----------------
    note_row = pivot_start_row - 2
    ws.merge_cells(f'A{note_row}:{last_col_letter}{note_row}')
    note = ws.cell(row=note_row, column=1)
    note.value = ('\u25bc  Live PivotTables (the engine behind every chart and KPI above).  '
                   'Click any cell below \u2192 Insert \u2192 Slicer to add interactive filters \u2014 '
                   'then Report Connections \u2192 select all PivotTables to cross-filter everything on this sheet.')
    note.font = Font(name='Calibri', size=10, italic=True, color=NAVY)
    note.alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[note_row].height = 22
    for cell in ws[f'A{note_row}:{last_col_letter}{note_row}'][0]:
        cell.fill = PatternFill('solid', start_color=LIGHT)

    for d in dims:
        h, g, n_items, idx, name, caption_row = pivot_layout[d]
        cap = ws.cell(row=caption_row, column=1)
        cap.value = d
        cap.font = Font(name='Calibri', size=12, bold=True, color=NAVY)
        ws.add_pivot(pivots[name])

    # ---------------- Column widths / row heights ----------------
    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 16
    for col in range(3, n_cols_layout + 1):
        ws.column_dimensions[get_column_letter(col)].width = 8.5
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 28
    ws.row_dimensions[3].height = 18
    for r in range(5, 9):
        ws.row_dimensions[r].height = 22
