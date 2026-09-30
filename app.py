"""
Tableau Metadata Extractor V2
Multi-source extraction: REST API + Metadata API (GraphQL) + File Parsing
"""

from flask import Flask, render_template, request, jsonify, send_file
import json
import os
os.environ["NO_PROXY"] = "insights.connect.te.com"
import sys
import getpass
import tableauserverclient as TSC
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import tempfile
import shutil
from io import BytesIO
from openpyxl.utils import get_column_letter
import threading
import queue
import uuid
import re
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
from admin_insights_analyzer import AdminInsightsAnalyzer
from export_utils import build_export, clean_cell, write_json_export
import lineage_engine
import workbook_analyzer
import time

# Configure logging - never log secrets
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Handle PyInstaller frozen executable paths
def get_base_path():
    """Get the base path for resources (handles both dev and frozen exe)"""
    if getattr(sys, 'frozen', False):
        # Running as compiled .exe
        return sys._MEIPASS
    else:
        # Running as script
        return os.path.dirname(os.path.abspath(__file__))

def get_app_data_path():
    """Get path for app data (extractions, uploads, etc.) - writable location"""
    if getattr(sys, 'frozen', False):
        # Running as compiled .exe - use user's app data folder
        if sys.platform == 'win32':
            app_data = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
            data_path = os.path.join(app_data, 'TableauMetadataExtractor')
        else:
            data_path = os.path.join(os.path.expanduser('~'), '.tableau_metadata_extractor')
    else:
        # Running as script - use current directory
        data_path = os.path.dirname(os.path.abspath(__file__))

    # Ensure directory exists
    if not os.path.exists(data_path):
        os.makedirs(data_path)
    return data_path

BASE_PATH = get_base_path()
APP_DATA_PATH = get_app_data_path()

# Initialize Flask with correct template and static paths
app = Flask(__name__,
            template_folder=os.path.join(BASE_PATH, 'templates'),
            static_folder=os.path.join(BASE_PATH, 'static'))
app.secret_key = os.urandom(24)  # Secure random key
# Flask's jsonify() alphabetizes dict keys by default, which silently reorders every column-ordered
# dataset this app returns (Lineage Info, Overall Lineage, Single Workbook Analysis, ...) into alphabetical
# order in the JSON response. Keep the order the code actually builds instead.
app.json.sort_keys = False

# ==================== Job Queue System ====================
job_queue = queue.Queue()
job_results = {}
job_status = {}
analytics_status = {}  # Alias for job_status (analytics specific)
worker_thread = None
admin_insights_data = None  # Storage for uploaded Admin Insights CSV

# Only the most recent results are kept in memory - a huge site can produce millions of rows per job.
MAX_STORED_JOB_RESULTS = 3
# Max rows per dataset sent to the browser. The full data always stays on the server
# (export / save use the full data; /api/job_result/<id>/dataset/<name> pages through it).
PREVIEW_ROWS_PER_DATASET = 5000
# Above this many rows the styled pandas exporters are too slow / memory hungry - use the streaming bucketed export.
LARGE_EXPORT_ROWS = 100_000


def _evict_old_job_results(keep_job_id):
    """Drop the oldest stored results (and their status entries) beyond MAX_STORED_JOB_RESULTS."""
    try:
        stale = [j for j in list(job_results.keys()) if j != keep_job_id]
        excess = len(job_results) - MAX_STORED_JOB_RESULTS
        for old in stale[:max(0, excess)]:
            _release_result(job_results.pop(old, None))
            logger.info(f"Evicted stored result of job {old} to free memory")
    except Exception as e:
        logger.warning(f"Could not evict old job results: {e}")


def job_worker():
    """Background worker for heavy extraction jobs"""
    while True:
        try:
            job_id, func, args, kwargs = job_queue.get()
            logger.info(f"Worker starting job {job_id}")
            job_status[job_id] = {'status': 'running', 'progress': 0, 'message': 'Processing...'}
            try:
                result = func(*args, **kwargs, job_id=job_id)
                job_results[job_id] = result
                _evict_old_job_results(job_id)
                message = 'Completed'
                if isinstance(result, dict) and result.get('warnings'):
                    message = f"Completed with {len(result['warnings'])} warning(s) - see the Warnings sheet in the export"
                job_status[job_id] = {'status': 'completed', 'progress': 100, 'message': message}
                logger.info(f"Job {job_id} completed successfully")
            except Exception as e:
                logger.error(f"Job {job_id} failed: {str(e)}")
                import traceback
                logger.error(traceback.format_exc())
                job_status[job_id] = {'status': 'failed', 'progress': 0, 'message': str(e)}
            finally:
                job_queue.task_done()
        except Exception as e:
            logger.error(f"Worker error: {str(e)}")

def ensure_worker_running():
    """Ensure the background worker thread is running"""
    global worker_thread
    if worker_thread is None or not worker_thread.is_alive():
        logger.info("Starting background worker thread")
        worker_thread = threading.Thread(target=job_worker, daemon=True)
        worker_thread.start()

# Start background worker
ensure_worker_running()


# ==================== Excel Export Helpers ====================

def sanitize_dataframe_for_excel(df):
    """
    Sanitize a DataFrame for safe Excel export.
    - Replaces NaN with empty string in object/string columns
    - Keeps numeric columns numeric but replaces NaN with empty string for display
    - Handles mixed types safely
    """
    if df is None or df.empty:
        return df

    df = df.copy()

    for col in df.columns:
        # Check if column is numeric (int, float)
        if pd.api.types.is_numeric_dtype(df[col]):
            # For numeric columns, fillna with empty string for display
            # But first convert to object type to allow mixed types
            df[col] = df[col].apply(lambda x: '' if pd.isna(x) else x)
        elif pd.api.types.is_datetime64_any_dtype(df[col]):
            # For datetime columns, convert NaT to empty string
            df[col] = df[col].apply(lambda x: '' if pd.isna(x) else x)
        else:
            # For object/string columns: NaN -> '', strip characters Excel/openpyxl reject
            # (control chars in SQL/formulas raise IllegalCharacterError) and cap cell length at 32k
            df[col] = df[col].apply(clean_cell)

    return df


def get_safe_column_width(series, col_name, max_width=50):
    """
    Safely calculate column width for Excel auto-sizing.
    Handles NaN, floats, empty series, and edge cases.

    Args:
        series: pandas Series to calculate width for
        col_name: column name (for header width comparison)
        max_width: maximum width to cap at

    Returns:
        int: safe column width
    """
    try:
        header_len = len(str(col_name)) if col_name is not None else 0

        if series is None or len(series) == 0:
            return min(header_len + 2, max_width)

        # Convert to string safely, handling NaN and other types
        def safe_len(val):
            if val is None:
                return 0
            if isinstance(val, float) and pd.isna(val):
                return 0
            try:
                return len(str(val))
            except:
                return 0

        lengths = series.apply(safe_len)
        max_data_len = lengths.max() if len(lengths) > 0 else 0

        # Handle case where max() returns NaN
        if pd.isna(max_data_len):
            max_data_len = 0

        max_length = max(int(max_data_len), header_len) + 2
        return min(max_length, max_width)

    except Exception:
        # Fallback to a reasonable default
        return min(len(str(col_name)) + 2 if col_name else 10, max_width)


def apply_excel_formatting(worksheet, df, header_fill=None, header_font=None, freeze_panes=True):
    """
    Apply standard formatting to an Excel worksheet.
    Includes auto-sizing columns and styling headers.

    Args:
        worksheet: openpyxl worksheet object
        df: pandas DataFrame that was written to the worksheet
        header_fill: optional PatternFill for headers
        header_font: optional Font for headers
        freeze_panes: whether to freeze the header row
    """
    if header_fill is None:
        header_fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
    if header_font is None:
        header_font = Font(color='FFFFFF', bold=True)

    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Auto-adjust column widths safely
    for idx, col in enumerate(df.columns):
        width = get_safe_column_width(df[col], col)
        worksheet.column_dimensions[get_column_letter(idx + 1)].width = width

    # Format headers
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border

    # Format data cells
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
        for cell in row:
            cell.border = thin_border

    # Freeze header row
    if freeze_panes:
        worksheet.freeze_panes = 'A2'


def create_analysis_sheet(writer, results):
    """
    Create an Analysis sheet with summary KPIs and insights from metadata.

    Args:
        writer: pandas ExcelWriter object
        results: dictionary of query results
    """
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    # Parse data from results - safely handle missing data (using new query types)
    ds_data = results.get('published_datasources', [])
    wb_data = results.get('workbooks', [])
    wb_dependency_data = results.get('workbook_dependency', [])
    wb_field_lineage_data = results.get('workbook_field_lineage', [])
    ds_impact_data = results.get('datasource_impact', [])

    # Convert to DataFrames
    df_ds = pd.DataFrame(ds_data) if ds_data else pd.DataFrame()
    df_wb = pd.DataFrame(wb_data) if wb_data else pd.DataFrame()
    df_wb_dep = pd.DataFrame(wb_dependency_data) if wb_dependency_data else pd.DataFrame()
    df_wb_field = pd.DataFrame(wb_field_lineage_data) if wb_field_lineage_data else pd.DataFrame()
    df_ds_impact = pd.DataFrame(ds_impact_data) if ds_impact_data else pd.DataFrame()

    # Helper function to safely get column values
    def safe_get(df, col, default=''):
        if df.empty or col not in df.columns:
            return pd.Series(dtype=object)
        return df[col].fillna(default)

    # Helper to safely convert to string
    def safe_str(val):
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return ''
        return str(val)

    # ==================== SECTION 1: Summary KPIs ====================
    total_datasources = len(df_ds)
    total_workbooks = len(df_wb)

    # Datasources with downstream workbooks (from datasource_impact)
    if not df_ds_impact.empty and 'Downstream Workbook Count' in df_ds_impact.columns:
        ds_with_downstream = len(df_ds_impact[df_ds_impact['Downstream Workbook Count'] > 0])
    else:
        ds_with_downstream = 0
    ds_without_downstream = total_datasources - ds_with_downstream if total_datasources > 0 else 0

    # Unique upstream tables
    upstream_tables_set = set()
    if not df_ds.empty and 'Upstream Tables' in df_ds.columns:
        for tables in safe_get(df_ds, 'Upstream Tables'):
            if tables:
                for t in str(tables).split(', '):
                    if t.strip():
                        upstream_tables_set.add(t.strip())
    total_upstream_tables = len(upstream_tables_set)

    # Total fields from workbook field lineage
    total_fields = len(df_wb_field) if not df_wb_field.empty else 0

    # Fields with formulas (calculated fields)
    if not df_wb_field.empty and 'Formula' in df_wb_field.columns:
        fields_with_formula = len(df_wb_field[safe_get(df_wb_field, 'Formula').str.strip() != ''])
    else:
        fields_with_formula = 0

    # Workbooks with embedded datasources
    if not df_wb.empty and 'Embedded Datasources' in df_wb.columns:
        wb_with_embedded = len(df_wb[safe_get(df_wb, 'Embedded Datasources').str.strip() != ''])
    else:
        wb_with_embedded = 0

    # Workbooks with upstream published datasources (from workbook_dependency)
    if not df_wb_dep.empty and 'Upstream Published DS Count' in df_wb_dep.columns:
        wb_with_upstream_ds = len(df_wb_dep[df_wb_dep['Upstream Published DS Count'] > 0])
    else:
        wb_with_upstream_ds = 0

    kpis = [
        ['SECTION 1: SUMMARY KPIs', ''],
        ['', ''],
        ['Metric', 'Value'],
        ['Total Published Datasources', total_datasources],
        ['Total Workbooks', total_workbooks],
        ['Datasources with Downstream Workbooks', ds_with_downstream],
        ['Datasources with No Downstream Workbooks', ds_without_downstream],
        ['Total Unique Upstream Tables', total_upstream_tables],
        ['Total Fields in Workbook Field Lineage', total_fields],
        ['Calculated Fields (with Formula)', fields_with_formula],
        ['Workbooks with Embedded Datasources', wb_with_embedded],
        ['Workbooks with Upstream Published DS', wb_with_upstream_ds],
        ['', ''],
    ]

    # ==================== SECTION 2: Datasource Impact Summary ====================
    section2_header = [['SECTION 2: PUBLISHED DATASOURCE IMPACT SUMMARY', '', '', '', '', '', '', '', '']]
    section2_cols = ['Datasource ID', 'Datasource Name', 'Project', 'Connection Types',
                     'Downstream Workbook Count', 'Downstream Workbooks', 'Downstream Owners',
                     'Upstream Tables', 'Status']

    section2_data = []

    if not df_ds_impact.empty:
        for _, row in df_ds_impact.iterrows():
            ds_id = safe_str(row.get('Datasource ID', ''))
            ds_name = safe_str(row.get('Datasource Name', ''))
            project = safe_str(row.get('Project', ''))
            conn_types = safe_str(row.get('Connection Types', ''))
            downstream_count = row.get('Downstream Workbook Count', 0)
            downstream_wbs = safe_str(row.get('Downstream Workbooks', ''))
            downstream_owners = safe_str(row.get('Downstream Workbook Owners', ''))
            upstream_tables = safe_str(row.get('Upstream Tables', ''))

            # Convert to int safely
            try:
                downstream_count = int(downstream_count) if downstream_count else 0
            except:
                downstream_count = 0

            status = 'In Use' if downstream_count > 0 else 'Unused / No downstream workbook'

            section2_data.append([
                ds_id, ds_name, project, conn_types,
                downstream_count, downstream_wbs, downstream_owners,
                upstream_tables, status
            ])

    # ==================== SECTION 3: Workbook Dependency Summary ====================
    section3_header = [['SECTION 3: WORKBOOK DEPENDENCY SUMMARY', '', '', '', '', '', '', '', '', '', '']]
    section3_cols = ['Workbook ID', 'Workbook Name', 'Project', 'Owner Email',
                     'Embedded DS Count', 'Embedded Datasources',
                     'Upstream Published DS Count', 'Upstream Published Datasources',
                     'Upstream Tables', 'Connection Types', 'Sheet Count']

    section3_data = []

    if not df_wb_dep.empty:
        for _, row in df_wb_dep.iterrows():
            wb_id = safe_str(row.get('Workbook ID', ''))
            wb_name = safe_str(row.get('Workbook Name', ''))
            project = safe_str(row.get('Project', ''))
            owner_email = safe_str(row.get('Owner Email', ''))
            embedded_count = row.get('Embedded Datasource Count', 0)
            embedded_ds = safe_str(row.get('Embedded Datasources', ''))
            upstream_ds_count = row.get('Upstream Published DS Count', 0)
            upstream_ds = safe_str(row.get('Upstream Published Datasources', ''))
            upstream_tables = safe_str(row.get('Upstream Tables', ''))
            conn_types = safe_str(row.get('Connection Types', ''))
            sheet_count = row.get('Sheet Count', 0)

            # Convert to int safely
            try:
                embedded_count = int(embedded_count) if embedded_count else 0
                upstream_ds_count = int(upstream_ds_count) if upstream_ds_count else 0
                sheet_count = int(sheet_count) if sheet_count else 0
            except:
                embedded_count = 0
                upstream_ds_count = 0
                sheet_count = 0

            section3_data.append([
                wb_id, wb_name, project, owner_email,
                embedded_count, embedded_ds,
                upstream_ds_count, upstream_ds,
                upstream_tables, conn_types, sheet_count
            ])

    # ==================== SECTION 4: Workbook Field Summary ====================
    section4_header = [['SECTION 4: WORKBOOK FIELD SUMMARY', '', '', '', '']]
    section4_cols = ['Workbook Name', 'Embedded Datasource', 'Total Fields', 'Calculated Fields', 'Regular Fields']

    section4_data = []

    # Group by workbook and datasource
    if not df_wb_field.empty and 'Workbook Name' in df_wb_field.columns:
        grouped = df_wb_field.groupby(['Workbook Name', 'Embedded Datasource Name'])
        for (wb_name, ds_name), group in grouped:
            total_fields_grp = len(group)
            if 'Formula' in group.columns:
                calc_fields = len(group[safe_get(group, 'Formula').str.strip() != ''])
            else:
                calc_fields = 0
            regular_fields = total_fields_grp - calc_fields

            section4_data.append([safe_str(wb_name), safe_str(ds_name), total_fields_grp, calc_fields, regular_fields])

    # ==================== SECTION 5: Risk / Review Flags ====================
    section5_header = [['SECTION 5: RISK / REVIEW FLAGS', '', '', '', '']]
    section5_cols = ['Object Type', 'Name', 'Project', 'Issue', 'Details']

    section5_data = []

    # Flag 1: Datasource has no downstream workbook (from datasource_impact)
    if not df_ds_impact.empty and 'Downstream Workbook Count' in df_ds_impact.columns:
        for _, row in df_ds_impact.iterrows():
            downstream_count = row.get('Downstream Workbook Count', 0)
            try:
                downstream_count = int(downstream_count) if downstream_count else 0
            except:
                downstream_count = 0

            if downstream_count == 0:
                section5_data.append([
                    'Datasource',
                    safe_str(row.get('Datasource Name', '')),
                    safe_str(row.get('Project', '')),
                    'No downstream workbook',
                    'This datasource is not used by any workbook'
                ])

    # Flag 2: Workbook has no upstream published datasource (from workbook_dependency)
    if not df_wb_dep.empty and 'Upstream Published DS Count' in df_wb_dep.columns:
        for _, row in df_wb_dep.iterrows():
            upstream_count = row.get('Upstream Published DS Count', 0)
            try:
                upstream_count = int(upstream_count) if upstream_count else 0
            except:
                upstream_count = 0

            if upstream_count == 0:
                section5_data.append([
                    'Workbook',
                    safe_str(row.get('Workbook Name', '')),
                    safe_str(row.get('Project', '')),
                    'No upstream published datasource',
                    'This workbook does not use any published datasource'
                ])

    # Flag 3: Blank upstream tables (from datasource_impact)
    if not df_ds_impact.empty and 'Upstream Tables' in df_ds_impact.columns:
        for _, row in df_ds_impact.iterrows():
            upstream = safe_str(row.get('Upstream Tables', ''))
            if not upstream.strip():
                section5_data.append([
                    'Datasource',
                    safe_str(row.get('Datasource Name', '')),
                    safe_str(row.get('Project', '')),
                    'Blank upstream tables',
                    'No upstream tables defined'
                ])

    # Flag 4: Blank upstream tables (from workbook_dependency)
    if not df_wb_dep.empty and 'Upstream Tables' in df_wb_dep.columns:
        for _, row in df_wb_dep.iterrows():
            upstream = safe_str(row.get('Upstream Tables', ''))
            if not upstream.strip():
                section5_data.append([
                    'Workbook',
                    safe_str(row.get('Workbook Name', '')),
                    safe_str(row.get('Project', '')),
                    'Blank upstream tables',
                    'No upstream tables defined'
                ])

    # Flag 5: Duplicate datasource names
    if not df_ds.empty and 'Name' in df_ds.columns:
        name_counts = df_ds['Name'].value_counts()
        duplicates = name_counts[name_counts > 1]
        for name in duplicates.index:
            dup_rows = df_ds[df_ds['Name'] == name]
            ids = ', '.join(dup_rows['ID'].astype(str).tolist()) if 'ID' in dup_rows.columns else ''
            section5_data.append([
                'Datasource', safe_str(name), '',
                'Duplicate name', f'Found {len(dup_rows)} datasources with same name. IDs: {ids}'
            ])

    # Flag 6: Duplicate workbook names
    if not df_wb.empty and 'Name' in df_wb.columns:
        name_counts = df_wb['Name'].value_counts()
        duplicates = name_counts[name_counts > 1]
        for name in duplicates.index:
            dup_rows = df_wb[df_wb['Name'] == name]
            ids = ', '.join(dup_rows['ID'].astype(str).tolist()) if 'ID' in dup_rows.columns else ''
            section5_data.append([
                'Workbook', safe_str(name), '',
                'Duplicate name', f'Found {len(dup_rows)} workbooks with same name. IDs: {ids}'
            ])

    # ==================== BUILD THE ANALYSIS SHEET ====================

    # Create worksheet
    workbook = writer.book
    ws = workbook.create_sheet('Analysis')

    # Styles
    header_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True, size=11)
    section_fill = PatternFill(start_color='2E75B6', end_color='2E75B6', fill_type='solid')
    section_font = Font(color='FFFFFF', bold=True, size=12)
    kpi_label_font = Font(bold=True)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    current_row = 1

    # SECTION 1: KPIs
    for row_data in kpis:
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            if current_row == 1:  # Section header
                cell.fill = section_fill
                cell.font = section_font
            elif current_row == 3:  # Column headers
                cell.fill = header_fill
                cell.font = header_font
            elif col_idx == 1 and current_row > 3:  # KPI labels
                cell.font = kpi_label_font
            cell.border = thin_border
        current_row += 1

    current_row += 1  # Empty row

    # SECTION 2: Datasource Impact Summary
    for row_data in section2_header:
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=val if col_idx == 1 else '')
            cell.fill = section_fill
            cell.font = section_font
        current_row += 1

    # Column headers
    for col_idx, col_name in enumerate(section2_cols, 1):
        cell = ws.cell(row=current_row, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
    current_row += 1

    # Data rows
    for row_data in section2_data:
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            cell.border = thin_border
        current_row += 1

    current_row += 1  # Empty row

    # SECTION 3: Workbook Dependency Summary
    for row_data in section3_header:
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=val if col_idx == 1 else '')
            cell.fill = section_fill
            cell.font = section_font
        current_row += 1

    for col_idx, col_name in enumerate(section3_cols, 1):
        cell = ws.cell(row=current_row, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
    current_row += 1

    for row_data in section3_data:
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            cell.border = thin_border
        current_row += 1

    current_row += 1

    # SECTION 4: Field Lineage Summary
    for row_data in section4_header:
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=val if col_idx == 1 else '')
            cell.fill = section_fill
            cell.font = section_font
        current_row += 1

    for col_idx, col_name in enumerate(section4_cols, 1):
        cell = ws.cell(row=current_row, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
    current_row += 1

    for row_data in section4_data:
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            cell.border = thin_border
        current_row += 1

    current_row += 1

    # SECTION 5: Risk / Review Flags
    for row_data in section5_header:
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=val if col_idx == 1 else '')
            cell.fill = section_fill
            cell.font = section_font
        current_row += 1

    for col_idx, col_name in enumerate(section5_cols, 1):
        cell = ws.cell(row=current_row, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
    current_row += 1

    for row_data in section5_data:
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            cell.border = thin_border
            # Highlight risk rows
            if col_idx == 4:  # Issue column
                cell.font = Font(color='C00000')  # Red text for issues
        current_row += 1

    # Auto-size columns (safely)
    for col_idx in range(1, 13):
        max_width = 12
        for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
            for cell in row:
                try:
                    if cell.value:
                        cell_len = len(str(cell.value))
                        if cell_len > max_width:
                            max_width = min(cell_len, 50)
                except:
                    pass
        ws.column_dimensions[get_column_letter(col_idx)].width = max_width + 2

    # Freeze first row
    ws.freeze_panes = 'A2'

    return ws


class MetadataAPIClient:
    """Client for Tableau Metadata API (GraphQL)"""

    def __init__(self, server_url, auth_token, site_id):
        self.server_url = server_url.rstrip('/')
        self.auth_token = auth_token
        self.site_id = site_id
        self.graphql_endpoint = f"{self.server_url}/api/metadata/graphql"
        self.available = False
        self._tls = threading.local()   # last_error / schema_error are per thread (parallel lineage fetch)
        self.last_error = ''
        self.schema_error = False
        self.fetch_errors = []  # non-fatal problems (partial results) surfaced to the user
        self._check_availability()

    def _check_availability(self):
        """Check if Metadata API is available"""
        try:
            response = self._execute_query("{ __typename }")
            self.available = response is not None
            if self.available:
                logger.info("Metadata API available")
            else:
                logger.warning("Metadata API unavailable - will use fallback")
        except Exception as e:
            logger.warning(f"Metadata API unavailable: {str(e)}")
            self.available = False

    # Tableau rejects a query that touches more than 20,000 nodes or runs too long.
    # Big sites hit both, so large lists are fetched page by page (see _fetch_all).
    QUERY_TIMEOUT = 300
    QUERY_RETRIES = 3

    @property
    def last_error(self):
        return getattr(self._tls, 'last_error', '')

    @last_error.setter
    def last_error(self, value):
        self._tls.last_error = value

    @property
    def schema_error(self):
        """True when the last query was rejected because a requested field/type does not exist on this server."""
        return getattr(self._tls, 'schema_error', False)

    @schema_error.setter
    def schema_error(self, value):
        self._tls.schema_error = value

    DEFAULT_PAGE_SIZE = 200
    MIN_PAGE_SIZE = 5

    def _execute_query(self, query, variables=None):
        """Execute GraphQL query (returns data dict or None). Failure reason is kept in self.last_error."""
        data, _ = self._execute_query_ex(query, variables)
        return data

    def _execute_query_ex(self, query, variables=None):
        """Execute GraphQL query with retries. Returns (data, retryable_by_smaller_page)."""
        headers = {
            'Content-Type': 'application/json',
            'X-Tableau-Auth': self.auth_token
        }
        payload = {'query': query}
        if variables:
            payload['variables'] = variables

        self.last_error = ''
        self.schema_error = False
        too_big = False
        for attempt in range(1, self.QUERY_RETRIES + 1):
            try:
                response = requests.post(self.graphql_endpoint, json=payload, headers=headers,
                                         timeout=self.QUERY_TIMEOUT)
                if response.status_code == 200:
                    data = response.json()
                    errors = data.get('errors')
                    if errors:
                        text = json.dumps(errors, default=str)[:500]
                        self.last_error = f'GraphQL errors: {text}'
                        logger.warning(self.last_error)
                        lowered = text.lower()
                        if ('cannot query field' in lowered or 'validation' in lowered
                                or 'unknown type' in lowered or 'unknown argument' in lowered):
                            self.schema_error = True
                        # NODE_LIMIT_EXCEEDED / TIMEOUT -> caller should retry with a smaller page
                        too_big = ('node_limit' in lowered or 'node limit' in lowered
                                   or 'timeout' in lowered or 'timed out' in lowered)
                        if too_big:
                            return None, True
                    return data.get('data'), False
                self.last_error = f'GraphQL request failed: HTTP {response.status_code}'
                logger.warning(self.last_error)
                if response.status_code in (502, 503, 504, 408, 429):
                    too_big = True  # gateway timeouts are usually the query being too heavy
                    time.sleep(2 * attempt)
                    continue
                return None, too_big
            except Exception as e:
                self.last_error = f'GraphQL request error: {str(e)}'
                logger.warning(self.last_error)
                too_big = True
                time.sleep(2 * attempt)
        return None, too_big

    def fetch_collection(self, root, variants):
        """
        Fetch a whole Metadata API collection (cursor paginated). `variants` are field selections from richest
        to minimal; the next one is tried only when the server rejects a field (schema error).
        Returns (nodes | None, variant_index, error_text).
        """
        last_err = ''
        for idx, fields in enumerate(variants):
            res = self._fetch_all(root, f'{root}Connection', fields,
                                  record_errors=(idx == len(variants) - 1))
            if res is not None:
                return res[root], idx, ''
            last_err = self.last_error
            if not self.schema_error:
                break                 # not a schema problem - a smaller field set would not help
        return None, -1, last_err or 'no response'

    def _paginated(self, query, root_key):
        """
        Run a single-shot collection query (`{ <root_key> { ...fields } }`) with cursor pagination,
        reusing the same field selection. Falls back to the single-shot query if it cannot be parsed
        or the server does not support the `<root_key>Connection` form.
        """
        try:
            start = query.index(root_key) + len(root_key)

            def match(text, open_idx, open_ch, close_ch):
                depth, i = 0, open_idx
                while i < len(text):
                    if text[i] == open_ch:
                        depth += 1
                    elif text[i] == close_ch:
                        depth -= 1
                        if depth == 0:
                            return i
                    i += 1
                raise ValueError('unbalanced brackets')

            filter_args = ''
            rest = query[start:].lstrip()
            offset = len(query) - len(rest)
            if rest.startswith('('):
                close = match(query, offset, '(', ')')
                filter_args = query[offset + 1:close].strip()
                offset = close + 1
            open_idx = query.index('{', offset)
            close_idx = match(query, open_idx, '{', '}')
            fields = query[open_idx + 1:close_idx].strip()
            if not fields:
                raise ValueError('empty selection')
        except Exception as e:
            logger.warning(f"Could not derive paginated query for {root_key}: {e}")
            return self._execute_query(query)
        return self._fetch_all(root_key, f'{root_key}Connection', fields,
                               legacy_query=query, filter_args=filter_args)

    def _fetch_all(self, list_key, connection_key, node_fields, legacy_query=None, filter_args='',
                   record_errors=True):
        """
        Fetch every node of a Metadata API collection with cursor pagination.

        Returns {list_key: [nodes...]} (same shape the old single-shot queries returned) or None
        when nothing could be fetched. If a page is rejected (node limit / timeout) the page size is
        halved and the same page retried. `legacy_query` is used only if the paginated form is
        unsupported by the server at all.
        """
        page_size = self.DEFAULT_PAGE_SIZE
        cursor = None
        nodes = []
        pages = 0
        while True:
            after = f', after: "{cursor}"' if cursor else ''
            extra = f', {filter_args}' if filter_args else ''
            query = (
                "query { %s(first: %d%s%s) { nodes { %s } "
                "pageInfo { hasNextPage endCursor } } }" % (connection_key, page_size, after, extra, node_fields)
            )
            data, too_big = self._execute_query_ex(query)
            conn = (data or {}).get(connection_key) if data else None

            if conn is None:
                if too_big and page_size > self.MIN_PAGE_SIZE:
                    page_size = max(self.MIN_PAGE_SIZE, page_size // 2)
                    logger.info(f"Metadata API {connection_key}: reducing page size to {page_size}")
                    continue
                if pages == 0 and legacy_query:
                    logger.info(f"Metadata API {connection_key}: paginated query unavailable, using legacy query")
                    return self._execute_query(legacy_query)
                # Partial result - keep what we have and record it
                if record_errors or pages > 0:
                    self.fetch_errors.append(
                        f"{connection_key}: stopped after {len(nodes)} records ({self.last_error or 'no response'})")
                break

            nodes.extend(conn.get('nodes') or [])
            pages += 1
            page_info = conn.get('pageInfo') or {}
            if not page_info.get('hasNextPage') or not page_info.get('endCursor'):
                break
            cursor = page_info['endCursor']

        if not nodes and pages == 0:
            return None
        logger.info(f"Metadata API {connection_key}: fetched {len(nodes)} records in {pages} page(s)")
        return {list_key: nodes}

    def get_workbook_lineage(self, workbook_luid):
        """Get lineage for a workbook"""
        query = """
        query WorkbookLineage($luid: String!) {
            workbooks(filter: {luid: $luid}) {
                id
                luid
                name
                upstreamDatasources {
                    id
                    luid
                    name
                    upstreamTables {
                        id
                        name
                        schema
                        fullName
                        database {
                            name
                            connectionType
                        }
                        columns {
                            id
                            name
                            remoteType
                        }
                    }
                }
                sheets {
                    id
                    name
                    upstreamDatasources {
                        name
                    }
                    upstreamFields {
                        name
                        upstreamColumns {
                            name
                            table {
                                name
                            }
                        }
                    }
                }
            }
        }
        """
        return self._execute_query(query, {'luid': workbook_luid})

    def get_datasource_lineage(self, datasource_luid):
        """Get lineage for a datasource"""
        query = """
        query DatasourceLineage($luid: String!) {
            publishedDatasources(filter: {luid: $luid}) {
                id
                luid
                name
                upstreamTables {
                    id
                    name
                    schema
                    fullName
                    database {
                        name
                        connectionType
                    }
                    columns {
                        id
                        name
                        remoteType
                    }
                }
                downstreamWorkbooks {
                    id
                    name
                    luid
                }
                fields {
                    id
                    name
                    description
                    isCalculated
                    formula
                    upstreamColumns {
                        name
                        table {
                            name
                        }
                    }
                }
            }
        }
        """
        return self._execute_query(query, {'luid': datasource_luid})

    def get_all_tables(self):
        """Get all database tables with comprehensive lineage information"""
        query = """
        query AllTables {
            databaseTables {
                id
                name
                schema
                fullName
                isEmbedded
                database {
                    id
                    name
                    connectionType
                    isEmbedded
                }
                columns {
                    id
                    name
                    remoteType
                    description
                    isNullable
                }
                downstreamDatasources {
                    id
                    name
                    luid
                    projectName
                    hasExtracts
                    extractLastRefreshTime
                }
                downstreamWorkbooks {
                    id
                    name
                    luid
                    projectName
                    createdAt
                    updatedAt
                }
                downstreamOwners {
                    name
                    email
                }
            }
        }
        """
        return self._paginated(query, 'databaseTables')

    def get_calculated_fields(self):
        """Get all calculated fields"""
        query = """
        query CalculatedFields {
            calculatedFields {
                id
                name
                formula
                dataCategory
                role
                datasource {
                    name
                    luid
                }
            }
        }
        """
        return self._paginated(query, 'calculatedFields')

    def get_workbooks_with_stats(self):
        """Get all workbooks with sheet counts and usage statistics"""
        query = """
        query WorkbooksWithStats {
            workbooks {
                id
                luid
                name
                projectName
                owner {
                    name
                    email
                }
                createdAt
                updatedAt
                sheets {
                    id
                    name
                }
                dashboards {
                    id
                    name
                }
                upstreamDatasources {
                    id
                    luid
                    name
                }
            }
        }
        """
        return self._paginated(query, 'workbooks')

    def get_views_with_usage(self):
        """Get all views/sheets with usage statistics"""
        query = """
        query ViewsWithUsage {
            sheets {
                id
                name
                workbook {
                    luid
                    name
                }
            }
            dashboards {
                id
                name
                workbook {
                    luid
                    name
                }
            }
        }
        """
        return self._execute_query(query)

    def get_custom_sql_tables(self):
        """Get custom SQL tables"""
        query = """
        query CustomSQLTables {
            customSQLTables {
                id
                name
                query
                database {
                    name
                    connectionType
                }
                downstreamDatasources {
                    name
                    luid
                }
            }
        }
        """
        return self._paginated(query, 'customSQLTables')

    def get_database_servers(self):
        """Get all database servers with connection/gateway details"""
        query = """
        query DatabaseServers {
            databaseServers {
                id
                name
                hostName
                port
                connectionType
                extendedConnectionType
                service
                description
                isEmbedded
                projectName
                isCertified
                hasActiveWarning
                contact {
                    name
                    email
                }
                downstreamDatasources {
                    id
                    name
                    luid
                    projectName
                }
                downstreamWorkbooks {
                    id
                    name
                    luid
                    projectName
                }
            }
        }
        """
        return self._paginated(query, 'databaseServers')

    def get_virtual_connections(self):
        """Get all virtual connections (centralized connection management)"""
        query = """
        query VirtualConnections {
            virtualConnections {
                id
                luid
                name
                description
                projectName
                createdAt
                updatedAt
                isCertified
                hasActiveWarning
                owner {
                    name
                    email
                }
                tables {
                    id
                    name
                }
                upstreamDatabases {
                    id
                    name
                    connectionType
                }
                upstreamTables {
                    id
                    name
                    schema
                    fullName
                }
                downstreamDatasources {
                    id
                    name
                    luid
                    projectName
                }
                downstreamWorkbooks {
                    id
                    name
                    luid
                    projectName
                }
            }
        }
        """
        return self._paginated(query, 'virtualConnections')

    def fetch_published_datasources(self):
        """Fetch all published datasources with upstream table lineage"""
        query = """
        query {
            publishedDatasources {
                id
                site { name }
                luid
                name
                projectName
                uri
                updatedAt
                owner {
                    username
                    name
                    email
                }
                upstreamTables {
                    name
                    schema
                    database {
                        name
                        connectionType
                    }
                }
            }
        }
        """
        data = self._paginated(query, 'publishedDatasources')
        if not data or 'publishedDatasources' not in data:
            return None

        # Flatten the response into tabular format
        flattened = []
        for ds in data['publishedDatasources']:
            # Extract upstream tables info
            upstream_tables = ds.get('upstreamTables') or []
            connection_types = set()
            table_names = []

            for table in upstream_tables:
                db = table.get('database') or {}
                db_name = db.get('name') or ''
                schema = table.get('schema') or ''
                table_name = table.get('name') or ''
                conn_type = db.get('connectionType') or ''

                if conn_type:
                    connection_types.add(conn_type)

                # Build full table path: database.schema.table
                parts = [p for p in [db_name, schema, table_name] if p]
                if parts:
                    table_names.append('.'.join(parts))

            # Extract owner info safely
            owner = ds.get('owner') or {}
            site = ds.get('site') or {}

            flattened.append({
                'ID': ds.get('id') or '',
                'LUID': ds.get('luid') or '',
                'Site': site.get('name') or '',
                'Name': ds.get('name') or '',
                'Project': ds.get('projectName') or '',
                'URI': ds.get('uri') or '',
                'Updated At': ds.get('updatedAt') or '',
                'Owner Username': owner.get('username') or '',
                'Owner Name': owner.get('name') or '',
                'Owner Email': owner.get('email') or '',
                'Connection Type': ', '.join(sorted(connection_types)) if connection_types else '',
                'Upstream Tables': ', '.join(table_names) if table_names else ''
            })

        return flattened

    # ==================== Enhanced Metadata GraphQL Methods ====================

    def fetch_metadata(self, query_type, project_filter=None, name_filter=None):
        """
        Unified metadata fetch method supporting multiple query types with filtering.

        Args:
            query_type: One of 'published_datasources', 'workbooks', 'workbook_dependency',
                       'workbook_field_lineage', 'datasource_impact'
            project_filter: Optional project name to filter by
            name_filter: Optional name filter for specific item

        Returns:
            dict with 'data' (flattened list) and 'error' (if any)
        """
        try:
            if query_type == 'published_datasources':
                return self._fetch_published_datasources_filtered(project_filter, name_filter)
            elif query_type == 'workbooks':
                return self._fetch_workbooks_filtered(project_filter, name_filter)
            elif query_type == 'workbook_dependency':
                return self._fetch_workbook_dependency_backbone(project_filter, name_filter)
            elif query_type == 'workbook_field_lineage':
                return self._fetch_workbook_field_lineage(name_filter)
            elif query_type == 'datasource_impact':
                return self._fetch_published_datasource_impact(project_filter)
            elif query_type == 'database_servers':
                return self._fetch_database_servers()
            elif query_type == 'virtual_connections':
                return self._fetch_virtual_connections()
            else:
                return {'data': [], 'error': f'Unknown query type: {query_type}'}
        except Exception as e:
            error_msg = str(e)
            if '20000' in error_msg or 'limit' in error_msg.lower():
                return {
                    'data': [],
                    'error': 'Response exceeds GraphQL 20,000 row limit. Please select a specific project to narrow the scope.'
                }
            return {'data': [], 'error': error_msg}

    def _build_filter(self, project_filter=None, name_filter=None):
        """Build GraphQL filter string"""
        filters = []
        if project_filter and project_filter != 'all':
            filters.append(f'projectName: "{project_filter}"')
        if name_filter:
            filters.append(f'name: "{name_filter}"')

        if filters:
            return f'(filter: {{ {", ".join(filters)} }})'
        return ''

    def _fetch_published_datasources_filtered(self, project_filter=None, name_filter=None):
        """Fetch published datasources with optional filtering"""
        filter_str = self._build_filter(project_filter, name_filter)

        query = f"""
        query {{
            publishedDatasources{filter_str} {{
                id
                site {{ name }}
                luid
                name
                projectName
                uri
                updatedAt
                owner {{
                    username
                    name
                    email
                }}
                upstreamTables {{
                    name
                    schema
                    database {{
                        name
                        connectionType
                    }}
                }}
            }}
        }}
        """

        data = self._paginated(query, 'publishedDatasources')
        if not data:
            return {'data': [], 'error': 'No response from Metadata API'}

        if 'publishedDatasources' not in data:
            return {'data': [], 'error': 'Invalid response structure'}

        flattened = []
        for ds in data['publishedDatasources']:
            upstream_tables = ds.get('upstreamTables') or []
            connection_types = set()
            table_names = []

            for table in upstream_tables:
                db = table.get('database') or {}
                db_name = db.get('name') or ''
                schema = table.get('schema') or ''
                table_name = table.get('name') or ''
                conn_type = db.get('connectionType') or ''

                if conn_type:
                    connection_types.add(conn_type)

                parts = [p for p in [db_name, schema, table_name] if p]
                if parts:
                    table_names.append('.'.join(parts))

            owner = ds.get('owner') or {}
            site = ds.get('site') or {}

            flattened.append({
                'ID': ds.get('id') or '',
                'LUID': ds.get('luid') or '',
                'Site': site.get('name') or '',
                'Name': ds.get('name') or '',
                'Project': ds.get('projectName') or '',
                'URI': ds.get('uri') or '',
                'Updated At': ds.get('updatedAt') or '',
                'Owner Username': owner.get('username') or '',
                'Owner Name': owner.get('name') or '',
                'Owner Email': owner.get('email') or '',
                'Connection Type': ', '.join(sorted(connection_types)) if connection_types else '',
                'Upstream Tables': ', '.join(table_names) if table_names else ''
            })

        return {'data': flattened, 'error': None}

    def _fetch_workbooks_filtered(self, project_filter=None, name_filter=None):
        """Fetch workbooks with optional filtering"""
        filter_str = self._build_filter(project_filter, name_filter)

        query = f"""
        query {{
            workbooks{filter_str} {{
                name
                projectName
                luid
                uri
                id
                projectLuid
                updatedAt
                owner {{
                    email
                }}
                embeddedDatasources {{
                    id
                    name
                    upstreamTables {{
                        name
                        schema
                        database {{
                            name
                            connectionType
                        }}
                    }}
                }}
            }}
        }}
        """

        data = self._paginated(query, 'workbooks')
        if not data:
            return {'data': [], 'error': 'No response from Metadata API'}

        if 'workbooks' not in data:
            return {'data': [], 'error': 'Invalid response structure'}

        flattened = []
        for wb in data['workbooks']:
            owner = wb.get('owner') or {}
            embedded_ds = wb.get('embeddedDatasources') or []

            # Collect all upstream tables from embedded datasources
            all_tables = []
            connection_types = set()
            ds_names = []

            for eds in embedded_ds:
                ds_names.append(eds.get('name') or '')
                upstream_tables = eds.get('upstreamTables') or []
                for table in upstream_tables:
                    db = table.get('database') or {}
                    db_name = db.get('name') or ''
                    schema = table.get('schema') or ''
                    table_name = table.get('name') or ''
                    conn_type = db.get('connectionType') or ''

                    if conn_type:
                        connection_types.add(conn_type)

                    parts = [p for p in [db_name, schema, table_name] if p]
                    if parts:
                        all_tables.append('.'.join(parts))

            flattened.append({
                'ID': wb.get('id') or '',
                'LUID': wb.get('luid') or '',
                'Name': wb.get('name') or '',
                'Project': wb.get('projectName') or '',
                'Project LUID': wb.get('projectLuid') or '',
                'URI': wb.get('uri') or '',
                'Updated At': wb.get('updatedAt') or '',
                'Owner Email': owner.get('email') or '',
                'Embedded Datasources': ', '.join([n for n in ds_names if n]) if ds_names else '',
                'Connection Types': ', '.join(sorted(connection_types)) if connection_types else '',
                'Upstream Tables': ', '.join(all_tables) if all_tables else ''
            })

        return {'data': flattened, 'error': None}

    def _fetch_workbook_dependency_backbone(self, project_filter=None, name_filter=None):
        """
        Fetch workbook dependency backbone - main workbook lineage showing identity,
        owner, embedded datasources, upstream published datasources, upstream tables, and sheets.
        """
        # Build filter with project and workbook name
        filters = []
        if project_filter and project_filter != 'all':
            filters.append(f'projectName: "{project_filter}"')
        if name_filter:
            filters.append(f'name: "{name_filter}"')
        filter_str = f'(filter: {{ {", ".join(filters)} }})' if filters else ''

        query = f"""
        query {{
            workbooks{filter_str} {{
                id
                luid
                name
                projectName
                uri
                updatedAt
                owner {{
                    username
                    name
                    email
                }}
                embeddedDatasources {{
                    id
                    name
                    upstreamTables {{
                        name
                        schema
                        database {{
                            name
                            connectionType
                        }}
                    }}
                }}
                upstreamDatasources {{
                    id
                    name
                }}
                sheets {{
                    id
                    name
                }}
            }}
        }}
        """

        data = self._paginated(query, 'workbooks')
        if not data:
            return {'data': [], 'error': 'No response from Metadata API'}

        if 'workbooks' not in data:
            return {'data': [], 'error': 'Invalid response structure'}

        flattened = []
        for wb in data['workbooks']:
            owner = wb.get('owner') or {}
            embedded_ds = wb.get('embeddedDatasources') or []
            upstream_ds = wb.get('upstreamDatasources') or []
            sheets = wb.get('sheets') or []

            # Process embedded datasources
            embedded_names = []
            all_tables = []
            connection_types = set()

            for eds in embedded_ds:
                eds_name = eds.get('name') or ''
                if eds_name:
                    embedded_names.append(eds_name)

                for table in (eds.get('upstreamTables') or []):
                    db = table.get('database') or {}
                    db_name = db.get('name') or ''
                    schema = table.get('schema') or ''
                    table_name = table.get('name') or ''
                    conn_type = db.get('connectionType') or ''

                    if conn_type:
                        connection_types.add(conn_type)

                    parts = [p for p in [db_name, schema, table_name] if p]
                    if parts:
                        all_tables.append('.'.join(parts))

            # Process upstream published datasources
            upstream_pub_ds_names = []
            for uds in upstream_ds:
                uds_name = uds.get('name') or ''
                if uds_name:
                    upstream_pub_ds_names.append(uds_name)

            # Process sheets
            sheet_names = [s.get('name') or '' for s in sheets if s.get('name')]

            flattened.append({
                'Workbook ID': wb.get('id') or '',
                'Workbook LUID': wb.get('luid') or '',
                'Workbook Name': wb.get('name') or '',
                'Project': wb.get('projectName') or '',
                'URI': wb.get('uri') or '',
                'Updated At': wb.get('updatedAt') or '',
                'Owner Username': owner.get('username') or '',
                'Owner Name': owner.get('name') or '',
                'Owner Email': owner.get('email') or '',
                'Embedded Datasource Count': len(embedded_names),
                'Embedded Datasources': ', '.join(embedded_names) if embedded_names else '',
                'Upstream Published DS Count': len(upstream_pub_ds_names),
                'Upstream Published Datasources': ', '.join(upstream_pub_ds_names) if upstream_pub_ds_names else '',
                'Upstream Tables': ', '.join(list(set(all_tables))) if all_tables else '',
                'Connection Types': ', '.join(sorted(connection_types)) if connection_types else '',
                'Sheet Count': len(sheet_names),
                'Sheets': ', '.join(sheet_names) if sheet_names else ''
            })

        return {'data': flattened, 'error': None}

    def _fetch_workbook_field_lineage(self, name_filter=None):
        """
        Fetch workbook field lineage - field-level lineage for workbook embedded datasources
        including calculated field formulas.
        """
        # Build filter with workbook name only
        filter_str = f'(filter: {{ name: "{name_filter}" }})' if name_filter else ''

        query = f"""
        query {{
            workbooks{filter_str} {{
                name
                embeddedDatasources {{
                    name
                    fields {{
                        id
                        name
                        ... on CalculatedField {{
                            formula
                        }}
                    }}
                }}
            }}
        }}
        """

        data = self._paginated(query, 'workbooks')
        if not data:
            return {'data': [], 'error': 'No response from Metadata API'}

        if 'workbooks' not in data:
            return {'data': [], 'error': 'Invalid response structure'}

        flattened = []
        for wb in data['workbooks']:
            wb_name = wb.get('name') or ''
            embedded_ds = wb.get('embeddedDatasources') or []

            for eds in embedded_ds:
                eds_name = eds.get('name') or ''
                fields = eds.get('fields') or []

                if fields:
                    for field in fields:
                        flattened.append({
                            'Workbook Name': wb_name,
                            'Embedded Datasource Name': eds_name,
                            'Field ID': field.get('id') or '',
                            'Field Name': field.get('name') or '',
                            'Formula': field.get('formula') or ''
                        })
                else:
                    flattened.append({
                        'Workbook Name': wb_name,
                        'Embedded Datasource Name': eds_name,
                        'Field ID': '',
                        'Field Name': '(No fields)',
                        'Formula': ''
                    })

        return {'data': flattened, 'error': None}

    def _fetch_published_datasource_impact(self, project_filter=None):
        """
        Fetch published datasource impact - shows which published datasources feed
        downstream workbooks and their upstream tables.
        """
        # Build filter with project only
        filter_str = f'(filter: {{ projectName: "{project_filter}" }})' if project_filter and project_filter != 'all' else ''

        query = f"""
        query {{
            publishedDatasources{filter_str} {{
                id
                luid
                name
                projectName
                downstreamWorkbooks {{
                    id
                    luid
                    name
                    projectName
                    owner {{
                        username
                    }}
                }}
                upstreamTables {{
                    name
                    schema
                    database {{
                        name
                        connectionType
                    }}
                }}
            }}
        }}
        """

        data = self._paginated(query, 'publishedDatasources')
        if not data:
            return {'data': [], 'error': 'No response from Metadata API'}

        if 'publishedDatasources' not in data:
            return {'data': [], 'error': 'Invalid response structure'}

        flattened = []
        for ds in data['publishedDatasources']:
            downstream_wbs = ds.get('downstreamWorkbooks') or []
            upstream_tables = ds.get('upstreamTables') or []

            # Process downstream workbooks
            wb_names = []
            wb_owners = []
            for dwb in downstream_wbs:
                wb_name = dwb.get('name') or ''
                if wb_name:
                    wb_names.append(wb_name)
                owner = dwb.get('owner') or {}
                owner_name = owner.get('username') or ''
                if owner_name:
                    wb_owners.append(owner_name)

            # Process upstream tables
            table_names = []
            connection_types = set()
            for table in upstream_tables:
                db = table.get('database') or {}
                db_name = db.get('name') or ''
                schema = table.get('schema') or ''
                table_name = table.get('name') or ''
                conn_type = db.get('connectionType') or ''

                if conn_type:
                    connection_types.add(conn_type)

                parts = [p for p in [db_name, schema, table_name] if p]
                if parts:
                    table_names.append('.'.join(parts))

            flattened.append({
                'Datasource ID': ds.get('id') or '',
                'Datasource LUID': ds.get('luid') or '',
                'Datasource Name': ds.get('name') or '',
                'Project': ds.get('projectName') or '',
                'Downstream Workbook Count': len(wb_names),
                'Downstream Workbooks': ', '.join(wb_names) if wb_names else '',
                'Downstream Workbook Owners': ', '.join(list(set(wb_owners))) if wb_owners else '',
                'Upstream Tables': ', '.join(table_names) if table_names else '',
                'Connection Types': ', '.join(sorted(connection_types)) if connection_types else ''
            })

        return {'data': flattened, 'error': None}

    def _fetch_database_servers(self):
        """
        Fetch database servers with connection/gateway details.
        Returns server hostnames, ports, connection types, and downstream usage.
        """
        data = self.get_database_servers()
        if not data:
            return {'data': [], 'error': 'No response from Metadata API'}

        if 'databaseServers' not in data:
            return {'data': [], 'error': 'Invalid response structure'}

        flattened = []
        for server in data['databaseServers']:
            contact = server.get('contact') or {}
            downstream_ds = server.get('downstreamDatasources') or []
            downstream_wb = server.get('downstreamWorkbooks') or []

            flattened.append({
                'Server ID': server.get('id') or '',
                'Server Name': server.get('name') or '',
                'Host Name': server.get('hostName') or '',
                'Port': server.get('port') or '',
                'Connection Type': server.get('connectionType') or '',
                'Extended Connection Type': server.get('extendedConnectionType') or '',
                'Service': server.get('service') or '',
                'Description': server.get('description') or '',
                'Is Embedded': server.get('isEmbedded', False),
                'Project Name': server.get('projectName') or '',
                'Is Certified': server.get('isCertified', False),
                'Has Active Warning': server.get('hasActiveWarning', False),
                'Contact Name': contact.get('name') or '',
                'Contact Email': contact.get('email') or '',
                'Downstream Datasource Count': len(downstream_ds),
                'Downstream Datasources': ', '.join([ds.get('name', '') for ds in downstream_ds]),
                'Downstream Workbook Count': len(downstream_wb),
                'Downstream Workbooks': ', '.join([wb.get('name', '') for wb in downstream_wb])
            })

        return {'data': flattened, 'error': None}

    def _fetch_virtual_connections(self):
        """
        Fetch virtual connections (centralized connection management).
        Returns VCs with upstream/downstream dependencies.
        """
        data = self.get_virtual_connections()
        if not data:
            return {'data': [], 'error': 'No response from Metadata API'}

        if 'virtualConnections' not in data:
            return {'data': [], 'error': 'Invalid response structure or no virtual connections found'}

        flattened = []
        for vc in data['virtualConnections']:
            owner = vc.get('owner') or {}
            tables = vc.get('tables') or []
            upstream_dbs = vc.get('upstreamDatabases') or []
            upstream_tables = vc.get('upstreamTables') or []
            downstream_ds = vc.get('downstreamDatasources') or []
            downstream_wb = vc.get('downstreamWorkbooks') or []

            flattened.append({
                'VC ID': vc.get('id') or '',
                'VC LUID': vc.get('luid') or '',
                'VC Name': vc.get('name') or '',
                'Description': vc.get('description') or '',
                'Project Name': vc.get('projectName') or '',
                'Created At': vc.get('createdAt') or '',
                'Updated At': vc.get('updatedAt') or '',
                'Is Certified': vc.get('isCertified', False),
                'Has Active Warning': vc.get('hasActiveWarning', False),
                'Owner Name': owner.get('name') or '',
                'Owner Email': owner.get('email') or '',
                'Table Count': len(tables),
                'Tables': ', '.join([t.get('name', '') for t in tables]),
                'Upstream Databases': ', '.join([db.get('name', '') for db in upstream_dbs]),
                'Upstream Connection Types': ', '.join(set([db.get('connectionType', '') for db in upstream_dbs if db.get('connectionType')])),
                'Upstream Tables': ', '.join([t.get('fullName', t.get('name', '')) for t in upstream_tables]),
                'Downstream Datasource Count': len(downstream_ds),
                'Downstream Datasources': ', '.join([ds.get('name', '') for ds in downstream_ds]),
                'Downstream Workbook Count': len(downstream_wb),
                'Downstream Workbooks': ', '.join([wb.get('name', '') for wb in downstream_wb])
            })

        return {'data': flattened, 'error': None}


class RepositoryConnector:
    """
    Connector for Tableau Server PostgreSQL Repository (workgroup database).
    Provides actual usage metrics from historical_events table.

    For Tableau Server (On-Prem):
        - Connects to PostgreSQL workgroup database (read-only)
        - Database: workgroup, Port: 8060, User: readonly

    For Tableau Cloud:
        - Uses Admin Insights CSV upload
    """

    def __init__(self):
        self.connection = None
        self.available = False
        self.usage_data = {}
        self.data_source = None  # 'repository', 'admin_insights', or None

    def connect(self, host, port=8060, database='workgroup', user='readonly', password=''):
        """Connect to Tableau Server PostgreSQL repository"""
        try:
            import psycopg2
            self.connection = psycopg2.connect(
                host=host,
                port=port,
                database=database,
                user=user,
                password=password,
                connect_timeout=10
            )
            self.available = True
            self.data_source = 'repository'
            logger.info(f"Repository: Connected to {host}:{port}/{database}")
            return True
        except ImportError:
            logger.warning("Repository: psycopg2 not installed. Install with: pip install psycopg2-binary")
            return False
        except Exception as e:
            logger.error(f"Repository: Connection failed - {str(e)}")
            self.available = False
            return False

    def disconnect(self):
        """Close repository connection"""
        if self.connection:
            try:
                self.connection.close()
                logger.info("Repository: Disconnected")
            except:
                pass
        self.connection = None
        self.available = False

    def get_workbook_usage_metrics(self):
        """
        Query actual usage metrics from historical_events table.
        Returns dict keyed by workbook_id (luid) with usage stats.
        """
        if not self.available or not self.connection:
            return {}

        try:
            cursor = self.connection.cursor()

            # Query usage data from historical_events
            query = """
            SELECT
                w.luid AS workbook_id,
                w.name AS workbook_name,
                MAX(h.created_at) AS last_viewed,
                COUNT(*) FILTER (WHERE h.created_at >= NOW() - INTERVAL '7 days') AS views_7d,
                COUNT(*) FILTER (WHERE h.created_at >= NOW() - INTERVAL '30 days') AS views_30d,
                COUNT(*) FILTER (WHERE h.created_at >= NOW() - INTERVAL '90 days') AS views_90d,
                COUNT(*) AS total_views,
                COUNT(DISTINCT h.user_id) FILTER (WHERE h.created_at >= NOW() - INTERVAL '7 days') AS unique_users_7d,
                COUNT(DISTINCT h.user_id) FILTER (WHERE h.created_at >= NOW() - INTERVAL '30 days') AS unique_users_30d,
                COUNT(DISTINCT h.user_id) FILTER (WHERE h.created_at >= NOW() - INTERVAL '90 days') AS unique_users_90d,
                COUNT(DISTINCT h.user_id) AS unique_users_total
            FROM historical_events h
            JOIN historical_event_types t ON h.event_type_id = t.type_id
            JOIN views v ON h.hist_view_id = v.id
            JOIN workbooks w ON v.workbook_id = w.id
            WHERE t.name IN ('Access View', 'View Accessed')
            GROUP BY w.luid, w.name;
            """

            cursor.execute(query)
            rows = cursor.fetchall()

            usage_data = {}
            for row in rows:
                workbook_id = row[0]
                usage_data[workbook_id] = {
                    'workbook_name': row[1],
                    'last_viewed': row[2].isoformat() if row[2] else None,
                    'views_7d': row[3] or 0,
                    'views_30d': row[4] or 0,
                    'views_90d': row[5] or 0,
                    'total_views': row[6] or 0,
                    'unique_users_7d': row[7] or 0,
                    'unique_users_30d': row[8] or 0,
                    'unique_users_90d': row[9] or 0,
                    'unique_users_total': row[10] or 0,
                    'active_7d': (row[3] or 0) > 0,
                    'active_30d': (row[4] or 0) > 0,
                    'active_90d': (row[5] or 0) > 0,
                    'dormant_90d': row[2] < (datetime.now() - pd.Timedelta(days=90)) if row[2] else True
                }

            cursor.close()
            self.usage_data = usage_data
            logger.info(f"Repository: Retrieved usage data for {len(usage_data)} workbooks")
            return usage_data

        except Exception as e:
            logger.error(f"Repository: Query failed - {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return {}

    def get_view_usage_metrics(self):
        """
        Query view-level usage metrics from historical_events table.
        Returns dict keyed by view_id with usage stats.
        """
        if not self.available or not self.connection:
            return {}

        try:
            cursor = self.connection.cursor()

            query = """
            SELECT
                v.luid AS view_id,
                v.name AS view_name,
                w.luid AS workbook_id,
                MAX(h.created_at) AS last_viewed,
                COUNT(*) FILTER (WHERE h.created_at >= NOW() - INTERVAL '7 days') AS views_7d,
                COUNT(*) FILTER (WHERE h.created_at >= NOW() - INTERVAL '30 days') AS views_30d,
                COUNT(*) AS total_views,
                COUNT(DISTINCT h.user_id) FILTER (WHERE h.created_at >= NOW() - INTERVAL '30 days') AS unique_users_30d
            FROM historical_events h
            JOIN historical_event_types t ON h.event_type_id = t.type_id
            JOIN views v ON h.hist_view_id = v.id
            JOIN workbooks w ON v.workbook_id = w.id
            WHERE t.name IN ('Access View', 'View Accessed')
            GROUP BY v.luid, v.name, w.luid;
            """

            cursor.execute(query)
            rows = cursor.fetchall()

            view_usage = {}
            for row in rows:
                view_id = row[0]
                view_usage[view_id] = {
                    'view_name': row[1],
                    'workbook_id': row[2],
                    'last_viewed': row[3].isoformat() if row[3] else None,
                    'views_7d': row[4] or 0,
                    'views_30d': row[5] or 0,
                    'total_views': row[6] or 0,
                    'unique_users_30d': row[7] or 0
                }

            cursor.close()
            logger.info(f"Repository: Retrieved view usage for {len(view_usage)} views")
            return view_usage

        except Exception as e:
            logger.error(f"Repository: View query failed - {str(e)}")
            return {}

    def get_datasource_usage_metrics(self):
        """
        Query datasource usage metrics based on workbook dependencies.
        """
        if not self.available or not self.connection:
            return {}

        try:
            cursor = self.connection.cursor()

            query = """
            SELECT
                d.luid AS datasource_id,
                d.name AS datasource_name,
                COUNT(DISTINCT w.id) AS workbook_count,
                COUNT(DISTINCT w.id) FILTER (
                    WHERE EXISTS (
                        SELECT 1 FROM historical_events h
                        JOIN historical_event_types t ON h.event_type_id = t.type_id
                        JOIN views v ON h.hist_view_id = v.id
                        WHERE v.workbook_id = w.id
                        AND t.name IN ('Access View', 'View Accessed')
                        AND h.created_at >= NOW() - INTERVAL '30 days'
                    )
                ) AS active_workbooks_30d
            FROM datasources d
            LEFT JOIN data_connections dc ON d.id = dc.datasource_id
            LEFT JOIN workbooks w ON dc.owner_id = w.id AND dc.owner_type = 'Workbook'
            GROUP BY d.luid, d.name;
            """

            cursor.execute(query)
            rows = cursor.fetchall()

            ds_usage = {}
            for row in rows:
                ds_id = row[0]
                ds_usage[ds_id] = {
                    'datasource_name': row[1],
                    'workbook_count': row[2] or 0,
                    'active_workbooks_30d': row[3] or 0
                }

            cursor.close()
            logger.info(f"Repository: Retrieved datasource usage for {len(ds_usage)} datasources")
            return ds_usage

        except Exception as e:
            logger.error(f"Repository: Datasource query failed - {str(e)}")
            return {}

    def load_admin_insights_csv(self, traffic_to_views_data):
        """
        Load usage data from Admin Insights CSV export (for Tableau Cloud).

        Expected columns:
        - Item Name (workbook name)
        - View Name
        - User Email
        - Timestamp
        - Item LUID (workbook id)
        """
        try:
            if isinstance(traffic_to_views_data, str):
                # It's a file path
                df = pd.read_csv(traffic_to_views_data)
            else:
                # It's file content (from upload)
                df = pd.read_csv(BytesIO(traffic_to_views_data))

            # Normalize column names
            df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')

            # Map common column variations
            column_mapping = {
                'item_name': 'workbook_name',
                'workbook_name': 'workbook_name',
                'item_luid': 'workbook_id',
                'workbook_luid': 'workbook_id',
                'workbook_id': 'workbook_id',
                'timestamp': 'timestamp',
                'event_time': 'timestamp',
                'created_at': 'timestamp',
                'user_email': 'user_email',
                'actor_user_name': 'user_email',
                'user_name': 'user_email'
            }

            df = df.rename(columns={k: v for k, v in column_mapping.items() if k in df.columns})

            if 'timestamp' not in df.columns or 'workbook_id' not in df.columns:
                logger.error("Admin Insights: Missing required columns (timestamp, workbook_id)")
                return False

            # Convert timestamp
            df['timestamp'] = pd.to_datetime(df['timestamp'], errors='coerce')
            now = pd.Timestamp.now()

            # Group by workbook
            usage_data = {}
            for workbook_id, group in df.groupby('workbook_id'):
                last_viewed = group['timestamp'].max()
                views_7d = len(group[group['timestamp'] >= now - pd.Timedelta(days=7)])
                views_30d = len(group[group['timestamp'] >= now - pd.Timedelta(days=30)])
                views_90d = len(group[group['timestamp'] >= now - pd.Timedelta(days=90)])

                if 'user_email' in df.columns:
                    unique_users_7d = group[group['timestamp'] >= now - pd.Timedelta(days=7)]['user_email'].nunique()
                    unique_users_30d = group[group['timestamp'] >= now - pd.Timedelta(days=30)]['user_email'].nunique()
                else:
                    unique_users_7d = 0
                    unique_users_30d = 0

                usage_data[workbook_id] = {
                    'workbook_name': group['workbook_name'].iloc[0] if 'workbook_name' in group.columns else '',
                    'last_viewed': last_viewed.isoformat() if pd.notna(last_viewed) else None,
                    'views_7d': views_7d,
                    'views_30d': views_30d,
                    'views_90d': views_90d,
                    'total_views': len(group),
                    'unique_users_7d': unique_users_7d,
                    'unique_users_30d': unique_users_30d,
                    'unique_users_90d': 0,
                    'unique_users_total': group['user_email'].nunique() if 'user_email' in group.columns else 0,
                    'active_7d': views_7d > 0,
                    'active_30d': views_30d > 0,
                    'active_90d': views_90d > 0,
                    'dormant_90d': last_viewed < (now - pd.Timedelta(days=90)) if pd.notna(last_viewed) else True
                }

            self.usage_data = usage_data
            self.available = True
            self.data_source = 'admin_insights'
            logger.info(f"Admin Insights: Loaded usage data for {len(usage_data)} workbooks")
            return True

        except Exception as e:
            logger.error(f"Admin Insights: Failed to load CSV - {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return False

    def get_usage_source_label(self):
        """Get label for UI badge"""
        if self.data_source == 'repository':
            return 'Repository'
        elif self.data_source == 'admin_insights':
            return 'Admin Insights'
        else:
            return 'REST API (Estimated)'


class AdminInsightsConnector:
    """
    Connector for Tableau Cloud Admin Insights.
    Uses TSC to access built-in Admin Insights datasources for accurate usage metrics.

    Admin Insights Datasources:
    - TS Events: View access events, login events
    - TS Background Tasks: Extract refreshes, subscriptions runs
    - TS Subscriptions: Subscription configurations
    - TS Content: Workbook/datasource metadata
    """

    # Admin Insights datasource names
    # Comprehensive patterns for Admin Insights datasource discovery
    DATASOURCE_PATTERNS = {
        'events': ['ts events', 'ts_events', 'tsevent', 'ts event'],
        'content': ['site content', 'site_content', 'sitecontent', 'ts content', 'ts_content'],
        'tasks': ['background task', 'ts background', 'ts_background', 'job performance'],
        'subscriptions': ['subscription', 'ts subscription', 'ts_subscription'],
        'users': ['ts users', 'ts_users', 'tsuser', 'ts user']
    }

    def __init__(self, server, site_id):
        """Initialize with authenticated TSC server connection"""
        self.server = server
        self.site_id = site_id
        self.available = False
        self.datasources = {}
        self.usage_data = {}
        self.data_source = 'admin_insights'

    def discover_datasources(self):
        """Find Admin Insights datasources on the site"""
        try:
            logger.info("AdminInsights: Discovering datasources...")
            all_datasources, _ = self.server.datasources.get()

            # Log all datasources for debugging
            admin_insights_candidates = []
            for ds in all_datasources:
                ds_name_lower = ds.name.lower()
                project_lower = (ds.project_name or '').lower()

                # Log datasources that might be Admin Insights related
                if 'admin' in project_lower or 'insight' in project_lower or \
                   'ts ' in ds_name_lower or ds_name_lower.startswith('ts_') or \
                   'event' in ds_name_lower or 'traffic' in ds_name_lower:
                    admin_insights_candidates.append(f"{ds.name} (project: {ds.project_name})")

                # Check if this is an Admin Insights datasource (flexible matching)
                # Also check if it's in Admin Insights project
                is_admin_insights_project = 'admin insight' in project_lower or 'admin_insight' in project_lower

                for key, patterns in self.DATASOURCE_PATTERNS.items():
                    for pattern in patterns:
                        # Match if pattern is in name OR if in Admin Insights project and contains a pattern
                        if pattern in ds_name_lower or \
                           (is_admin_insights_project and any(p in ds_name_lower for p in patterns)):
                            if key not in self.datasources:  # Only store first match for each type
                                self.datasources[key] = {
                                    'id': ds.id,
                                    'name': ds.name,
                                    'project': ds.project_name,
                                    'updated_at': str(ds.updated_at) if ds.updated_at else ''
                                }
                                logger.info(f"AdminInsights: Found {key} datasource: {ds.name} (project: {ds.project_name})")
                            break

            if admin_insights_candidates:
                logger.info(f"AdminInsights: Potential candidates found: {admin_insights_candidates}")
            else:
                logger.info(f"AdminInsights: No Admin Insights candidates found among {len(all_datasources)} datasources")
                # List some datasource names for debugging
                sample_names = [ds.name for ds in all_datasources[:10]]
                logger.info(f"AdminInsights: Sample datasource names: {sample_names}")

            if 'events' in self.datasources:
                self.available = True
                logger.info(f"AdminInsights: Available - found {len(self.datasources)} datasources")
            else:
                logger.warning("AdminInsights: TS Events datasource not found on this site")
                logger.info("AdminInsights: Admin Insights may need to be enabled. Go to Tableau Cloud > Settings > Extensions > Admin Insights")

            return self.available

        except Exception as e:
            logger.error(f"AdminInsights: Error discovering datasources - {str(e)}")
            return False

    def get_usage_metrics_from_events(self):
        """
        Query TS Events datasource to compute usage metrics.
        Uses the REST API to get workbook view events.
        """
        if not self.available or 'events' not in self.datasources:
            logger.warning("AdminInsights: TS Events datasource not available")
            return {}

        try:
            events_ds_id = self.datasources['events']['id']
            logger.info(f"AdminInsights: Fetching events from datasource {events_ds_id}")

            # Get the datasource details
            datasource = self.server.datasources.get_by_id(events_ds_id)

            # Try to download and parse the datasource
            # Note: This requires the datasource to be an extract
            temp_dir = tempfile.mkdtemp()
            try:
                file_path = self.server.datasources.download(
                    events_ds_id,
                    filepath=temp_dir,
                    include_extract=True
                )
                logger.info(f"AdminInsights: Downloaded datasource to {file_path}")

                # Parse the downloaded file
                usage_data = self._parse_events_extract(file_path)
                return usage_data

            finally:
                # Cleanup temp directory
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass

        except Exception as e:
            logger.error(f"AdminInsights: Error fetching events - {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return {}

    def _parse_events_extract(self, file_path):
        """
        Parse the downloaded TS Events extract file.
        Supports .tdsx (packaged datasource) and .hyper (extract) formats.
        """
        usage_data = {}

        try:
            # Check if it's a packaged datasource (.tdsx)
            if file_path.endswith('.tdsx'):
                # Extract the .hyper file from the .tdsx
                hyper_path = self._extract_hyper_from_tdsx(file_path)
                if hyper_path:
                    usage_data = self._query_hyper_file(hyper_path)
            elif file_path.endswith('.hyper'):
                usage_data = self._query_hyper_file(file_path)
            else:
                logger.warning(f"AdminInsights: Unsupported file format: {file_path}")

        except Exception as e:
            logger.error(f"AdminInsights: Error parsing extract - {str(e)}")

        return usage_data

    def _extract_hyper_from_tdsx(self, tdsx_path):
        """Extract .hyper file from .tdsx package"""
        try:
            temp_dir = tempfile.mkdtemp()
            with zipfile.ZipFile(tdsx_path, 'r') as z:
                for name in z.namelist():
                    if name.endswith('.hyper'):
                        z.extract(name, temp_dir)
                        return os.path.join(temp_dir, name)
            return None
        except Exception as e:
            logger.error(f"AdminInsights: Error extracting hyper file - {str(e)}")
            return None

    def _query_hyper_file(self, hyper_path):
        """
        Query the Hyper file using tableauhyperapi.
        Falls back to pandas if Hyper API not available.
        """
        usage_data = {}

        try:
            # Try using tableauhyperapi
            from tableauhyperapi import HyperProcess, Connection, Telemetry, TableName

            with HyperProcess(telemetry=Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU) as hyper:
                with Connection(endpoint=hyper.endpoint, database=hyper_path) as connection:
                    # Get table names and discover schema
                    schemas = connection.catalog.get_schema_names()
                    logger.info(f"AdminInsights: Found schemas: {schemas}")

                    # Find the data table
                    table = None
                    for schema in schemas:
                        tables = connection.catalog.get_table_names(schema=schema)
                        if tables:
                            table = tables[0]
                            logger.info(f"AdminInsights: Using table {table}")
                            break

                    if not table:
                        logger.warning("AdminInsights: No tables found in Hyper file")
                        return {}

                    # Discover columns
                    table_def = connection.catalog.get_table_definition(table)
                    columns = [col.name.unescaped for col in table_def.columns]
                    columns_lower = {c.lower(): c for c in columns}
                    logger.info(f"AdminInsights: Columns found: {columns[:15]}...")

                    # Find key columns with flexible matching
                    def find_col(candidates):
                        for c in candidates:
                            c_lower = c.lower().replace(' ', '_').replace('-', '_')
                            for col_lower, col_orig in columns_lower.items():
                                col_norm = col_lower.replace(' ', '_').replace('-', '_')
                                if c_lower == col_norm or c_lower in col_norm:
                                    return col_orig
                        return None

                    # Find workbook LUID - prefer Workbook Luid over Item Luid
                    wb_luid_col = find_col(['workbook_luid', 'workbook luid', 'workbook_id'])
                    item_luid_col = find_col(['item_luid', 'item luid', 'item_id', 'content_luid'])
                    item_name_col = find_col(['item_name', 'item name', 'content_name', 'workbook_name'])
                    item_type_col = find_col(['item_type', 'item type', 'content_type'])
                    timestamp_col = find_col(['event_date', 'timestamp', 'event_time', 'created_at'])
                    event_col = find_col(['event_name', 'event_type', 'event'])
                    user_col = find_col(['actor_user_id', 'actor_user_name', 'user_id', 'user_name', 'user_luid'])

                    logger.info(f"AdminInsights: Column mapping - wb_luid: {wb_luid_col}, item_luid: {item_luid_col}, item_name: {item_name_col}, item_type: {item_type_col}, timestamp: {timestamp_col}, event: {event_col}")

                    # Determine which ID column to use
                    id_col = wb_luid_col or item_luid_col
                    if not id_col:
                        logger.warning("AdminInsights: No workbook/item LUID column found")
                        return self._query_hyper_alternative_schema(connection)

                    # Build query based on available columns
                    # If we have workbook_luid, use it directly
                    # If we only have item_luid, filter for workbooks or aggregate views by workbook
                    if wb_luid_col:
                        # Direct workbook LUID available - best case
                        group_col = f'"{wb_luid_col}"'
                        name_expr = f'MAX("{item_name_col}")' if item_name_col else "'Unknown'"
                        filter_clause = ""  # All events have workbook context
                    else:
                        # Item LUID only - filter for workbook events or use item as workbook
                        group_col = f'"{item_luid_col}"'
                        name_expr = f'MAX("{item_name_col}")' if item_name_col else "'Unknown'"
                        # Filter for workbook-type items if item_type column exists
                        if item_type_col:
                            filter_clause = f"""WHERE LOWER("{item_type_col}") IN ('workbook', 'view', 'dashboard', 'sheet')"""
                        else:
                            filter_clause = ""

                    # First, let's see what event types exist in the data
                    if event_col:
                        try:
                            event_query = f'SELECT DISTINCT "{event_col}" FROM {table} LIMIT 50'
                            event_result = connection.execute_query(event_query)
                            event_types = [str(row[0]) for row in event_result if row[0]]
                            logger.info(f"AdminInsights: Event types found: {event_types[:20]}")
                        except Exception as e:
                            logger.warning(f"AdminInsights: Could not query event types: {e}")
                            event_types = []

                    # Add event type filter - be more inclusive
                    # Tableau Cloud event names: "Access View", "Access Published View", "Access Workbook", "Sign In", etc.
                    if event_col and filter_clause:
                        filter_clause += f""" AND (LOWER("{event_col}") LIKE '%access%' OR LOWER("{event_col}") LIKE '%view%' OR LOWER("{event_col}") LIKE '%open%' OR LOWER("{event_col}") LIKE '%load%')"""
                    elif event_col:
                        filter_clause = f"""WHERE (LOWER("{event_col}") LIKE '%access%' OR LOWER("{event_col}") LIKE '%view%' OR LOWER("{event_col}") LIKE '%open%' OR LOWER("{event_col}") LIKE '%load%')"""

                    # Build time-based aggregation
                    if timestamp_col:
                        time_agg = f"""
                            MAX("{timestamp_col}") as last_viewed,
                            COUNT(*) FILTER (WHERE "{timestamp_col}" >= CURRENT_DATE - INTERVAL '7' DAY) as views_7d,
                            COUNT(*) FILTER (WHERE "{timestamp_col}" >= CURRENT_DATE - INTERVAL '30' DAY) as views_30d,
                            COUNT(*) FILTER (WHERE "{timestamp_col}" >= CURRENT_DATE - INTERVAL '90' DAY) as views_90d,
                            COUNT(*) as total_views
                        """
                        if user_col:
                            user_agg = f", COUNT(DISTINCT \"{user_col}\") FILTER (WHERE \"{timestamp_col}\" >= CURRENT_DATE - INTERVAL '30' DAY) as unique_users_30d"
                        else:
                            user_agg = ", 0 as unique_users_30d"
                    else:
                        time_agg = "NULL as last_viewed, COUNT(*) as views_7d, COUNT(*) as views_30d, COUNT(*) as views_90d, COUNT(*) as total_views"
                        if user_col:
                            user_agg = f", COUNT(DISTINCT \"{user_col}\") as unique_users_30d"
                        else:
                            user_agg = ", 0 as unique_users_30d"

                    query = f"""
                    SELECT
                        {group_col} as workbook_id,
                        {name_expr} as workbook_name,
                        {time_agg}
                        {user_agg}
                    FROM {table}
                    {filter_clause}
                    GROUP BY {group_col}
                    """

                    logger.info(f"AdminInsights: Executing query...")

                    try:
                        result = connection.execute_query(query)
                        for row in result:
                            wb_id = str(row[0]) if row[0] else ''
                            if wb_id:
                                last_viewed = row[2]
                                views_7d = row[3] or 0
                                views_30d = row[4] or 0
                                views_90d = row[5] or 0
                                total_views = row[6] or 0
                                unique_users_30d = row[7] or 0

                                usage_data[wb_id] = {
                                    'workbook_name': row[1] or '',
                                    'last_viewed': last_viewed.isoformat() if hasattr(last_viewed, 'isoformat') else str(last_viewed) if last_viewed else None,
                                    'views_7d': views_7d,
                                    'views_30d': views_30d,
                                    'views_90d': views_90d,
                                    'total_views': total_views,
                                    'unique_users_30d': unique_users_30d,
                                    'active_7d': views_7d > 0,
                                    'active_30d': views_30d > 0,
                                    'active_90d': views_90d > 0,
                                    'dormant_90d': views_90d == 0
                                }
                        logger.info(f"AdminInsights: Extracted usage for {len(usage_data)} items")

                        # If we used item_luid and have workbooks data, try to merge view events to workbooks
                        if not wb_luid_col and item_type_col:
                            usage_data = self._aggregate_views_to_workbooks(connection, table, columns_lower, usage_data)

                    except Exception as qe:
                        logger.warning(f"AdminInsights: Query failed - {str(qe)}, trying alternative schema")
                        import traceback
                        logger.debug(traceback.format_exc())
                        usage_data = self._query_hyper_alternative_schema(connection)

        except ImportError:
            logger.warning("AdminInsights: tableauhyperapi not installed. Install with: pip install tableauhyperapi")
            # Fall back to simpler parsing if possible
            usage_data = self._parse_hyper_without_api(hyper_path)

        except Exception as e:
            logger.error(f"AdminInsights: Error querying Hyper file - {str(e)}")
            import traceback
            logger.error(traceback.format_exc())

        return usage_data

    def _aggregate_views_to_workbooks(self, connection, table, columns_lower, item_usage):
        """
        If we have view-level data, try to aggregate it to workbook level.
        This is needed when the events data is keyed by view but we need workbook-level metrics.
        """
        try:
            # Check if there's a workbook LUID column we can use
            def find_col(candidates):
                for c in candidates:
                    c_lower = c.lower().replace(' ', '_').replace('-', '_')
                    for col_lower, col_orig in columns_lower.items():
                        col_norm = col_lower.replace(' ', '_').replace('-', '_')
                        if c_lower == col_norm:
                            return col_orig
                return None

            wb_luid_col = find_col(['workbook_luid', 'workbook luid'])
            if not wb_luid_col:
                # No workbook LUID column, return item-level data as-is
                return item_usage

            # Query for workbook-level aggregation
            item_luid_col = find_col(['item_luid', 'item luid'])
            timestamp_col = find_col(['event_date', 'timestamp', 'event_time'])
            user_col = find_col(['actor_user_id', 'actor_user_name', 'user_luid'])

            if not timestamp_col:
                return item_usage

            # Build unique users clause separately to avoid f-string backslash issue
            unique_users_clause = f", COUNT(DISTINCT \"{user_col}\") FILTER (WHERE \"{timestamp_col}\" >= CURRENT_DATE - INTERVAL '30' DAY) as unique_users_30d" if user_col else ", 0 as unique_users_30d"

            query = f"""
            SELECT
                "{wb_luid_col}" as workbook_id,
                MAX("{timestamp_col}") as last_viewed,
                COUNT(*) FILTER (WHERE "{timestamp_col}" >= CURRENT_DATE - INTERVAL '7' DAY) as views_7d,
                COUNT(*) FILTER (WHERE "{timestamp_col}" >= CURRENT_DATE - INTERVAL '30' DAY) as views_30d,
                COUNT(*) FILTER (WHERE "{timestamp_col}" >= CURRENT_DATE - INTERVAL '90' DAY) as views_90d,
                COUNT(*) as total_views
                {unique_users_clause}
            FROM {table}
            WHERE "{wb_luid_col}" IS NOT NULL
            GROUP BY "{wb_luid_col}"
            """

            result = connection.execute_query(query)
            workbook_usage = {}

            for row in result:
                wb_id = str(row[0]) if row[0] else ''
                if wb_id:
                    last_viewed = row[1]
                    views_7d = row[2] or 0
                    views_30d = row[3] or 0
                    views_90d = row[4] or 0
                    total_views = row[5] or 0
                    unique_users_30d = row[6] or 0

                    workbook_usage[wb_id] = {
                        'last_viewed': last_viewed.isoformat() if hasattr(last_viewed, 'isoformat') else str(last_viewed) if last_viewed else None,
                        'views_7d': views_7d,
                        'views_30d': views_30d,
                        'views_90d': views_90d,
                        'total_views': total_views,
                        'unique_users_30d': unique_users_30d,
                        'active_7d': views_7d > 0,
                        'active_30d': views_30d > 0,
                        'active_90d': views_90d > 0,
                        'dormant_90d': views_90d == 0
                    }

            if workbook_usage:
                logger.info(f"AdminInsights: Aggregated to {len(workbook_usage)} workbook-level records")
                return workbook_usage

        except Exception as e:
            logger.warning(f"AdminInsights: Could not aggregate to workbook level - {str(e)}")

        return item_usage

    def _query_hyper_alternative_schema(self, connection):
        """Try alternative column names for Admin Insights"""
        usage_data = {}

        # Try different possible column name variations
        column_variations = [
            {
                'workbook_id': 'Item Luid',
                'workbook_name': 'Item Name',
                'timestamp': 'Timestamp',
                'event_type': 'Event Type',
                'user': 'Actor User Name'
            },
            {
                'workbook_id': 'item_luid',
                'workbook_name': 'item_name',
                'timestamp': 'timestamp',
                'event_type': 'event_type',
                'user': 'actor_user_name'
            },
            {
                'workbook_id': 'Workbook LUID',
                'workbook_name': 'Workbook Name',
                'timestamp': 'Event Time',
                'event_type': 'Event Type Name',
                'user': 'User Email'
            }
        ]

        for cols in column_variations:
            try:
                query = f"""
                SELECT
                    "{cols['workbook_id']}" as workbook_id,
                    "{cols['workbook_name']}" as workbook_name,
                    "{cols['timestamp']}" as event_time,
                    "{cols['user']}" as user_name
                FROM "Extract"."Extract"
                WHERE "{cols['event_type']}" LIKE '%View%'
                OR "{cols['event_type']}" LIKE '%Access%'
                """

                result = connection.execute_query(query)
                events = []
                for row in result:
                    events.append({
                        'workbook_id': str(row[0]) if row[0] else '',
                        'workbook_name': row[1] or '',
                        'timestamp': row[2],
                        'user': row[3] or ''
                    })

                if events:
                    usage_data = self._aggregate_events(events)
                    logger.info(f"AdminInsights: Aggregated {len(events)} events for {len(usage_data)} workbooks")
                    break

            except Exception as e:
                continue

        return usage_data

    def _aggregate_events(self, events):
        """Aggregate raw events into usage metrics per workbook"""
        from collections import defaultdict

        now = datetime.now()
        day_7_ago = now - pd.Timedelta(days=7)
        day_30_ago = now - pd.Timedelta(days=30)
        day_90_ago = now - pd.Timedelta(days=90)

        workbook_events = defaultdict(list)
        for event in events:
            wb_id = event['workbook_id']
            if wb_id:
                workbook_events[wb_id].append(event)

        usage_data = {}
        for wb_id, wb_events in workbook_events.items():
            timestamps = []
            users_7d = set()
            users_30d = set()
            users_90d = set()

            for event in wb_events:
                ts = event['timestamp']
                if ts:
                    if isinstance(ts, str):
                        try:
                            ts = pd.to_datetime(ts)
                        except:
                            continue
                    timestamps.append(ts)

                    user = event.get('user', '')
                    if ts >= day_7_ago:
                        users_7d.add(user)
                    if ts >= day_30_ago:
                        users_30d.add(user)
                    if ts >= day_90_ago:
                        users_90d.add(user)

            if timestamps:
                last_viewed = max(timestamps)
                views_7d = sum(1 for ts in timestamps if ts >= day_7_ago)
                views_30d = sum(1 for ts in timestamps if ts >= day_30_ago)
                views_90d = sum(1 for ts in timestamps if ts >= day_90_ago)

                usage_data[wb_id] = {
                    'workbook_name': wb_events[0].get('workbook_name', ''),
                    'last_viewed': last_viewed.isoformat() if last_viewed else None,
                    'views_7d': views_7d,
                    'views_30d': views_30d,
                    'views_90d': views_90d,
                    'total_views': len(timestamps),
                    'unique_users_7d': len(users_7d),
                    'unique_users_30d': len(users_30d),
                    'unique_users_90d': len(users_90d),
                    'active_7d': views_7d > 0,
                    'active_30d': views_30d > 0,
                    'active_90d': views_90d > 0,
                    'dormant_90d': last_viewed < day_90_ago if last_viewed else True
                }

        return usage_data

    def _parse_hyper_without_api(self, hyper_path):
        """Fallback parsing without Hyper API - limited functionality"""
        logger.warning("AdminInsights: Hyper API not available, usage metrics will be limited")
        return {}

    def get_background_task_metrics(self):
        """Get refresh metrics from TS Background Tasks datasource"""
        if 'tasks' not in self.datasources:
            return {}

        try:
            tasks_ds_id = self.datasources['tasks']['id']
            logger.info(f"AdminInsights: Fetching background tasks from {tasks_ds_id}")

            # Similar download and parse logic
            temp_dir = tempfile.mkdtemp()
            try:
                file_path = self.server.datasources.download(
                    tasks_ds_id,
                    filepath=temp_dir,
                    include_extract=True
                )

                # Parse for refresh metrics
                return self._parse_background_tasks(file_path)
            finally:
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass

        except Exception as e:
            logger.error(f"AdminInsights: Error fetching background tasks - {str(e)}")
            return {}

    def _parse_background_tasks(self, file_path):
        """Parse background tasks for refresh metrics"""
        task_metrics = {}

        try:
            if file_path.endswith('.tdsx'):
                hyper_path = self._extract_hyper_from_tdsx(file_path)
                if not hyper_path:
                    return {}
            else:
                hyper_path = file_path

            from tableauhyperapi import HyperProcess, Connection, Telemetry

            with HyperProcess(telemetry=Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU) as hyper:
                with Connection(endpoint=hyper.endpoint, database=hyper_path) as connection:
                    query = """
                    SELECT
                        "Item Luid" as item_id,
                        MAX("Completed At") as last_refresh,
                        COUNT(*) FILTER (WHERE "Task Status" = 'Failed' AND "Completed At" >= CURRENT_DATE - INTERVAL '30' DAY) as failures_30d
                    FROM "Extract"."Extract"
                    WHERE "Task Type" LIKE '%Extract%'
                    GROUP BY "Item Luid"
                    """

                    try:
                        result = connection.execute_query(query)
                        for row in result:
                            item_id = str(row[0]) if row[0] else ''
                            if item_id:
                                task_metrics[item_id] = {
                                    'last_refresh': row[1].isoformat() if row[1] else None,
                                    'refresh_failures_30d': row[2] or 0
                                }
                    except:
                        pass

        except ImportError:
            logger.warning("AdminInsights: tableauhyperapi not available for background tasks")
        except Exception as e:
            logger.error(f"AdminInsights: Error parsing background tasks - {str(e)}")

        return task_metrics

    def get_subscription_metrics(self):
        """Get subscription counts from TS Subscriptions datasource"""
        if 'subscriptions' not in self.datasources:
            return {}

        try:
            subs_ds_id = self.datasources['subscriptions']['id']
            logger.info(f"AdminInsights: Fetching subscriptions from {subs_ds_id}")

            temp_dir = tempfile.mkdtemp()
            try:
                file_path = self.server.datasources.download(
                    subs_ds_id,
                    filepath=temp_dir,
                    include_extract=True
                )

                return self._parse_subscriptions(file_path)
            finally:
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass

        except Exception as e:
            logger.error(f"AdminInsights: Error fetching subscriptions - {str(e)}")
            return {}

    def _parse_subscriptions(self, file_path):
        """Parse subscriptions datasource"""
        subscription_metrics = {}

        try:
            if file_path.endswith('.tdsx'):
                hyper_path = self._extract_hyper_from_tdsx(file_path)
                if not hyper_path:
                    return {}
            else:
                hyper_path = file_path

            from tableauhyperapi import HyperProcess, Connection, Telemetry

            with HyperProcess(telemetry=Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU) as hyper:
                with Connection(endpoint=hyper.endpoint, database=hyper_path) as connection:
                    query = """
                    SELECT
                        "Workbook Luid" as workbook_id,
                        COUNT(*) as subscriber_count
                    FROM "Extract"."Extract"
                    WHERE "Workbook Luid" IS NOT NULL
                    GROUP BY "Workbook Luid"
                    """

                    try:
                        result = connection.execute_query(query)
                        for row in result:
                            wb_id = str(row[0]) if row[0] else ''
                            if wb_id:
                                subscription_metrics[wb_id] = {
                                    'subscriber_count': row[1] or 0
                                }
                    except:
                        pass

        except ImportError:
            logger.warning("AdminInsights: tableauhyperapi not available for subscriptions")
        except Exception as e:
            logger.error(f"AdminInsights: Error parsing subscriptions - {str(e)}")

        return subscription_metrics

    def get_all_usage_metrics(self):
        """
        Get comprehensive usage metrics from all Admin Insights sources.
        Returns a combined dictionary keyed by workbook_id.
        """
        if not self.available:
            if not self.discover_datasources():
                return {}

        # Get usage from events
        usage_data = self.get_usage_metrics_from_events()

        # Enrich with background task data
        task_metrics = self.get_background_task_metrics()
        for wb_id, task_data in task_metrics.items():
            if wb_id in usage_data:
                usage_data[wb_id].update(task_data)
            else:
                usage_data[wb_id] = task_data

        # Enrich with subscription data
        sub_metrics = self.get_subscription_metrics()
        for wb_id, sub_data in sub_metrics.items():
            if wb_id in usage_data:
                usage_data[wb_id].update(sub_data)
            else:
                usage_data[wb_id] = sub_data

        self.usage_data = usage_data
        logger.info(f"AdminInsights: Combined usage data for {len(usage_data)} workbooks")
        return usage_data


class XMLParser:
    """Parser for Tableau .twb/.twbx/.tds/.tdsx files"""

    @staticmethod
    def extract_from_twbx(file_path):
        """Extract .twb from .twbx (zip file)"""
        temp_dir = tempfile.mkdtemp()
        try:
            with zipfile.ZipFile(file_path, 'r') as z:
                for name in z.namelist():
                    if name.endswith('.twb'):
                        z.extract(name, temp_dir)
                        return os.path.join(temp_dir, name), temp_dir
            return None, temp_dir
        except Exception as e:
            logger.error(f"Error extracting twbx: {str(e)}")
            return None, temp_dir

    @staticmethod
    def extract_from_tdsx(file_path):
        """Extract .tds from .tdsx (zip file)"""
        temp_dir = tempfile.mkdtemp()
        try:
            with zipfile.ZipFile(file_path, 'r') as z:
                for name in z.namelist():
                    if name.endswith('.tds'):
                        z.extract(name, temp_dir)
                        return os.path.join(temp_dir, name), temp_dir
            return None, temp_dir
        except Exception as e:
            logger.error(f"Error extracting tdsx: {str(e)}")
            return None, temp_dir

    @staticmethod
    def parse_connection_info(datasource):
        """Parse detailed connection information including type detection"""
        connections = []
        ds_name = datasource.get('name', '')
        ds_caption = datasource.get('caption', ds_name)

        connection = datasource.find('.//connection')
        if connection is not None:
            conn_class = connection.get('class', '')

            # Detect connection type
            if conn_class in ['sqlproxy', 'remote', 'dataengine']:
                conn_type = 'published'
                repo_loc = datasource.find('repository-location')      # sibling of <connection> in real workbooks
                if repo_loc is None:
                    repo_loc = connection.find('.//repository-location')
                if repo_loc is not None:
                    connections.append({
                        'ds_name': ds_name,
                        'ds_caption': ds_caption,
                        'connection_class': conn_class,
                        'connection_type': conn_type,
                        'server': repo_loc.get('server', ''),
                        'repository_location': repo_loc.get('datasource', ''),
                        'dbname': '',
                        'schema': '',
                        'authentication': ''
                    })
                else:
                    connections.append({
                        'ds_name': ds_name,
                        'ds_caption': ds_caption,
                        'connection_class': conn_class,
                        'connection_type': conn_type,
                        'server': connection.get('server', ''),
                        'dbname': connection.get('dbname', ''),
                        'schema': connection.get('schema', ''),
                        'authentication': connection.get('authentication', '')
                    })

            elif conn_class == 'federated':
                # Federated connections have multiple named connections
                for named_conn in datasource.findall('.//named-connection'):
                    inner_conn = named_conn.find('.//connection')
                    if inner_conn is not None:
                        connections.append({
                            'ds_name': ds_name,
                            'ds_caption': ds_caption,
                            'conn_name': named_conn.get('name', ''),
                            'connection_class': inner_conn.get('class', ''),
                            'connection_type': 'federated-named',
                            'server': inner_conn.get('server', ''),
                            'dbname': inner_conn.get('dbname', ''),
                            'schema': inner_conn.get('schema', ''),
                            'authentication': inner_conn.get('authentication', '')
                        })
            else:
                # Embedded connection
                connections.append({
                    'ds_name': ds_name,
                    'ds_caption': ds_caption,
                    'connection_class': conn_class,
                    'connection_type': 'embedded',
                    'server': connection.get('server', ''),
                    'dbname': connection.get('dbname', ''),
                    'schema': connection.get('schema', ''),
                    'authentication': connection.get('authentication', '')
                })
        else:
            # Check for named connections without parent connection
            for named_conn in datasource.findall('.//named-connection'):
                inner_conn = named_conn.find('.//connection')
                if inner_conn is not None:
                    connections.append({
                        'ds_name': ds_name,
                        'ds_caption': ds_caption,
                        'conn_name': named_conn.get('name', ''),
                        'connection_class': inner_conn.get('class', ''),
                        'connection_type': 'named',
                        'server': inner_conn.get('server', ''),
                        'dbname': inner_conn.get('dbname', ''),
                        'schema': inner_conn.get('schema', ''),
                        'authentication': inner_conn.get('authentication', '')
                    })

        return connections

    @staticmethod
    def parse_tables(datasource, ds_name, ds_caption):
        """Extract table/relation information from datasource"""
        tables = []
        unique_tables = set()

        for relation in datasource.findall('.//relation'):
            table_name = relation.get('name', '')
            table_type = relation.get('type', '')

            if table_type == 'text':
                continue  # Skip custom SQL, handled separately

            table_key = (ds_caption, table_name)
            if table_key not in unique_tables and table_name:
                unique_tables.add(table_key)
                # Get columns in this relation
                columns = [col.get('name', '') for col in relation.findall('.//column')]
                tables.append({
                    'datasource_name': ds_caption,
                    'datasource_internal': ds_name,
                    'table_name': table_name,
                    'table_type': table_type,
                    'columns': json.dumps(columns) if columns else ''
                })

        return tables

    # ---- Tables used per workbook / data source -------------------------------------------------

    CONNECTION_LABELS = lineage_engine.CONNECTION_LABELS

    _IDENT = r'(?:\[[^\]]+\]|"[^"]+"|`[^`]+`|[\w#@$]+)'
    _SQL_TABLE_REF = re.compile(r'\b(?:FROM|JOIN)\s+((?:%s\.){0,3}%s)' % (_IDENT, _IDENT), re.IGNORECASE)
    _SQL_CTE = re.compile(r'(?:\bWITH\s+(?:RECURSIVE\s+)?|,\s*)(%s)\s+AS\s*\(' % _IDENT, re.IGNORECASE)
    _SQL_NOT_TABLES = {'select', 'lateral', 'unnest', 'values', 'dual', 'table', 'only', 'final'}

    @staticmethod
    def split_table_reference(ref):
        """'[db].[dbo].[v_fact]' -> ('db', 'dbo', 'v_fact'); missing parts are ''."""
        ref = (ref or '').strip()
        parts = re.findall(r'\[([^\]]+)\]', ref)
        if not parts:
            parts = [p.strip('"`') for p in re.split(r'\.(?=(?:[^"]*"[^"]*")*[^"]*$)', ref) if p.strip()]
        parts = [p.strip() for p in parts if p.strip()]
        if not parts:
            return '', '', ''
        table = parts[-1]
        if table.endswith('$'):  # Excel sheet reference: [Sheet1$]
            table = table[:-1]
        schema = parts[-2] if len(parts) >= 2 else ''
        database = parts[-3] if len(parts) >= 3 else ''
        return database, schema, table

    @staticmethod
    def parse_sql_tables(sql):
        """Best-effort list of tables referenced by FROM/JOIN in a custom SQL text (CTE names excluded)."""
        if not sql:
            return []
        text = re.sub(r'--[^\n]*', ' ', sql)
        text = re.sub(r'/\*.*?\*/', ' ', text, flags=re.DOTALL)
        ctes = {c.strip('[]"`').lower() for c in XMLParser._SQL_CTE.findall(text)}
        found = []
        for ref in XMLParser._SQL_TABLE_REF.findall(text):
            name = ref.strip()
            bare = name.split('.')[-1].strip('[]"`').lower()
            if bare in XMLParser._SQL_NOT_TABLES or (bare in ctes and '.' not in name):
                continue
            if name not in found:
                found.append(name)
        return found

    @staticmethod
    def parse_object_tables(datasource_elements, object_id, object_name, object_type):
        """
        One row per (data source, connection, table) used by a workbook / published data source:
        connection type, server string, database, schema, table. Handles federated (multi-connection),
        legacy, custom SQL (tables parsed from the SQL text), stored procedures and published data
        sources. Tables that only belong to a Hyper extract are ignored.
        """
        rows, seen = [], set()
        source = 'Workbook XML' if object_type == 'Workbook' else 'Data Source XML'

        def add(ds_caption, conn, table_kind, table_name, alias='', ref='', database='', schema=''):
            cclass = conn.get('class', '') if conn is not None else ''
            server = ''
            if conn is not None:
                server = (conn.get('server') or conn.get('filename') or conn.get('directory')
                          or conn.get('project') or conn.get('url') or '')
                database = database or conn.get('dbname') or conn.get('catalog') or conn.get('dataset') or ''
                schema = schema or conn.get('schema') or ''
            key = (ds_caption, cclass, server, database, schema, table_kind, table_name, alias)
            if key in seen:
                return
            seen.add(key)
            rows.append({
                'object_id': object_id,
                'object_name': object_name,
                'object_type': object_type,
                'connection_type': XMLParser.CONNECTION_LABELS.get(cclass, cclass),
                'connection_server': server,
                'port': (conn.get('port') or '') if conn is not None else '',
                'warehouse': (conn.get('warehouse') or '') if conn is not None else '',
                'catalog': ((conn.get('catalog') or conn.get('dbname') if cclass == 'databricks' else conn.get('catalog'))
                            or '') if conn is not None else '',
                'authentication': (conn.get('authentication') or conn.get('oauth-config-id') and 'oauth' or '')
                                  if conn is not None else '',
                'database': database,
                'schema': schema,
                'table_name': table_name,
                'full_table_name': '.'.join(x for x in (schema, table_name) if x) if table_kind in (
                    'Table', 'Stored Procedure', 'Custom SQL (table parsed from SQL)') else table_name,
                'table_kind': table_kind,
                'datasource_name': ds_caption,
                'connection_class': cclass,
                'table_alias': alias,
                'table_reference': ref,
                'source': source,
            })

        for ds in datasource_elements:
            ds_internal = ds.get('name', '')
            if ds_internal == 'Parameters':
                continue
            ds_caption = ds.get('caption', ds_internal)

            # Relations inside <extract> are the Hyper extract's own tables, not the source tables
            in_extract = set()
            for ex in ds.iter('extract'):
                for el in ex.iter():
                    in_extract.add(id(el))

            named = {}
            for nc in ds.iter('named-connection'):
                inner = nc.find('.//connection')
                if inner is not None:
                    named[nc.get('name', '')] = inner
            top = ds.find('connection')
            if top is None and len(named) == 1:
                top = next(iter(named.values()))

            # A workbook that uses a published data source
            if top is not None and top.get('class') == 'sqlproxy':
                repo = ds.find('repository-location')          # sibling of <connection> in real workbooks
                if repo is None:
                    repo = top.find('.//repository-location')
                published_name = (repo.get('id') if repo is not None and repo.get('id') else ds_caption)
                add(ds_caption, top, 'Published Data Source', published_name,
                    ref=repo.get('derived-from', '') if repo is not None else '',
                    database=top.get('dbname', ''))
                continue

            for rel in ds.iter('relation'):
                if id(rel) in in_extract:
                    continue
                rtype = rel.get('type', '')
                if rel.get('join') or rtype in ('join', 'collection', 'union'):
                    continue  # containers - the tables inside are visited on their own

                link = rel.get('connection', '')
                if link and link in named:
                    conn = named[link]
                elif top is not None and top.get('class') != 'federated':
                    conn = top
                else:
                    conn = next(iter(named.values()), top)
                if conn is not None and conn.get('class') == 'sqlproxy':
                    continue

                ref = rel.get('table', '')
                alias = rel.get('name', '')

                if rtype == 'text':
                    sql = (rel.text or '').strip()
                    add(ds_caption, conn, 'Custom SQL', alias or 'Custom SQL', alias=alias, ref='')
                    for tref in XMLParser.parse_sql_tables(sql):
                        db, sch, tbl = XMLParser.split_table_reference(tref)
                        add(ds_caption, conn, 'Custom SQL (table parsed from SQL)', tbl, alias=alias,
                            ref=tref, database=db, schema=sch)
                elif rtype == 'stored-proc':
                    db, sch, proc = XMLParser.split_table_reference(ref or alias)
                    add(ds_caption, conn, 'Stored Procedure', proc, alias=alias, ref=ref, database=db, schema=sch)
                else:
                    db, sch, tbl = XMLParser.split_table_reference(ref or alias)
                    if not tbl:
                        continue
                    add(ds_caption, conn, 'Table', tbl, alias=alias if alias != tbl else '', ref=ref,
                        database=db, schema=sch)

        return rows

    @staticmethod
    def parse_actions(root, workbook_name, workbook_id):
        """Extract dashboard actions (filter, highlight, URL, etc.) with URL detection"""
        actions = []
        for action in root.findall('.//action'):
            action_name = action.get('name', '')
            action_caption = action.get('caption', '')

            # Get source worksheets
            source_sheets = []
            for source in action.findall('.//source'):
                ws = source.get('worksheet', '')
                if ws:
                    source_sheets.append(ws)

            # Get target worksheets
            target_sheets = []
            for target in action.findall('.//target'):
                ws = target.get('worksheet', '')
                if ws:
                    target_sheets.append(ws)

            # Get activation type
            activation = action.find('.//activation')
            activation_type = activation.get('type', 'select') if activation is not None else 'select'

            # Detect action type (URL vs Filter/Highlight)
            link = action.find('.//link')
            if link is not None:
                action_type = 'URL'
                url_target = link.get('target', '')
            else:
                action_type = 'Filter/Highlight'
                url_target = ''

            actions.append({
                'object_id': workbook_id,
                'object_name': workbook_name,
                'action_name': action_name,
                'action_caption': action_caption,
                'action_type': action_type,
                'activation_type': activation_type,
                'source_sheets': json.dumps(source_sheets),
                'target_sheets': json.dumps(target_sheets),
                'url_target': url_target
            })

        return actions

    @staticmethod
    def parse_column_info(column, ds_name, ds_caption, tables, dbname, connection_type):
        """Extract detailed column metadata including Description, Aggregation, Hidden"""
        col_name = column.get('name', '')
        col_datatype = column.get('datatype', '')
        col_role = column.get('role', '')
        col_description = column.get('description', '')
        col_aggregation = column.get('aggregation', '')
        col_default_value = column.get('default-value', '')
        col_hidden = column.get('hidden', 'false')

        # Check for calculation formula
        calc_elem = column.find('.//calculation')
        formula = calc_elem.get('formula', '') if calc_elem is not None else ''

        results = []

        if tables:
            for table in tables:
                col_data = {
                    'datasource_caption': ds_caption,
                    'datasource_name': ds_name,
                    'db_name': dbname,
                    'connection_type': connection_type,
                    'from_table': table.get('table_name', ''),
                    'table_type': table.get('table_type', ''),
                    'field_name': col_name,
                    'field_type': col_datatype,
                    'measure_or_attribute': col_role,
                    'formula': formula,
                    'description': col_description,
                    'aggregation': col_aggregation,
                    'default_value': col_default_value,
                    'hidden': col_hidden
                }
                results.append(col_data)
        else:
            # Handle case where no tables are found
            col_data = {
                'datasource_caption': ds_caption,
                'datasource_name': ds_name,
                'db_name': dbname,
                'connection_type': connection_type,
                'from_table': '',
                'table_type': '',
                'field_name': col_name,
                'field_type': col_datatype,
                'measure_or_attribute': col_role,
                'formula': formula,
                'description': col_description,
                'aggregation': col_aggregation,
                'default_value': col_default_value,
                'hidden': col_hidden
            }
            results.append(col_data)

        return results

    @staticmethod
    def parse_custom_visuals(root, workbook_name, workbook_id):
        """Extract custom/non-standard visualizations from worksheets with full mark analysis"""
        custom_visuals = []
        standard_marks = {'text', 'shape', 'line', 'area', 'bar', 'circle', 'square',
                         'cross', 'triangle', 'automatic', 'map', 'pie', 'gantt', 'polygon'}

        for worksheet in root.findall('.//worksheet'):
            ws_name = worksheet.get('name', '')

            # Find all mark types used in worksheet
            mark_types = []
            for mark in worksheet.findall('.//mark'):
                mark_class = mark.get('class', '')
                if mark_class:
                    mark_types.append(mark_class)

            # Identify custom vs standard marks
            custom_marks = [m for m in mark_types if m.lower() not in standard_marks]
            standard_used = [m for m in mark_types if m.lower() in standard_marks]

            if mark_types:
                custom_visuals.append({
                    'object_id': workbook_id,
                    'object_name': workbook_name,
                    'worksheet_name': ws_name,
                    'all_mark_types': json.dumps(list(set(mark_types))),
                    'standard_marks': json.dumps(list(set(standard_used))) if standard_used else '',
                    'custom_marks': json.dumps(list(set(custom_marks))) if custom_marks else '',
                    'has_custom_visuals': 'Yes' if custom_marks else 'No'
                })

        return custom_visuals

    @staticmethod
    def _parse_filters_for_sheet(sheet):
        """Extract filters for a single sheet (helper for parallel processing)"""
        filters = []
        for filter_elem in sheet.findall('.//filter'):
            filter_column = filter_elem.get('column', '')
            for groupfilter in filter_elem.findall('.//groupfilter[@function="member"]'):
                member_value = groupfilter.get('member', '')
                filters.append({
                    'sheet_name': sheet.get('name', ''),
                    'field_name': filter_column,
                    'filtered_item': member_value
                })
        return filters

    @staticmethod
    def parse_filters_parallel(root, workbook_name, workbook_id):
        """Parallelize the extraction of sheet filters using ThreadPoolExecutor"""
        sheets = root.findall('.//worksheet')
        all_filters = []

        with ThreadPoolExecutor(max_workers=4) as executor:
            results = executor.map(XMLParser._parse_filters_for_sheet, sheets)
            for result in results:
                for f in result:
                    f['object_id'] = workbook_id
                    f['object_name'] = workbook_name
                    all_filters.append(f)

        return all_filters

    @staticmethod
    def parse_parameters_enhanced(root, workbook_name, workbook_id):
        """Extract parameters with domain type, current values, and range constraints"""
        parameters = []
        for param in root.findall(".//column[@param-domain-type]"):
            param_name = param.get('name', '').strip('[]')
            param_caption = param.get('caption', param_name)
            param_type = param.get('param-domain-type', '')
            data_type = param.get('datatype', '')

            # Get current value from calculation
            calc = param.find('.//calculation')
            current_value = calc.get('formula', '') if calc is not None else ''

            # Get allowable values
            allowable = []
            for member in param.findall('.//member'):
                allowable.append(member.get('value', ''))

            # Get range constraints if any
            range_elem = param.find('.//range')
            min_value = range_elem.get('min', '') if range_elem is not None else ''
            max_value = range_elem.get('max', '') if range_elem is not None else ''
            step_size = range_elem.get('granularity', '') if range_elem is not None else ''

            parameters.append({
                'object_id': workbook_id,
                'object_name': workbook_name,
                'parameter_name': param_caption,
                'internal_name': param_name,
                'param_domain_type': param_type,
                'datatype': data_type,
                'current_value': current_value,
                'allowable_values': json.dumps(allowable) if allowable else '',
                'min_value': min_value,
                'max_value': max_value,
                'step_size': step_size
            })

        return parameters

    @staticmethod
    def parse_workbook(file_path, workbook_name, workbook_id):
        """Parse .twb file and extract all metadata"""
        result = {
            'custom_sql': [],
            'calculations': [],
            'connections': [],
            'parameters': [],
            'filters': [],
            'sets': [],
            'groups': [],
            'worksheets': [],
            'dashboards': [],
            'relationships': [],
            'tables': [],
            'actions': [],
            'columns': [],
            'custom_visuals': [],
            'parameters_enhanced': [],
            'object_tables': []
        }

        try:
            tree = ET.parse(file_path)
            root = tree.getroot()

            # Tables used by this workbook, with their connection type / server
            result['object_tables'] = XMLParser.parse_object_tables(
                root.findall('.//datasource'), workbook_id, workbook_name, 'Workbook')

            # Build dashboard-worksheet map
            dashboard_map = {}
            for dashboard in root.findall('.//dashboard'):
                dash_name = dashboard.get('name', '')
                for zone in dashboard.findall('.//zone'):
                    sheet_name = zone.get('name', '')
                    if sheet_name:
                        dashboard_map[sheet_name] = dash_name

            # Parse dashboard actions
            result['actions'] = XMLParser.parse_actions(root, workbook_name, workbook_id)

            # Parse datasources
            for datasource in root.findall('.//datasource'):
                ds_internal = datasource.get('name', 'Unknown')
                ds_name = datasource.get('caption', ds_internal)

                # Extract connections with type detection
                conn_details = XMLParser.parse_connection_info(datasource)
                for conn_info in conn_details:
                    result['connections'].append({
                        'object_id': workbook_id,
                        'object_name': workbook_name,
                        'object_type': 'Workbook',
                        'datasource_name': ds_name,
                        'datasource_internal': ds_internal,
                        'connection_class': conn_info.get('connection_class', ''),
                        'connection_type': conn_info.get('connection_type', 'embedded'),
                        'server': conn_info.get('server', ''),
                        'dbname': conn_info.get('dbname', ''),
                        'schema': conn_info.get('schema', ''),
                        'authentication': conn_info.get('authentication', ''),
                        'repository_location': conn_info.get('repository_location', '')
                    })

                # Extract tables
                tables = XMLParser.parse_tables(datasource, ds_internal, ds_name)
                for table in tables:
                    table['object_id'] = workbook_id
                    table['object_name'] = workbook_name
                    result['tables'].append(table)

                # Extract detailed column metadata
                connection_type = conn_details[0].get('connection_type', '') if conn_details else ''
                dbname = conn_details[0].get('dbname', '') if conn_details else ''
                for column in datasource.findall('.//column'):
                    col_info_list = XMLParser.parse_column_info(
                        column, ds_internal, ds_name, tables, dbname, connection_type
                    )
                    for col_info in col_info_list:
                        col_info['object_id'] = workbook_id
                        col_info['object_name'] = workbook_name
                        result['columns'].append(col_info)

                # Extract Custom SQL
                for relation in datasource.findall('.//relation'):
                    rel_type = relation.get('type', '')
                    if rel_type == 'text':
                        sql_text = relation.text or ''
                        if sql_text.strip():
                            result['custom_sql'].append({
                                'object_id': workbook_id,
                                'object_name': workbook_name,
                                'file_type': 'Workbook',
                                'datasource_name': ds_name,
                                'sql_name': relation.get('name', 'Custom SQL'),
                                'sql_text': sql_text.strip()
                            })

                # Extract calculations/formulas with hidden field
                for column in datasource.findall('.//column'):
                    calc = column.find('.//calculation')
                    if calc is not None:
                        formula = calc.get('formula', '')
                        if formula:
                            result['calculations'].append({
                                'object_id': workbook_id,
                                'object_name': workbook_name,
                                'field_name': column.get('name', '').strip('[]'),
                                'field_caption': column.get('caption', ''),
                                'datatype': column.get('datatype', ''),
                                'role': column.get('role', ''),
                                'formula_text': formula,
                                'is_hidden': column.get('hidden', 'false'),
                                'is_datasource_calc': False,
                                'is_workbook_calc': True,
                                'datasource_name': ds_name
                            })

                # Extract parameters
                if datasource.get('name') == 'Parameters':
                    for column in datasource.findall('.//column'):
                        calc = column.find('.//calculation')
                        result['parameters'].append({
                            'object_id': workbook_id,
                            'object_name': workbook_name,
                            'parameter_name': column.get('caption', column.get('name', '')),
                            'datatype': column.get('datatype', ''),
                            'value': calc.get('formula', '') if calc is not None else '',
                            'role': column.get('role', '')
                        })

                # Extract sets with based_on_field
                for group in datasource.findall('.//group'):
                    if group.get('name-style') == 'set':
                        # Get the field the set is based on
                        groupfilter = group.find('.//groupfilter')
                        based_on_field = groupfilter.get('field', '') if groupfilter is not None else ''
                        result['sets'].append({
                            'object_id': workbook_id,
                            'object_name': workbook_name,
                            'set_name': group.get('caption', group.get('name', '')),
                            'datasource_name': ds_name,
                            'based_on_field': based_on_field
                        })

                # Extract groups with based_on_field and members
                for group in datasource.findall('.//group'):
                    if group.get('name-style') != 'set':
                        # Get the field the group is based on
                        groupfilter = group.find('.//groupfilter')
                        based_on_field = groupfilter.get('field', '') if groupfilter is not None else ''
                        # Get group members
                        members = []
                        for member in group.findall('.//groupfilter[@member]'):
                            members.append(member.get('member', ''))
                        result['groups'].append({
                            'object_id': workbook_id,
                            'object_name': workbook_name,
                            'group_name': group.get('caption', group.get('name', '')),
                            'datasource_name': ds_name,
                            'based_on_field': based_on_field,
                            'members': json.dumps(members) if members else ''
                        })

                # Extract relationships/joins
                for relation in datasource.findall('.//relation'):
                    join_type = relation.get('join', '')
                    if join_type:
                        for clause in relation.findall('.//clause[@type="join"]'):
                            expressions = clause.findall('.//expression[@op="="]//expression')
                            if len(expressions) >= 2:
                                result['relationships'].append({
                                    'object_id': workbook_id,
                                    'object_name': workbook_name,
                                    'datasource_name': ds_name,
                                    'join_type': join_type,
                                    'left_field': expressions[0].get('op', ''),
                                    'right_field': expressions[1].get('op', '')
                                })

            # Parse worksheets
            for worksheet in root.findall('.//worksheet'):
                ws_name = worksheet.get('name', '')

                # Get fields used in worksheet
                fields_used = []
                for datasource_dep in worksheet.findall('.//datasource-dependencies'):
                    ds_name = datasource_dep.get('datasource', '')
                    for col in datasource_dep.findall('.//column'):
                        fields_used.append({
                            'datasource': ds_name,
                            'field': col.get('name', '')
                        })

                # Get mark type (visualization type)
                mark = worksheet.find('.//mark')
                mark_class = mark.get('class', 'automatic') if mark is not None else 'automatic'

                result['worksheets'].append({
                    'object_id': workbook_id,
                    'object_name': workbook_name,
                    'worksheet_name': ws_name,
                    'visualization_type': mark_class,
                    'fields_used': json.dumps(fields_used)
                })

                # Extract worksheet-level filters
                for filter_elem in worksheet.findall('.//filter'):
                    filter_col = filter_elem.get('column', '')
                    for gf in filter_elem.findall('.//groupfilter'):
                        result['filters'].append({
                            'object_id': workbook_id,
                            'object_name': workbook_name,
                            'worksheet_name': ws_name,
                            'filter_field': filter_col,
                            'filter_function': gf.get('function', ''),
                            'filter_member': gf.get('member', '')
                        })

            # Parse dashboards with size information
            for dashboard in root.findall('.//dashboard'):
                dash_name = dashboard.get('name', '')

                # Get worksheets in dashboard
                sheets_in_dash = []
                for zone in dashboard.findall('.//zone'):
                    sheet_name = zone.get('name', '')
                    if sheet_name:
                        sheets_in_dash.append(sheet_name)

                # Get dashboard size
                size = dashboard.find('.//size')
                width = size.get('maxwidth', '') if size is not None else ''
                height = size.get('maxheight', '') if size is not None else ''

                result['dashboards'].append({
                    'object_id': workbook_id,
                    'object_name': workbook_name,
                    'dashboard_name': dash_name,
                    'worksheets': json.dumps(sheets_in_dash),
                    'worksheet_count': len(sheets_in_dash),
                    'width': width,
                    'height': height
                })

            # Extract custom/non-standard visualizations
            result['custom_visuals'] = XMLParser.parse_custom_visuals(root, workbook_name, workbook_id)

            # Extract enhanced parameters with domain info
            result['parameters_enhanced'] = XMLParser.parse_parameters_enhanced(root, workbook_name, workbook_id)

        except Exception as e:
            logger.error(f"Error parsing workbook {workbook_name}: {str(e)}")

        return result

    @staticmethod
    def parse_datasource(file_path, datasource_name, datasource_id):
        """Parse .tds file and extract all metadata"""
        result = {
            'custom_sql': [],
            'calculations': [],
            'connections': [],
            'relationships': [],
            'object_tables': []
        }

        try:
            tree = ET.parse(file_path)
            root = tree.getroot()

            # Find datasource element
            datasource = root.find('.//datasource')
            if datasource is None:
                datasource = root

            # Tables used by this published data source, with their connection type / server
            result['object_tables'] = XMLParser.parse_object_tables(
                [datasource], datasource_id, datasource_name, 'Data Source')

            ds_name = datasource.get('caption', datasource.get('name', datasource_name))

            # Extract connections
            for conn in datasource.findall('.//connection'):
                result['connections'].append({
                    'object_id': datasource_id,
                    'object_name': datasource_name,
                    'object_type': 'Data Source',
                    'datasource_name': ds_name,
                    'connection_class': conn.get('class', ''),
                    'server': conn.get('server', ''),
                    'port': conn.get('port', ''),
                    'dbname': conn.get('dbname', ''),
                    'schema': conn.get('schema', ''),
                    'username': conn.get('username', ''),
                    'authentication': conn.get('authentication', '')
                })

            # Extract Custom SQL
            for relation in datasource.findall('.//relation'):
                rel_type = relation.get('type', '')
                if rel_type == 'text':
                    sql_text = relation.text or ''
                    if sql_text.strip():
                        result['custom_sql'].append({
                            'object_id': datasource_id,
                            'object_name': datasource_name,
                            'file_type': 'Data Source',
                            'datasource_name': ds_name,
                            'sql_name': relation.get('name', 'Custom SQL'),
                            'sql_text': sql_text.strip()
                        })

            # Extract calculations
            for column in datasource.findall('.//column'):
                calc = column.find('.//calculation')
                if calc is not None:
                    formula = calc.get('formula', '')
                    if formula:
                        result['calculations'].append({
                            'object_id': datasource_id,
                            'object_name': datasource_name,
                            'field_name': column.get('name', '').strip('[]'),
                            'caption': column.get('caption', ''),
                            'datatype': column.get('datatype', ''),
                            'role': column.get('role', ''),
                            'formula_text': formula,
                            'is_datasource_calc': True,
                            'is_workbook_calc': False,
                            'datasource_name': ds_name
                        })

            # Extract relationships/joins
            for relation in datasource.findall('.//relation'):
                join_type = relation.get('join', '')
                if join_type:
                    for clause in relation.findall('.//clause[@type="join"]'):
                        expressions = clause.findall('.//expression[@op="="]//expression')
                        if len(expressions) >= 2:
                            result['relationships'].append({
                                'object_id': datasource_id,
                                'object_name': datasource_name,
                                'datasource_name': ds_name,
                                'join_type': join_type,
                                'left_field': expressions[0].get('op', ''),
                                'right_field': expressions[1].get('op', '')
                            })

        except Exception as e:
            logger.error(f"Error parsing datasource {datasource_name}: {str(e)}")

        return result


class DeepMetadataExtractor:
    """Comprehensive metadata extractor using REST + Metadata API + File Parsing"""

    def __init__(self, config_data):
        self.config_data = config_data
        self.server = None
        self.auth_token = None
        self.metadata_api = None
        self.temp_dirs = []
        self.project_filter = config_data.get('project_filter', 'all')
        self.project_ids = config_data.get('project_ids', [])
        self.project_name = config_data.get('project_name', 'All Projects')
        self.content_types = config_data.get('content_types', [])
        self.object_ids = config_data.get('object_ids', [])
        self.filtered_project_ids = set()  # Will store project IDs to filter by
        self.filtered_object_ids = set()   # Will store specific object IDs to filter by
        # None = inventory has not run yet; a list (possibly empty) = the exact scope to process.
        self.filtered_workbooks = None
        self.filtered_datasources = None
        self.filtered_flows = None
        self.lineage_only = bool(config_data.get('lineage_only'))   # skip file downloads/permissions/admin data
        self.warnings = []                 # non-fatal problems, exported on the "Warnings" sheet
        self._auth = None
        self.job_id = None

    MAX_WARNINGS = 1000
    RETRIES = 3

    # Prep flow content (.tfl/.tflx) is parsed with workbook_analyzer.FlowAnalyzer (the same engine behind
    # Single Workbook Analysis) and merged into these separate flow_* result keys - kept apart from the
    # workbook/datasource sheets, and from Lineage Info's site-wide 'lineage_flows' (which comes from the
    # Metadata API graph, not from downloading the flow file itself).
    FLOW_SCHEMA_MAP = {
        'workbook_summary': 'flow_summary', 'datasource_inventory': 'flow_io',
        'connection_inventory': 'flow_connections', 'table_inventory': 'flow_tables',
        'join_analysis': 'flow_joins', 'relationship_analysis': 'flow_unions',
        'custom_sql': 'flow_custom_sql', 'calculated_fields': 'flow_calculations',
        'parameters': 'flow_parameters', 'filters': 'flow_filters', 'field_usage': 'flow_field_usage',
        'workbook_lineage': 'flow_content_lineage', 'impact': 'flow_impact',
        'data_quality': 'flow_data_quality', 'flow_steps': 'flow_steps',
    }

    def authenticate(self):
        """Authenticate with Tableau Server"""
        server = TSC.Server(self.config_data['server_url'], use_server_version=True)
        auth = TSC.PersonalAccessTokenAuth(
            self.config_data['pat_name'],
            self.config_data['pat_token'],
            self.config_data['site_name']
        )
        self.server = server
        self._auth = auth
        return server, auth

    def _warn(self, stage, message, item=''):
        """Record a non-fatal problem instead of silently swallowing it."""
        logger.warning(f"[{stage}] {item} {message}".strip())
        if len(self.warnings) < self.MAX_WARNINGS:
            self.warnings.append({'stage': stage, 'item': str(item), 'message': str(message)[:500]})
        elif len(self.warnings) == self.MAX_WARNINGS:
            self.warnings.append({'stage': 'warnings', 'item': '',
                                  'message': f'More than {self.MAX_WARNINGS} warnings - the rest are only in the log'})

    def _set_status(self, progress, message):
        if self.job_id:
            job_status[self.job_id] = {'status': 'running', 'progress': progress, 'message': message}

    def _reauth(self):
        """Sign in again - Tableau auth tokens expire (default 240 min), long extractions outlive them."""
        self.server.auth.sign_in(self._auth)
        self.auth_token = self.server.auth_token
        if self.metadata_api:
            self.metadata_api.auth_token = self.auth_token
        logger.info("Re-authenticated with Tableau")

    def _retry(self, func, *args, stage='api', item='', **kwargs):
        """
        Call func with retries: re-authenticates on an expired session (401) and backs off on
        throttling/timeouts/5xx. Raises the last error when all attempts fail.
        """
        last = None
        for attempt in range(1, self.RETRIES + 1):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                last = e
                code = str(getattr(e, 'code', '') or '')
                text = str(e)
                if (code.startswith('401') or '401002' in text or 'Unauthorized' in text
                        or 'NotSignedIn' in type(e).__name__):
                    try:
                        self._reauth()
                    except Exception as re_err:
                        logger.warning(f"Re-authentication failed: {re_err}")
                elif code[:1] == '4' and not code.startswith('429') and not code.startswith('408'):
                    break  # 400/403/404/409...: permanent for this item - retrying will not help
                elif not code and not isinstance(e, (requests.exceptions.RequestException, OSError)):
                    break  # not an API/network error (e.g. a parsing bug) - retrying will not help
                if attempt < self.RETRIES:
                    time.sleep(2 * attempt)
        raise last

    def _pager_list(self, endpoint, stage, item=''):
        """Fetch all pages from a TSC endpoint; failures become warnings and return what was read."""
        items = []
        try:
            for x in TSC.Pager(endpoint):
                items.append(x)
        except Exception as e:
            self._warn(stage, f'stopped after {len(items)} items: {e}', item)
        return items

    def cleanup(self):
        """Clean up temporary directories"""
        for temp_dir in self.temp_dirs:
            try:
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
            except:
                pass
        self.temp_dirs = []

    def extract_all(self, job_id=None):
        """
        Main extraction method - combines REST + Metadata API + File Parsing
        """
        results = {
            'inventory': [],
            'connections': [],
            'database_servers': [],
            'virtual_connections': [],
            'bridge_connections': [],
            'custom_sql': [],
            'calculations': [],
            'lineage': [],
            'permissions': [],
            'parameters': [],
            'filters': [],
            'sets': [],
            'groups': [],
            'worksheets': [],
            'dashboards': [],
            'relationships': [],
            'users': [],
            'groups_admin': [],
            'schedules': [],
            'jobs': [],
            'tasks': [],
            'subscriptions': [],
            'tables': [],
            'actions': [],
            'columns': [],
            'custom_visuals': [],
            'parameters_enhanced': [],
            'object_tables': [],
            'object_tables_summary': [],
            'overall_lineage': [],
            'lineage_summary': [],
            'lineage_tables_used': [],
            'lineage_workbooks': [],
            'lineage_datasources': [],
            'lineage_flows': [],
            'lineage_custom_sql': [],
            'lineage_table_matrix': [],
            'lineage_impact': [],
            'flow_summary': [], 'flow_io': [], 'flow_connections': [], 'flow_tables': [], 'flow_joins': [],
            'flow_unions': [], 'flow_custom_sql': [], 'flow_calculations': [], 'flow_parameters': [],
            'flow_filters': [], 'flow_field_usage': [], 'flow_content_lineage': [], 'flow_impact': [],
            'flow_data_quality': [], 'flow_steps': [],
            'warnings': [],
            'stats': {
                'custom_sql_count': 0,
                'calculations_count': 0,
                'lineage_edges': 0,
                'tables_count': 0,
                'actions_count': 0,
                'columns_count': 0,
                'custom_visuals_count': 0
            }
        }
        self.job_id = job_id

        try:
            server, auth = self.authenticate()

            with server.auth.sign_in(auth):
                self.auth_token = server.auth_token
                site_id = server.site_id

                # Initialize Metadata API client
                self.metadata_api = MetadataAPIClient(
                    self.config_data['server_url'],
                    self.auth_token,
                    site_id
                )

                self._set_status(10, 'Fetching inventory...')

                # 1. REST API: Get inventory
                self._extract_inventory(results)

                # Lineage Info: start the (parallel, cursor-paginated) Metadata API fetch in the background so
                # it overlaps with the REST calls and file downloads below.
                lineage_pool, lineage_future = None, None
                if self.metadata_api.available:
                    lineage_pool = ThreadPoolExecutor(max_workers=1)
                    progress = None
                    if self.lineage_only:
                        progress = lambda done, total, name: self._set_status(
                            25 + int(60 * done / max(total, 1)), f'Lineage Info: fetched {name} ({done}/{total})')
                    lineage_future = lineage_pool.submit(lineage_engine.fetch_lineage_data, self.metadata_api,
                                                         progress, 4)
                else:
                    self._warn('lineage', 'Metadata API unavailable - Lineage Info could not be built')

                try:
                    if not self.lineage_only:
                        self._set_status(18, 'Extracting Bridge connection metadata...')

                        # 1b. REST API: Get Bridge connection metadata
                        self._extract_bridge_connections(results)

                        self._set_status(25, 'Extracting from Metadata API (paged)...')

                        # 2. Metadata API: Get lineage and calculated fields
                        if self.metadata_api.available:
                            self._extract_from_metadata_api(results)
                            for problem in self.metadata_api.fetch_errors:
                                self._warn('metadata_api', problem)
                        else:
                            self._warn('metadata_api', 'Metadata API unavailable - lineage/calculation data from the API was skipped')

                        self._set_status(40, 'Downloading and parsing files...')

                        # 3. File Parsing: Download and parse workbooks/datasources
                        self._extract_from_files(results, job_id)

                        self._set_status(80, 'Parsing Prep flows...')

                        # 3b. Flow content: download and parse each in-scope .tfl/.tflx (steps, inputs/outputs,
                        # joins, custom SQL, calculations, lineage, impact, data-quality findings)
                        self._extract_flow_content(results, job_id)

                        self._set_status(85, 'Extracting permissions...')

                        # 4. REST API: Get permissions
                        self._extract_permissions(results)

                        self._set_status(90, 'Extracting admin data...')

                        # 5. REST API: Get admin metadata
                        self._extract_admin_data(results)

                        # Metadata API data is site-wide; keep only what belongs to the selected scope
                        self._apply_scope_to_api_rows(results)

                        # Tables used per workbook / data source (XML + Metadata API fallback) and the roll-up
                        self._add_metadata_api_tables(results)
                        results['object_tables_summary'] = self._build_object_tables_summary(results['object_tables'])

                        # Overall Lineage: one row per (object, table) across the WHOLE site - every workbook,
                        # data source and flow together, with a Custom SQL / Direct indicator per table.
                        results['overall_lineage'] = self._build_overall_lineage(results['object_tables'])

                    # 6. Lineage Info (graph, impact analysis, custom SQL analysis)
                    self._set_status(92 if not self.lineage_only else 88, 'Building lineage graph...')
                    self._build_lineage_info(results, lineage_future)
                finally:
                    if lineage_pool:
                        lineage_pool.shutdown(wait=False)

                results['warnings'] = self.warnings

                # Update stats
                results['stats']['custom_sql_count'] = len(results['custom_sql'])
                results['stats']['calculations_count'] = len(results['calculations'])
                results['stats']['lineage_edges'] = len(results['lineage'])
                results['stats']['tables_count'] = len(results['tables'])
                results['stats']['actions_count'] = len(results['actions'])
                results['stats']['columns_count'] = len(results['columns'])
                results['stats']['custom_visuals_count'] = len(results['custom_visuals'])
                results['stats']['object_tables_count'] = len(results['object_tables'])
                results['stats']['lineage_rows'] = len(results['lineage_summary'])
                results['stats']['flows_analyzed'] = len(results['flow_summary'])
                results['stats']['warnings_count'] = len(self.warnings)

        except Exception as e:
            logger.error(f"Extraction error: {str(e)}")
            raise
        finally:
            self.cleanup()

        return results

    def _lineage_scope(self):
        """Which content ids the emitted lineage rows are limited to (impact is always site-wide)."""
        if not self._filters_active:
            return {}
        return {
            'workbooks': {wb.id for wb in (self.filtered_workbooks or [])},
            'datasources': {ds.id for ds in (self.filtered_datasources or [])},
            'flows': {f.id for f in (self.filtered_flows or [])},
        }

    def _build_lineage_info(self, results, lineage_future):
        """Join the background Metadata API fetch and build the Lineage Info datasets."""
        if lineage_future is None:
            return
        try:
            raw = lineage_future.result()
        except Exception as e:
            self._warn('lineage', f'Lineage fetch failed: {e}')
            return
        for problem in raw.get('errors', []):
            self._warn('lineage', problem)
        for problem in self.metadata_api.fetch_errors:
            self._warn('lineage', problem)
        if self.lineage_only:
            self._warn('lineage', 'Lineage-only run: server/port/warehouse/authentication come from Metadata API '
                                  'database info only. Run the full Deep Extraction to add them from workbook XML.')
        try:
            built = lineage_engine.build_lineage(
                raw, results['object_tables'], results['custom_sql'], results['bridge_connections'],
                scope=self._lineage_scope())
        except Exception as e:
            logger.error(f"Lineage build failed: {e}", exc_info=True)
            self._warn('lineage', f'Lineage build failed: {e}')
            return
        finally:
            raw = None
        results['stats']['lineage_info'] = built.pop('lineage_stats', {})
        results.update(built)

    @property
    def _filters_active(self):
        return bool(
            self.project_ids or self.object_ids or self.content_types
            or (self.project_filter and self.project_filter != 'all')
        )

    def _add_metadata_api_tables(self, results):
        """
        Workbooks/data sources whose file could not be downloaded/parsed still get their tables from
        the Metadata API lineage (no server string is available from that source).
        """
        covered = {r['object_id'] for r in results['object_tables']}
        added = 0
        for edge in results['lineage']:
            if edge.get('upstream_type') != 'Table':
                continue
            luid = edge.get('downstream_luid')
            if not luid or luid in covered:
                continue
            db, schema, table = XMLParser.split_table_reference(edge.get('upstream_object', ''))
            results['object_tables'].append({
                'object_id': luid,
                'object_name': edge.get('downstream_object', ''),
                'object_type': edge.get('downstream_type', ''),
                'connection_type': edge.get('connection_type', ''),
                'connection_server': '',
                'database': edge.get('upstream_database', '') or db,
                'schema': edge.get('upstream_schema', '') or schema,
                'table_name': table or edge.get('upstream_object', ''),
                'full_table_name': '.'.join(x for x in ((edge.get('upstream_schema', '') or schema), table) if x)
                                   or edge.get('upstream_object', ''),
                'table_kind': 'Table',
                'datasource_name': '',
                'connection_class': '',
                'table_alias': '',
                'table_reference': edge.get('upstream_object', ''),
                'source': 'Metadata API',
            })
            added += 1
        if added:
            logger.info(f"Added {added} table rows from Metadata API lineage for objects without parsed files")

    @staticmethod
    def _build_object_tables_summary(object_tables):
        """One row per workbook/data source/flow + connection: ID, name, connection type, server, tables used."""
        groups = {}
        for r in object_tables:
            key = (r['object_id'], r['connection_type'], r['connection_server'], r['database'])
            g = groups.get(key)
            if g is None:
                g = groups[key] = {
                    'object_id': r['object_id'],
                    'object_name': r['object_name'],
                    'object_type': r['object_type'],
                    'connection_type': r['connection_type'],
                    'connection_server': r['connection_server'],
                    'database': r['database'],
                    'table_count': 0,
                    'tables_used': [],
                }
            label = r.get('full_table_name') or r['table_name']
            if r['table_kind'] == 'Custom SQL':
                label = f"[Custom SQL: {r['table_name']}]"
            elif r['table_kind'] == 'Published Data Source':
                label = f"[Published DS: {r['table_name']}]"
            if label and label not in g['tables_used']:
                g['tables_used'].append(label)
        summary = []
        for g in groups.values():
            g['table_count'] = len(g['tables_used'])
            g['tables_used'] = ', '.join(g['tables_used'])
            summary.append(g)
        return summary

    # table_kind values that mean "this table came from custom SQL", not a direct/physical connection
    CUSTOM_SQL_TABLE_KINDS = {'Custom SQL', 'Custom SQL (table parsed from SQL)'}

    @staticmethod
    def _build_overall_lineage(object_tables):
        """
        Site-wide, one row per (object, table): every table used by every workbook, data source and flow
        extracted in this run, in one flat list - covers direct/physical tables, published-datasource
        references, stored procedures, and tables found by parsing Custom SQL (whether XML-sourced or the
        Metadata API fallback). See DeepMetadataExtractor.CUSTOM_SQL_TABLE_KINDS for the Custom SQL / Direct split.
        """
        rows = []
        for r in object_tables:
            kind = r.get('table_kind', '')
            is_custom_sql = (kind in DeepMetadataExtractor.CUSTOM_SQL_TABLE_KINDS
                             or 'Custom SQL' in str(r.get('source', '')))
            rows.append({
                'Object ID': r.get('object_id', ''),
                'Object Name': r.get('object_name', ''),
                'Object Type': r.get('object_type', ''),
                'Datasource Name': r.get('datasource_name', ''),
                'Connection Type': r.get('connection_type', ''),
                'Server': r.get('connection_server', ''),
                'Database': r.get('database', ''),
                'Schema': r.get('schema', ''),
                'Table Name': r.get('table_name', ''),
                'Full Table Name': r.get('full_table_name', ''),
                'Table Kind': kind,
                'Custom SQL / Direct': 'Custom SQL' if is_custom_sql else 'Direct',
                'Source': r.get('source', ''),
            })
        return rows

    def _apply_scope_to_api_rows(self, results):
        """
        The Metadata API queries return the whole site. When a project/object/type filter was
        selected, drop lineage / calculation / custom-SQL rows that belong to other content so the
        export matches the chosen scope.
        """
        if not self._filters_active:
            return
        wb_ids = {wb.id for wb in (self.filtered_workbooks or [])}
        ds_ids = {ds.id for ds in (self.filtered_datasources or [])}
        keep = wb_ids | ds_ids

        def in_scope(row, *keys):
            for k in keys:
                v = row.get(k)
                if v and v in keep:
                    return True
            return False

        before = (len(results['calculations']), len(results['custom_sql']), len(results['lineage']))
        results['calculations'] = [r for r in results['calculations']
                                   if r.get('source') != 'Metadata API' or in_scope(r, 'object_id')]
        results['custom_sql'] = [r for r in results['custom_sql']
                                 if r.get('source') != 'Metadata API' or in_scope(r, 'object_id')]
        results['lineage'] = [r for r in results['lineage'] if in_scope(r, 'downstream_luid')]
        after = (len(results['calculations']), len(results['custom_sql']), len(results['lineage']))
        logger.info(f"Scoped Metadata API rows to the selected filters: {before} -> {after}")

    def _extract_inventory(self, results):
        """Extract inventory from REST API with optional project/object filtering"""
        try:
            # Initialize filtered lists early to ensure they exist even if later API calls fail
            self.filtered_workbooks = []
            self.filtered_datasources = []
            self.filtered_flows = []

            # Build set of project IDs to filter by (including child projects)
            # Use Pager to get ALL projects (not just first 100)
            projects = list(TSC.Pager(self.server.projects))
            project_map = {p.id: p for p in projects}

            # Handle multiple project IDs or single project filter
            if self.project_ids and len(self.project_ids) > 0:
                # Multiple projects selected - get all with their children
                self.filtered_project_ids = set()
                for pid in self.project_ids:
                    self.filtered_project_ids.update(self._get_project_with_children(pid, projects))
            elif self.project_filter and self.project_filter != 'all':
                # Single project selected
                self.filtered_project_ids = self._get_project_with_children(self.project_filter, projects)
            else:
                self.filtered_project_ids = set()  # Empty means all projects

            # Build sets of object IDs by type (if specific objects selected)
            self.filtered_wb_ids = set()
            self.filtered_ds_ids = set()
            self.filtered_flow_ids = set()
            self.filtered_view_ids = set()
            has_object_filter = False

            if self.object_ids and len(self.object_ids) > 0:
                has_object_filter = True
                for oid in self.object_ids:
                    if oid.startswith('wb_'):
                        self.filtered_wb_ids.add(oid[3:])
                    elif oid.startswith('ds_'):
                        self.filtered_ds_ids.add(oid[3:])
                    elif oid.startswith('flow_'):
                        self.filtered_flow_ids.add(oid[5:])
                    elif oid.startswith('view_'):
                        self.filtered_view_ids.add(oid[5:])
                    elif '_' in oid:
                        # Fallback: extract raw ID
                        self.filtered_wb_ids.add(oid.split('_', 1)[1])
                    else:
                        # No prefix - add to all
                        self.filtered_wb_ids.add(oid)
                        self.filtered_ds_ids.add(oid)

            # Check if specific content types are selected
            extract_workbooks = not self.content_types or 'Workbook' in self.content_types
            extract_datasources = not self.content_types or 'Data Source' in self.content_types
            extract_flows = not self.content_types or 'Flow' in self.content_types
            extract_views = (not self.content_types or 'View' in self.content_types) and not self.lineage_only

            # If specific objects are selected, only extract those types
            if has_object_filter:
                extract_workbooks = extract_workbooks and len(self.filtered_wb_ids) > 0
                extract_datasources = extract_datasources and len(self.filtered_ds_ids) > 0
                extract_flows = extract_flows and len(self.filtered_flow_ids) > 0
                extract_views = extract_views and len(self.filtered_view_ids) > 0

            # Projects
            for p in projects:
                if self.filtered_project_ids and p.id not in self.filtered_project_ids:
                    continue
                results['inventory'].append({
                    'id': p.id,
                    'type': 'Project',
                    'name': p.name,
                    'parent_id': p.parent_id if hasattr(p, 'parent_id') else '',
                    'owner_id': p.owner_id if hasattr(p, 'owner_id') else '',
                    'created_at': str(p.created_at) if hasattr(p, 'created_at') and p.created_at else '',
                    'updated_at': str(p.updated_at) if hasattr(p, 'updated_at') and p.updated_at else ''
                })

            # Workbooks - filter by project and object IDs
            # Use Pager to get ALL workbooks (not just first 100)
            workbooks = self._pager_list(self.server.workbooks, 'inventory', 'workbooks')
            self.filtered_workbooks = []
            if extract_workbooks:
                for wb in workbooks:
                    # Filter by project
                    if self.filtered_project_ids and wb.project_id not in self.filtered_project_ids:
                        continue
                    # Filter by specific workbook IDs
                    if self.filtered_wb_ids and wb.id not in self.filtered_wb_ids:
                        continue
                    self.filtered_workbooks.append(wb)
                    results['inventory'].append({
                        'id': wb.id,
                        'type': 'Workbook',
                        'name': wb.name,
                        'project_id': wb.project_id,
                        'project_name': wb.project_name,
                        'owner_id': wb.owner_id if hasattr(wb, 'owner_id') else '',
                        'created_at': str(wb.created_at) if hasattr(wb, 'created_at') and wb.created_at else '',
                        'updated_at': str(wb.updated_at) if hasattr(wb, 'updated_at') and wb.updated_at else '',
                        'webpage_url': wb.webpage_url if hasattr(wb, 'webpage_url') else ''
                    })

            # Views - wrapped in try/except to prevent blocking datasources setup
            try:
                if extract_views:
                    # Use Pager to get ALL views (not just first 100)
                    views = self._pager_list(self.server.views, 'inventory', 'views')
                    parent_wb_ids = {wb.id for wb in self.filtered_workbooks}
                    for v in views:
                        # Filter by parent workbook (if workbooks are filtered)
                        if parent_wb_ids and hasattr(v, 'workbook_id') and v.workbook_id not in parent_wb_ids:
                            continue
                        # Filter by specific view IDs
                        if self.filtered_view_ids and v.id not in self.filtered_view_ids:
                            continue
                        results['inventory'].append({
                            'id': v.id,
                            'type': 'View',
                            'name': v.name,
                            'workbook_id': v.workbook_id if hasattr(v, 'workbook_id') else '',
                            'owner_id': v.owner_id if hasattr(v, 'owner_id') else '',
                            'created_at': str(v.created_at) if hasattr(v, 'created_at') and v.created_at else '',
                            'total_views': v.total_views if hasattr(v, 'total_views') else 0
                        })
            except Exception as e:
                self._warn('inventory', f'Error extracting views: {e}')

            # Data Sources - filter by project and object IDs
            # Use Pager to get ALL datasources (not just first 100)
            datasources = self._pager_list(self.server.datasources, 'inventory', 'datasources')
            if extract_datasources:
                for ds in datasources:
                    # Filter by project
                    if self.filtered_project_ids and ds.project_id not in self.filtered_project_ids:
                        continue
                    # Filter by specific datasource IDs
                    if self.filtered_ds_ids and ds.id not in self.filtered_ds_ids:
                        continue
                    self.filtered_datasources.append(ds)
                    results['inventory'].append({
                        'id': ds.id,
                        'type': 'Data Source',
                        'name': ds.name,
                        'project_id': ds.project_id,
                        'project_name': ds.project_name,
                        'owner_id': ds.owner_id if hasattr(ds, 'owner_id') else '',
                        'created_at': str(ds.created_at) if hasattr(ds, 'created_at') and ds.created_at else '',
                        'updated_at': str(ds.updated_at) if hasattr(ds, 'updated_at') and ds.updated_at else '',
                        'webpage_url': ds.webpage_url if hasattr(ds, 'webpage_url') else ''
                    })

            # Flows - filter by project and object IDs
            try:
                # Use Pager to get ALL flows (not just first 100)
                flows = self._pager_list(self.server.flows, 'inventory', 'flows') if extract_flows else []
                if extract_flows:
                    for f in flows:
                        if self.filtered_project_ids and f.project_id not in self.filtered_project_ids:
                            continue
                        if self.filtered_flow_ids and f.id not in self.filtered_flow_ids:
                            continue
                        self.filtered_flows.append(f)
                        results['inventory'].append({
                            'id': f.id,
                            'type': 'Flow',
                            'name': f.name,
                            'project_id': f.project_id,
                            'project_name': f.project_name,
                            'owner_id': f.owner_id if hasattr(f, 'owner_id') else '',
                        'created_at': str(f.created_at) if hasattr(f, 'created_at') and f.created_at else '',
                        'updated_at': str(f.updated_at) if hasattr(f, 'updated_at') and f.updated_at else ''
                    })
            except Exception as e:
                self._warn('inventory', f'Error extracting flows: {e}')

        except Exception as e:
            self._warn('inventory', f'Inventory extraction error: {e}')

    def _extract_bridge_connections(self, results):
        """Extract Tableau Bridge connection metadata from datasources and workbooks"""
        try:
            logger.info("Extracting Bridge connection metadata...")

            # Extract connections from datasources
            # Always the scoped list from the inventory step (an empty scope means nothing to check)
            datasources_to_check = self.filtered_datasources or []
            total_ds = len(datasources_to_check)

            for ds_idx, ds in enumerate(datasources_to_check):
                if ds_idx % 50 == 0:
                    self._set_status(18 + int(3 * ds_idx / max(total_ds, 1)),
                                     f'Bridge connections: data source {ds_idx + 1}/{total_ds}')
                try:
                    # Populate connections for this datasource
                    self._retry(self.server.datasources.populate_connections, ds,
                                stage='bridge_connections', item=ds.name)

                    # Check Bridge setting (use_remote_query_agent)
                    use_bridge = getattr(ds, 'use_remote_query_agent', None)

                    # Get connection details
                    if hasattr(ds, 'connections') and ds.connections:
                        for conn in ds.connections:
                            results['bridge_connections'].append({
                                'object_type': 'Data Source',
                                'object_id': ds.id,
                                'object_name': ds.name,
                                'project_name': ds.project_name if hasattr(ds, 'project_name') else '',
                                'connection_id': conn.id if hasattr(conn, 'id') else '',
                                'connection_type': conn.connection_type if hasattr(conn, 'connection_type') else '',
                                'server_address': conn.server_address if hasattr(conn, 'server_address') else '',
                                'server_port': conn.server_port if hasattr(conn, 'server_port') else '',
                                'username': conn.username if hasattr(conn, 'username') else '',
                                'embed_password': conn.embed_password if hasattr(conn, 'embed_password') else '',
                                'use_tableau_bridge': use_bridge if use_bridge is not None else '',
                                'datasource_content_url': ds.content_url if hasattr(ds, 'content_url') else '',
                                'has_extracts': ds.has_extracts if hasattr(ds, 'has_extracts') else '',
                                'extract_encryption_mode': ds.extract_encryption_mode if hasattr(ds, 'extract_encryption_mode') else ''
                            })
                    else:
                        # Record datasource even without detailed connections
                        results['bridge_connections'].append({
                            'object_type': 'Data Source',
                            'object_id': ds.id,
                            'object_name': ds.name,
                            'project_name': ds.project_name if hasattr(ds, 'project_name') else '',
                            'connection_id': '',
                            'connection_type': '',
                            'server_address': '',
                            'server_port': '',
                            'username': '',
                            'embed_password': '',
                            'use_tableau_bridge': use_bridge if use_bridge is not None else '',
                            'datasource_content_url': ds.content_url if hasattr(ds, 'content_url') else '',
                            'has_extracts': ds.has_extracts if hasattr(ds, 'has_extracts') else '',
                            'extract_encryption_mode': ds.extract_encryption_mode if hasattr(ds, 'extract_encryption_mode') else ''
                        })
                except Exception as e:
                    self._warn('bridge_connections', f'Could not get connections: {e}', f'Data Source {ds.name}')

            # Extract connections from workbooks
            workbooks_to_check = self.filtered_workbooks or []
            total_wb = len(workbooks_to_check)

            for wb_idx, wb in enumerate(workbooks_to_check):
                if wb_idx % 50 == 0:
                    self._set_status(21 + int(4 * wb_idx / max(total_wb, 1)),
                                     f'Bridge connections: workbook {wb_idx + 1}/{total_wb}')
                try:
                    # Populate connections for this workbook
                    self._retry(self.server.workbooks.populate_connections, wb,
                                stage='bridge_connections', item=wb.name)

                    if hasattr(wb, 'connections') and wb.connections:
                        for conn in wb.connections:
                            results['bridge_connections'].append({
                                'object_type': 'Workbook',
                                'object_id': wb.id,
                                'object_name': wb.name,
                                'project_name': wb.project_name if hasattr(wb, 'project_name') else '',
                                'connection_id': conn.id if hasattr(conn, 'id') else '',
                                'connection_type': conn.connection_type if hasattr(conn, 'connection_type') else '',
                                'server_address': conn.server_address if hasattr(conn, 'server_address') else '',
                                'server_port': conn.server_port if hasattr(conn, 'server_port') else '',
                                'username': conn.username if hasattr(conn, 'username') else '',
                                'embed_password': conn.embed_password if hasattr(conn, 'embed_password') else '',
                                'use_tableau_bridge': '',  # Bridge setting is on datasource level
                                'datasource_content_url': '',
                                'has_extracts': '',
                                'extract_encryption_mode': ''
                            })
                except Exception as e:
                    self._warn('bridge_connections', f'Could not get connections: {e}', f'Workbook {wb.name}')

            logger.info(f"Extracted {len(results['bridge_connections'])} Bridge connection records")

        except Exception as e:
            self._warn('bridge_connections', f'Bridge connection extraction error: {e}')

    def _get_project_with_children(self, project_id, all_projects):
        """Get a project and all its child projects recursively"""
        result = {project_id}
        for p in all_projects:
            parent_id = getattr(p, 'parent_id', None)
            if parent_id and parent_id in result:
                result.add(p.id)
        # Run again to get nested children
        changed = True
        while changed:
            changed = False
            for p in all_projects:
                parent_id = getattr(p, 'parent_id', None)
                if parent_id and parent_id in result and p.id not in result:
                    result.add(p.id)
                    changed = True
        return result

    def _extract_from_metadata_api(self, results):
        """Extract lineage and metadata from Metadata API (GraphQL)"""
        try:
            # Get calculated fields
            calc_data = self.metadata_api.get_calculated_fields()
            if calc_data and 'calculatedFields' in calc_data:
                for field in calc_data['calculatedFields']:
                    ds_info = field.get('datasource', {}) or {}
                    results['calculations'].append({
                        'object_id': ds_info.get('luid', ''),
                        'object_name': ds_info.get('name', ''),
                        'field_name': field.get('name', ''),
                        'formula_text': field.get('formula', ''),
                        'data_category': field.get('dataCategory', ''),
                        'role': field.get('role', ''),
                        'is_datasource_calc': True,
                        'is_workbook_calc': False,
                        'source': 'Metadata API'
                    })

            # Get custom SQL tables
            sql_data = self.metadata_api.get_custom_sql_tables()
            if sql_data and 'customSQLTables' in sql_data:
                for table in sql_data['customSQLTables']:
                    db_info = table.get('database', {}) or {}
                    ds_list = table.get('downstreamDatasources', []) or []
                    for ds in ds_list:
                        results['custom_sql'].append({
                            'object_id': ds.get('luid', ''),
                            'object_name': ds.get('name', ''),
                            'file_type': 'Data Source',
                            'datasource_name': ds.get('name', ''),
                            'sql_name': table.get('name', ''),
                            'sql_text': table.get('query', ''),
                            'database': db_info.get('name', ''),
                            'connection_type': db_info.get('connectionType', ''),
                            'source': 'Metadata API'
                        })

            # Get all tables for lineage
            tables_data = self.metadata_api.get_all_tables()
            if tables_data and 'databaseTables' in tables_data:
                for table in tables_data['databaseTables']:
                    db_info = table.get('database', {}) or {}
                    table_schema = table.get('schema', '')
                    is_embedded = table.get('isEmbedded', False)

                    # Downstream datasources
                    for ds in (table.get('downstreamDatasources', []) or []):
                        results['lineage'].append({
                            'upstream_object': table.get('fullName', table.get('name', '')),
                            'upstream_type': 'Table',
                            'upstream_database': db_info.get('name', ''),
                            'upstream_schema': table_schema,
                            'upstream_is_embedded': is_embedded,
                            'downstream_object': ds.get('name', ''),
                            'downstream_type': 'Data Source',
                            'downstream_luid': ds.get('luid', ''),
                            'downstream_project': ds.get('projectName', ''),
                            'downstream_has_extracts': ds.get('hasExtracts', False),
                            'downstream_extract_refresh': ds.get('extractLastRefreshTime', ''),
                            'connection_type': db_info.get('connectionType', '')
                        })

                    # Downstream workbooks
                    for wb in (table.get('downstreamWorkbooks', []) or []):
                        results['lineage'].append({
                            'upstream_object': table.get('fullName', table.get('name', '')),
                            'upstream_type': 'Table',
                            'upstream_database': db_info.get('name', ''),
                            'upstream_schema': table_schema,
                            'upstream_is_embedded': is_embedded,
                            'downstream_object': wb.get('name', ''),
                            'downstream_type': 'Workbook',
                            'downstream_luid': wb.get('luid', ''),
                            'downstream_project': wb.get('projectName', ''),
                            'downstream_created': wb.get('createdAt', ''),
                            'downstream_updated': wb.get('updatedAt', ''),
                            'connection_type': db_info.get('connectionType', '')
                        })

                    # Column lineage
                    for col in (table.get('columns', []) or []):
                        for ds in (table.get('downstreamDatasources', []) or []):
                            results['lineage'].append({
                                'upstream_object': f"{table.get('name', '')}.{col.get('name', '')}",
                                'upstream_type': 'Column',
                                'upstream_database': db_info.get('name', ''),
                                'upstream_schema': table_schema,
                                'downstream_object': ds.get('name', ''),
                                'downstream_type': 'Data Source',
                                'downstream_luid': ds.get('luid', ''),
                                'downstream_project': ds.get('projectName', ''),
                                'column_type': col.get('remoteType', ''),
                                'column_description': col.get('description', ''),
                                'column_nullable': col.get('isNullable', ''),
                                'connection_type': db_info.get('connectionType', '')
                            })

            # Get database servers (gateway/connection details)
            db_servers_data = self.metadata_api.get_database_servers()
            if db_servers_data and 'databaseServers' in db_servers_data:
                for server in db_servers_data['databaseServers']:
                    contact = server.get('contact') or {}
                    downstream_ds = server.get('downstreamDatasources') or []
                    downstream_wb = server.get('downstreamWorkbooks') or []

                    results['database_servers'].append({
                        'server_id': server.get('id', ''),
                        'server_name': server.get('name', ''),
                        'host_name': server.get('hostName', ''),
                        'port': server.get('port', ''),
                        'connection_type': server.get('connectionType', ''),
                        'extended_connection_type': server.get('extendedConnectionType', ''),
                        'service': server.get('service', ''),
                        'description': server.get('description', ''),
                        'is_embedded': server.get('isEmbedded', False),
                        'project_name': server.get('projectName', ''),
                        'is_certified': server.get('isCertified', False),
                        'has_active_warning': server.get('hasActiveWarning', False),
                        'contact_name': contact.get('name', ''),
                        'contact_email': contact.get('email', ''),
                        'downstream_datasources': ', '.join([ds.get('name', '') for ds in downstream_ds]),
                        'downstream_datasource_count': len(downstream_ds),
                        'downstream_workbooks': ', '.join([wb.get('name', '') for wb in downstream_wb]),
                        'downstream_workbook_count': len(downstream_wb)
                    })

            # Get virtual connections (centralized connection management)
            vc_data = self.metadata_api.get_virtual_connections()
            if vc_data and 'virtualConnections' in vc_data:
                for vc in vc_data['virtualConnections']:
                    owner = vc.get('owner') or {}
                    tables = vc.get('tables') or []
                    upstream_dbs = vc.get('upstreamDatabases') or []
                    upstream_tables = vc.get('upstreamTables') or []
                    downstream_ds = vc.get('downstreamDatasources') or []
                    downstream_wb = vc.get('downstreamWorkbooks') or []

                    results['virtual_connections'].append({
                        'vc_id': vc.get('id', ''),
                        'vc_luid': vc.get('luid', ''),
                        'vc_name': vc.get('name', ''),
                        'description': vc.get('description', ''),
                        'project_name': vc.get('projectName', ''),
                        'created_at': vc.get('createdAt', ''),
                        'updated_at': vc.get('updatedAt', ''),
                        'is_certified': vc.get('isCertified', False),
                        'has_active_warning': vc.get('hasActiveWarning', False),
                        'owner_name': owner.get('name', ''),
                        'owner_email': owner.get('email', ''),
                        'tables': ', '.join([t.get('name', '') for t in tables]),
                        'table_count': len(tables),
                        'upstream_databases': ', '.join([db.get('name', '') for db in upstream_dbs]),
                        'upstream_connection_types': ', '.join(set([db.get('connectionType', '') for db in upstream_dbs if db.get('connectionType')])),
                        'upstream_tables': ', '.join([t.get('fullName', t.get('name', '')) for t in upstream_tables]),
                        'downstream_datasources': ', '.join([ds.get('name', '') for ds in downstream_ds]),
                        'downstream_datasource_count': len(downstream_ds),
                        'downstream_workbooks': ', '.join([wb.get('name', '') for wb in downstream_wb]),
                        'downstream_workbook_count': len(downstream_wb)
                    })

        except Exception as e:
            logger.error(f"Metadata API extraction error: {str(e)}")

    def _download_and_parse(self, kind, item, results):
        """
        Download ONE workbook/datasource into its own temp dir, parse it, and always delete the files.
        (Previously every download was kept until the very end - on a big site that fills the disk.)
        """
        temp_dir = tempfile.mkdtemp(prefix='tme_')
        extract_dir = None
        try:
            # File name = id (names can contain characters that are invalid in file names / dots)
            target = os.path.join(temp_dir, item.id)
            endpoint = self.server.workbooks if kind == 'workbook' else self.server.datasources
            downloaded = self._retry(endpoint.download, item.id, filepath=target, include_extract=False,
                                     stage='files', item=item.name)
            lower = str(downloaded).lower()

            if kind == 'workbook':
                if lower.endswith('.twbx'):
                    path, extract_dir = XMLParser.extract_from_twbx(downloaded)
                elif lower.endswith('.twb'):
                    path = downloaded
                else:
                    self._warn('files', f'Unsupported workbook file type: {os.path.splitext(lower)[1]}', item.name)
                    return
                if not path:
                    self._warn('files', 'No .twb found inside the package', item.name)
                    return
                parsed = XMLParser.parse_workbook(path, item.name, item.id)
            else:
                if lower.endswith('.tdsx'):
                    path, extract_dir = XMLParser.extract_from_tdsx(downloaded)
                elif lower.endswith('.tds'):
                    path = downloaded
                else:
                    self._warn('files', f'Unsupported data source file type: {os.path.splitext(lower)[1]}', item.name)
                    return
                if not path:
                    self._warn('files', 'No .tds found inside the package', item.name)
                    return
                parsed = XMLParser.parse_datasource(path, item.name, item.id)

            self._merge_parsed_results(results, parsed)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
            if extract_dir:
                shutil.rmtree(extract_dir, ignore_errors=True)

    def _extract_from_files(self, results, job_id=None):
        """Download and parse workbook/datasource files (only the scoped items from the inventory step)"""
        try:
            filters_active = self._filters_active
            filter_msg = " (filtered)" if filters_active else ""

            workbooks = self.filtered_workbooks or []
            datasources = self.filtered_datasources or []
            logger.info(f"[Deep Extraction] filters_active={filters_active} "
                        f"workbooks={len(workbooks)} datasources={len(datasources)}")

            total = len(workbooks)
            failed = 0
            for idx, wb in enumerate(workbooks):
                self._set_status(40 + int((idx / max(total, 1)) * 30),
                                 f'Parsing workbook {idx+1}/{total}{filter_msg}: {wb.name}')
                try:
                    self._download_and_parse('workbook', wb, results)
                except Exception as e:
                    failed += 1
                    self._warn('files', f'Error processing workbook: {e}', wb.name)

            total_ds = len(datasources)
            for idx, ds in enumerate(datasources):
                self._set_status(70 + int((idx / max(total_ds, 1)) * 10),
                                 f'Parsing datasource {idx+1}/{total_ds}{filter_msg}: {ds.name}')
                try:
                    self._download_and_parse('datasource', ds, results)
                except Exception as e:
                    failed += 1
                    self._warn('files', f'Error processing data source: {e}', ds.name)

            if failed:
                self._warn('files', f'{failed} of {total + total_ds} files could not be downloaded/parsed')

        except Exception as e:
            self._warn('files', f'File parsing error: {e}')

    def _merge_parsed_results(self, results, parsed):
        """Merge parsed results into main results"""
        for key in parsed:
            if key in results and isinstance(results[key], list):
                results[key].extend(parsed[key])

    def _extract_flow_content(self, results, job_id=None):
        """
        Download and parse each in-scope Prep flow (.tfl/.tflx) with workbook_analyzer.FlowAnalyzer -
        the same engine used by Single Workbook Analysis - and merge its output into the flow_* result
        keys (see FLOW_SCHEMA_MAP). One flow at a time, its temp file always deleted afterwards.
        """
        flows = self.filtered_flows or []
        total = len(flows)
        if not total:
            return
        failed = 0
        for idx, f in enumerate(flows):
            self._set_status(80 + int((idx / max(total, 1)) * 5), f'Parsing flow {idx + 1}/{total}: {f.name}')
            temp_dir = tempfile.mkdtemp(prefix='tme_flow_')
            try:
                target = os.path.join(temp_dir, f.id)
                downloaded = self._retry(self.server.flows.download, f.id, filepath=target,
                                         stage='flows', item=f.name)
                path = str(downloaded)
                if os.path.splitext(path)[1].lower() not in workbook_analyzer.FLOW_EXTS:
                    # TSC didn't give us a recognizable extension - detect from content and rename
                    renamed = path + ('.tflx' if zipfile.is_zipfile(path) else '.tfl')
                    os.replace(path, renamed)
                    path = renamed

                out = workbook_analyzer.analyze_file(path, display_name=f.name, use_cache=False)

                summary = (out.get('workbook_summary') or [{}])[0]
                if str(summary.get('Status', '')).startswith('FAILED'):
                    failed += 1
                    self._warn('flows', summary['Status'], f.name)

                for src_key, dst_key in self.FLOW_SCHEMA_MAP.items():
                    rows = out.get(src_key)
                    if rows:
                        results[dst_key].extend(rows)

                # Also fold flow tables into object_tables (object_tables_summary is built from this list
                # right after this method returns) so flows show up in the "Tables Used" tab/export
                # alongside workbooks and data sources, not just in the separate Flow_Tables sheet.
                for row in (out.get('table_inventory') or []):
                    results['object_tables'].append({
                        'object_id': f.id, 'object_name': f.name, 'object_type': 'Flow',
                        'connection_type': row.get('Connection Type', ''), 'connection_server': row.get('Server', ''),
                        'port': '', 'warehouse': '', 'catalog': '', 'authentication': '',
                        'database': row.get('Database', ''), 'schema': row.get('Schema', ''),
                        'table_name': row.get('Table Name', ''), 'full_table_name': row.get('Full Table Name', ''),
                        'table_kind': row.get('Table Kind', ''), 'datasource_name': row.get('Datasource', ''),
                        'connection_class': '', 'table_alias': row.get('Table Alias', ''),
                        'table_reference': row.get('Physical Table', ''), 'source': row.get('Source', 'Flow'),
                    })

                errors = [d for d in (out.get('data_quality') or []) if d.get('Severity') == 'Error']
                if errors:
                    self._warn('flows', f'{len(errors)} data-quality issue(s) found - see Flow_Data_Quality sheet',
                               f.name)
            except Exception as e:
                failed += 1
                self._warn('flows', f'Error processing flow: {e}', f.name)
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        if failed:
            self._warn('flows', f'{failed} of {total} flows could not be downloaded/parsed')

    def _extract_permissions(self, results):
        """Extract permissions from REST API for the scoped workbooks, data sources and projects"""
        try:
            targets = []
            for wb in (self.filtered_workbooks or []):
                targets.append(('Workbook', wb, self.server.workbooks))
            for ds in (self.filtered_datasources or []):
                targets.append(('Data Source', ds, self.server.datasources))

            projects = self._pager_list(self.server.projects, 'permissions', 'projects')
            if self.filtered_project_ids:
                projects = [p for p in projects if p.id in self.filtered_project_ids]
            for p in projects:
                targets.append(('Project', p, self.server.projects))

            total = len(targets)
            failed = 0
            for idx, (object_type, obj, endpoint) in enumerate(targets):
                if idx % 25 == 0:
                    self._set_status(85 + int(5 * idx / max(total, 1)),
                                     f'Permissions {idx + 1}/{total}: {obj.name}')
                try:
                    self._retry(endpoint.populate_permissions, obj, stage='permissions', item=obj.name)
                    for rule in obj.permissions:
                        grantee_type = 'User' if rule.grantee.tag_name == 'user' else 'Group'
                        for cap in rule.capabilities:
                            results['permissions'].append({
                                'object_id': obj.id,
                                'object_type': object_type,
                                'object_name': obj.name,
                                'grantee_type': grantee_type,
                                'grantee_id': rule.grantee.id,
                                'capability': cap.name,
                                'mode': cap.mode
                            })
                except Exception as e:
                    failed += 1
                    self._warn('permissions', f'Could not read permissions: {e}', f'{object_type} {obj.name}')

            if failed:
                self._warn('permissions', f'Permissions could not be read for {failed} of {total} objects')

        except Exception as e:
            self._warn('permissions', f'Permissions extraction error: {e}')

    def _manual_pages(self, fetch, stage, max_pages=500, **kwargs):
        """Page through endpoints whose get() does not work with TSC.Pager (jobs, tasks, subscriptions)."""
        items = []
        try:
            page = 1
            while page <= max_pages:
                opts = TSC.RequestOptions(pagenumber=page, pagesize=1000)
                batch, pagination = fetch(req_options=opts, **kwargs)
                items.extend(batch)
                if not batch or page * pagination.page_size >= pagination.total_available:
                    break
                page += 1
        except Exception as e:
            self._warn(stage, f'stopped after {len(items)} items: {e}')
        return items

    def _extract_admin_data(self, results):
        """Extract administrative data from REST API"""
        try:
            for u in self._pager_list(self.server.users, 'admin', 'users'):
                results['users'].append({
                    'id': u.id,
                    'name': u.name,
                    'fullname': u.fullname if hasattr(u, 'fullname') else '',
                    'email': u.email if hasattr(u, 'email') else '',
                    'site_role': u.site_role,
                    'auth_setting': u.auth_setting if hasattr(u, 'auth_setting') else '',
                    'last_login': str(u.last_login) if hasattr(u, 'last_login') and u.last_login else ''
                })

            for g in self._pager_list(self.server.groups, 'admin', 'groups'):
                results['groups_admin'].append({
                    'id': g.id,
                    'name': g.name,
                    'domain_name': g.domain_name if hasattr(g, 'domain_name') else '',
                    'license_mode': g.license_mode if hasattr(g, 'license_mode') else ''
                })

            for s in self._pager_list(self.server.schedules, 'admin', 'schedules'):
                results['schedules'].append({
                    'id': s.id,
                    'name': s.name,
                    'schedule_type': s.schedule_type if hasattr(s, 'schedule_type') else '',
                    'state': s.state if hasattr(s, 'state') else '',
                    'priority': s.priority if hasattr(s, 'priority') else '',
                    'frequency': s.frequency if hasattr(s, 'frequency') else '',
                    'next_run_at': str(s.next_run_at) if hasattr(s, 'next_run_at') and s.next_run_at else ''
                })

            for j in self._manual_pages(self.server.jobs.get, 'admin_jobs'):
                results['jobs'].append({
                    'id': j.id,
                    'job_type': j.type if hasattr(j, 'type') else '',
                    'status': j.status if hasattr(j, 'status') else '',
                    'progress': j.progress if hasattr(j, 'progress') else '',
                    'created_at': str(j.created_at) if hasattr(j, 'created_at') and j.created_at else '',
                    'completed_at': str(j.completed_at) if hasattr(j, 'completed_at') and j.completed_at else ''
                })

            for t in self._manual_pages(self.server.tasks.get, 'admin_tasks'):
                results['tasks'].append({
                    'id': t.id,
                    'task_type': t.task_type if hasattr(t, 'task_type') else '',
                    'priority': t.priority if hasattr(t, 'priority') else '',
                    'schedule_id': t.schedule_id if hasattr(t, 'schedule_id') else ''
                })

            for s in self._manual_pages(self.server.subscriptions.get, 'admin_subscriptions'):
                results['subscriptions'].append({
                    'id': s.id,
                    'subject': s.subject if hasattr(s, 'subject') else '',
                    'user_id': s.user_id if hasattr(s, 'user_id') else '',
                    'content_type': s.content_type if hasattr(s, 'content_type') else '',
                    'schedule_id': s.schedule_id if hasattr(s, 'schedule_id') else ''
                })

        except Exception as e:
            self._warn('admin', f'Admin data extraction error: {e}')


# (sheet name, key in the deep-extraction results) - order = order in the export
DEEP_EXPORT_SHEETS = [
    ('Inventory', 'inventory'), ('Overall_Lineage', 'overall_lineage'),
    ('Object_Tables_Summary', 'object_tables_summary'), ('Object_Tables', 'object_tables'),
    ('Lineage_Summary', 'lineage_summary'), ('Lineage_Tables_Used', 'lineage_tables_used'),
    ('Workbook_Lineage', 'lineage_workbooks'),
    ('Datasource_Lineage', 'lineage_datasources'), ('Flow_Lineage', 'lineage_flows'),
    ('Custom_SQL_Analysis', 'lineage_custom_sql'), ('Table_Usage_Matrix', 'lineage_table_matrix'),
    ('Impact_Analysis', 'lineage_impact'),
    ('Flow_Summary', 'flow_summary'), ('Flow_IO', 'flow_io'), ('Flow_Connections', 'flow_connections'),
    ('Flow_Tables', 'flow_tables'), ('Flow_Joins', 'flow_joins'), ('Flow_Unions', 'flow_unions'),
    ('Flow_Custom_SQL', 'flow_custom_sql'), ('Flow_Calculations', 'flow_calculations'),
    ('Flow_Parameters', 'flow_parameters'), ('Flow_Filters', 'flow_filters'),
    ('Flow_Field_Usage', 'flow_field_usage'), ('Flow_Content_Lineage', 'flow_content_lineage'),
    ('Flow_Impact', 'flow_impact'), ('Flow_Data_Quality', 'flow_data_quality'), ('Flow_Steps', 'flow_steps'),
    ('Connections', 'connections'), ('Database_Servers', 'database_servers'),
    ('Virtual_Connections', 'virtual_connections'), ('Bridge_Connections', 'bridge_connections'),
    ('Custom_SQL', 'custom_sql'), ('Calculations', 'calculations'), ('Columns', 'columns'),
    ('Tables', 'tables'), ('Lineage', 'lineage'), ('Permissions', 'permissions'),
    ('Parameters', 'parameters'), ('Parameters_Enhanced', 'parameters_enhanced'), ('Filters', 'filters'),
    ('Sets', 'sets'), ('Groups', 'groups'), ('Worksheets', 'worksheets'), ('Dashboards', 'dashboards'),
    ('Actions', 'actions'), ('Custom_Visuals', 'custom_visuals'), ('Relationships', 'relationships'),
    ('Users', 'users'), ('Groups_Admin', 'groups_admin'), ('Schedules', 'schedules'), ('Jobs', 'jobs'),
    ('Tasks', 'tasks'), ('Subscriptions', 'subscriptions'), ('Warnings', 'warnings'),
]


def export_to_comprehensive_excel(results, filename_prefix='tableau_metadata', fmt='xlsx'):
    """
    Export all deep-extraction results. Returns (file_path, download_name, mimetype).

    Streams rows to disk and buckets the output, so a huge site never hits Excel's row limit or runs
    out of memory: one .xlsx when it fits, otherwise a .zip of several .xlsx files (fmt='csv' gives
    a .zip of CSVs). See export_utils.build_export.
    """
    datasets = [(sheet, results.get(key) or []) for sheet, key in DEEP_EXPORT_SHEETS]
    empty = [{'Dataset': name, 'Rows': 0} for name, rows in datasets if not rows]
    return build_export(datasets, prefix=filename_prefix, fmt=fmt, extra_summary=empty)


def _send_temp_file(path, download_name, mimetype):
    """send_file for a temporary export file that is deleted once the response has been sent."""
    response = send_file(path, mimetype=mimetype, as_attachment=True, download_name=download_name)

    def _cleanup():
        try:
            os.remove(path)
        except OSError:
            pass
    response.call_on_close(_cleanup)
    return response


def _is_deep_result(result):
    return isinstance(result, dict) and 'inventory' in result and 'stats' in result


def _preview_result(result, limit=None):
    """
    Copy of a deep-extraction result that is safe to send to the browser: every dataset is capped at
    `limit` rows, and `counts` carries the true totals. The full data stays on the server.
    """
    limit = limit or PREVIEW_ROWS_PER_DATASET
    preview, counts, truncated = {}, {}, False
    for key, value in result.items():
        if isinstance(value, list):
            counts[key] = len(value)
            if len(value) > limit:
                preview[key] = value[:limit]
                truncated = True
            else:
                preview[key] = value
        else:
            preview[key] = value
    preview['counts'] = counts
    preview['truncated'] = truncated
    preview['preview_limit'] = limit
    return preview



# ==================== Simple Extractor for Quick Operations ====================
class TableauMetadataExtractor:
    """Simple extractor for dashboard stats and quick operations"""

    def __init__(self):
        self.config_data = {}

    def authenticate(self):
        server = TSC.Server(self.config_data['server_url'], use_server_version=True)
        auth = TSC.PersonalAccessTokenAuth(
            self.config_data['pat_name'],
            self.config_data['pat_token'],
            self.config_data['site_name']
        )
        return server, auth

    def get_dashboard_stats(self):
        server, auth = self.authenticate()
        stats = {
            'projects': 0, 'workbooks': 0, 'datasources': 0, 'flows': 0,
            'views': 0, 'users': 0, 'groups': 0, 'schedules': 0,
            'custom_sql_count': 0, 'calculations_count': 0, 'lineage_edges': 0
        }

        with server.auth.sign_in(auth):
            try:
                # Use Pager to get ALL items (not just first 100)
                stats['projects'] = sum(1 for _ in TSC.Pager(server.projects))
            except: pass

            try:
                stats['workbooks'] = sum(1 for _ in TSC.Pager(server.workbooks))
            except: pass

            try:
                stats['datasources'] = sum(1 for _ in TSC.Pager(server.datasources))
            except: pass

            try:
                stats['flows'] = sum(1 for _ in TSC.Pager(server.flows))
            except: pass

            try:
                stats['views'] = sum(1 for _ in TSC.Pager(server.views))
            except: pass

            try:
                stats['users'] = sum(1 for _ in TSC.Pager(server.users))
            except: pass

            try:
                stats['groups'] = sum(1 for _ in TSC.Pager(server.groups))
            except: pass

            try:
                stats['schedules'] = sum(1 for _ in TSC.Pager(server.schedules))
            except: pass

        return stats

    def fetch_metadata(self):
        server, auth = self.authenticate()
        data = []

        with server.auth.sign_in(auth):
            # Use Pager to get ALL items (not just first 100)
            projects = list(TSC.Pager(server.projects))
            workbooks = list(TSC.Pager(server.workbooks))
            datasources = list(TSC.Pager(server.datasources))

            try:
                flows = list(TSC.Pager(server.flows))
            except:
                flows = []

            # Site
            data.append({
                "ID": "site_" + self.config_data['site_name'],
                "Type": "Site",
                "Name": self.config_data['site_name'],
                "Parent": "",
                "URL": self.config_data['server_url']
            })

            # Projects
            for p in projects:
                data.append({
                    "ID": f"proj_{p.id}",
                    "Type": "Project",
                    "Name": p.name,
                    "Parent": self.config_data['site_name'],
                    "URL": ""
                })

            # Workbooks
            for wb in workbooks:
                data.append({
                    "ID": f"wb_{wb.id}",
                    "Type": "Workbook",
                    "Name": wb.name,
                    "Parent": wb.project_name,
                    "URL": wb.webpage_url if hasattr(wb, 'webpage_url') else ""
                })

            # Datasources
            for ds in datasources:
                data.append({
                    "ID": f"ds_{ds.id}",
                    "Type": "Data Source",
                    "Name": ds.name,
                    "Parent": ds.project_name,
                    "URL": ds.webpage_url if hasattr(ds, 'webpage_url') else ""
                })

            # Flows
            for f in flows:
                data.append({
                    "ID": f"flow_{f.id}",
                    "Type": "Flow",
                    "Name": f.name,
                    "Parent": f.project_name,
                    "URL": ""
                })

            # Views - Use Pager to get ALL views
            try:
                views = list(TSC.Pager(server.views))
                for v in views:
                    # Find parent workbook name
                    parent_wb = next((wb.name for wb in workbooks if wb.id == getattr(v, 'workbook_id', None)), '')
                    data.append({
                        "ID": f"view_{v.id}",
                        "Type": "View",
                        "Name": v.name,
                        "Parent": parent_wb,
                        "URL": v.content_url if hasattr(v, 'content_url') else ""
                    })
            except Exception as e:
                logger.warning(f"Error fetching views: {str(e)}")

        return data

    def get_users(self):
        server, auth = self.authenticate()
        users_data = []
        from datetime import datetime, timezone
        with server.auth.sign_in(auth):
            # Use Pager to get ALL users (not just first 100)
            users = list(TSC.Pager(server.users))
            for u in users:
                last_login = getattr(u, 'last_login', None)
                last_login_str = str(last_login) if last_login else ''

                # Calculate days since last login and activity status
                days_since_login = None
                activity_status = 'Never'
                if last_login:
                    try:
                        # Handle timezone-aware datetime
                        now = datetime.now(timezone.utc)
                        if hasattr(last_login, 'tzinfo') and last_login.tzinfo:
                            days_since_login = (now - last_login).days
                        else:
                            days_since_login = (datetime.now() - last_login).days

                        if days_since_login <= 7:
                            activity_status = 'Active'
                        elif days_since_login <= 30:
                            activity_status = 'Recent'
                        elif days_since_login <= 90:
                            activity_status = 'Inactive'
                        else:
                            activity_status = 'Dormant'
                    except:
                        pass

                users_data.append({
                    'id': u.id,
                    'name': u.name,
                    'fullname': getattr(u, 'fullname', ''),
                    'email': getattr(u, 'email', ''),
                    'site_role': u.site_role,
                    'last_login': last_login_str,
                    'days_since_login': days_since_login,
                    'activity_status': activity_status,
                    'auth_setting': getattr(u, 'auth_setting', '')
                })
        return users_data

    def get_groups(self):
        server, auth = self.authenticate()
        groups_data = []
        with server.auth.sign_in(auth):
            # Use Pager to get ALL groups (not just first 100)
            groups = list(TSC.Pager(server.groups))
            for g in groups:
                groups_data.append({
                    'id': g.id,
                    'name': g.name,
                    'domain_name': getattr(g, 'domain_name', ''),
                    'license_mode': getattr(g, 'license_mode', '')
                })
        return groups_data

    def get_schedules(self):
        server, auth = self.authenticate()
        schedules_data = []
        with server.auth.sign_in(auth):
            try:
                # Use Pager to get ALL schedules (not just first 100)
                schedules = list(TSC.Pager(server.schedules))
                for s in schedules:
                    schedules_data.append({
                        'id': s.id,
                        'name': s.name,
                        'schedule_type': getattr(s, 'schedule_type', ''),
                        'state': getattr(s, 'state', ''),
                        'frequency': getattr(s, 'frequency', '')
                    })
            except:
                pass
        return schedules_data

    def get_jobs(self):
        server, auth = self.authenticate()
        jobs_data = []
        with server.auth.sign_in(auth):
            try:
                jobs, _ = server.jobs.get()
                for j in jobs:
                    jobs_data.append({
                        'id': j.id,
                        'job_type': getattr(j, 'type', ''),
                        'status': getattr(j, 'status', ''),
                        'progress': getattr(j, 'progress', '')
                    })
            except:
                pass
        return jobs_data

    def get_projects(self):
        server, auth = self.authenticate()
        projects_data = []
        with server.auth.sign_in(auth):
            try:
                # Use Pager to get ALL projects (not just first 100)
                projects = list(TSC.Pager(server.projects))
                for p in projects:
                    projects_data.append({
                        'id': p.id,
                        'name': p.name,
                        'description': getattr(p, 'description', ''),
                        'parent_id': getattr(p, 'parent_id', '')
                    })
            except:
                pass
        return projects_data


extractor = TableauMetadataExtractor()


# ==================== Flask Routes ====================

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/dashboard')
def dashboard():
    return render_template('dashboard.html')


@app.route('/ai-intelligence')
def ai_intelligence():
    """Render the AI Intelligence workspace page"""
    # Get current connection info from extractor if available
    server_url = getattr(extractor, 'server_url', None)
    username = getattr(extractor, 'username', None)
    site_name = getattr(extractor, 'site_name', None)

    return render_template('ai_intelligence.html',
                         server_url=server_url,
                         username=username,
                         site_name=site_name)


@app.route('/api/load_config', methods=['POST'])
def load_config():
    try:
        config_data = request.get_json()
        if not config_data:
            return jsonify({'error': 'No configuration data provided'})

        missing = [k for k in ['server_url', 'site_name', 'pat_name', 'pat_token']
                   if k not in config_data or not config_data[k].strip()]
        if missing:
            return jsonify({'error': f"Missing fields: {', '.join(missing)}"})

        # Detect if this is Tableau Cloud
        server_url = config_data.get('server_url', '').lower()
        is_cloud = (
            '.online.tableau.com' in server_url or
            'prod-useast-a.online.tableau.com' in server_url or
            'prod-useast-b.online.tableau.com' in server_url or
            'prod-apnortheast-a.online.tableau.com' in server_url or
            'prod-uk-a.online.tableau.com' in server_url or
            'dub01.online.tableau.com' in server_url or
            '10ax.online.tableau.com' in server_url or
            '10ay.online.tableau.com' in server_url or
            '10az.online.tableau.com' in server_url or
            'us-east-1.online.tableau.com' in server_url or
            'us-west-2.online.tableau.com' in server_url or
            'eu-west-1.online.tableau.com' in server_url or
            'ap-southeast-2.online.tableau.com' in server_url
        )
        config_data['is_cloud'] = is_cloud
        logger.info(f"Config: Server URL = {server_url}, is_cloud = {is_cloud}")

        # Clear any cached analytics data from previous session
        global admin_insights_analyzer, job_results, job_status
        admin_insights_analyzer = AdminInsightsAnalyzer()
        for _old in list(job_results.values()):
            _release_result(_old)
        job_results.clear()
        job_status.clear()
        logger.info("Cleared cached analytics data for new connection")

        extractor.config_data = config_data
        return jsonify({'success': True, 'is_cloud': is_cloud})
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/api/test_connection', methods=['POST'])
def test_connection():
    try:
        config_data = request.get_json()
        extractor.config_data = config_data
        server, auth = extractor.authenticate()
        with server.auth.sign_in(auth):
            return jsonify({'success': True, 'message': 'Connection successful!'})
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/api/dashboard_stats', methods=['GET'])
def dashboard_stats():
    try:
        stats = extractor.get_dashboard_stats()
        return jsonify({'success': True, 'data': stats})
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/api/fetch_metadata', methods=['GET'])
def fetch_metadata():
    try:
        data = extractor.fetch_metadata()
        return jsonify({'success': True, 'data': data})
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/api/users', methods=['GET'])
def get_users():
    try:
        data = extractor.get_users()
        return jsonify({'success': True, 'data': data})
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/api/groups', methods=['GET'])
def get_groups():
    try:
        data = extractor.get_groups()
        return jsonify({'success': True, 'data': data})
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/api/schedules', methods=['GET'])
def get_schedules():
    try:
        data = extractor.get_schedules()
        return jsonify({'success': True, 'data': data})
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/api/jobs', methods=['GET'])
def get_jobs():
    try:
        data = extractor.get_jobs()
        return jsonify({'success': True, 'data': data})
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/api/extract_deep_metadata', methods=['POST'])
def extract_deep_metadata():
    """Start deep metadata extraction as background job"""
    try:
        # Ensure worker thread is running
        ensure_worker_running()

        request_data = request.get_json() or {}

        # Try to get config from request first (frontend sends from sessionStorage)
        # Fall back to extractor.config_data
        config_data = None
        if request_data.get('config'):
            config_data = request_data.get('config')
            # Update extractor config as well
            extractor.config_data = config_data
        elif extractor.config_data:
            config_data = extractor.config_data.copy()

        if not config_data:
            return jsonify({'error': 'No configuration available. Please reconnect to the server.'})

        # Add all filter parameters to config
        project_filter = request_data.get('project_id', 'all')
        project_ids = request_data.get('project_ids', [])
        project_name = request_data.get('project_name', 'All Projects')
        content_types = request_data.get('content_types', [])
        object_ids = request_data.get('object_ids', [])

        # Handle comma-separated project IDs
        if isinstance(project_filter, str) and ',' in project_filter:
            project_ids = project_filter.split(',')
            project_filter = 'multiple'

        config_data['project_filter'] = project_filter
        config_data['project_ids'] = project_ids
        config_data['project_name'] = project_name
        config_data['content_types'] = content_types
        config_data['object_ids'] = object_ids
        config_data['lineage_only'] = bool(request_data.get('lineage_only'))

        job_id = str(uuid.uuid4())
        job_status[job_id] = {'status': 'queued', 'progress': 0, 'message': 'Queued for processing', 'project_name': project_name}

        def run_extraction(config, job_id):
            deep_extractor = DeepMetadataExtractor(config)
            return deep_extractor.extract_all(job_id=job_id)

        job_queue.put((job_id, run_extraction, (config_data,), {}))

        return jsonify({'success': True, 'job_id': job_id, 'project_name': project_name})
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/api/job_status/<job_id>', methods=['GET'])
def get_job_status(job_id):
    """Get status of a background job"""
    if job_id in job_status:
        return jsonify({'success': True, 'status': job_status[job_id]})
    else:
        return jsonify({'error': 'Job not found'})

@app.route('/api/job_result/<job_id>', methods=['GET'])
def get_job_result(job_id):
    """Get result of a completed job (deep-extraction results are capped to a preview - see _preview_result)"""
    if job_id in job_results:
        result = job_results[job_id]
        if isinstance(result, dict) and result.get('kind') == SW_KIND:
            result = _sw_public(result)
        elif _is_deep_result(result):
            result = _preview_result(result)
        return jsonify({'success': True, 'data': result})
    elif job_id in job_status:
        return jsonify({'error': 'Job not completed yet', 'status': job_status[job_id]})
    else:
        return jsonify({'error': 'Job not found'})

@app.route('/api/job_result/<job_id>/dataset/<name>', methods=['GET'])
def get_job_result_dataset(job_id, name):
    """Page through one full dataset of a stored result: ?offset=0&limit=5000"""
    result = job_results.get(job_id)
    if result is None:
        return jsonify({'error': 'Job result not found (it may have been replaced by a newer extraction)'}), 404
    rows = result.get(name) if isinstance(result, dict) else None
    if not isinstance(rows, list):
        return jsonify({'error': f'Unknown dataset: {name}'}), 404
    try:
        offset = max(0, int(request.args.get('offset', 0)))
        limit = min(50000, max(1, int(request.args.get('limit', PREVIEW_ROWS_PER_DATASET))))
    except ValueError:
        return jsonify({'error': 'offset and limit must be integers'}), 400
    return jsonify({'success': True, 'dataset': name, 'total': len(rows), 'offset': offset,
                    'rows': rows[offset:offset + limit]})

LINEAGE_EXPORT_SHEETS = [
    ('Lineage_Summary', 'lineage_summary'), ('Lineage_Tables_Used', 'lineage_tables_used'),
    ('Workbook_Lineage', 'lineage_workbooks'),
    ('Datasource_Lineage', 'lineage_datasources'), ('Flow_Lineage', 'lineage_flows'),
    ('Custom_SQL_Analysis', 'lineage_custom_sql'), ('Table_Usage_Matrix', 'lineage_table_matrix'),
    ('Impact_Analysis', 'lineage_impact'),
]


@app.route('/api/export_lineage/<job_id>', methods=['GET'])
def export_lineage(job_id):
    """Download the seven Lineage Info sheets (bucketed into several files when very large)."""
    try:
        result = job_results.get(job_id)
        if not isinstance(result, dict):
            return jsonify({'error': 'Job result not found (it may have been replaced by a newer extraction)'}), 404
        if request.args.get('check'):
            return jsonify({'success': True})
        fmt = 'csv' if request.args.get('format', 'xlsx').lower() == 'csv' else 'xlsx'
        path, name, mimetype = build_export(
            [(sheet, result.get(key) or []) for sheet, key in LINEAGE_EXPORT_SHEETS],
            prefix='tableau_lineage_info', fmt=fmt)
        return _send_temp_file(path, name, mimetype)
    except Exception as e:
        logger.error(f"Lineage export error: {e}", exc_info=True)
        return jsonify({'error': str(e)}), 500


@app.route('/api/export_object_tables/<job_id>', methods=['GET'])
def export_object_tables(job_id):
    """Download only the 'tables used per workbook / data source' lists (summary + one row per table)."""
    try:
        result = job_results.get(job_id)
        if not isinstance(result, dict):
            return jsonify({'error': 'Job result not found (it may have been replaced by a newer extraction)'}), 404
        if request.args.get('check'):
            return jsonify({'success': True})
        fmt = 'csv' if request.args.get('format', 'xlsx').lower() == 'csv' else 'xlsx'
        path, name, mimetype = build_export(
            [('Object_Tables_Summary', result.get('object_tables_summary') or []),
             ('Object_Tables', result.get('object_tables') or [])],
            prefix='tableau_tables_used', fmt=fmt)
        return _send_temp_file(path, name, mimetype)
    except Exception as e:
        logger.error(f"Object tables export error: {e}", exc_info=True)
        return jsonify({'error': str(e)}), 500


@app.route('/api/export_job_result/<job_id>', methods=['GET'])
def export_job_result(job_id):
    """
    Export the full job result. ?format=xlsx (default) or ?format=csv.
    Returns one .xlsx when it fits, otherwise a .zip with the data split into several files.
    """
    try:
        if job_id not in job_results:
            return jsonify({'error': 'Job result not found (it may have been replaced by a newer extraction)'}), 404

        if request.args.get('check'):
            return jsonify({'success': True})

        fmt = 'csv' if request.args.get('format', 'xlsx').lower() == 'csv' else 'xlsx'
        path, name, mimetype = export_to_comprehensive_excel(job_results[job_id], 'tableau_deep_metadata', fmt)
        return _send_temp_file(path, name, mimetype)
    except Exception as e:
        logger.error(f"Export error: {e}", exc_info=True)
        return jsonify({'error': str(e)}), 500

@app.route('/api/export_excel', methods=['POST'])
def export_excel():
    try:
        data = request.json.get('data', [])
        export_type = request.json.get('type', 'metadata')

        if not data:
            return jsonify({'error': 'No data to export'}), 400

        if len(data) > LARGE_EXPORT_ROWS:
            path, name, mimetype = build_export([('Data', data)], prefix=str(export_type))
            return _send_temp_file(path, name, mimetype)

        df = pd.DataFrame(data)
        df = sanitize_dataframe_for_excel(df)  # Sanitize for safe export
        output = BytesIO()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Data')
            worksheet = writer.sheets['Data']

            # Use safe column width calculation and formatting
            apply_excel_formatting(worksheet, df)

        output.seek(0)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'{export_type}_{timestamp}.xlsx'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ==================== Metadata GraphQL API ====================

def get_metadata_api_client():
    """Helper to get authenticated Metadata API client"""
    if not extractor or not extractor.config_data:
        return None, 'Not connected to Tableau Server. Please connect first.'

    config_data = extractor.config_data
    server_url = config_data.get('server_url', '').rstrip('/')

    if not server_url:
        return None, 'Server URL not configured'

    try:
        server = TSC.Server(server_url, use_server_version=True)
        auth = TSC.PersonalAccessTokenAuth(
            config_data['pat_name'],
            config_data['pat_token'],
            config_data['site_name']
        )

        server.auth.sign_in(auth)
        metadata_api = MetadataAPIClient(server_url, server.auth_token, server.site_id)

        if not metadata_api.available:
            server.auth.sign_out()
            return None, 'Metadata API access is not enabled or user lacks permission.'

        return (metadata_api, server), None
    except Exception as e:
        return None, str(e)


@app.route('/api/metadata/projects', methods=['GET'])
def get_metadata_projects():
    """Get list of projects for filtering"""
    try:
        if not extractor or not extractor.config_data:
            return jsonify({'error': 'Not connected to Tableau Server'}), 400

        config_data = extractor.config_data
        server = TSC.Server(config_data['server_url'], use_server_version=True)
        auth = TSC.PersonalAccessTokenAuth(
            config_data['pat_name'],
            config_data['pat_token'],
            config_data['site_name']
        )

        with server.auth.sign_in(auth):
            # Use Pager to get ALL projects (not just first 100)
            projects = list(TSC.Pager(server.projects))
            project_list = [{'id': p.id, 'name': p.name} for p in projects]
            project_list.sort(key=lambda x: x['name'].lower())

        return jsonify({
            'success': True,
            'projects': project_list,
            'count': len(project_list)
        })

    except Exception as e:
        logger.error(f"Error fetching projects: {str(e)}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/metadata/query', methods=['POST'])
def run_metadata_query():
    """
    Run metadata queries with filtering support.

    Request body:
    {
        "query_type": "published_datasources" (single query)
        OR
        "queries": ["published_datasources", "workbooks", ...],
        "project_filter": "all" or "ProjectName",
        "name_filter": "" (optional)
    }
    """
    try:
        data = request.get_json() or {}

        # Support both single query_type and multiple queries
        queries = data.get('queries', [])
        single_query = data.get('query_type')
        if single_query and not queries:
            queries = [single_query]

        project_filter = data.get('project_filter')
        name_filter = data.get('name_filter', '').strip() if data.get('name_filter') else None

        if not queries:
            return jsonify({'error': 'No queries selected'}), 400

        result = get_metadata_api_client()
        if result[1]:  # Error
            return jsonify({'error': result[1]}), 400

        metadata_api, server = result[0]

        results = {}
        errors = {}

        try:
            for query_type in queries:
                query_result = metadata_api.fetch_metadata(
                    query_type,
                    project_filter if project_filter != 'all' else None,
                    name_filter if name_filter else None
                )

                if query_result.get('error'):
                    errors[query_type] = query_result['error']
                    results[query_type] = []
                else:
                    results[query_type] = query_result.get('data', [])

        finally:
            try:
                server.auth.sign_out()
            except:
                pass

        # For single query, return simpler format
        if single_query and len(queries) == 1:
            query_type = queries[0]
            if errors.get(query_type):
                return jsonify({
                    'success': False,
                    'error': errors[query_type]
                })
            return jsonify({
                'success': True,
                'data': results.get(query_type, []),
                'count': len(results.get(query_type, []))
            })

        # For multiple queries, return full results
        return jsonify({
            'success': True,
            'results': results,
            'errors': errors,
            'counts': {k: len(v) for k, v in results.items()}
        })

    except Exception as e:
        error_msg = str(e)
        if '20000' in error_msg or 'limit' in error_msg.lower():
            return jsonify({
                'error': 'Response exceeds GraphQL 20,000 row limit. Please select a specific project to narrow the scope.'
            }), 400
        logger.error(f"Error running metadata query: {error_msg}")
        return jsonify({'error': error_msg}), 500


@app.route('/api/metadata/published_datasources', methods=['GET'])
def get_published_datasources_graphql():
    """Fetch published datasources via Metadata GraphQL API (legacy endpoint)"""
    try:
        result = get_metadata_api_client()
        if result[1]:
            return jsonify({'error': result[1]}), 400

        metadata_api, server = result[0]

        try:
            query_result = metadata_api.fetch_metadata('published_datasources')

            if query_result.get('error'):
                return jsonify({'error': query_result['error']}), 500

            return jsonify({
                'success': True,
                'data': query_result.get('data', []),
                'count': len(query_result.get('data', []))
            })
        finally:
            try:
                server.auth.sign_out()
            except:
                pass

    except Exception as e:
        logger.error(f"Error fetching published datasources via GraphQL: {str(e)}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/metadata/export', methods=['POST'])
def export_metadata_results():
    """
    Export multiple query results to a single Excel file with multiple sheets.

    Request body:
    {
        "results": {
            "published_datasources": [...],
            "workbooks": [...],
            ...
        }
    }
    """
    try:
        data = request.json or {}
        results = data.get('results', {})

        if not results:
            return jsonify({'error': 'No data to export'}), 400

        if sum(len(r) for r in results.values() if isinstance(r, list)) > LARGE_EXPORT_ROWS:
            names = {'published_datasources': 'Published_Datasources', 'workbooks': 'Workbooks',
                     'workbook_dependency': 'Workbook_Dependency', 'workbook_field_lineage': 'Workbook_Field_Lineage',
                     'datasource_impact': 'Datasource_Impact', 'database_servers': 'Database_Servers',
                     'virtual_connections': 'Virtual_Connections'}
            datasets = [(names.get(k, k[:31]), v) for k, v in results.items() if isinstance(v, list)]
            path, name, mimetype = build_export(datasets, prefix='metadata_export')
            return _send_temp_file(path, name, mimetype)

        output = BytesIO()

        # Sheet name mapping
        sheet_names = {
            'published_datasources': 'Published_Datasources',
            'workbooks': 'Workbooks',
            'workbook_dependency': 'Workbook_Dependency',
            'workbook_field_lineage': 'Workbook_Field_Lineage',
            'datasource_impact': 'Datasource_Impact',
            'database_servers': 'Database_Servers',
            'virtual_connections': 'Virtual_Connections'
        }

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            sheets_written = 0

            for query_type, rows in results.items():
                if not rows:
                    continue

                sheet_name = sheet_names.get(query_type, query_type[:31])  # Excel sheet name limit

                df = pd.DataFrame(rows)
                df = sanitize_dataframe_for_excel(df)
                df.to_excel(writer, index=False, sheet_name=sheet_name)

                worksheet = writer.sheets[sheet_name]
                apply_excel_formatting(worksheet, df)
                sheets_written += 1

            # If no sheets were written, add a placeholder
            if sheets_written == 0:
                pd.DataFrame({'Message': ['No data to export']}).to_excel(
                    writer, index=False, sheet_name='Info'
                )
            else:
                # Create the Analysis sheet with summary and insights
                try:
                    create_analysis_sheet(writer, results)
                except Exception as analysis_error:
                    logger.warning(f"Could not create Analysis sheet: {str(analysis_error)}")
                    # Continue without the analysis sheet if it fails

        output.seek(0)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'metadata_export_{timestamp}.xlsx'
        )

    except Exception as e:
        logger.error(f"Error exporting metadata: {str(e)}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/metadata/export_datasources', methods=['POST'])
def export_published_datasources():
    """Export published datasources to Excel (legacy endpoint)"""
    try:
        data = request.json.get('data', [])

        if not data:
            return jsonify({'error': 'No data to export'}), 400

        if len(data) > LARGE_EXPORT_ROWS:
            path, name, mimetype = build_export([('Published Datasources', data)], prefix='published_datasources')
            return _send_temp_file(path, name, mimetype)

        df = pd.DataFrame(data)
        df = sanitize_dataframe_for_excel(df)
        output = BytesIO()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Published Datasources')
            worksheet = writer.sheets['Published Datasources']
            apply_excel_formatting(worksheet, df)

        output.seek(0)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'published_datasources_{timestamp}.xlsx'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ==================== Save/Load Extraction Results ====================

EXTRACTIONS_DIR = os.path.join(APP_DATA_PATH, 'extractions')

def ensure_extractions_dir():
    """Ensure the extractions directory exists"""
    if not os.path.exists(EXTRACTIONS_DIR):
        os.makedirs(EXTRACTIONS_DIR)

@app.route('/api/save_extraction', methods=['POST'])
def save_extraction():
    """Save extraction results to a JSON file"""
    try:
        ensure_extractions_dir()
        data = request.get_json()
        job_id = data.get('job_id')
        filename = data.get('filename')

        if not job_id or job_id not in job_results:
            return jsonify({'error': 'No extraction results to save'})

        results = job_results[job_id]

        # Generate filename if not provided
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            project_name = data.get('project_name', 'all_projects')
            safe_project = "".join(c for c in project_name if c.isalnum() or c in (' ', '-', '_')).strip()
            safe_project = safe_project.replace(' ', '_')[:50]
            filename = f"extraction_{safe_project}_{timestamp}.json"

        filepath = os.path.join(EXTRACTIONS_DIR, filename)

        # Add metadata to saved file
        save_data = {
            'metadata': {
                'saved_at': datetime.now().isoformat(),
                'project_filter': data.get('project_name', 'All Projects'),
                'server_url': extractor.config_data.get('server_url', '') if extractor.config_data else ''
            },
            'results': results
        }

        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(save_data, f, separators=(',', ':'), default=str)

        return jsonify({'success': True, 'filename': filename, 'filepath': filepath})
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/api/load_extraction', methods=['POST'])
def load_extraction():
    """Load extraction results from a JSON file"""
    try:
        if 'file' in request.files:
            # Load from uploaded file
            file = request.files['file']
            if file.filename == '':
                return jsonify({'error': 'No file selected'})

            content = file.read().decode('utf-8')
            data = json.loads(content)
        else:
            # Load from filename
            filename = request.json.get('filename')
            if not filename:
                return jsonify({'error': 'No filename provided'})

            filepath = os.path.join(EXTRACTIONS_DIR, filename)
            if not os.path.exists(filepath):
                return jsonify({'error': 'File not found'})

            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)

        # Extract results from saved format
        if 'results' in data:
            results = data['results']
            metadata = data.get('metadata', {})
        else:
            # Backwards compatibility - old format without metadata wrapper
            results = data
            metadata = {}

        # Create a new job ID for loaded results
        job_id = str(uuid.uuid4())
        job_results[job_id] = results
        job_status[job_id] = {'status': 'completed', 'progress': 100, 'message': 'Loaded from file'}

        _evict_old_job_results(job_id)
        return jsonify({
            'success': True,
            'job_id': job_id,
            'metadata': metadata,
            'data': _preview_result(results) if _is_deep_result(results) else results
        })
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/api/list_extractions', methods=['GET'])
def list_extractions():
    """List all saved extraction files"""
    try:
        ensure_extractions_dir()
        files = []
        for filename in os.listdir(EXTRACTIONS_DIR):
            if filename.endswith('.json'):
                filepath = os.path.join(EXTRACTIONS_DIR, filename)
                stat = os.stat(filepath)
                files.append({
                    'filename': filename,
                    'size': stat.st_size,
                    'modified': datetime.fromtimestamp(stat.st_mtime).isoformat()
                })

        # Sort by modified date descending
        files.sort(key=lambda x: x['modified'], reverse=True)
        return jsonify({'success': True, 'files': files})
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/api/delete_extraction/<filename>', methods=['DELETE'])
def delete_extraction(filename):
    """Delete a saved extraction file"""
    try:
        filepath = os.path.join(EXTRACTIONS_DIR, filename)
        if os.path.exists(filepath):
            os.remove(filepath)
            return jsonify({'success': True})
        else:
            return jsonify({'error': 'File not found'})
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/api/projects', methods=['GET'])
def get_projects():
    """Get list of projects for filtering"""
    try:
        data = extractor.get_projects()
        return jsonify({'success': True, 'data': data})
    except Exception as e:
        return jsonify({'error': str(e)})


@app.route('/api/system_user', methods=['GET'])
def get_system_user():
    """Get the current system username"""
    try:
        # Try different methods to get the username
        username = None
        try:
            username = os.getlogin()
        except:
            pass
        if not username:
            try:
                username = getpass.getuser()
            except:
                pass
        if not username:
            username = os.environ.get('USER') or os.environ.get('USERNAME') or 'User'

        # Get initials for avatar
        parts = username.split('.')
        if len(parts) >= 2:
            initials = (parts[0][0] + parts[1][0]).upper()
        else:
            initials = username[:2].upper()

        return jsonify({'success': True, 'username': username, 'initials': initials})
    except Exception as e:
        return jsonify({'success': False, 'username': 'User', 'initials': 'U'})


# ==================== Global Filter Context ====================

# In-memory filter context (per-session would be better in production)
filter_context = {
    'project_ids': [],           # List of project IDs to filter by
    'content_type': '',          # Workbook, Data Source, Flow, View
    'object_ids': [],            # List of specific object IDs
    'search_term': '',           # Global search term
    'deep_search': False,        # Search in SQL and formulas too
    'active': False              # Whether filters are active
}

def apply_filter_to_content(data, filter_ctx):
    """Apply filter context to content data"""
    if not filter_ctx.get('active', False):
        return data

    filtered = data.copy() if isinstance(data, list) else list(data)

    # Filter by project IDs
    project_ids = filter_ctx.get('project_ids', [])
    if project_ids:
        filtered = [item for item in filtered if
                    item.get('Parent', '') in [p.get('name', '') for p in extractor.get_projects() if p['id'] in project_ids] or
                    item.get('project_name', '') in [p.get('name', '') for p in extractor.get_projects() if p['id'] in project_ids] or
                    (item.get('Type') == 'Project' and item.get('ID', '').replace('proj_', '') in project_ids)]

    # Filter by content type
    content_type = filter_ctx.get('content_type', '')
    if content_type:
        filtered = [item for item in filtered if item.get('Type', '') == content_type]

    # Filter by object IDs
    object_ids = filter_ctx.get('object_ids', [])
    if object_ids:
        filtered = [item for item in filtered if item.get('ID', '') in object_ids]

    # Filter by search term
    search_term = filter_ctx.get('search_term', '').lower()
    if search_term:
        filtered = [item for item in filtered if
                    search_term in item.get('Name', '').lower() or
                    search_term in item.get('name', '').lower() or
                    search_term in str(item.get('object_name', '')).lower()]

    return filtered

def apply_filter_to_deep_data(data_key, data, filter_ctx):
    """Apply filter context to deep extraction data"""
    if not filter_ctx.get('active', False):
        return data

    filtered = data.copy() if isinstance(data, list) else list(data)
    project_ids = filter_ctx.get('project_ids', [])
    object_ids = filter_ctx.get('object_ids', [])
    search_term = filter_ctx.get('search_term', '').lower()
    deep_search = filter_ctx.get('deep_search', False)

    # Filter by project IDs if specified
    if project_ids and extractor and extractor.config_data:
        try:
            server, auth = extractor.authenticate()
            with server.auth.sign_in(auth):
                # Get all projects and expand to include children
                projects, _ = server.projects.get()
                expanded_project_ids = set(project_ids)
                changed = True
                while changed:
                    changed = False
                    for p in projects:
                        parent_id = getattr(p, 'parent_id', None)
                        if parent_id and parent_id in expanded_project_ids and p.id not in expanded_project_ids:
                            expanded_project_ids.add(p.id)
                            changed = True

                # Get workbooks and datasources in those projects
                workbooks, _ = server.workbooks.get()
                datasources, _ = server.datasources.get()
                valid_object_ids = set()
                for wb in workbooks:
                    if wb.project_id in expanded_project_ids:
                        valid_object_ids.add(wb.id)
                for ds in datasources:
                    if ds.project_id in expanded_project_ids:
                        valid_object_ids.add(ds.id)

                # Filter by object_id
                filtered = [item for item in filtered if item.get('object_id', '') in valid_object_ids]
        except Exception as e:
            logger.warning(f"Error filtering deep data by project: {str(e)}")

    # Filter by object IDs if specified
    if object_ids:
        # Extract base IDs without prefixes
        base_ids = []
        for oid in object_ids:
            if '_' in oid:
                base_ids.append(oid.split('_', 1)[1])
            else:
                base_ids.append(oid)

        filtered = [item for item in filtered if
                    item.get('object_id', '') in base_ids or
                    f"wb_{item.get('object_id', '')}" in object_ids or
                    f"ds_{item.get('object_id', '')}" in object_ids or
                    f"flow_{item.get('object_id', '')}" in object_ids]

    # Apply search term
    if search_term:
        def matches_search(item):
            # Basic name matching
            if search_term in str(item.get('object_name', '')).lower():
                return True
            if search_term in str(item.get('name', '')).lower():
                return True
            if search_term in str(item.get('datasource_name', '')).lower():
                return True

            # Deep search in SQL and formulas
            if deep_search:
                if data_key == 'custom_sql' and search_term in str(item.get('sql_text', '')).lower():
                    return True
                if data_key == 'calculations' and search_term in str(item.get('formula_text', '')).lower():
                    return True

            return False

        filtered = [item for item in filtered if matches_search(item)]

    return filtered

def apply_filter_to_lineage(lineage_data, filter_ctx):
    """Apply filter context to lineage data - show selected nodes plus upstream/downstream"""
    if not filter_ctx.get('active', False):
        return lineage_data

    object_ids = filter_ctx.get('object_ids', [])
    if not object_ids:
        return lineage_data

    # Extract object names from IDs
    object_names = set()
    for oid in object_ids:
        # Try to find the object name from content
        for item in extractor.fetch_metadata():
            if item.get('ID') == oid:
                object_names.add(item.get('Name', ''))
                break

    if not object_names:
        return lineage_data

    # Find all lineage edges that involve selected objects
    relevant_edges = []
    connected_objects = set(object_names)

    # First pass: find direct connections
    for edge in lineage_data:
        upstream = edge.get('upstream_object', '')
        downstream = edge.get('downstream_object', '')

        if upstream in object_names or downstream in object_names:
            relevant_edges.append(edge)
            connected_objects.add(upstream)
            connected_objects.add(downstream)

    # Second pass: include edges between connected objects
    for edge in lineage_data:
        upstream = edge.get('upstream_object', '')
        downstream = edge.get('downstream_object', '')

        if upstream in connected_objects and downstream in connected_objects:
            if edge not in relevant_edges:
                relevant_edges.append(edge)

    return relevant_edges


@app.route('/api/set_filter_context', methods=['POST'])
def set_filter_context():
    """Set the global filter context"""
    global filter_context
    try:
        data = request.get_json() or {}

        filter_context = {
            'project_ids': data.get('project_ids', []),
            'content_type': data.get('content_type', ''),
            'content_types': data.get('content_types', []),
            'object_ids': data.get('object_ids', []),
            'search_term': data.get('search_term', ''),
            'deep_search': data.get('deep_search', False),
            'active': bool(data.get('project_ids') or data.get('object_ids') or data.get('search_term') or data.get('content_type') or data.get('content_types'))
        }

        logger.info(f"Filter context set: project_ids={len(filter_context['project_ids'])}, object_ids={len(filter_context['object_ids'])}, content_types={filter_context['content_types']}, active={filter_context['active']}")
        return jsonify({'success': True, 'filter_context': filter_context})
    except Exception as e:
        return jsonify({'error': str(e)})


@app.route('/api/get_filter_context', methods=['GET'])
def get_filter_context():
    """Get the current filter context"""
    return jsonify({'success': True, 'filter_context': filter_context})


@app.route('/api/clear_filter_context', methods=['POST'])
def clear_filter_context():
    """Clear the global filter context"""
    global filter_context
    filter_context = {
        'project_ids': [],
        'content_type': '',
        'object_ids': [],
        'search_term': '',
        'deep_search': False,
        'active': False
    }
    return jsonify({'success': True, 'filter_context': filter_context})


@app.route('/api/content_by_type', methods=['GET'])
def content_by_type():
    """Get content filtered by type and optionally by project IDs"""
    try:
        content_type = request.args.get('type', '')
        project_ids = request.args.getlist('project_ids')

        server, auth = extractor.authenticate()
        data = []

        with server.auth.sign_in(auth):
            if content_type == 'Workbook' or not content_type:
                workbooks, _ = server.workbooks.get()
                for wb in workbooks:
                    if project_ids and wb.project_id not in project_ids:
                        continue
                    data.append({
                        'id': f"wb_{wb.id}",
                        'raw_id': wb.id,
                        'name': wb.name,
                        'type': 'Workbook',
                        'project_id': wb.project_id,
                        'project_name': wb.project_name
                    })

            if content_type == 'Data Source' or not content_type:
                datasources, _ = server.datasources.get()
                for ds in datasources:
                    if project_ids and ds.project_id not in project_ids:
                        continue
                    data.append({
                        'id': f"ds_{ds.id}",
                        'raw_id': ds.id,
                        'name': ds.name,
                        'type': 'Data Source',
                        'project_id': ds.project_id,
                        'project_name': ds.project_name
                    })

            if content_type == 'Flow' or not content_type:
                try:
                    flows, _ = server.flows.get()
                    for f in flows:
                        if project_ids and f.project_id not in project_ids:
                            continue
                        data.append({
                            'id': f"flow_{f.id}",
                            'raw_id': f.id,
                            'name': f.name,
                            'type': 'Flow',
                            'project_id': f.project_id,
                            'project_name': f.project_name
                        })
                except:
                    pass

            if content_type == 'View' or not content_type:
                try:
                    views, _ = server.views.get()
                    for v in views:
                        # Views don't have direct project_id, they're part of workbooks
                        data.append({
                            'id': f"view_{v.id}",
                            'raw_id': v.id,
                            'name': v.name,
                            'type': 'View',
                            'project_id': '',
                            'project_name': ''
                        })
                except:
                    pass

        # Sort by name
        data.sort(key=lambda x: x['name'].lower())

        return jsonify({'success': True, 'data': data})
    except Exception as e:
        return jsonify({'error': str(e)})


@app.route('/api/filtered_stats', methods=['GET'])
def filtered_stats():
    """Get dashboard stats with filter context applied"""
    try:
        base_stats = extractor.get_dashboard_stats()

        if not filter_context.get('active', False):
            return jsonify({'success': True, 'data': base_stats, 'filtered': False})

        # Get filtered content counts
        server, auth = extractor.authenticate()
        stats = base_stats.copy()

        project_ids = filter_context.get('project_ids', [])
        object_ids = filter_context.get('object_ids', [])
        content_types = filter_context.get('content_types', [])

        with server.auth.sign_in(auth):
            # Get all projects to build hierarchy
            all_projects, _ = server.projects.get()

            # Expand project_ids to include child projects
            expanded_project_ids = set()
            if project_ids:
                for pid in project_ids:
                    expanded_project_ids.add(pid)
                    # Add child projects recursively
                    changed = True
                    while changed:
                        changed = False
                        for p in all_projects:
                            parent_id = getattr(p, 'parent_id', None)
                            if parent_id and parent_id in expanded_project_ids and p.id not in expanded_project_ids:
                                expanded_project_ids.add(p.id)
                                changed = True

            if expanded_project_ids or object_ids:
                # Recalculate stats based on filter
                workbooks, _ = server.workbooks.get()
                datasources, _ = server.datasources.get()

                # Filter by expanded project IDs
                if expanded_project_ids:
                    workbooks = [wb for wb in workbooks if wb.project_id in expanded_project_ids]
                    datasources = [ds for ds in datasources if ds.project_id in expanded_project_ids]

                # Filter by specific object IDs if selected
                if object_ids:
                    wb_ids = [oid.replace('wb_', '') for oid in object_ids if oid.startswith('wb_')]
                    ds_ids = [oid.replace('ds_', '') for oid in object_ids if oid.startswith('ds_')]
                    view_ids = [oid.replace('view_', '') for oid in object_ids if oid.startswith('view_')]
                    flow_ids = [oid.replace('flow_', '') for oid in object_ids if oid.startswith('flow_')]

                    if wb_ids:
                        workbooks = [wb for wb in workbooks if wb.id in wb_ids]
                    if ds_ids:
                        datasources = [ds for ds in datasources if ds.id in ds_ids]

                stats['workbooks'] = len(workbooks)
                stats['datasources'] = len(datasources)

                # Recalculate views based on filtered workbooks
                view_count = 0
                for wb in workbooks:
                    try:
                        server.workbooks.populate_views(wb, usage=True)
                        view_count += len(wb.views) if wb.views else 0
                    except:
                        pass
                stats['views'] = view_count

                # Filter flows
                try:
                    flows, _ = server.flows.get()
                    if expanded_project_ids:
                        flows = [f for f in flows if f.project_id in expanded_project_ids]
                    if object_ids:
                        flow_ids = [oid.replace('flow_', '') for oid in object_ids if oid.startswith('flow_')]
                        if flow_ids:
                            flows = [f for f in flows if f.id in flow_ids]
                    stats['flows'] = len(flows)
                except:
                    pass

                # Update projects count to show selected + children
                if expanded_project_ids:
                    stats['projects'] = len(expanded_project_ids)

        return jsonify({'success': True, 'data': stats, 'filtered': True})
    except Exception as e:
        logger.error(f"Error in filtered_stats: {str(e)}")
        return jsonify({'error': str(e)})


@app.route('/api/filtered_job_result/<job_id>', methods=['GET'])
def filtered_job_result(job_id):
    """Get job result with filter context applied"""
    try:
        if job_id not in job_results:
            return jsonify({'error': 'Job result not found'})

        results = job_results[job_id]

        def _safe(payload):
            # Never send an entire huge extraction to the browser in one response
            return _preview_result(payload) if _is_deep_result(payload) else payload

        if not filter_context.get('active', False):
            return jsonify({'success': True, 'data': _safe(results), 'filtered': False})

        # Apply filters to each data type
        filtered_results = {}

        for key, data in results.items():
            if isinstance(data, list):
                if key == 'lineage':
                    filtered_results[key] = apply_filter_to_lineage(data, filter_context)
                else:
                    filtered_results[key] = apply_filter_to_deep_data(key, data, filter_context)
            else:
                filtered_results[key] = data

        return jsonify({'success': True, 'data': _safe(filtered_results), 'filtered': True})
    except Exception as e:
        return jsonify({'error': str(e)})


# ==================== Single Workbook Analysis (TWB / TWBX) ====================

SW_KIND = 'single_workbook'
_sw_uploads = {}          # upload_id -> temp dir with uploaded workbook files


def _release_result(result):
    """Free resources (spool files) held by a stored job result."""
    try:
        if isinstance(result, dict) and result.get('_bulk') is not None:
            result['_bulk'].cleanup()
    except Exception as e:
        logger.warning(f"Could not release result resources: {e}")


def _sw_safe_name(name):
    base = os.path.basename((name or '').replace('\\', '/')) or 'workbook'
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', base)


def _sw_native_browse(kind):
    """Open the OS file/folder chooser on this machine (the app runs locally). Returns (path, error)."""
    import subprocess
    try:
        if sys.platform == 'win32':
            if kind == 'folder':
                dlg = ("$d = New-Object System.Windows.Forms.FolderBrowserDialog; "
                       "$d.Description = 'Select a folder containing Tableau files (.twb .twbx .tds .tdsx .tfl .tflx)'; "
                       "$r = $d.ShowDialog($o); if ($r -eq 'OK') { [Console]::Out.Write($d.SelectedPath) }")
            else:
                dlg = ("$d = New-Object System.Windows.Forms.OpenFileDialog; "
                       "$d.Filter = 'Tableau files (workbooks, datasources, flows)|*.twb;*.twbx;*.tds;*.tdsx;*.tfl;*.tflx|All files (*.*)|*.*'; "
                       "$r = $d.ShowDialog($o); if ($r -eq 'OK') { [Console]::Out.Write($d.FileName) }")
            script = ("Add-Type -AssemblyName System.Windows.Forms; "
                      "$o = New-Object System.Windows.Forms.Form; $o.TopMost = $true; " + dlg)
            cmd = ['powershell', '-NoProfile', '-STA', '-Command', script]
            flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=900, creationflags=flags)
        elif sys.platform == 'darwin':
            what = 'choose folder' if kind == 'folder' else 'choose file of type {"twb","twbx","tds","tdsx","tfl","tflx"}'
            res = subprocess.run(['osascript', '-e', f'POSIX path of ({what})'], capture_output=True, text=True,
                                 timeout=900)
        else:
            cmd = ['zenity', '--file-selection'] + (['--directory'] if kind == 'folder' else [])
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=900)
        path = (res.stdout or '').strip()
        return (path, '') if path else ('', '')
    except FileNotFoundError:
        return '', 'No native file chooser is available here - type or paste the path instead.'
    except Exception as e:
        return '', f'Could not open the file chooser: {e}'


@app.route('/api/single_workbook/browse', methods=['POST'])
def sw_browse():
    """Native Browse File / Browse Folder dialog (returns the chosen absolute path; empty if cancelled)."""
    kind = 'folder' if (request.get_json(silent=True) or {}).get('kind') == 'folder' else 'file'
    path, err = _sw_native_browse(kind)
    if err:
        return jsonify({'success': False, 'error': err})
    return jsonify({'success': True, 'path': path})


@app.route('/api/single_workbook/upload', methods=['POST'])
def sw_upload():
    """Receive dragged / browsed .twb/.twbx/.tds/.tdsx/.tfl/.tflx files (several calls may share one upload_id)."""
    upload_id = request.form.get('upload_id') or str(uuid.uuid4())
    folder = _sw_uploads.get(upload_id)
    if folder is None or not os.path.isdir(folder):
        folder = tempfile.mkdtemp(prefix='sw_upload_')
        _sw_uploads[upload_id] = folder
    saved, skipped = [], []
    for f in request.files.getlist('files'):
        name = _sw_safe_name(f.filename)
        if not name.lower().endswith(workbook_analyzer.SUPPORTED_EXTS):
            skipped.append(name)
            continue
        target = os.path.join(folder, name)
        n = 1
        while os.path.exists(target):
            stem, ext = os.path.splitext(name)
            target = os.path.join(folder, f'{stem} ({n}){ext}')
            n += 1
        f.save(target)
        saved.append(os.path.basename(target))
    return jsonify({'success': True, 'upload_id': upload_id, 'saved': saved, 'skipped': skipped})


def _sw_display_names(files):
    """
    Unique workbook names: the file stem; when stems collide the relative path (without, then with, the extension).
    """
    stems = {}
    for p in files:
        stems.setdefault(os.path.splitext(os.path.basename(p))[0].lower(), []).append(p)
    try:
        base = os.path.commonpath([os.path.dirname(p) for p in files]) if len(files) > 1 else ''
    except ValueError:
        base = ''
    names = {}
    for p in files:
        stem = os.path.splitext(os.path.basename(p))[0]
        if len(stems[stem.lower()]) == 1:
            names[p] = stem
        else:
            rel = os.path.relpath(p, base) if base else os.path.basename(p)
            names[p] = os.path.splitext(rel)[0].replace('\\', '/')
    counts = {}
    for n in names.values():
        counts[n.lower()] = counts.get(n.lower(), 0) + 1
    for p, n in list(names.items()):
        if counts[n.lower()] > 1:                          # same folder and stem, different extension
            rel = os.path.relpath(p, base) if base else os.path.basename(p)
            names[p] = rel.replace('\\', '/')
    return names


def run_single_workbook_analysis(files, source_label, workers, upload_dir, job_id=None):
    """Analyse every workbook (bounded batches, threaded); rows are spooled to disk as they are produced."""
    bulk = workbook_analyzer.BulkAnalysis()
    names = _sw_display_names(files)
    total = len(files)
    done, failed = 0, 0
    batch_size = max(8, workers * 4)
    started = time.time()
    try:
        with ThreadPoolExecutor(max_workers=workers) as pool:
            for i in range(0, total, batch_size):
                batch = files[i:i + batch_size]
                results = list(pool.map(
                    lambda p: workbook_analyzer.analyze_file(p, names[p], use_cache=not upload_dir), batch))
                for p, res in zip(batch, results):
                    if upload_dir:
                        res['workbook_summary'][0]['File Path'] = f'(uploaded) {os.path.basename(p)}'
                    if str(res['workbook_summary'][0].get('Status', '')).startswith('FAILED'):
                        failed += 1
                    bulk.add(names[p], res)
                    done += 1
                job_status[job_id] = {
                    'status': 'running', 'progress': min(97, int(95 * done / max(total, 1))),
                    'message': f'Analysed {done}/{total} workbook(s)' + (f' ({failed} failed)' if failed else '')}
        job_status[job_id] = {'status': 'running', 'progress': 98, 'message': 'Consolidating results...'}
        bulk.finish()
    except Exception:
        bulk.cleanup()
        raise
    finally:
        if upload_dir:
            shutil.rmtree(upload_dir, ignore_errors=True)
    stats = {'workbooks': done, 'failed': failed, 'seconds': round(time.time() - started, 1),
             'rows': {k: len(s) for k, s in bulk.spools.items()}}
    return {'kind': SW_KIND, 'source': source_label, 'stats': stats, '_bulk': bulk}


def _start_thread_job(job_id, func, *args):
    """Run a job in its own thread (independent of the single extraction worker queue)."""
    def target():
        job_status[job_id] = {'status': 'running', 'progress': 0, 'message': 'Starting...'}
        try:
            result = func(*args, job_id=job_id)
            job_results[job_id] = result
            _evict_old_job_results(job_id)
            job_status[job_id] = {'status': 'completed', 'progress': 100, 'message': 'Completed'}
        except Exception as e:
            logger.error(f"Job {job_id} failed: {e}", exc_info=True)
            job_status[job_id] = {'status': 'failed', 'progress': 0, 'message': str(e)}
    threading.Thread(target=target, daemon=True).start()


@app.route('/api/single_workbook/analyze', methods=['POST'])
def sw_analyze():
    """Start analysing a file, a folder (recursive scan) or previously uploaded files."""
    try:
        data = request.get_json(silent=True) or {}
        upload_id = data.get('upload_id')
        path = (data.get('path') or '').strip().strip('"')
        recursive = data.get('recursive', True) is not False
        upload_dir = None
        if upload_id:
            upload_dir = _sw_uploads.pop(upload_id, None)
            if not upload_dir or not os.path.isdir(upload_dir):
                return jsonify({'error': 'Uploaded files were not found - please upload again.'})
            files = workbook_analyzer.find_workbooks(upload_dir, recursive=False)
            label = f'{len(files)} uploaded file(s)'
        elif path:
            if not os.path.exists(path):
                return jsonify({'error': f'Path not found: {path}'})
            files = workbook_analyzer.find_workbooks(path, recursive=recursive)
            label = path
        else:
            return jsonify({'error': 'Provide a file or folder path, or upload workbook files.'})
        if not files:
            return jsonify({'error': 'No Tableau files found (.twb .twbx .tds .tdsx .tfl .tflx).'})
        workers = int(data.get('workers') or min(8, (os.cpu_count() or 4)))
        job_id = str(uuid.uuid4())
        job_status[job_id] = {'status': 'queued', 'progress': 0, 'message': f'Found {len(files)} workbook(s)'}
        _start_thread_job(job_id, run_single_workbook_analysis, files, label, max(1, min(workers, 16)), upload_dir)
        return jsonify({'success': True, 'job_id': job_id, 'files': len(files)})
    except Exception as e:
        logger.error(f"Single workbook analyze error: {e}", exc_info=True)
        return jsonify({'error': str(e)})


def _sw_result(job_id):
    result = job_results.get(job_id)
    if isinstance(result, dict) and result.get('kind') == SW_KIND and result.get('_bulk') is not None:
        return result
    return None


SW_LIST_COLUMNS = ['Workbook', 'File Path', 'File Type', 'Content Type', 'File Size (MB)', 'Tableau Release', 'Site Name',
                   'Datasources', 'Tables', 'Custom SQL', 'Joins', 'Calculated Fields', 'Worksheets', 'Dashboards',
                   'Unused Fields', 'Broken References', 'Findings', 'Status']


def _sw_public(result):
    bulk = result['_bulk']
    return {
        'kind': SW_KIND, 'source': result.get('source', ''), 'stats': result['stats'],
        'datasets': [{'key': k, 'sheet': sheet, 'count': len(bulk.spools[k])}
                     for k, (sheet, _cols) in workbook_analyzer.SCHEMAS.items()],
        'workbooks': [{c: r.get(c, '') for c in SW_LIST_COLUMNS} for r in bulk.workbooks],
    }


@app.route('/api/single_workbook/summary/<job_id>', methods=['GET'])
def sw_summary(job_id):
    result = _sw_result(job_id)
    if result is None:
        return jsonify({'error': 'Result not found (it may have been replaced by a newer analysis).'}), 404
    return jsonify({'success': True, **_sw_public(result)})


@app.route('/api/single_workbook/rows/<job_id>', methods=['GET'])
def sw_rows(job_id):
    """Incremental loading: one page of one dataset, optionally for a single workbook and/or a search text."""
    result = _sw_result(job_id)
    if result is None:
        return jsonify({'error': 'Result not found'}), 404
    bulk = result['_bulk']
    key = request.args.get('dataset', 'workbook_summary')
    if key not in bulk.spools:
        return jsonify({'error': f'Unknown dataset: {key}'}), 404
    try:
        offset = max(0, int(request.args.get('offset', 0)))
        limit = min(1000, max(1, int(request.args.get('limit', 50))))
    except ValueError:
        return jsonify({'error': 'offset and limit must be integers'}), 400
    workbook = request.args.get('workbook', '')
    start, stop = 0, None
    if workbook:
        rng = bulk.ranges[key].get(workbook)
        start, stop = rng if rng else (0, 0)
    total, rows = bulk.spools[key].search(request.args.get('q', ''), start, stop, offset, limit)
    return jsonify({'success': True, 'dataset': key, 'columns': workbook_analyzer.SCHEMAS[key][1],
                    'total': total, 'offset': offset, 'rows': rows})


def _sw_tree(bulk, wb):
    """Workbook tree: datasources > connections/tables/custom SQL/fields, worksheets, dashboards, stories, ..."""
    def rows(k):
        return bulk.workbook_rows(k, wb)

    def node(label, type_='item', children=None, detail=''):
        return {'label': label, 'type': type_, 'children': children or [], 'detail': detail}

    summary = (rows('workbook_summary') or [{}])[0]
    conns, tabs, custom = rows('connection_inventory'), rows('table_inventory'), rows('custom_sql')
    joins, rels, calcs = rows('join_analysis'), rows('relationship_analysis'), rows('calculated_fields')
    ds_nodes = []
    for d in rows('datasource_inventory'):
        if d['Datasource Type'] == 'Parameters':
            continue                                    # the internal "Parameters" datasource is listed under Parameters
        cap = d['Caption']
        kids = []
        c_nodes = [node(f"{c['Connection Type']} | {c['Server'] or c['File / Directory'] or '-'} | {c['Database'] or '-'}"
                        f" [{c['Connection Role']}]", 'connection',
                        detail=f"auth: {c['Authentication Mode'] or '-'}  port: {c['Port'] or '-'}  ssl: {c['SSL Settings'] or '-'}")
                   for c in conns if c['Datasource'] == cap]
        if c_nodes:
            kids.append(node(f'Connections ({len(c_nodes)})', 'group', c_nodes))
        t_nodes = [node(t['Full Table Name'], 'table', detail=f"{t['Table Kind']} | {t['Source']}")
                   for t in tabs if t['Datasource'] == cap]
        if t_nodes:
            kids.append(node(f'Tables ({len(t_nodes)})', 'group', t_nodes[:500]))
        seen_sql = {}
        for c in custom:
            if c['Datasource'] == cap:
                seen_sql.setdefault(c['Custom SQL Name'], []).append(c)
        if seen_sql:
            kids.append(node(f'Custom SQL ({len(seen_sql)})', 'group', [
                node(n, 'sql', [node(r['Referenced Table'] or '(no tables found)', 'table') for r in rs],
                     detail=(rs[0]['SQL Query'] or '')[:600]) for n, rs in seen_sql.items()]))
        j_nodes = [node(f"{j['Left Table']} {j['Join Type']} JOIN {j['Right Table']}", 'join', detail=j['Join Clause'])
                   for j in joins if j['Datasource'] == cap]
        j_nodes += [node(f"{r['Type']}: {r['Left Table'] or r['Name']} - {r['Right Table']}", 'join',
                         detail=r['Clause'] or r['Details']) for r in rels if r['Datasource'] == cap]
        if j_nodes:
            kids.append(node(f'Joins / Relationships ({len(j_nodes)})', 'group', j_nodes))
        c2 = [node(c['Calculation Name'], 'calc', detail=c['Formula']) for c in calcs if c['Datasource'] == cap]
        if c2:
            kids.append(node(f'Calculated fields ({len(c2)})', 'group', c2[:500]))
        ds_nodes.append(node(f"{cap}  [{d['Datasource Type']}]", 'datasource', kids,
                             f"{d['Fields']} fields, {d['Unused Fields']} unused"))
    ws_nodes = [node(w['Worksheet'] + (' (hidden)' if w['Hidden'] else ''), 'worksheet', [n for n in [
        node('Dashboards: ' + w['Dashboards'], 'link') if w['Dashboards'] else None,
        node('Dimensions: ' + w['Dimensions Used'], 'field') if w['Dimensions Used'] else None,
        node('Measures: ' + w['Measures Used'], 'field') if w['Measures Used'] else None,
        node('Tables: ' + w['Tables Used'], 'table') if w['Tables Used'] else None] if n],
        detail=w['Mark Type']) for w in rows('worksheets')]
    d_nodes = [node(d['Dashboard'] + (' (hidden)' if d['Hidden'] else ''), 'dashboard', [n for n in [
        node('Worksheets: ' + d['Worksheets Included'], 'link') if d['Worksheets Included'] else None,
        node('Filters: ' + d['Filters Included'], 'field') if d['Filters Included'] else None,
        node('Parameters: ' + d['Parameters Included'], 'field') if d['Parameters Included'] else None] if n],
        detail=f"{d['Objects Count']} objects") for d in rows('dashboards')]
    stories = {}
    for s in rows('stories'):
        stories.setdefault(s['Story'], []).append(node(f"{s['Story Point Caption'] or s['Story Point']}: {s['Referenced Object']}", 'link'))
    findings = {}
    for f in rows('data_quality'):
        findings.setdefault(f['Finding Type'], []).append(
            node(f"{f['Object Name']}", 'finding', detail=f['Details']))
    steps = [node(f"{s['Step Name']}  [{s['Step Type']}]", 'step', detail=(f"{s['Details']}  in: {s['Inputs'] or '-'}  out: {s['Outputs'] or '-'}"))
             for s in rows('flow_steps')]
    children = [
        node(f'Flow steps ({len(steps)})', 'group', steps[:500]),
        node(f'Datasources ({len(ds_nodes)})', 'group', ds_nodes),
        node(f'Worksheets ({len(ws_nodes)})', 'group', ws_nodes),
        node(f'Dashboards ({len(d_nodes)})', 'group', d_nodes),
        node(f'Stories ({len(stories)})', 'group', [node(n, 'story', v) for n, v in stories.items()]),
        node(f"Parameters ({len(rows('parameters'))})", 'group', [
            node(p['Parameter Name'], 'param', detail=f"{p['Data Type']} = {p['Current Value']} ({p['Allowed Values']})")
            for p in rows('parameters')]),
        node(f"Findings ({sum(len(v) for v in findings.values())})", 'group', [
            node(f'{k} ({len(v)})', 'group', v[:200]) for k, v in sorted(findings.items())]),
    ]
    return node(wb, 'workbook', [c for c in children if c['children']],
                f"{summary.get('File Type', '')} | Tableau {summary.get('Tableau Release', '')} | {summary.get('Status', '')}")


@app.route('/api/single_workbook/tree/<job_id>', methods=['GET'])
def sw_tree(job_id):
    result = _sw_result(job_id)
    if result is None:
        return jsonify({'error': 'Result not found'}), 404
    wb = request.args.get('workbook', '')
    if wb not in result['_bulk'].ranges['workbook_summary']:
        return jsonify({'error': f'Unknown workbook: {wb}'}), 404
    return jsonify({'success': True, 'tree': _sw_tree(result['_bulk'], wb)})


@app.route('/api/single_workbook/export/<job_id>', methods=['GET'])
def sw_export(job_id):
    """Export everything (or one workbook): ?format=xlsx|csv|json&workbook=<name>. Large results are bucketed."""
    result = _sw_result(job_id)
    if result is None:
        return jsonify({'error': 'Result not found (it may have been replaced by a newer analysis).'}), 404
    if request.args.get('check'):
        return jsonify({'success': True})
    bulk = result['_bulk']
    workbook = request.args.get('workbook', '')
    fmt = request.args.get('format', 'xlsx').lower()
    try:
        datasets = []
        for key, (sheet, _cols) in workbook_analyzer.SCHEMAS.items():
            spool = bulk.spools[key]
            if workbook:
                rng = bulk.ranges[key].get(workbook, (0, 0))
                datasets.append((sheet, spool[rng[0]:rng[1]]))
            else:
                datasets.append((sheet, spool))
        prefix = 'single_workbook_' + _sw_safe_name(workbook) if workbook else 'single_workbook_analysis'
        if fmt == 'json':
            path, name, mimetype = write_json_export(datasets, prefix=prefix)
        else:
            path, name, mimetype = build_export(datasets, prefix=prefix, fmt='csv' if fmt == 'csv' else 'xlsx')
        return _send_temp_file(path, name, mimetype)
    except Exception as e:
        logger.error(f"Single workbook export error: {e}", exc_info=True)
        return jsonify({'error': str(e)}), 500

# ==================== Usage Analytics Engine ====================

class UsageAnalyticsEngine:
    """
    Enterprise-level Usage & Governance Analytics Engine combining REST + Metadata API
    Provides: Usage Analytics, Criticality Scoring, Sunset Scoring, Governance Metrics
    Supports governance decisions and rationalization planning.
    """

    # Criticality Score weights (0-100)
    CRITICALITY_WEIGHTS = {
        'usage': 0.30,        # Views and user engagement
        'unique_users': 0.20, # User breadth
        'lineage_impact': 0.20,  # Data dependencies
        'operational': 0.15,  # Schedules, subscriptions
        'complexity': 0.15    # Dashboard/field complexity
    }

    # Sunset Score weights (0-100) - higher = more likely to sunset
    SUNSET_WEIGHTS = {
        'no_views_90d': 30,       # No views in 90 days
        'no_subscribers': 15,     # No active subscriptions
        'no_refresh': 15,         # No refresh schedule
        'owner_inactive': 20,     # Owner is inactive
        'low_lineage': 10,        # Low downstream impact
        'no_description': 5,      # No description (poor governance)
        'low_usage_trend': 5      # Declining usage trend
    }

    def __init__(self, config_data):
        self.config_data = config_data
        self.server = None
        self.auth_token = None
        self.metadata_api = None
        self.repository = RepositoryConnector()  # Repository connector for actual usage data
        self.admin_insights = None  # Admin Insights connector for Tableau Cloud
        self.workbook_analytics = []
        self.view_analytics = []
        self.lineage_data = []
        self.sunset_candidates = []
        self.kpi_summary = {}
        self.user_map = {}
        self.project_map = {}
        # workbook_id -> [datasource luid, ...] it uses - built in _enrich_workbooks_from_metadata_api,
        # consumed by _compute_datasource_governance to know which workbooks use each published datasource
        # (REST content ids and Metadata API luids are the same GUID for the same object, so this maps
        # directly onto the ds.id keys used elsewhere in this engine)
        self.wb_datasource_map = {}
        self.usage_data_source = 'rest'  # 'rest', 'repository', or 'admin_insights'

    def authenticate(self):
        """Authenticate with Tableau Server"""
        server = TSC.Server(self.config_data['server_url'], use_server_version=True)
        auth = TSC.PersonalAccessTokenAuth(
            self.config_data['pat_name'],
            self.config_data['pat_token'],
            self.config_data['site_name']
        )
        self.server = server
        return server, auth

    def run_analysis(self, job_id=None, repository_config=None, admin_insights_data=None, use_admin_insights=False):
        """Run complete usage and criticality analysis"""
        results = {
            'workbook_analytics': [],
            'view_analytics': [],
            'datasource_analytics': [],  # NEW: Datasource governance
            'lineage_impact': [],
            'retirement_candidates': [],
            'kpi_summary': {},
            'data_sources': {'rest': [], 'metadata_api': [], 'repository': [], 'admin_insights': [], 'calculated': []},
            'usage_data_source': 'REST API (Estimated)',  # Will be updated if better source available
            'errors': []
        }

        try:
            server, auth = self.authenticate()

            with server.auth.sign_in(auth):
                self.auth_token = server.auth_token
                site_id = server.site_id

                # Initialize Metadata API
                self.metadata_api = MetadataAPIClient(
                    self.config_data['server_url'],
                    self.auth_token,
                    site_id
                )

                if job_id:
                    job_status[job_id] = {'status': 'running', 'progress': 10, 'message': 'Fetching workbooks and views...'}

                # 1. REST API: Get workbooks with usage stats
                logger.info("Analytics: Step 1 - Fetching workbooks from REST API...")
                workbook_data = self._get_workbook_usage_from_rest()
                logger.info(f"Analytics: Step 1 Complete - Got {len(workbook_data)} workbooks")
                results['data_sources']['rest'].append('workbooks')
                results['data_sources']['rest'].append('views')

                if job_id:
                    job_status[job_id] = {'status': 'running', 'progress': 20, 'message': 'Enriching with Metadata API...'}

                # 1b. Metadata API: Get accurate sheet/dashboard counts
                if self.metadata_api.available:
                    logger.info("Analytics: Step 1b - Enriching workbook data from Metadata API...")
                    workbook_data = self._enrich_workbooks_from_metadata_api(workbook_data)
                    logger.info(f"Analytics: Step 1b Complete - Enriched {len(workbook_data)} workbooks")
                    results['data_sources']['metadata_api'].append('workbooks')

                if job_id:
                    job_status[job_id] = {'status': 'running', 'progress': 30, 'message': 'Fetching view statistics...'}

                # 2. REST API: Get views with usage stats
                logger.info("Analytics: Step 2 - Fetching views...")
                view_data = self._get_view_usage_from_rest()
                logger.info(f"Analytics: Step 2 Complete - Got {len(view_data)} views")

                if job_id:
                    job_status[job_id] = {'status': 'running', 'progress': 40, 'message': 'Extracting lineage from Metadata API...'}

                # 3. Metadata API: Get lineage data
                logger.info("Analytics: Step 3 - Fetching lineage...")
                lineage_data = {}
                if self.metadata_api.available:
                    lineage_data = self._get_lineage_from_metadata_api()
                    logger.info(f"Analytics: Step 3 Complete - Got lineage data with {len(lineage_data.get('tables', []))} tables")
                    results['data_sources']['metadata_api'].append('lineage')
                    results['data_sources']['metadata_api'].append('calculated_fields')
                else:
                    logger.warning("Analytics: Metadata API not available")
                    results['errors'].append('Metadata API unavailable - lineage data limited')

                if job_id:
                    job_status[job_id] = {'status': 'running', 'progress': 55, 'message': 'Fetching operational data...'}

                # 4. REST API: Get operational data (schedules, jobs, tasks)
                logger.info("Analytics: Step 4 - Fetching operational data...")
                operational_data = self._get_operational_data_from_rest()
                logger.info(f"Analytics: Step 4 Complete - Got operational data")
                results['data_sources']['rest'].append('schedules')
                results['data_sources']['rest'].append('jobs')
                results['data_sources']['rest'].append('tasks')

                if job_id:
                    job_status[job_id] = {'status': 'running', 'progress': 65, 'message': 'Fetching actual usage data...'}

                # 4b. Repository/Admin Insights: Get actual usage data
                repository_usage = {}
                if repository_config:
                    # Try to connect to Tableau Server repository (PostgreSQL)
                    logger.info("Analytics: Step 4b - Connecting to Repository...")
                    if self.repository.connect(
                        host=repository_config.get('host', ''),
                        port=repository_config.get('port', 8060),
                        database=repository_config.get('database', 'workgroup'),
                        user=repository_config.get('user', 'readonly'),
                        password=repository_config.get('password', '')
                    ):
                        repository_usage = self.repository.get_workbook_usage_metrics()
                        if repository_usage:
                            self.usage_data_source = 'repository'
                            results['usage_data_source'] = 'Repository'
                            results['data_sources']['repository'].append('historical_events')
                            logger.info(f"Analytics: Step 4b Complete - Got usage data for {len(repository_usage)} workbooks from Repository")
                        self.repository.disconnect()
                elif use_admin_insights:
                    # Use Admin Insights via TSC (for Tableau Cloud)
                    logger.info("Analytics: Step 4b - Fetching Admin Insights via TSC...")
                    try:
                        self.admin_insights = AdminInsightsConnector(server, site_id)
                        if self.admin_insights.discover_datasources():
                            repository_usage = self.admin_insights.get_all_usage_metrics()
                            if repository_usage:
                                self.usage_data_source = 'admin_insights'
                                results['usage_data_source'] = 'Admin Insights'
                                results['data_sources']['admin_insights'].append('TS Events')
                                if 'tasks' in self.admin_insights.datasources:
                                    results['data_sources']['admin_insights'].append('TS Background Tasks')
                                if 'subscriptions' in self.admin_insights.datasources:
                                    results['data_sources']['admin_insights'].append('TS Subscriptions')
                                logger.info(f"Analytics: Step 4b Complete - Got usage data for {len(repository_usage)} workbooks from Admin Insights")
                            else:
                                logger.warning("Analytics: Admin Insights found but no usage data extracted")
                                results['errors'].append('Admin Insights found but no usage data could be extracted')
                        else:
                            logger.warning("Analytics: Admin Insights datasources not found on this site")
                            results['errors'].append('Admin Insights not enabled on this site. Go to Tableau Cloud Settings > Extensions > Admin Insights to enable it.')
                    except Exception as ai_error:
                        logger.error(f"Analytics: Admin Insights error - {str(ai_error)}")
                        results['errors'].append(f'Admin Insights error: {str(ai_error)}')
                elif admin_insights_data:
                    # Load from Admin Insights CSV upload (for Tableau Cloud)
                    logger.info("Analytics: Step 4b - Loading Admin Insights CSV data...")
                    if self.repository.load_admin_insights_csv(admin_insights_data):
                        repository_usage = self.repository.usage_data
                        self.usage_data_source = 'admin_insights'
                        results['usage_data_source'] = 'Admin Insights (CSV)'
                        results['data_sources']['admin_insights'].append('admin_insights_csv')
                        logger.info(f"Analytics: Step 4b Complete - Got usage data for {len(repository_usage)} workbooks from Admin Insights CSV")

                if job_id:
                    job_status[job_id] = {'status': 'running', 'progress': 70, 'message': 'Computing criticality scores...'}

                # 5. Compute analytics for each workbook
                logger.info(f"Analytics: Step 5 - Computing analytics for {len(workbook_data)} workbooks...")
                results['workbook_analytics'] = self._compute_workbook_analytics(
                    workbook_data, view_data, lineage_data, operational_data, repository_usage
                )
                logger.info(f"Analytics: Step 5 Complete - Generated {len(results['workbook_analytics'])} analytics records")
                results['data_sources']['calculated'].append('criticality_scores')
                results['data_sources']['calculated'].append('usage_ranks')

                if job_id:
                    job_status[job_id] = {'status': 'running', 'progress': 75, 'message': 'Computing view analytics...'}

                # 6. Compute view-level analytics
                results['view_analytics'] = self._compute_view_analytics(view_data, workbook_data)

                if job_id:
                    job_status[job_id] = {'status': 'running', 'progress': 80, 'message': 'Computing datasource governance...'}

                # 6b. Compute datasource governance analytics
                logger.info("Analytics: Step 6b - Computing datasource governance...")
                results['datasource_analytics'] = self._compute_datasource_governance(
                    results['workbook_analytics'], lineage_data, operational_data
                )
                logger.info(f"Analytics: Step 6b Complete - Generated {len(results['datasource_analytics'])} datasource records")
                results['data_sources']['calculated'].append('datasource_governance')

                if job_id:
                    job_status[job_id] = {'status': 'running', 'progress': 85, 'message': 'Identifying retirement candidates...'}

                # 7. Identify retirement candidates
                results['retirement_candidates'] = self._identify_retirement_candidates(
                    results['workbook_analytics'], operational_data
                )

                if job_id:
                    job_status[job_id] = {'status': 'running', 'progress': 90, 'message': 'Preparing lineage data...'}

                # 8. Prepare lineage impact data
                results['lineage_impact'] = self._prepare_lineage_impact(lineage_data)

                if job_id:
                    job_status[job_id] = {'status': 'running', 'progress': 95, 'message': 'Computing KPI summary...'}

                # 9. Compute KPI summary
                results['kpi_summary'] = self._compute_kpi_summary(results)

        except Exception as e:
            logger.error(f"Usage analytics error: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            results['errors'].append(str(e))

        return results

    def _get_workbook_usage_from_rest(self):
        """Get all workbooks with comprehensive usage and governance statistics from REST API"""
        workbooks_data = []
        try:
            logger.info("Analytics: Starting workbook fetch from REST API...")

            # Get all workbooks - Use Pager to get ALL (not just first 100)
            all_workbooks = list(TSC.Pager(self.server.workbooks))
            logger.info(f"Analytics: Found {len(all_workbooks)} workbooks from REST API")

            # Build project lookup - Use Pager to get ALL
            projects = list(TSC.Pager(self.server.projects))
            self.project_map = {p.id: p.name for p in projects}
            logger.info(f"Analytics: Built project lookup with {len(self.project_map)} projects")

            # Build user lookup with activity status - Use Pager to get ALL
            users = list(TSC.Pager(self.server.users))
            self.user_map = {}
            for u in users:
                last_login = getattr(u, 'last_login', None)
                is_active = True
                if last_login:
                    try:
                        login_dt = last_login if isinstance(last_login, datetime) else datetime.fromisoformat(str(last_login).replace('Z', '+00:00'))
                        days_since = (datetime.now(login_dt.tzinfo) if login_dt.tzinfo else datetime.now() - login_dt).days
                        is_active = days_since <= 90
                    except:
                        pass
                self.user_map[u.id] = {
                    'name': u.name,
                    'email': getattr(u, 'email', ''),
                    'site_role': getattr(u, 'site_role', ''),
                    'last_login': str(last_login) if last_login else '',
                    'is_active': is_active
                }
            logger.info(f"Analytics: Built user lookup with {len(self.user_map)} users")

            for idx, wb in enumerate(all_workbooks):
                try:
                    # Calculate total views across all views in workbook
                    total_views = 0
                    view_count = 0
                    top_view_name = ''
                    top_view_count = 0
                    view_details = []

                    # Try to populate views
                    try:
                        self.server.workbooks.populate_views(wb, usage=True)
                        if wb.views:
                            for view in wb.views:
                                view_count += 1
                                view_total = getattr(view, 'total_views', 0) or 0
                                total_views += view_total
                                view_details.append({
                                    'name': view.name,
                                    'total_views': view_total
                                })
                                if view_total > top_view_count:
                                    top_view_count = view_total
                                    top_view_name = view.name
                    except Exception as ve:
                        logger.warning(f"Analytics: Could not populate views for {wb.name}: {str(ve)}")

                    # Get project name from lookup
                    project_name = self.project_map.get(wb.project_id, '')

                    # Get owner info from lookup
                    owner_id = getattr(wb, 'owner_id', '')
                    owner_info = self.user_map.get(owner_id, {})
                    owner_name = owner_info.get('name', '')
                    owner_active = owner_info.get('is_active', True)

                    # Get description for governance check
                    description = getattr(wb, 'description', '') or ''
                    has_description = len(description.strip()) > 0

                    # Get tags
                    tags = [t.name for t in getattr(wb, 'tags', [])] if hasattr(wb, 'tags') and wb.tags else []
                    tag_count = len(tags)

                    # Calculate top view percentage
                    top_view_percentage = round((top_view_count / total_views * 100), 1) if total_views > 0 else 0

                    workbooks_data.append({
                        'id': wb.id,
                        'name': wb.name,
                        'project_id': wb.project_id,
                        'project_name': project_name,
                        'owner_id': owner_id,
                        'owner_name': owner_name,
                        'owner_active': owner_active,
                        'created_at': str(wb.created_at) if wb.created_at else '',
                        'updated_at': str(wb.updated_at) if wb.updated_at else '',
                        'description': description,
                        'has_description': has_description,
                        'total_views_all_time': total_views,
                        'view_count': view_count,
                        'view_details': view_details,
                        'top_view_name': top_view_name,
                        'top_view_views': top_view_count,
                        'top_view_percentage': top_view_percentage,
                        'size': getattr(wb, 'size', 0) or 0,
                        'tags': tags,
                        'tag_count': tag_count,
                        'content_url': getattr(wb, 'content_url', ''),
                        'webpage_url': getattr(wb, 'webpage_url', ''),
                        'data_source': 'REST API'
                    })
                except Exception as e:
                    logger.warning(f"Analytics: Error processing workbook {wb.name}: {str(e)}")
                    continue

            logger.info(f"Analytics: Successfully processed {len(workbooks_data)} workbooks")

        except Exception as e:
            import traceback
            logger.error(f"Analytics: Error fetching workbooks: {str(e)}")
            logger.error(f"Analytics: Traceback: {traceback.format_exc()}")

        return workbooks_data

    def _get_view_usage_from_rest(self):
        """Get all views with usage statistics from REST API"""
        views_data = []
        try:
            logger.info("Analytics: Starting view fetch from REST API...")
            # Use Pager to get ALL views (not just first 100)
            all_views = list(TSC.Pager(self.server.views))
            logger.info(f"Analytics: Found {len(all_views)} views from REST API")

            for view in all_views:
                try:
                    views_data.append({
                        'id': view.id,
                        'name': view.name,
                        'workbook_id': getattr(view, 'workbook_id', ''),
                        'owner_id': getattr(view, 'owner_id', ''),
                        'total_views': getattr(view, 'total_views', 0) or 0,
                        'created_at': str(view.created_at) if hasattr(view, 'created_at') and view.created_at else '',
                        'updated_at': str(view.updated_at) if hasattr(view, 'updated_at') and view.updated_at else '',
                        'tags': [t.name for t in view.tags] if hasattr(view, 'tags') and view.tags else [],
                        'data_source': 'REST API'
                    })
                except Exception as ve:
                    logger.warning(f"Error processing view {view.name}: {str(ve)}")

            logger.info(f"Analytics: Successfully processed {len(views_data)} views")

        except Exception as e:
            import traceback
            logger.error(f"Analytics: Error fetching views: {str(e)}")
            logger.error(f"Analytics: Traceback: {traceback.format_exc()}")

        return views_data

    def _enrich_workbooks_from_metadata_api(self, workbook_data):
        """Enrich workbook data with accurate sheet/dashboard counts from Metadata API"""
        if not self.metadata_api or not self.metadata_api.available:
            return workbook_data

        try:
            # Get workbook stats from Metadata API
            wb_stats_result = self.metadata_api.get_workbooks_with_stats()

            if wb_stats_result and 'workbooks' in wb_stats_result:
                # Build lookup by luid
                metadata_lookup = {}
                self.wb_datasource_map = {}
                for wb in wb_stats_result['workbooks']:
                    luid = wb.get('luid', '')
                    if luid:
                        sheets_count = len(wb.get('sheets', []))
                        dashboards_count = len(wb.get('dashboards', []))
                        metadata_lookup[luid] = {
                            'sheets_count': sheets_count,
                            'dashboards_count': dashboards_count,
                            'total_views_count': sheets_count + dashboards_count,
                            'datasource_count': len(wb.get('upstreamDatasources', [])),
                            'metadata_owner': wb.get('owner', {}).get('name', '') if wb.get('owner') else '',
                            'metadata_owner_email': wb.get('owner', {}).get('email', '') if wb.get('owner') else ''
                        }
                        # datasource luid -> which workbooks use it (feeds _compute_datasource_governance)
                        self.wb_datasource_map[luid] = [
                            ds.get('luid', '') for ds in (wb.get('upstreamDatasources') or []) if ds.get('luid')
                        ]

                logger.info(f"Analytics: Built metadata lookup for {len(metadata_lookup)} workbooks, "
                           f"{sum(len(v) for v in self.wb_datasource_map.values())} workbook-datasource edges")

                # Enrich workbook data
                for wb in workbook_data:
                    wb_id = wb.get('id', '')
                    if wb_id in metadata_lookup:
                        meta = metadata_lookup[wb_id]
                        # Update dashboard count with accurate data from Metadata API
                        wb['dashboard_count'] = meta['total_views_count']
                        wb['sheets_count'] = meta['sheets_count']
                        wb['dashboards_only_count'] = meta['dashboards_count']
                        wb['datasource_count'] = meta['datasource_count']
                        # Use metadata owner if REST API owner is empty
                        if not wb.get('owner_name') and meta['metadata_owner']:
                            wb['owner_name'] = meta['metadata_owner']
                        wb['data_source'] = 'REST API + Metadata API'

                logger.info(f"Analytics: Enriched workbook data with Metadata API stats")

        except Exception as e:
            logger.warning(f"Analytics: Error enriching from Metadata API: {str(e)}")
            import traceback
            logger.warning(f"Analytics: Traceback: {traceback.format_exc()}")

        return workbook_data

    def _get_lineage_from_metadata_api(self):
        """Get comprehensive lineage data from Metadata API"""
        lineage_data = {
            'workbook_upstream': {},
            'datasource_upstream': {},
            'calculated_fields': [],
            'custom_sql': [],
            'tables': []
        }

        if not self.metadata_api or not self.metadata_api.available:
            return lineage_data

        try:
            # Get all tables with lineage
            tables_result = self.metadata_api.get_all_tables()
            if tables_result and 'databaseTables' in tables_result:
                for table in tables_result['databaseTables']:
                    table_info = {
                        'id': table.get('id', ''),
                        'name': table.get('name', ''),
                        'schema': table.get('schema', ''),
                        'full_name': table.get('fullName', ''),
                        'database_name': table.get('database', {}).get('name', '') if table.get('database') else '',
                        'connection_type': table.get('database', {}).get('connectionType', '') if table.get('database') else '',
                        'downstream_workbooks': [w.get('name', '') for w in table.get('downstreamWorkbooks', [])],
                        'downstream_datasources': [d.get('name', '') for d in table.get('downstreamDatasources', [])],
                        'column_count': len(table.get('columns', [])),
                        'data_source': 'Metadata API'
                    }
                    lineage_data['tables'].append(table_info)

                    # Build upstream mapping for workbooks
                    for wb in table.get('downstreamWorkbooks', []):
                        wb_luid = wb.get('luid', '')
                        if wb_luid:
                            if wb_luid not in lineage_data['workbook_upstream']:
                                lineage_data['workbook_upstream'][wb_luid] = {'tables': [], 'databases': set()}
                            lineage_data['workbook_upstream'][wb_luid]['tables'].append(table_info['name'])
                            if table_info['database_name']:
                                lineage_data['workbook_upstream'][wb_luid]['databases'].add(table_info['database_name'])

                    # Build upstream mapping for published datasources (feeds datasource governance's
                    # "upstream tables" / lineage-impact count - previously always empty, see luid below)
                    for ds in table.get('downstreamDatasources', []):
                        ds_luid = ds.get('luid', '')
                        if ds_luid:
                            if ds_luid not in lineage_data['datasource_upstream']:
                                lineage_data['datasource_upstream'][ds_luid] = {'tables': [], 'databases': set()}
                            lineage_data['datasource_upstream'][ds_luid]['tables'].append(table_info['name'])
                            if table_info['database_name']:
                                lineage_data['datasource_upstream'][ds_luid]['databases'].add(table_info['database_name'])

            # Get calculated fields
            calc_result = self.metadata_api.get_calculated_fields()
            if calc_result and 'calculatedFields' in calc_result:
                lineage_data['calculated_fields'] = [{
                    'id': cf.get('id', ''),
                    'name': cf.get('name', ''),
                    'formula': cf.get('formula', ''),
                    'datasource_name': cf.get('datasource', {}).get('name', '') if cf.get('datasource') else '',
                    'data_source': 'Metadata API'
                } for cf in calc_result['calculatedFields']]

            # Get custom SQL
            sql_result = self.metadata_api.get_custom_sql_tables()
            if sql_result and 'customSQLTables' in sql_result:
                lineage_data['custom_sql'] = [{
                    'id': sql.get('id', ''),
                    'name': sql.get('name', ''),
                    'query': sql.get('query', ''),
                    'database_name': sql.get('database', {}).get('name', '') if sql.get('database') else '',
                    'connection_type': sql.get('database', {}).get('connectionType', '') if sql.get('database') else '',
                    'data_source': 'Metadata API'
                } for sql in sql_result['customSQLTables']]

        except Exception as e:
            logger.error(f"Error getting lineage from Metadata API: {str(e)}")

        return lineage_data

    def _get_operational_data_from_rest(self):
        """Get operational data: schedules, tasks, jobs, subscriptions from REST API"""
        operational_data = {
            'schedules': [],
            'tasks': {},
            'jobs': [],
            'job_failures': {},  # Track failures by workbook
            'subscriptions': {},  # Track subscriptions by workbook
            'users': {},
            'datasource_tasks': {},  # Track tasks by datasource
            'datasource_job_failures': {}  # Track failures by datasource
        }

        try:
            # Get schedules - Use Pager to get ALL
            schedules = list(TSC.Pager(self.server.schedules))
            operational_data['schedules'] = [{
                'id': s.id,
                'name': s.name,
                'schedule_type': s.schedule_type,
                'state': s.state,
                'frequency': getattr(s, 'frequency', '')
            } for s in schedules]

            # Use already built user map if available
            if self.user_map:
                operational_data['users'] = self.user_map
            else:
                # Use Pager to get ALL users
                users = list(TSC.Pager(self.server.users))
                for u in users:
                    last_login = getattr(u, 'last_login', None)
                    is_active = True
                    if last_login:
                        try:
                            days_since = (datetime.now() - last_login).days if not hasattr(last_login, 'tzinfo') else 0
                            is_active = days_since <= 90
                        except:
                            pass
                    operational_data['users'][u.id] = {
                        'name': u.name,
                        'email': getattr(u, 'email', ''),
                        'site_role': getattr(u, 'site_role', ''),
                        'last_login': str(last_login) if last_login else '',
                        'is_active': is_active
                    }

            # Get tasks (extract refresh tasks) for workbooks
            try:
                # Use Pager to get ALL workbooks
                all_workbooks = list(TSC.Pager(self.server.workbooks))
                for wb in all_workbooks:
                    try:
                        self.server.workbooks.populate_extract_refresh_tasks(wb)
                        if hasattr(wb, 'extract_refresh_tasks') and wb.extract_refresh_tasks:
                            operational_data['tasks'][wb.id] = {
                                'type': 'workbook',
                                'name': wb.name,
                                'task_count': len(wb.extract_refresh_tasks),
                                'has_refresh_schedule': True,
                                'tasks': [{
                                    'id': t.id,
                                    'priority': getattr(t, 'priority', 0),
                                    'schedule_id': getattr(t, 'schedule_id', '')
                                } for t in wb.extract_refresh_tasks]
                            }
                        else:
                            operational_data['tasks'][wb.id] = {
                                'type': 'workbook',
                                'name': wb.name,
                                'task_count': 0,
                                'has_refresh_schedule': False,
                                'tasks': []
                            }
                    except:
                        operational_data['tasks'][wb.id] = {
                            'type': 'workbook',
                            'name': wb.name,
                            'task_count': 0,
                            'has_refresh_schedule': False,
                            'tasks': []
                        }
            except Exception as e:
                logger.warning(f"Error getting tasks: {str(e)}")

            # Get subscriptions for workbooks (if API supports it)
            try:
                for wb in all_workbooks:
                    try:
                        # Try to get subscriptions for this workbook's views
                        self.server.workbooks.populate_views(wb, usage=True)
                        subscriber_count = 0
                        if hasattr(wb, 'views') and wb.views:
                            for view in wb.views:
                                # Subscriptions are typically at the view level
                                # Note: TSC may not expose subscription counts directly
                                pass
                        operational_data['subscriptions'][wb.id] = {
                            'subscriber_count': subscriber_count,
                            'has_subscriptions': subscriber_count > 0
                        }
                    except:
                        operational_data['subscriptions'][wb.id] = {
                            'subscriber_count': 0,
                            'has_subscriptions': False
                        }
            except Exception as e:
                logger.warning(f"Error getting subscriptions: {str(e)}")

            # Get recent jobs for failure analysis (last 30 days)
            try:
                job_opts = TSC.RequestOptions()
                jobs, _ = self.server.jobs.get(req_options=job_opts)
                cutoff_30d = datetime.now() - timedelta(days=30)

                for j in jobs[:200]:  # Analyze recent 200 jobs
                    job_created = None
                    try:
                        if hasattr(j, 'created_at') and j.created_at:
                            job_created = j.created_at if isinstance(j.created_at, datetime) else datetime.fromisoformat(str(j.created_at).replace('Z', '+00:00'))
                    except:
                        pass

                    job_info = {
                        'id': j.id,
                        'job_type': getattr(j, 'job_type', ''),
                        'status': j.status if hasattr(j, 'status') else '',
                        'created_at': str(j.created_at) if hasattr(j, 'created_at') and j.created_at else '',
                        'completed_at': str(j.completed_at) if hasattr(j, 'completed_at') and j.completed_at else ''
                    }
                    operational_data['jobs'].append(job_info)

                    # Track failures by workbook (if we can identify the workbook)
                    if hasattr(j, 'status') and j.status in ['Failed', 'Cancelled']:
                        wb_id = getattr(j, 'workbook_id', None) or getattr(j, 'content_id', None)
                        if wb_id and job_created and job_created > cutoff_30d:
                            if wb_id not in operational_data['job_failures']:
                                operational_data['job_failures'][wb_id] = 0
                            operational_data['job_failures'][wb_id] += 1

                        # Also track by datasource if applicable
                        ds_id = getattr(j, 'datasource_id', None)
                        if ds_id and job_created and job_created > cutoff_30d:
                            if ds_id not in operational_data['datasource_job_failures']:
                                operational_data['datasource_job_failures'][ds_id] = 0
                            operational_data['datasource_job_failures'][ds_id] += 1

            except Exception as e:
                logger.warning(f"Error getting jobs: {str(e)}")

            # Get datasource tasks (extract refresh tasks)
            try:
                # Use Pager to get ALL datasources
                all_datasources = list(TSC.Pager(self.server.datasources))
                for ds in all_datasources:
                    try:
                        self.server.datasources.populate_extract_refresh_tasks(ds)
                        if hasattr(ds, 'extract_refresh_tasks') and ds.extract_refresh_tasks:
                            # Get last refresh time
                            last_refresh = ''
                            refresh_frequency = ''
                            for t in ds.extract_refresh_tasks:
                                if hasattr(t, 'schedule_id') and t.schedule_id:
                                    # Try to get schedule info
                                    for sched in operational_data['schedules']:
                                        if sched['id'] == t.schedule_id:
                                            refresh_frequency = sched.get('frequency', '')
                                            break

                            operational_data['datasource_tasks'][ds.id] = {
                                'type': 'datasource',
                                'name': ds.name,
                                'task_count': len(ds.extract_refresh_tasks),
                                'has_refresh_schedule': True,
                                'last_refresh': last_refresh,
                                'refresh_frequency': refresh_frequency
                            }
                        else:
                            operational_data['datasource_tasks'][ds.id] = {
                                'type': 'datasource',
                                'name': ds.name,
                                'task_count': 0,
                                'has_refresh_schedule': False,
                                'last_refresh': '',
                                'refresh_frequency': ''
                            }
                    except Exception as ds_err:
                        operational_data['datasource_tasks'][ds.id] = {
                            'type': 'datasource',
                            'name': ds.name,
                            'task_count': 0,
                            'has_refresh_schedule': False,
                            'last_refresh': '',
                            'refresh_frequency': ''
                        }
            except Exception as e:
                logger.warning(f"Error getting datasource tasks: {str(e)}")

        except Exception as e:
            logger.error(f"Error getting operational data: {str(e)}")

        return operational_data

    def _compute_workbook_analytics(self, workbook_data, view_data, lineage_data, operational_data, repository_usage=None):
        """Compute comprehensive governance analytics, criticality scores, and sunset scores for each workbook"""
        analytics = []
        repository_usage = repository_usage or {}

        # Debug logging for usage data matching
        if repository_usage:
            usage_ids = list(repository_usage.keys())[:5]
            wb_ids = [wb['id'] for wb in workbook_data[:5]]
            matched_count = sum(1 for wb in workbook_data if wb['id'] in repository_usage)
            logger.info(f"Analytics: repository_usage has {len(repository_usage)} entries, matched {matched_count}/{len(workbook_data)} workbooks")
            logger.info(f"Analytics: Sample usage IDs: {usage_ids}")
            logger.info(f"Analytics: Sample workbook IDs: {wb_ids}")
        else:
            logger.info("Analytics: No repository_usage data available - using fallback dates")

        # Build view lookup by workbook
        views_by_workbook = {}
        for view in view_data:
            wb_id = view['workbook_id']
            if wb_id not in views_by_workbook:
                views_by_workbook[wb_id] = []
            views_by_workbook[wb_id].append(view)

        # Calculate max values for normalization
        max_views = max([wb.get('total_views_all_time', 0) for wb in workbook_data], default=1) or 1
        max_tables = 1
        if lineage_data.get('workbook_upstream'):
            max_tables = max([len(v.get('tables', [])) for v in lineage_data['workbook_upstream'].values()], default=1) or 1

        for wb in workbook_data:
            wb_id = wb['id']
            wb_views = views_by_workbook.get(wb_id, [])

            # ==================== USAGE METRICS ====================
            total_views_all_time = wb.get('total_views_all_time', 0)
            view_count = wb.get('view_count', 0) or wb.get('dashboard_count', 0)

            # Lineage metrics
            upstream = lineage_data.get('workbook_upstream', {}).get(wb_id, {})
            upstream_tables = len(upstream.get('tables', []))
            upstream_databases = len(upstream.get('databases', set()))
            lineage_impact_count = upstream_tables + upstream_databases

            # Complexity from Metadata API enrichment
            sheet_count = wb.get('sheets_count', view_count)
            dashboard_count = wb.get('dashboards_only_count', 0)
            calc_field_count = wb.get('calculated_field_count', 0)
            custom_sql_count = wb.get('custom_sql_count', 0)
            datasource_count = wb.get('datasource_count', 0)

            # ==================== OPERATIONAL HEALTH ====================
            tasks = operational_data.get('tasks', {}).get(wb_id, {})
            has_refresh_schedule = tasks.get('has_refresh_schedule', False)
            refresh_task_count = tasks.get('task_count', 0)

            # Subscription data
            subs = operational_data.get('subscriptions', {}).get(wb_id, {})
            subscriber_count = subs.get('subscriber_count', 0)
            has_subscriptions = subs.get('has_subscriptions', False)

            # Job failures
            job_failures_30d = operational_data.get('job_failures', {}).get(wb_id, 0)

            # ==================== GOVERNANCE METRICS ====================
            owner_id = wb.get('owner_id', '')
            owner_name = wb.get('owner_name', '')
            owner_active = wb.get('owner_active', True)

            # Fallback to operational data for owner info
            if not owner_name or owner_active is None:
                owner_info = operational_data.get('users', {}).get(owner_id, {})
                if not owner_name:
                    owner_name = owner_info.get('name', '')
                if owner_active is None:
                    owner_active = owner_info.get('is_active', True)

            has_description = wb.get('has_description', False)
            tag_count = wb.get('tag_count', 0)

            # ==================== TIME-BASED METRICS ====================
            # Check for repository usage data (actual access data)
            repo_data = repository_usage.get(wb_id, {})
            has_repo_data = bool(repo_data)

            if has_repo_data:
                # Use actual access data from repository
                last_viewed = repo_data.get('last_viewed', '')
                views_7d = repo_data.get('views_7d', 0)
                views_30d = repo_data.get('views_30d', 0)
                views_90d = repo_data.get('views_90d', 0)
                unique_users_7d = repo_data.get('unique_users_7d', 0)
                unique_users_30d = repo_data.get('unique_users_30d', 0)
                unique_users_90d = repo_data.get('unique_users_90d', 0)
                active_7d = repo_data.get('active_7d', False)
                active_30d = repo_data.get('active_30d', False)
                active_90d = repo_data.get('active_90d', False)
                dormant_90d = repo_data.get('dormant_90d', True)

                # Override total_views if available
                if repo_data.get('total_views', 0) > 0:
                    total_views_all_time = repo_data['total_views']

                days_since_access = 999
                if last_viewed:
                    try:
                        last_dt = datetime.fromisoformat(last_viewed.replace('Z', '+00:00').replace('+00:00', ''))
                        days_since_access = (datetime.now() - last_dt).days
                    except:
                        pass
            else:
                # Fallback to updated_at (modification date, NOT access date)
                last_viewed = wb.get('updated_at', '')
                views_7d = 0
                views_30d = 0
                views_90d = 0
                unique_users_7d = 0
                unique_users_30d = 0
                unique_users_90d = 0
                active_7d = False
                active_30d = False
                active_90d = False
                dormant_90d = True

                days_since_access = 999
                if last_viewed:
                    try:
                        last_dt = datetime.fromisoformat(last_viewed.replace('Z', '+00:00').replace('+00:00', ''))
                        days_since_access = (datetime.now() - last_dt).days
                        # Estimate activity from modification date (less accurate)
                        active_7d = days_since_access <= 7
                        active_30d = days_since_access <= 30
                        active_90d = days_since_access <= 90
                        dormant_90d = days_since_access > 90
                    except:
                        pass

            # Usage trend (based on actual or estimated access)
            if active_7d:
                usage_trend = 'active'
            elif active_30d:
                usage_trend = 'stable'
            elif active_90d:
                usage_trend = 'declining'
            else:
                usage_trend = 'dormant'

            # ==================== ENGAGEMENT METRICS ====================
            top_view_name = wb.get('top_view_name', '')
            top_view_percentage = wb.get('top_view_percentage', 0)

            # ==================== CRITICALITY SCORE (0-100) ====================
            # Usage component (30%)
            usage_score = min((total_views_all_time / max_views) * 100, 100) if max_views > 0 else 0

            # Unique users component (20%)
            if has_repo_data and unique_users_30d > 0:
                # Use actual unique users from repository
                unique_users_score = min(unique_users_30d * 5, 100)  # Scale: 20 users = 100
            else:
                # Approximated from view count (less accurate)
                unique_users_score = min(view_count * 10, 100)

            # Lineage impact component (20%)
            impact_score = min((upstream_tables / max_tables) * 100, 100) if max_tables > 0 else 0

            # Operational component (15%)
            operational_score = 0
            if has_refresh_schedule:
                operational_score += 50
            if has_subscriptions:
                operational_score += 30
            if owner_active:
                operational_score += 20

            # Complexity component (15%) - higher complexity = more critical
            complexity_raw = sheet_count + dashboard_count + calc_field_count + custom_sql_count + datasource_count
            complexity_score = min(complexity_raw * 5, 100)

            # Total criticality score
            criticality_score = round(
                usage_score * self.CRITICALITY_WEIGHTS['usage'] +
                unique_users_score * self.CRITICALITY_WEIGHTS['unique_users'] +
                impact_score * self.CRITICALITY_WEIGHTS['lineage_impact'] +
                operational_score * self.CRITICALITY_WEIGHTS['operational'] +
                complexity_score * self.CRITICALITY_WEIGHTS['complexity']
            , 1)

            # Determine tier
            if criticality_score >= 70:
                tier = 'Tier 1'
                tier_label = 'Mission Critical'
            elif criticality_score >= 40:
                tier = 'Tier 2'
                tier_label = 'Important'
            else:
                tier = 'Tier 3'
                tier_label = 'Low Priority'

            # ==================== SUNSET SCORE (0-100) ====================
            sunset_score = 0
            sunset_reasons = []

            # No views in 90 days
            if days_since_access > 90:
                sunset_score += self.SUNSET_WEIGHTS['no_views_90d']
                sunset_reasons.append(f'No access in {days_since_access} days')

            # No subscribers
            if not has_subscriptions:
                sunset_score += self.SUNSET_WEIGHTS['no_subscribers']
                sunset_reasons.append('No subscribers')

            # No refresh schedule
            if not has_refresh_schedule:
                sunset_score += self.SUNSET_WEIGHTS['no_refresh']
                sunset_reasons.append('No refresh schedule')

            # Owner inactive
            if not owner_active:
                sunset_score += self.SUNSET_WEIGHTS['owner_inactive']
                sunset_reasons.append('Owner inactive')

            # Low lineage impact
            if lineage_impact_count < 2:
                sunset_score += self.SUNSET_WEIGHTS['low_lineage']
                sunset_reasons.append('Low data dependencies')

            # No description (poor governance)
            if not has_description:
                sunset_score += self.SUNSET_WEIGHTS['no_description']
                sunset_reasons.append('No description')

            # Declining usage trend
            if usage_trend in ['declining', 'dormant']:
                sunset_score += self.SUNSET_WEIGHTS['low_usage_trend']
                sunset_reasons.append(f'Usage trend: {usage_trend}')

            # Sunset candidate flag
            is_sunset_candidate = sunset_score >= 70

            # ==================== GOVERNANCE RISK ====================
            governance_risk_score = 0
            governance_issues = []

            if not owner_active:
                governance_risk_score += 30
                governance_issues.append('Inactive owner')
            if not has_description:
                governance_risk_score += 20
                governance_issues.append('No description')
            if tag_count == 0:
                governance_risk_score += 15
                governance_issues.append('No tags')
            if not has_refresh_schedule and datasource_count > 0:
                governance_risk_score += 20
                governance_issues.append('No refresh schedule')
            if job_failures_30d > 0:
                governance_risk_score += 15
                governance_issues.append(f'{job_failures_30d} job failures')

            if governance_risk_score >= 60:
                governance_risk = 'High'
            elif governance_risk_score >= 30:
                governance_risk = 'Medium'
            else:
                governance_risk = 'Low'

            analytics.append({
                'workbook_id': wb_id,
                'workbook_name': wb['name'],
                'project_name': wb.get('project_name', ''),

                # Owner & Governance
                'owner_name': owner_name,
                'owner_active': owner_active,
                'has_description': has_description,
                'tag_count': tag_count,
                'governance_risk': governance_risk,
                'governance_risk_score': governance_risk_score,
                'governance_issues': governance_issues,

                # Dates
                'created_at': wb.get('created_at', ''),
                'updated_at': wb.get('updated_at', ''),
                'last_viewed': last_viewed,  # Actual last access from repository or updated_at fallback
                'last_accessed': repo_data.get('last_viewed', '') if has_repo_data else '',  # Only from actual access data
                'days_since_access': days_since_access,
                'usage_trend': usage_trend,

                # Access metrics (from repository if available)
                'views_7d': views_7d,
                'views_30d': views_30d,
                'views_90d': views_90d,
                'active_7d': active_7d,
                'active_30d': active_30d,
                'active_90d': active_90d,
                'dormant_90d': dormant_90d,
                'usage_data_source': 'repository' if has_repo_data else 'estimated',

                # Usage metrics
                'total_views_all_time': total_views_all_time,
                'total_access_count': total_views_all_time,  # Backward compatibility
                'unique_users_7d': unique_users_7d,
                'unique_users_30d': unique_users_30d,
                'unique_users_90d': unique_users_90d,
                'top_view_name': top_view_name,
                'top_view_percentage': top_view_percentage,

                # Complexity metrics
                'sheet_count': sheet_count,
                'dashboard_count': dashboard_count,
                'calculated_field_count': calc_field_count,
                'custom_sql_count': custom_sql_count,
                'datasource_count': datasource_count,
                'complexity_score': round(complexity_score, 1),

                # Lineage metrics
                'upstream_tables': upstream_tables,
                'upstream_databases': upstream_databases,
                'lineage_impact_count': lineage_impact_count,
                'upstream_table_names': upstream.get('tables', [])[:5],

                # Operational metrics
                'has_refresh_schedule': has_refresh_schedule,
                'refresh_task_count': refresh_task_count,
                'subscriber_count': subscriber_count,
                'has_subscriptions': has_subscriptions,
                'job_failures_30d': job_failures_30d,

                # Scores
                'usage_score': round(usage_score, 1),
                'unique_users_score': round(unique_users_score, 1),
                'impact_score': round(impact_score, 1),
                'operational_score': round(operational_score, 1),
                'complexity_score': round(complexity_score, 1),
                'criticality_score': criticality_score,
                'total_score': criticality_score,  # Alias for backward compatibility
                'tier': tier,
                'tier_label': tier_label,

                # Engagement
                'top_view_views': wb.get('top_view_views', 0),

                # Sunset analysis
                'sunset_score': sunset_score,
                'sunset_reasons': sunset_reasons,
                'is_sunset_candidate': is_sunset_candidate,
                'sunset_recommendation': 'Sunset' if is_sunset_candidate else ('Review' if sunset_score >= 50 else 'Keep'),

                # Data sources
                'data_sources': ['REST API', 'Metadata API' if lineage_data.get('workbook_upstream') else 'REST API only']
            })

        # Sort by criticality score descending and add rank
        analytics.sort(key=lambda x: x['criticality_score'], reverse=True)
        for i, item in enumerate(analytics):
            item['rank'] = i + 1

        return analytics

    def _compute_view_analytics(self, view_data, workbook_data):
        """Compute view-level analytics"""
        # Build workbook lookup
        wb_lookup = {wb['id']: wb for wb in workbook_data}

        analytics = []
        for view in view_data:
            wb = wb_lookup.get(view['workbook_id'], {})
            analytics.append({
                'view_id': view['id'],
                'view_name': view['name'],
                'workbook_id': view['workbook_id'],
                'workbook_name': wb.get('name', ''),
                'project_name': wb.get('project_name', ''),
                'total_views': view['total_views'],
                'created_at': view.get('created_at', ''),
                'updated_at': view.get('updated_at', ''),
                'data_source': 'REST API'
            })

        # Sort by views and add rank
        analytics.sort(key=lambda x: x['total_views'], reverse=True)
        for i, item in enumerate(analytics):
            item['rank'] = i + 1

        return analytics

    def _compute_datasource_governance(self, workbook_analytics, lineage_data, operational_data):
        """Compute datasource governance analytics"""
        datasource_analytics = []

        try:
            # Get all published datasources - Use Pager to get ALL
            all_datasources = list(TSC.Pager(self.server.datasources))
            logger.info(f"Analytics: Found {len(all_datasources)} datasources")

            # Build datasource-to-workbook mapping (datasource_id -> list of workbook_ids that use it).
            # Built from self.wb_datasource_map (workbook luid -> upstream datasource luids), populated in
            # _enrich_workbooks_from_metadata_api. NOTE: previously this read a 'datasource_downstream' key
            # that lineage_data never set, so workbook_count/active_workbook_count_* were always 0 here.
            ds_workbook_map = {}  # datasource_id -> list of workbook_ids
            for wb_id, ds_ids in self.wb_datasource_map.items():
                for ds_id in ds_ids:
                    ds_workbook_map.setdefault(ds_id, []).append(wb_id)

            # Build workbook analytics lookup
            wb_analytics_lookup = {wb['workbook_id']: wb for wb in workbook_analytics}

            for ds in all_datasources:
                ds_id = ds.id
                ds_name = ds.name

                # Get owner info
                owner_id = getattr(ds, 'owner_id', '')
                owner_info = self.user_map.get(owner_id, {})
                owner_name = owner_info.get('name', '')
                owner_active = owner_info.get('is_active', True)

                # Get project info
                project_name = self.project_map.get(getattr(ds, 'project_id', ''), '')

                # Get description
                description = getattr(ds, 'description', '') or ''
                has_description = len(description.strip()) > 0

                # Get tags
                tags = [t.name for t in getattr(ds, 'tags', [])] if hasattr(ds, 'tags') and ds.tags else []
                tag_count = len(tags)

                # Get workbooks using this datasource
                using_workbook_ids = ds_workbook_map.get(ds_id, [])
                workbook_count = len(using_workbook_ids)

                # Calculate active workbook counts
                active_workbook_count_7d = 0
                active_workbook_count_30d = 0
                active_workbook_count_90d = 0

                for wb_id in using_workbook_ids:
                    wb_analytics = wb_analytics_lookup.get(wb_id)
                    if wb_analytics:
                        days_since = wb_analytics.get('days_since_access', 999)
                        if days_since <= 7:
                            active_workbook_count_7d += 1
                        if days_since <= 30:
                            active_workbook_count_30d += 1
                        if days_since <= 90:
                            active_workbook_count_90d += 1

                # Get operational data
                ds_tasks = operational_data.get('datasource_tasks', {}).get(ds_id, {})
                has_refresh_schedule = ds_tasks.get('has_refresh_schedule', False)
                refresh_task_count = ds_tasks.get('task_count', 0)
                last_refresh = ds_tasks.get('last_refresh', '')
                refresh_frequency = ds_tasks.get('refresh_frequency', '')

                # Get job failures
                ds_job_failures = operational_data.get('datasource_job_failures', {}).get(ds_id, 0)

                # Get lineage impact
                ds_lineage = lineage_data.get('datasource_upstream', {}).get(ds_id, {})
                upstream_tables = len(ds_lineage.get('tables', []))
                lineage_edge_count = upstream_tables + workbook_count

                # Calculate Health Score (0-100)
                health_score = 0

                # +20 if used by > 5 active workbooks (30d)
                if active_workbook_count_30d > 5:
                    health_score += 20
                elif active_workbook_count_30d > 0:
                    health_score += 10

                # +20 if refresh schedule exists
                if has_refresh_schedule:
                    health_score += 20

                # +20 if no refresh failures
                if ds_job_failures == 0:
                    health_score += 20
                elif ds_job_failures < 3:
                    health_score += 10

                # +20 if owner active
                if owner_active:
                    health_score += 20

                # +20 if description exists
                if has_description:
                    health_score += 20

                # Calculate Risk Level
                risk_score = 100 - health_score
                if risk_score >= 60:
                    risk_level = 'High'
                elif risk_score >= 30:
                    risk_level = 'Medium'
                else:
                    risk_level = 'Low'

                # ==================== Compute Sunset Score ====================
                # Sunset weights similar to workbooks
                SUNSET_WEIGHTS = {
                    'no_views_90d': 30,
                    'no_subscribers': 15,
                    'no_refresh': 15,
                    'owner_inactive': 20,
                    'low_lineage': 10,
                    'no_description': 5,
                    'low_usage_trend': 5
                }

                sunset_score = 0
                sunset_reasons = []

                # Check if no active workbooks in 90 days
                if active_workbook_count_90d == 0:
                    sunset_score += SUNSET_WEIGHTS['no_views_90d']
                    if workbook_count == 0:
                        sunset_reasons.append('No workbooks using datasource')
                    else:
                        sunset_reasons.append('No active workbooks in 90 days')
                elif active_workbook_count_30d == 0 and active_workbook_count_90d > 0:
                    sunset_score += SUNSET_WEIGHTS['no_views_90d'] // 2
                    sunset_reasons.append('No active workbooks in 30 days')

                # Check refresh schedule
                if not has_refresh_schedule:
                    sunset_score += SUNSET_WEIGHTS['no_refresh']
                    sunset_reasons.append('No refresh schedule')
                elif ds_job_failures > 3:
                    sunset_score += SUNSET_WEIGHTS['no_refresh'] // 2
                    sunset_reasons.append('High refresh failures')

                # Check owner status
                if not owner_active:
                    sunset_score += SUNSET_WEIGHTS['owner_inactive']
                    sunset_reasons.append('Owner inactive')

                # Check lineage impact
                if lineage_edge_count <= 1:
                    sunset_score += SUNSET_WEIGHTS['low_lineage']
                    sunset_reasons.append('Low lineage impact')

                # Check description
                if not has_description:
                    sunset_score += SUNSET_WEIGHTS['no_description']
                    sunset_reasons.append('No description')

                # Determine recommendation
                if sunset_score >= 70:
                    sunset_recommendation = 'Sunset'
                    is_sunset_candidate = True
                elif sunset_score >= 50:
                    sunset_recommendation = 'Review'
                    is_sunset_candidate = True
                else:
                    sunset_recommendation = 'Keep'
                    is_sunset_candidate = False

                datasource_analytics.append({
                    'datasource_id': ds_id,
                    'datasource_name': ds_name,
                    'project_name': project_name,
                    'content_url': getattr(ds, 'content_url', ''),

                    # Owner & Governance
                    'owner_name': owner_name,
                    'owner_active': owner_active,
                    'has_description': has_description,
                    'tag_count': tag_count,

                    # Usage
                    'workbook_count': workbook_count,
                    'active_workbook_count_7d': active_workbook_count_7d,
                    'active_workbook_count_30d': active_workbook_count_30d,
                    'active_workbook_count_90d': active_workbook_count_90d,

                    # Operational
                    'has_refresh_schedule': has_refresh_schedule,
                    'refresh_task_count': refresh_task_count,
                    'last_refresh': last_refresh,
                    'refresh_frequency': refresh_frequency,
                    'refresh_failures_30d': ds_job_failures,

                    # Impact
                    'upstream_tables': upstream_tables,
                    'lineage_edge_count': lineage_edge_count,

                    # Scores
                    'health_score': health_score,
                    'risk_score': risk_score,
                    'risk_level': risk_level,

                    # Sunset
                    'is_sunset_candidate': is_sunset_candidate,
                    'sunset_score': sunset_score,
                    'sunset_recommendation': sunset_recommendation,
                    'sunset_reasons': sunset_reasons,

                    'data_source': 'REST API + Lineage'
                })

        except Exception as e:
            logger.error(f"Error computing datasource governance: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())

        # Sort by health score descending
        datasource_analytics.sort(key=lambda x: x['health_score'], reverse=True)
        for i, item in enumerate(datasource_analytics):
            item['rank'] = i + 1

        return datasource_analytics

    def _identify_retirement_candidates(self, workbook_analytics, operational_data):
        """Identify workbooks that are candidates for retirement"""
        candidates = []

        for wb in workbook_analytics:
            reasons = []
            retirement_score = 0

            # Check criteria
            if wb['days_since_access'] > 90:
                reasons.append(f"Not accessed in {wb['days_since_access']} days")
                retirement_score += 30

            if not wb['has_refresh_schedule']:
                reasons.append("No refresh schedule")
                retirement_score += 15

            if not wb['owner_active']:
                reasons.append("Owner is inactive")
                retirement_score += 20

            if wb.get('total_access_count', 0) < 10:
                reasons.append("Very low usage (< 10 accesses)")
                retirement_score += 25

            if wb['total_score'] < 30:
                reasons.append("Low criticality score")
                retirement_score += 10

            # If multiple criteria met, flag as candidate
            if len(reasons) >= 2 or retirement_score >= 50:
                candidates.append({
                    'workbook_id': wb['workbook_id'],
                    'workbook_name': wb['workbook_name'],
                    'project_name': wb['project_name'],
                    'owner_name': wb['owner_name'],
                    'days_since_access': wb['days_since_access'],
                    'total_access_count': wb.get('total_access_count', 0),
                    'criticality_score': wb['total_score'],
                    'tier': wb['tier'],
                    'retirement_score': retirement_score,
                    'reasons': reasons
                })

        # Sort by retirement score
        candidates.sort(key=lambda x: x['retirement_score'], reverse=True)
        return candidates

    def _prepare_lineage_impact(self, lineage_data):
        """Prepare lineage impact data for export"""
        impact_data = []

        for table in lineage_data.get('tables', []):
            impact_data.append({
                'table_name': table['name'],
                'schema': table.get('schema', ''),
                'database': table.get('database_name', ''),
                'connection_type': table.get('connection_type', ''),
                'downstream_workbook_count': len(table.get('downstream_workbooks', [])),
                'downstream_datasource_count': len(table.get('downstream_datasources', [])),
                'downstream_workbooks': ', '.join(table.get('downstream_workbooks', [])[:5]),
                'downstream_datasources': ', '.join(table.get('downstream_datasources', [])[:5]),
                'column_count': table.get('column_count', 0),
                'data_source': 'Metadata API'
            })

        return impact_data

    def _compute_kpi_summary(self, results):
        """Compute comprehensive KPI summary for governance dashboard"""
        workbooks = results['workbook_analytics']
        views = results['view_analytics']
        datasources = results.get('datasource_analytics', [])

        # Usage metrics - based on user access, NOT extract refresh
        total_views_all_time = sum([wb.get('total_views_all_time', 0) for wb in workbooks])
        active_workbooks_30d = len([wb for wb in workbooks if wb['days_since_access'] <= 30])
        active_workbooks_7d = len([wb for wb in workbooks if wb['days_since_access'] <= 7])
        dormant_workbooks_90d = len([wb for wb in workbooks if wb['days_since_access'] > 90])
        avg_views_per_workbook = round(total_views_all_time / len(workbooks), 1) if workbooks else 0

        # Governance metrics
        workbooks_without_owner = len([wb for wb in workbooks if not wb.get('owner_active', True)])
        workbooks_without_refresh = len([wb for wb in workbooks if not wb.get('has_refresh_schedule', False)])
        workbooks_without_description = len([wb for wb in workbooks if not wb.get('has_description', False)])
        high_governance_risk = len([wb for wb in workbooks if wb.get('governance_risk') == 'High'])

        # Sunset candidates
        sunset_candidates = len([wb for wb in workbooks if wb.get('is_sunset_candidate', False)])
        review_candidates = len([wb for wb in workbooks if wb.get('sunset_recommendation') == 'Review'])

        # Tier distribution
        tier_counts = {'Tier 1': 0, 'Tier 2': 0, 'Tier 3': 0}
        for wb in workbooks:
            tier = wb.get('tier', 'Tier 3')
            if tier in tier_counts:
                tier_counts[tier] += 1

        # Top workbook by criticality
        top_workbook = workbooks[0] if workbooks else {}

        # Datasource governance metrics
        total_datasources = len(datasources)
        healthy_datasources = len([ds for ds in datasources if ds.get('health_score', 0) >= 70])
        high_risk_datasources = len([ds for ds in datasources if ds.get('risk_level') == 'High'])
        unused_datasources = len([ds for ds in datasources if ds.get('workbook_count', 0) == 0])
        datasource_sunset_candidates = len([ds for ds in datasources if ds.get('is_sunset_candidate', False)])

        return {
            # Basic counts
            'total_workbooks': len(workbooks),
            'total_views': len(views),
            'total_datasources': total_datasources,

            # Usage KPIs (user access only, NOT extract refresh)
            'total_views_all_time': total_views_all_time,
            'total_access_count': total_views_all_time,  # Backward compatibility
            'active_workbooks_30d': active_workbooks_30d,
            'active_workbooks_7d': active_workbooks_7d,
            'dormant_workbooks_90d': dormant_workbooks_90d,
            'avg_views_per_workbook': avg_views_per_workbook,

            # Governance KPIs
            'workbooks_without_owner': workbooks_without_owner,
            'workbooks_without_refresh': workbooks_without_refresh,
            'workbooks_without_description': workbooks_without_description,
            'high_governance_risk': high_governance_risk,

            # Sunset KPIs
            'sunset_candidates': sunset_candidates,
            'retirement_candidates': sunset_candidates,  # Backward compatibility
            'review_candidates': review_candidates,

            # Tier distribution
            'tier_1_count': tier_counts['Tier 1'],
            'tier_2_count': tier_counts['Tier 2'],
            'tier_3_count': tier_counts['Tier 3'],

            # Top workbook info
            'top_workbook_name': top_workbook.get('workbook_name', ''),
            'top_workbook_views': top_workbook.get('total_views_all_time', 0),
            'top_workbook_score': top_workbook.get('criticality_score', 0),

            # Lineage
            'total_lineage_edges': len(results['lineage_impact']),

            # Datasource KPIs
            'healthy_datasources': healthy_datasources,
            'high_risk_datasources': high_risk_datasources,
            'unused_datasources': unused_datasources,
            'datasource_sunset_candidates': datasource_sunset_candidates,

            # Meta
            'metadata_api_available': bool(results['data_sources']['metadata_api']),
            'data_sources_used': results['data_sources']
        }


# Analytics job results storage
analytics_results = {}
analytics_status = {}


@app.route('/api/run_usage_analytics', methods=['POST'])
def run_usage_analytics():
    """Start usage analytics job"""
    try:
        request_data = request.get_json() or {}
        config_data = request_data.get('config') or extractor.config_data

        if not config_data:
            return jsonify({'error': 'No configuration available. Please connect to server first.'})

        # Repository configuration (optional - for Tableau Server)
        repository_config = request_data.get('repository_config')

        # Admin Insights flag (for Tableau Cloud via TSC)
        use_admin_insights = request_data.get('use_admin_insights', False)

        # Log config details for debugging
        logger.info(f"Analytics: Config received with keys: {list(config_data.keys())}")
        logger.info(f"Analytics: Server URL: {config_data.get('server_url', 'N/A')}")
        logger.info(f"Analytics: Site Name: {config_data.get('site_name', 'N/A')}")
        if repository_config:
            logger.info(f"Analytics: Repository config provided for host: {repository_config.get('host', 'N/A')}")
        if use_admin_insights:
            logger.info("Analytics: Admin Insights mode enabled (Tableau Cloud)")

        job_id = str(uuid.uuid4())
        analytics_status[job_id] = {'status': 'queued', 'progress': 0, 'message': 'Starting analytics...'}

        def run_analytics(config, repo_config, use_ai, job_id):
            engine = UsageAnalyticsEngine(config)
            return engine.run_analysis(job_id=job_id, repository_config=repo_config, use_admin_insights=use_ai)

        job_queue.put((job_id, run_analytics, (config_data, repository_config, use_admin_insights), {}))

        return jsonify({'success': True, 'job_id': job_id})
    except Exception as e:
        logger.error(f"Error starting analytics: {str(e)}")
        return jsonify({'error': str(e)})


@app.route('/api/upload_admin_insights', methods=['POST'])
def upload_admin_insights():
    """Upload Admin Insights CSV for Tableau Cloud usage data"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'})

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'})

        if not file.filename.endswith('.csv'):
            return jsonify({'error': 'File must be a CSV'})

        # Read file content
        file_content = file.read()

        # Store in session for use in analytics
        # In production, you might want to store this differently
        global admin_insights_data
        admin_insights_data = file_content

        # Parse to get preview
        df = pd.read_csv(BytesIO(file_content))
        preview = {
            'rows': len(df),
            'columns': list(df.columns),
            'sample': df.head(5).to_dict('records')
        }

        return jsonify({
            'success': True,
            'message': f'Uploaded {len(df)} rows',
            'preview': preview
        })
    except Exception as e:
        logger.error(f"Error uploading admin insights: {str(e)}")
        return jsonify({'error': str(e)})


@app.route('/api/run_usage_analytics_with_insights', methods=['POST'])
def run_usage_analytics_with_insights():
    """Start usage analytics job with Admin Insights data"""
    try:
        request_data = request.get_json() or {}
        config_data = request_data.get('config') or extractor.config_data

        if not config_data:
            return jsonify({'error': 'No configuration available. Please connect to server first.'})

        # Check for uploaded admin insights
        global admin_insights_data
        insights_data = admin_insights_data

        if not insights_data:
            return jsonify({'error': 'No Admin Insights data uploaded. Please upload CSV first.'})

        logger.info(f"Analytics: Using Admin Insights data ({len(insights_data)} bytes)")

        job_id = str(uuid.uuid4())
        analytics_status[job_id] = {'status': 'queued', 'progress': 0, 'message': 'Starting analytics with Admin Insights...'}

        def run_analytics(config, insights, job_id):
            engine = UsageAnalyticsEngine(config)
            return engine.run_analysis(job_id=job_id, admin_insights_data=insights)

        job_queue.put((job_id, run_analytics, (config_data, insights_data), {}))

        return jsonify({'success': True, 'job_id': job_id})
    except Exception as e:
        logger.error(f"Error starting analytics with insights: {str(e)}")
        return jsonify({'error': str(e)})


@app.route('/api/test_repository_connection', methods=['POST'])
def test_repository_connection():
    """Test connection to Tableau Server PostgreSQL repository"""
    try:
        data = request.get_json() or {}

        host = data.get('host', '')
        port = data.get('port', 8060)
        database = data.get('database', 'workgroup')
        user = data.get('user', 'readonly')
        password = data.get('password', '')

        if not host:
            return jsonify({'error': 'Host is required'})

        connector = RepositoryConnector()
        if connector.connect(host, port, database, user, password):
            # Try a simple query
            try:
                cursor = connector.connection.cursor()
                cursor.execute("SELECT COUNT(*) FROM workbooks")
                count = cursor.fetchone()[0]
                cursor.close()
                connector.disconnect()
                return jsonify({
                    'success': True,
                    'message': f'Connected successfully. Found {count} workbooks in repository.'
                })
            except Exception as e:
                connector.disconnect()
                return jsonify({'error': f'Connected but query failed: {str(e)}'})
        else:
            return jsonify({'error': 'Failed to connect to repository'})
    except Exception as e:
        logger.error(f"Repository connection test error: {str(e)}")
        return jsonify({'error': str(e)})


@app.route('/api/analytics_status/<job_id>', methods=['GET'])
def get_analytics_status(job_id):
    """Get analytics job status"""
    if job_id in job_status:
        return jsonify({'success': True, **job_status[job_id]})
    return jsonify({'error': 'Job not found'})


@app.route('/api/analytics_result/<job_id>', methods=['GET'])
def get_analytics_result(job_id):
    """Get analytics job result"""
    if job_id in job_results:
        return jsonify({'success': True, 'data': job_results[job_id]})
    return jsonify({'error': 'Result not found'})


ANALYTICS_EXPORT_SHEETS = [
    ('Workbooks_Analytics', 'workbook_analytics'), ('Views_Analytics', 'view_analytics'),
    ('Lineage_Impact', 'lineage_impact'), ('Datasource_Governance', 'datasource_analytics'),
    ('Flow_Governance', 'flow_analytics'), ('User_Analytics', 'user_analytics'),
    ('Retirement_Candidates', 'retirement_candidates'),
]


def _build_analytics_export(analytics_data, fmt='xlsx'):
    """Bucketed/streamed export of a usage-analytics result (see export_utils.build_export)."""
    datasets = [(sheet, analytics_data.get(key) or []) for sheet, key in ANALYTICS_EXPORT_SHEETS]
    kpi = analytics_data.get('kpi_summary')
    if kpi and isinstance(kpi, dict):
        datasets.append(('KPI_Summary', [kpi]))
    errors = analytics_data.get('errors')
    if errors:
        datasets.append(('Errors', [{'message': str(e)} for e in errors]))
    return build_export(datasets, prefix='usage_analytics', fmt=fmt)


@app.route('/api/export_analytics/<job_id>', methods=['GET'])
def export_analytics_by_job(job_id):
    """Export a stored analytics result straight from the server (no need to post it back from the browser)."""
    try:
        result = job_results.get(job_id)
        if not isinstance(result, dict):
            return jsonify({'error': 'Analytics result not found (it may have been replaced by a newer run)'}), 404
        if request.args.get('check'):
            return jsonify({'success': True})
        fmt = 'csv' if request.args.get('format', 'xlsx').lower() == 'csv' else 'xlsx'
        path, name, mimetype = _build_analytics_export(result, fmt)
        return _send_temp_file(path, name, mimetype)
    except Exception as e:
        logger.error(f"Analytics export error: {e}", exc_info=True)
        return jsonify({'error': str(e)}), 500


@app.route('/api/export_analytics', methods=['POST'])
def export_analytics():
    """Export analytics posted from the browser (fallback when the server-side result is gone)"""
    try:
        data = request.get_json(silent=True) or {}
        job_id = data.get('job_id')
        analytics_data = job_results.get(job_id) if job_id else None
        if not isinstance(analytics_data, dict):
            analytics_data = data.get('data', {})

        fmt = 'csv' if str(data.get('format', 'xlsx')).lower() == 'csv' else 'xlsx'
        path, name, mimetype = _build_analytics_export(analytics_data, fmt)
        return _send_temp_file(path, name, mimetype)
    except Exception as e:
        logger.error(f"Export error: {str(e)}", exc_info=True)
        return jsonify({'error': str(e)}), 500


# ==================== Usage Analysis 2.0 API Endpoints ====================

# Global analyzer instance for Usage Analysis 2.0
usage_analyzer_2 = AdminInsightsAnalyzer()
usage_analysis_2_results = {}


@app.route('/api/run_usage_analysis_2_auto', methods=['POST'])
def run_usage_analysis_2_auto():
    """
    Run Usage Analysis 2.0 by auto-fetching Admin Insights via TSC.
    This is the primary method for Tableau Cloud.
    """
    try:
        request_data = request.get_json() or {}
        config_data = request_data.get('config') or extractor.config_data

        if not config_data:
            return jsonify({'error': 'No configuration available. Please connect to Tableau Cloud first.'})

        job_id = str(uuid.uuid4())
        job_status[job_id] = {'status': 'queued', 'progress': 0, 'message': 'Starting Usage Analysis 2.0...'}

        def run_auto_analysis(config, job_id):
            """Background job for auto-fetching Admin Insights"""
            global usage_analyzer_2
            usage_analyzer_2 = AdminInsightsAnalyzer()

            try:
                job_status[job_id] = {'status': 'running', 'progress': 5, 'message': 'Connecting to Tableau Cloud...'}

                # Authenticate
                server_url = config.get('server_url', '')
                site_name = config.get('site_name', '')
                token_name = config.get('pat_name', '')
                token_secret = config.get('pat_token', '')

                server = TSC.Server(server_url, use_server_version=True)
                auth = TSC.PersonalAccessTokenAuth(token_name, token_secret, site_name)

                with server.auth.sign_in(auth):
                    job_status[job_id] = {'status': 'running', 'progress': 10, 'message': 'Discovering Admin Insights datasources...'}

                    # Find Admin Insights datasources
                    all_datasources, _ = server.datasources.get()
                    admin_insights_ds = {}

                    # Admin Insights datasource patterns
                    ai_patterns = {
                        'ts_events': ['ts events', 'ts_events', 'tsevent'],
                        'site_content': ['site content', 'site_content', 'sitecontent', 'ts content'],
                        'subscriptions': ['subscription', 'ts subscription'],
                        'ts_users': ['ts users', 'ts_users', 'tsuser'],
                        'job_performance': ['job performance', 'background task', 'ts background']
                    }

                    for ds in all_datasources:
                        ds_name_lower = ds.name.lower()
                        project_lower = (ds.project_name or '').lower()

                        # Check if it's in Admin Insights project
                        is_admin_insights_project = 'admin insight' in project_lower or 'admin_insight' in project_lower

                        for key, patterns in ai_patterns.items():
                            for pattern in patterns:
                                if pattern in ds_name_lower or (is_admin_insights_project and any(p in ds_name_lower for p in patterns)):
                                    if key not in admin_insights_ds:
                                        admin_insights_ds[key] = ds
                                        logger.info(f"UsageAnalysis2 Auto: Found {key}: {ds.name} (project: {ds.project_name})")
                                    break

                    if not admin_insights_ds:
                        # Log available datasources for debugging
                        sample_ds = [f"{ds.name} ({ds.project_name})" for ds in all_datasources[:20]]
                        logger.warning(f"UsageAnalysis2 Auto: No Admin Insights found. Sample datasources: {sample_ds}")
                        job_status[job_id] = {'status': 'error', 'progress': 0,
                                            'message': 'Admin Insights datasources not found. Please enable Admin Insights in Tableau Cloud settings or use CSV upload.'}
                        return {'success': False, 'error': 'Admin Insights not found'}

                    job_status[job_id] = {'status': 'running', 'progress': 20,
                                         'message': f'Found {len(admin_insights_ds)} Admin Insights datasources. Downloading...'}

                    # Download and parse each datasource
                    all_data = {}
                    progress_per_ds = 50 // max(len(admin_insights_ds), 1)

                    for idx, (key, ds) in enumerate(admin_insights_ds.items()):
                        try:
                            job_status[job_id] = {'status': 'running',
                                                 'progress': 20 + (idx * progress_per_ds),
                                                 'message': f'Downloading {key}...'}

                            temp_dir = tempfile.mkdtemp()
                            try:
                                # Download datasource
                                file_path = server.datasources.download(
                                    ds.id,
                                    filepath=temp_dir,
                                    include_extract=True
                                )
                                logger.info(f"UsageAnalysis2 Auto: Downloaded {key} to {file_path}")

                                # Parse based on file type
                                df = parse_admin_insights_file(file_path)
                                if df is not None and len(df) > 0:
                                    all_data[key] = df
                                    logger.info(f"UsageAnalysis2 Auto: Parsed {key} with {len(df)} rows")

                            finally:
                                try:
                                    shutil.rmtree(temp_dir)
                                except:
                                    pass

                        except Exception as e:
                            logger.warning(f"UsageAnalysis2 Auto: Error downloading {key}: {str(e)}")

                    if not all_data:
                        job_status[job_id] = {'status': 'error', 'progress': 0,
                                            'message': 'Could not parse Admin Insights data. Try CSV upload.'}
                        return {'success': False, 'error': 'Could not parse data'}

                    # Load data into analyzer
                    job_status[job_id] = {'status': 'running', 'progress': 70, 'message': 'Processing data...'}

                    if 'ts_events' in all_data:
                        usage_analyzer_2.ts_events_df = all_data['ts_events']
                    if 'site_content' in all_data:
                        usage_analyzer_2.site_content_df = all_data['site_content']
                    if 'subscriptions' in all_data:
                        usage_analyzer_2.subscriptions_df = all_data['subscriptions']
                    if 'ts_users' in all_data:
                        usage_analyzer_2.ts_users_df = all_data['ts_users']
                    if 'job_performance' in all_data:
                        usage_analyzer_2.job_performance_df = all_data['job_performance']

                    usage_analyzer_2.site_name = site_name

                    # Run full analysis
                    results = usage_analyzer_2.run_full_analysis(job_id=job_id)
                    return results

            except Exception as e:
                logger.error(f"UsageAnalysis2 Auto: Error - {str(e)}")
                import traceback
                logger.error(traceback.format_exc())
                job_status[job_id] = {'status': 'error', 'progress': 0, 'message': str(e)}
                return {'success': False, 'error': str(e)}

        job_queue.put((job_id, run_auto_analysis, (config_data,), {}))

        return jsonify({'success': True, 'job_id': job_id})

    except Exception as e:
        logger.error(f"UsageAnalysis2 Auto: API error - {str(e)}")
        return jsonify({'error': str(e)})


def parse_admin_insights_file(file_path):
    """Parse Admin Insights file (tdsx/hyper/csv)"""
    try:
        if file_path.endswith('.csv'):
            return pd.read_csv(file_path)

        elif file_path.endswith('.tdsx'):
            # Extract hyper from tdsx
            temp_dir = tempfile.mkdtemp()
            try:
                with zipfile.ZipFile(file_path, 'r') as z:
                    for name in z.namelist():
                        if name.endswith('.hyper'):
                            z.extract(name, temp_dir)
                            hyper_path = os.path.join(temp_dir, name)
                            return parse_hyper_file(hyper_path)
            finally:
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass

        elif file_path.endswith('.hyper'):
            return parse_hyper_file(file_path)

    except Exception as e:
        logger.error(f"UsageAnalysis2: Error parsing file {file_path}: {str(e)}")

    return None


def parse_hyper_file(hyper_path):
    """Parse Hyper file to DataFrame"""
    try:
        from tableauhyperapi import HyperProcess, Connection, Telemetry, TableName

        with HyperProcess(telemetry=Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU) as hyper:
            with Connection(endpoint=hyper.endpoint, database=hyper_path) as connection:
                # Get all tables
                schemas = connection.catalog.get_schema_names()

                for schema in schemas:
                    tables = connection.catalog.get_table_names(schema=schema)
                    for table in tables:
                        # Read table into pandas
                        table_def = connection.catalog.get_table_definition(table)
                        columns = [col.name.unescaped for col in table_def.columns]

                        query = f'SELECT * FROM {table}'
                        result = connection.execute_query(query)
                        rows = list(result)

                        if rows:
                            df = pd.DataFrame(rows, columns=columns)
                            logger.info(f"UsageAnalysis2: Parsed hyper table {table} with {len(df)} rows, columns: {columns[:5]}...")
                            return df

    except ImportError:
        logger.warning("UsageAnalysis2: tableauhyperapi not installed")
    except Exception as e:
        logger.error(f"UsageAnalysis2: Error parsing hyper: {str(e)}")

    return None


@app.route('/api/upload_admin_insights_2', methods=['POST'])
def upload_admin_insights_2():
    """Upload Admin Insights CSV files for Usage Analysis 2.0"""
    try:
        uploaded_files = {}

        # Handle multiple file uploads
        for key in ['ts_events', 'site_content', 'subscriptions', 'ts_users', 'job_performance', 'combined']:
            if key in request.files:
                file = request.files[key]
                if file.filename and file.filename.endswith('.csv'):
                    uploaded_files[key] = file.read()
                    logger.info(f"UsageAnalysis2: Uploaded {key} ({len(uploaded_files[key])} bytes)")

        if not uploaded_files:
            return jsonify({'error': 'No CSV files uploaded'})

        # Initialize analyzer
        global usage_analyzer_2
        usage_analyzer_2 = AdminInsightsAnalyzer()

        # Load files
        if 'combined' in uploaded_files:
            usage_analyzer_2.load_combined_csv(uploaded_files['combined'])
        else:
            usage_analyzer_2.load_csv_data(
                ts_events_file=uploaded_files.get('ts_events'),
                site_content_file=uploaded_files.get('site_content'),
                subscriptions_file=uploaded_files.get('subscriptions'),
                ts_users_file=uploaded_files.get('ts_users'),
                job_performance_file=uploaded_files.get('job_performance')
            )

        # Get preview info
        preview = {
            'ts_events_rows': len(usage_analyzer_2.ts_events_df) if usage_analyzer_2.ts_events_df is not None else 0,
            'site_content_rows': len(usage_analyzer_2.site_content_df) if usage_analyzer_2.site_content_df is not None else 0,
            'subscriptions_rows': len(usage_analyzer_2.subscriptions_df) if usage_analyzer_2.subscriptions_df is not None else 0,
            'ts_users_rows': len(usage_analyzer_2.ts_users_df) if usage_analyzer_2.ts_users_df is not None else 0,
            'job_performance_rows': len(usage_analyzer_2.job_performance_df) if usage_analyzer_2.job_performance_df is not None else 0
        }

        return jsonify({
            'success': True,
            'message': 'Files uploaded successfully',
            'preview': preview
        })

    except Exception as e:
        logger.error(f"UsageAnalysis2: Upload error - {str(e)}")
        return jsonify({'error': str(e)})


@app.route('/api/run_usage_analysis_2', methods=['POST'])
def run_usage_analysis_2():
    """Run Usage Analysis 2.0"""
    try:
        global usage_analyzer_2, usage_analysis_2_results

        if usage_analyzer_2.ts_events_df is None:
            return jsonify({'error': 'No data loaded. Please upload Admin Insights CSV files first.'})

        job_id = str(uuid.uuid4())
        job_status[job_id] = {'status': 'queued', 'progress': 0, 'message': 'Starting Usage Analysis 2.0...'}

        def run_analysis(analyzer, job_id):
            return analyzer.run_full_analysis(job_id=job_id)

        job_queue.put((job_id, run_analysis, (usage_analyzer_2,), {}))

        return jsonify({'success': True, 'job_id': job_id})

    except Exception as e:
        logger.error(f"UsageAnalysis2: Run error - {str(e)}")
        return jsonify({'error': str(e)})


@app.route('/api/usage_analysis_2_status/<job_id>', methods=['GET'])
def get_usage_analysis_2_status(job_id):
    """Get Usage Analysis 2.0 job status"""
    status = job_status.get(job_id, {'status': 'not_found', 'progress': 0, 'message': 'Job not found'})
    return jsonify(status)


@app.route('/api/usage_analysis_2_debug', methods=['GET'])
def get_usage_analysis_2_debug():
    """
    Debug endpoint for Usage Analysis 2.0 - shows raw data counts, column mappings,
    and data quality information to help diagnose issues.
    """
    global usage_analyzer_2

    debug_data = {
        'data_sources': {},
        'column_info': {},
        'item_types': {},
        'event_types': {},
        'sample_data': {},
        'warnings': [],
        'analysis_info': {}
    }

    try:
        # TS Events info
        if usage_analyzer_2.ts_events_df is not None and len(usage_analyzer_2.ts_events_df) > 0:
            df = usage_analyzer_2.ts_events_df
            debug_data['data_sources']['ts_events'] = len(df)
            debug_data['column_info']['ts_events'] = list(df.columns)

            # Find event types
            event_col = None
            for candidate in ['event_name', 'event_type', 'event', 'action', 'Event Name']:
                if candidate in df.columns:
                    event_col = candidate
                    break
            if event_col:
                debug_data['event_types'] = df[event_col].value_counts().head(20).to_dict()

            # Sample data (first 3 rows, first 8 columns)
            sample = df.head(3).iloc[:, :min(8, len(df.columns))].fillna('').astype(str)
            debug_data['sample_data']['ts_events'] = sample.to_dict('records')

        # Site Content info
        if usage_analyzer_2.site_content_df is not None and len(usage_analyzer_2.site_content_df) > 0:
            df = usage_analyzer_2.site_content_df
            debug_data['data_sources']['site_content'] = len(df)
            debug_data['column_info']['site_content'] = list(df.columns)

            # Find item types
            type_col = None
            for candidate in ['item_type', 'content_type', 'type', 'Item Type', 'Content Type']:
                if candidate in df.columns:
                    type_col = candidate
                    break
            if type_col:
                raw_types = df[type_col].value_counts().to_dict()
                debug_data['item_types']['raw'] = raw_types
                # Show normalized types
                normalized_counts = {}
                for item_type, count in raw_types.items():
                    normalized = usage_analyzer_2._normalize_item_type(item_type)
                    normalized_counts[normalized] = normalized_counts.get(normalized, 0) + count
                debug_data['item_types']['normalized'] = normalized_counts

            # Sample data
            sample = df.head(3).iloc[:, :min(8, len(df.columns))].fillna('').astype(str)
            debug_data['sample_data']['site_content'] = sample.to_dict('records')

        # Subscriptions info
        if usage_analyzer_2.subscriptions_df is not None:
            debug_data['data_sources']['subscriptions'] = len(usage_analyzer_2.subscriptions_df)
            debug_data['column_info']['subscriptions'] = list(usage_analyzer_2.subscriptions_df.columns)

        # TS Users info
        if usage_analyzer_2.ts_users_df is not None:
            debug_data['data_sources']['ts_users'] = len(usage_analyzer_2.ts_users_df)
            debug_data['column_info']['ts_users'] = list(usage_analyzer_2.ts_users_df.columns)

        # Job Performance info
        if usage_analyzer_2.job_performance_df is not None:
            debug_data['data_sources']['job_performance'] = len(usage_analyzer_2.job_performance_df)
            debug_data['column_info']['job_performance'] = list(usage_analyzer_2.job_performance_df.columns)

        # Analysis results info
        if hasattr(usage_analyzer_2, 'workbook_intelligence') and usage_analyzer_2.workbook_intelligence is not None:
            debug_data['analysis_info']['workbooks'] = len(usage_analyzer_2.workbook_intelligence)
            if len(usage_analyzer_2.workbook_intelligence) > 0:
                wb = usage_analyzer_2.workbook_intelligence
                debug_data['analysis_info']['workbooks_with_views'] = int((wb['views_30d'] > 0).sum()) if 'views_30d' in wb.columns else 0

        if hasattr(usage_analyzer_2, 'datasource_intelligence') and usage_analyzer_2.datasource_intelligence is not None:
            debug_data['analysis_info']['datasources'] = len(usage_analyzer_2.datasource_intelligence)
            if len(usage_analyzer_2.datasource_intelligence) > 0:
                ds = usage_analyzer_2.datasource_intelligence
                debug_data['analysis_info']['datasources_with_access'] = int((ds['access_30d'] > 0).sum()) if 'access_30d' in ds.columns else 0

        # Warnings from analyzer
        if hasattr(usage_analyzer_2, 'debug_info'):
            debug_data['warnings'] = usage_analyzer_2.debug_info.get('warnings', [])
            debug_data['column_mappings'] = usage_analyzer_2.debug_info.get('column_mappings', {})
            debug_data['dedup_counts'] = usage_analyzer_2.debug_info.get('dedup_counts', {})

        return jsonify(debug_data)

    except Exception as e:
        logger.error(f"UsageAnalysis2 Debug: Error - {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({'error': str(e)})


def convert_timestamps_to_strings(obj):
    """Recursively convert Timestamp/datetime objects to ISO strings for JSON serialization"""
    import numpy as np
    if isinstance(obj, dict):
        return {k: convert_timestamps_to_strings(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [convert_timestamps_to_strings(item) for item in obj]
    elif isinstance(obj, pd.Timestamp):
        return obj.isoformat() if pd.notna(obj) else None
    elif isinstance(obj, datetime):
        return obj.isoformat()
    elif isinstance(obj, (np.integer, np.floating)):
        return obj.item()  # Convert numpy types to Python native types
    elif obj is None:
        return None
    else:
        try:
            if pd.isna(obj):
                return None
        except (ValueError, TypeError):
            pass
        return obj


@app.route('/api/usage_analysis_2_result/<job_id>', methods=['GET'])
def get_usage_analysis_2_result(job_id):
    """Get Usage Analysis 2.0 results"""
    import numpy as np
    from flask import Response

    if job_id not in job_results:
        return jsonify({'error': 'Results not found'})

    def json_serial(obj):
        """JSON serializer for objects not serializable by default json code"""
        if isinstance(obj, (datetime, pd.Timestamp)):
            return obj.isoformat() if pd.notna(obj) else None
        if isinstance(obj, (np.integer,)):
            return int(obj)
        if isinstance(obj, (np.floating,)):
            return float(obj) if not np.isnan(obj) else None
        if isinstance(obj, (np.ndarray,)):
            return obj.tolist()
        # Handle tableauhyperapi Timestamp type
        if hasattr(obj, '__class__') and 'Timestamp' in obj.__class__.__name__:
            try:
                return str(obj)
            except:
                return None
        # Handle any date/time-like objects
        if hasattr(obj, 'isoformat'):
            return obj.isoformat()
        if hasattr(obj, '__str__') and ('date' in type(obj).__name__.lower() or 'time' in type(obj).__name__.lower()):
            return str(obj)
        try:
            if pd.isna(obj):
                return None
        except:
            pass
        raise TypeError(f"Type {type(obj)} not serializable")

    def clean_nan_values(obj):
        """Recursively replace NaN/Infinity with None for valid JSON"""
        import math
        if isinstance(obj, dict):
            return {k: clean_nan_values(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [clean_nan_values(item) for item in obj]
        elif isinstance(obj, float):
            if math.isnan(obj) or math.isinf(obj):
                return None
            return obj
        elif isinstance(obj, (np.floating,)):
            if np.isnan(obj) or np.isinf(obj):
                return None
            return float(obj)
        return obj

    try:
        # First clean NaN values, then serialize
        cleaned_results = clean_nan_values(job_results[job_id])
        json_str = json.dumps(cleaned_results, default=json_serial)
        return Response(json_str, mimetype='application/json')
    except Exception as e:
        logger.error(f"Error serializing results: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({'error': f'Serialization error: {str(e)}'})


@app.route('/api/export_governance_pack', methods=['POST'])
def export_governance_pack():
    """Export detailed governance pack Excel file"""
    try:
        global usage_analyzer_2

        if usage_analyzer_2.workbook_intelligence is None:
            return jsonify({'error': 'No analysis data available. Run Usage Analysis 2.0 first.'})

        # Get site name from config
        config = extractor.config_data or {}
        usage_analyzer_2.site_name = config.get('site_name', 'Unknown')

        output = usage_analyzer_2.export_governance_pack()

        site_name = usage_analyzer_2.site_name.replace(' ', '_')
        filename = f'Governance_Export_{site_name}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        logger.error(f"UsageAnalysis2: Export governance pack error - {str(e)}")
        return jsonify({'error': str(e)})


@app.route('/api/export_sunset_candidates', methods=['POST'])
def export_sunset_candidates():
    """Export sunset candidates only"""
    try:
        global usage_analyzer_2

        if usage_analyzer_2.sunset_candidates is None or len(usage_analyzer_2.sunset_candidates) == 0:
            return jsonify({'error': 'No sunset candidates available. Run Usage Analysis 2.0 first.'})

        output = BytesIO()
        sunset_export = usage_analyzer_2.sunset_candidates.copy()

        # Convert list columns to strings
        if 'sunset_reasons' in sunset_export.columns:
            sunset_export['sunset_reasons'] = sunset_export['sunset_reasons'].apply(
                lambda x: '; '.join(x) if isinstance(x, list) else str(x)
            )

        # Sanitize for safe Excel export
        sunset_export = sanitize_dataframe_for_excel(sunset_export)

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            sunset_export.to_excel(writer, sheet_name='Sunset Candidates', index=False)

        output.seek(0)

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'Sunset_Candidates_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )

    except Exception as e:
        logger.error(f"UsageAnalysis2: Export sunset candidates error - {str(e)}")
        return jsonify({'error': str(e)})


@app.route('/api/export_scoring_model', methods=['GET'])
def export_scoring_model():
    """Export scoring model definition"""
    try:
        scoring_def = pd.DataFrame([
            {'Score Type': 'Sunset Score', 'Factor': 'No views in 90+ days', 'Weight': 30, 'Description': 'Content not accessed in over 90 days'},
            {'Score Type': 'Sunset Score', 'Factor': 'No views in 30 days', 'Weight': 20, 'Description': 'Content not accessed in 30 days'},
            {'Score Type': 'Sunset Score', 'Factor': 'Owner inactive', 'Weight': 15, 'Description': 'Owner has not logged in for 90+ days'},
            {'Score Type': 'Sunset Score', 'Factor': 'No refresh schedule', 'Weight': 10, 'Description': 'No extract refresh configured'},
            {'Score Type': 'Sunset Score', 'Factor': 'Not certified', 'Weight': 10, 'Description': 'Content not certified'},
            {'Score Type': 'Sunset Score', 'Factor': 'No description', 'Weight': 5, 'Description': 'Missing documentation'},
            {'Score Type': 'Sunset Score', 'Factor': 'No subscribers', 'Weight': 5, 'Description': 'No active subscriptions'},
            {'Score Type': 'Sunset Score', 'Factor': 'Stale content', 'Weight': 5, 'Description': 'Not modified in 180+ days'},
            {'Score Type': 'Health Score', 'Factor': 'Recent activity', 'Weight': 25, 'Description': 'Views in last 7/30/90 days'},
            {'Score Type': 'Health Score', 'Factor': 'User adoption', 'Weight': 25, 'Description': 'Unique users in 30 days'},
            {'Score Type': 'Health Score', 'Factor': 'Data freshness', 'Weight': 20, 'Description': 'Has refresh schedule'},
            {'Score Type': 'Health Score', 'Factor': 'Certification', 'Weight': 15, 'Description': 'Content is certified'},
            {'Score Type': 'Health Score', 'Factor': 'Documentation', 'Weight': 15, 'Description': 'Has description'},
            {'Score Type': 'Tier Classification', 'Factor': 'Tier 1 - Mission Critical', 'Weight': 'N/A', 'Description': 'High usage, certified, active owner'},
            {'Score Type': 'Tier Classification', 'Factor': 'Tier 2 - Important', 'Weight': 'N/A', 'Description': 'Moderate usage, some governance'},
            {'Score Type': 'Tier Classification', 'Factor': 'Tier 3 - Low Priority', 'Weight': 'N/A', 'Description': 'Low usage, governance gaps'},
        ])

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            scoring_def.to_excel(writer, sheet_name='Scoring Model', index=False)

        output.seek(0)

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'Scoring_Model_Definition.xlsx'
        )

    except Exception as e:
        logger.error(f"UsageAnalysis2: Export scoring model error - {str(e)}")
        return jsonify({'error': str(e)})


@app.route('/api/quick_export_2', methods=['POST'])
def quick_export_2():
    """Quick export of current filtered data"""
    try:
        data = request.get_json() or {}
        export_type = data.get('type', 'workbooks')

        global usage_analyzer_2

        output = BytesIO()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if export_type == 'workbooks' and usage_analyzer_2.workbook_intelligence is not None:
                df = sanitize_dataframe_for_excel(usage_analyzer_2.workbook_intelligence)
                df.to_excel(writer, sheet_name='Workbooks', index=False)
            elif export_type == 'datasources' and usage_analyzer_2.datasource_intelligence is not None:
                df = sanitize_dataframe_for_excel(usage_analyzer_2.datasource_intelligence)
                df.to_excel(writer, sheet_name='Datasources', index=False)
            elif export_type == 'flows' and usage_analyzer_2.flow_intelligence is not None:
                df = sanitize_dataframe_for_excel(usage_analyzer_2.flow_intelligence)
                df.to_excel(writer, sheet_name='Flows', index=False)
            elif export_type == 'users' and usage_analyzer_2.user_intelligence is not None:
                user_export = usage_analyzer_2.user_intelligence.copy()
                if 'user_email' in user_export.columns and 'user_email_masked' in user_export.columns:
                    user_export['user_email'] = user_export['user_email_masked']
                user_export = sanitize_dataframe_for_excel(user_export)
                user_export.to_excel(writer, sheet_name='Users', index=False)
            else:
                return jsonify({'error': f'No data available for {export_type}'})

        output.seek(0)

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'{export_type}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )

    except Exception as e:
        logger.error(f"UsageAnalysis2: Quick export error - {str(e)}")
        return jsonify({'error': str(e)})



# ==================== Admin Insights CSV Analytics ====================

# Global Admin Insights analyzer instance
admin_insights_analyzer = AdminInsightsAnalyzer()


@app.route('/api/admin-insights/upload', methods=['POST'])
def upload_admin_insights_csv():
    """Upload Admin Insights CSV files"""
    try:
        csv_type = request.form.get('csv_type')
        valid_types = [
            'site_content', 'ts_events', 'ts_users', 'subscriptions',
            'permissions', 'groups', 'viz_load_times', 'job_performance'
        ]

        if csv_type not in valid_types:
            return jsonify({'error': f'Invalid csv_type. Must be one of: {", ".join(valid_types)}'})

        if 'file' not in request.files:
            return jsonify({'error': 'No file provided'})

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'})

        file_content = file.read()
        filename = file.filename

        global admin_insights_analyzer

        # Load based on type
        load_methods = {
            'site_content': admin_insights_analyzer.load_site_content,
            'ts_events': admin_insights_analyzer.load_ts_events,
            'ts_users': admin_insights_analyzer.load_ts_users,
            'subscriptions': admin_insights_analyzer.load_subscriptions,
            'permissions': admin_insights_analyzer.load_permissions,
            'groups': admin_insights_analyzer.load_groups,
            'viz_load_times': admin_insights_analyzer.load_viz_load_times,
            'job_performance': admin_insights_analyzer.load_job_performance
        }

        success = load_methods[csv_type](file_content)

        if not success:
            return jsonify({'error': 'Failed to parse CSV file'})

        # Get info about loaded file
        status = admin_insights_analyzer.get_upload_status()
        file_info = status.get(csv_type, {})

        return jsonify({
            'ok': True,
            'csv_type': csv_type,
            'filename': filename,
            'rows': file_info.get('rows', 0),
            'columns': file_info.get('columns', []),
            'date_range': file_info.get('date_range', {})
        })

    except Exception as e:
        logger.error(f"Admin Insights Upload Error: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({'error': str(e)})


@app.route('/api/admin-insights/status', methods=['GET'])
def get_admin_insights_status():
    """Get current Admin Insights CSV upload status"""
    global admin_insights_analyzer
    return jsonify(admin_insights_analyzer.get_upload_status())


@app.route('/api/admin-insights/clear', methods=['POST'])
def clear_admin_insights():
    """Clear all Admin Insights CSV uploads"""
    global admin_insights_analyzer
    admin_insights_analyzer = AdminInsightsAnalyzer()
    return jsonify({'ok': True, 'message': 'All Admin Insights data cleared'})


@app.route('/api/admin-insights/run', methods=['POST'])
def run_admin_insights_analysis():
    """Run comprehensive Admin Insights analysis"""
    try:
        global admin_insights_analyzer

        # Check if Site_Content is uploaded (required)
        if admin_insights_analyzer.site_content_df is None:
            return jsonify({'error': 'Site_Content.csv is required. Please upload it first.'})

        job_id = str(uuid.uuid4())
        job_status[job_id] = {'status': 'running', 'progress': 10, 'message': 'Initializing analysis...'}

        def run_analysis(analyzer, job_id):
            try:
                job_status[job_id] = {'status': 'running', 'progress': 20, 'message': 'Building content master...'}
                
                job_status[job_id] = {'status': 'running', 'progress': 40, 'message': 'Computing governance scores...'}
                
                job_status[job_id] = {'status': 'running', 'progress': 60, 'message': 'Building workbook governance...'}
                
                job_status[job_id] = {'status': 'running', 'progress': 80, 'message': 'Computing summary metrics...'}
                
                results = analyzer.run_full_analysis()

                job_status[job_id] = {'status': 'completed', 'progress': 100, 'message': 'Analysis complete!'}
                job_results[job_id] = results
                return results

            except Exception as e:
                logger.error(f"Admin Insights Analysis Error: {str(e)}")
                import traceback
                logger.error(traceback.format_exc())
                job_status[job_id] = {'status': 'error', 'progress': 0, 'message': str(e)}
                return {'success': False, 'error': str(e)}

        job_queue.put((job_id, run_analysis, (admin_insights_analyzer,), {}))
        return jsonify({'success': True, 'job_id': job_id})

    except Exception as e:
        logger.error(f"Admin Insights Run Error: {str(e)}")
        return jsonify({'error': str(e)})


@app.route('/api/admin-insights/export', methods=['GET'])
def export_admin_insights():
    """Export Admin Insights analysis to Excel"""
    try:
        global admin_insights_analyzer

        if admin_insights_analyzer.content_master is None:
            return jsonify({'error': 'No analysis data available. Run analysis first.'})

        include_raw = request.args.get('include_raw', 'false').lower() == 'true'
        output = admin_insights_analyzer.export_to_excel(include_raw=include_raw)

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'Admin_Insights_Governance_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )

    except Exception as e:
        logger.error(f"Admin Insights Export Error: {str(e)}")
        return jsonify({'error': str(e)})


# Legacy endpoints for backward compatibility with existing frontend
@app.route('/api/usage/csv/upload', methods=['POST'])
def upload_usage_csv_legacy():
    """Legacy endpoint - redirects to Admin Insights upload"""
    # Map old types to new types
    type_map = {
        'traffic_views': 'ts_events',
        'traffic_datasources': 'ts_events',
        'space_usage': 'site_content'
    }
    
    csv_type = request.form.get('csv_type', '')
    if csv_type in type_map:
        # Modify the request to use the new type
        request.form = request.form.copy()
        request.form['csv_type'] = type_map[csv_type]
    
    return upload_admin_insights_csv()


@app.route('/api/usage/csv/status', methods=['GET'])
def get_csv_upload_status_legacy():
    """Legacy endpoint for CSV status"""
    return get_admin_insights_status()


@app.route('/api/usage/csv/clear', methods=['POST'])
def clear_csv_uploads_legacy():
    """Legacy endpoint to clear uploads"""
    return clear_admin_insights()


@app.route('/api/usage/csv/run', methods=['POST'])
def run_csv_analysis_legacy():
    """Legacy endpoint to run analysis"""
    return run_admin_insights_analysis()


@app.route('/api/usage/csv/export', methods=['GET'])
def export_csv_analysis_legacy():
    """Legacy endpoint for export"""
    return export_admin_insights()


# ==================== AI Assistant (Multi-Provider Support) ====================

# AI Configuration - Supports Claude, OpenAI, and Ollama
AI_CONFIG = {
    'provider': 'ollama',  # 'claude', 'openai', or 'ollama'
    'claude_api_key': '',  # User's Claude API key
    'openai_api_key': '',  # User's OpenAI API key
    'claude_model': 'claude-sonnet-4-20250514',  # Default Claude model
    'openai_model': 'gpt-4-turbo-preview',  # Default OpenAI model
    'ollama_url': 'http://localhost:11434',
    'ollama_model': 'llama3.2',  # Default Ollama model
    'model': 'llama3.2',  # Active model (for backward compat)
    'enabled': True
}

# Conversation history storage (in-memory, per session concept)
ai_conversations = {}


def query_claude(messages, model=None, api_key=None):
    """Send a query to Claude API and get response"""
    model = model or AI_CONFIG['claude_model']
    api_key = api_key or AI_CONFIG['claude_api_key']

    if not api_key:
        return {'success': False, 'error': 'Claude API key not configured'}

    try:
        # Convert messages to Claude format
        system_message = ""
        claude_messages = []
        for msg in messages:
            if msg['role'] == 'system':
                system_message = msg['content']
            else:
                claude_messages.append({
                    'role': msg['role'],
                    'content': msg['content']
                })

        response = requests.post(
            'https://api.anthropic.com/v1/messages',
            headers={
                'Content-Type': 'application/json',
                'x-api-key': api_key,
                'anthropic-version': '2023-06-01'
            },
            json={
                'model': model,
                'max_tokens': 4096,
                'system': system_message,
                'messages': claude_messages
            },
            timeout=120
        )

        if response.status_code == 200:
            result = response.json()
            content = result.get('content', [])
            text = content[0].get('text', '') if content else ''
            return {
                'success': True,
                'response': text,
                'model': model,
                'provider': 'claude'
            }
        else:
            error_msg = response.json().get('error', {}).get('message', f'Status {response.status_code}')
            return {'success': False, 'error': f"Claude API error: {error_msg}"}
    except requests.exceptions.Timeout:
        return {'success': False, 'error': 'Request timed out'}
    except Exception as e:
        return {'success': False, 'error': str(e)}


def query_openai(messages, model=None, api_key=None):
    """Send a query to OpenAI API and get response"""
    model = model or AI_CONFIG['openai_model']
    api_key = api_key or AI_CONFIG['openai_api_key']

    if not api_key:
        return {'success': False, 'error': 'OpenAI API key not configured'}

    try:
        response = requests.post(
            'https://api.openai.com/v1/chat/completions',
            headers={
                'Content-Type': 'application/json',
                'Authorization': f'Bearer {api_key}'
            },
            json={
                'model': model,
                'messages': messages,
                'max_tokens': 4096
            },
            timeout=120
        )

        if response.status_code == 200:
            result = response.json()
            text = result.get('choices', [{}])[0].get('message', {}).get('content', '')
            return {
                'success': True,
                'response': text,
                'model': model,
                'provider': 'openai'
            }
        else:
            error_msg = response.json().get('error', {}).get('message', f'Status {response.status_code}')
            return {'success': False, 'error': f"OpenAI API error: {error_msg}"}
    except requests.exceptions.Timeout:
        return {'success': False, 'error': 'Request timed out'}
    except Exception as e:
        return {'success': False, 'error': str(e)}


def query_ai(messages, provider=None, model=None, api_key=None):
    """Universal AI query function - routes to appropriate provider"""
    provider = provider or AI_CONFIG['provider']

    if provider == 'claude':
        return query_claude(messages, model or AI_CONFIG['claude_model'], api_key or AI_CONFIG['claude_api_key'])
    elif provider == 'openai':
        return query_openai(messages, model or AI_CONFIG['openai_model'], api_key or AI_CONFIG['openai_api_key'])
    else:
        # Default to Ollama
        return query_ollama(messages, model or AI_CONFIG['ollama_model'])

def check_ollama_status():
    """Check if Ollama is running and available"""
    try:
        response = requests.get(f"{AI_CONFIG['ollama_url']}/api/tags", timeout=5)
        if response.status_code == 200:
            models = response.json().get('models', [])
            return {
                'available': True,
                'models': [m['name'] for m in models],
                'current_model': AI_CONFIG['model']
            }
    except Exception as e:
        logger.warning(f"Ollama not available: {e}")
    return {'available': False, 'models': [], 'current_model': None}

def get_context_summary():
    """Get a summary of current data context for AI including detailed extraction data"""
    context = {
        'has_extraction': bool(job_results),
        'has_analytics': False,
        'stats': {},
        'datasources': [],
        'connections': [],
        'workbooks': []
    }

    # Get basic stats from current state
    if hasattr(app, 'current_stats'):
        context['stats'] = app.current_stats

    # Check for analytics data and extract detailed information
    for job_id, result in job_results.items():
        if isinstance(result, dict):
            if 'workbook_analytics' in result or 'inventory' in result:
                context['has_analytics'] = True
                context['workbook_count'] = len(result.get('workbook_analytics', result.get('inventory', [])))
                context['datasource_count'] = len(result.get('datasource_analytics', []))
                context['user_count'] = len(result.get('user_analytics', []))

            # Extract detailed datasource and connection information
            if 'bridge_connections' in result:
                seen_datasources = set()
                for conn in result['bridge_connections']:
                    ds_name = conn.get('object_name', '')
                    obj_type = conn.get('object_type', '')
                    conn_type = conn.get('connection_type', '')
                    server_addr = conn.get('server_address', '')

                    # Track unique datasources
                    if obj_type == 'Data Source' and ds_name and ds_name not in seen_datasources:
                        seen_datasources.add(ds_name)
                        context['datasources'].append({
                            'name': ds_name,
                            'project': conn.get('project_name', ''),
                            'has_extracts': conn.get('has_extracts', '')
                        })

                    # Track all connections with server info
                    if conn_type or server_addr:
                        context['connections'].append({
                            'datasource': ds_name,
                            'type': obj_type,
                            'connection_type': conn_type,
                            'server': server_addr,
                            'port': conn.get('server_port', ''),
                            'uses_bridge': conn.get('use_tableau_bridge', '')
                        })

            # Extract workbook names if available
            if 'inventory' in result:
                for item in result['inventory'][:50]:  # Limit to 50
                    if isinstance(item, dict):
                        context['workbooks'].append({
                            'name': item.get('name', item.get('Name', '')),
                            'project': item.get('project_name', item.get('Project', '')),
                            'owner': item.get('owner', item.get('Owner', ''))
                        })

            # Get custom SQL info if available
            if 'custom_sql' in result:
                context['custom_sql_count'] = len(result['custom_sql'])
                context['custom_sql_samples'] = []
                for sql in result['custom_sql'][:10]:  # First 10 samples
                    context['custom_sql_samples'].append({
                        'datasource': sql.get('datasource_name', ''),
                        'sql_name': sql.get('sql_name', ''),
                        'database': sql.get('database', ''),
                        'connection_type': sql.get('connection_type', '')
                    })

            # Break after first result with data
            if context['datasources'] or context['connections']:
                break

    return context

def build_system_prompt(context):
    """Build the system prompt for AI with current context (legacy)"""
    return build_system_prompt_v2(context)


def build_system_prompt_v2(context):
    """Build an enhanced system prompt for AI with current context"""
    workbooks = context.get('workbook_count', 0)
    datasources = context.get('datasource_count', 0)
    projects = context.get('project_count', 0)
    users = context.get('user_count', 0)
    views = context.get('view_count', 0)

    # Build detailed datasource list
    datasource_details = ""
    if context.get('datasources'):
        ds_list = context['datasources'][:30]  # Limit to 30 for prompt size
        datasource_details = "\n\nDATASOURCE DETAILS:\n"
        for i, ds in enumerate(ds_list, 1):
            datasource_details += f"{i}. {ds.get('name', 'Unknown')}"
            if ds.get('project'):
                datasource_details += f" (Project: {ds['project']})"
            if ds.get('has_extracts'):
                datasource_details += f" [Extracts: {ds['has_extracts']}]"
            datasource_details += "\n"
        if len(context['datasources']) > 30:
            datasource_details += f"... and {len(context['datasources']) - 30} more datasources\n"

    # Build connection details
    connection_details = ""
    if context.get('connections'):
        conn_list = context['connections'][:30]  # Limit to 30
        connection_details = "\n\nCONNECTION DETAILS:\n"
        for conn in conn_list:
            ds_name = conn.get('datasource', 'Unknown')
            conn_type = conn.get('connection_type', '')
            server = conn.get('server', '')
            if conn_type or server:
                connection_details += f"- {ds_name}: "
                if conn_type:
                    connection_details += f"Type={conn_type}"
                if server:
                    connection_details += f", Server={server}"
                if conn.get('port'):
                    connection_details += f":{conn['port']}"
                if conn.get('uses_bridge'):
                    connection_details += f" [Bridge: {conn['uses_bridge']}]"
                connection_details += "\n"
        if len(context['connections']) > 30:
            connection_details += f"... and {len(context['connections']) - 30} more connections\n"

    # Build workbook details
    workbook_details = ""
    if context.get('workbooks'):
        wb_list = context['workbooks'][:20]  # Limit to 20
        workbook_details = "\n\nWORKBOOK DETAILS:\n"
        for i, wb in enumerate(wb_list, 1):
            workbook_details += f"{i}. {wb.get('name', 'Unknown')}"
            if wb.get('project'):
                workbook_details += f" (Project: {wb['project']})"
            if wb.get('owner'):
                workbook_details += f" [Owner: {wb['owner']}]"
            workbook_details += "\n"
        if len(context['workbooks']) > 20:
            workbook_details += f"... and {len(context['workbooks']) - 20} more workbooks\n"

    system_prompt = f"""You are a Metadata Intelligence Agent for Tableau. You are an expert BI consultant who helps users understand and optimize their Tableau environment.

CURRENT TABLEAU ENVIRONMENT:
- Total Workbooks: {workbooks}
- Total Datasources: {datasources}
- Total Projects: {projects}
- Total Users: {users}
- Total Views: {views}
{datasource_details}{connection_details}{workbook_details}
YOUR ROLE:
You are NOT a generic chatbot. You are a specialized Tableau metadata analyst. When users ask questions:

1. ALWAYS use the actual data provided above when answering questions
2. If asked for datasource names, connection types, or servers - use the DETAILED DATA above
3. Provide specific, actionable insights based on the data
4. Think like a senior BI consultant or Tableau governance specialist
5. Give direct answers using the data above - don't say "I don't have the data"

RESPONSE STYLE:
- Be conversational but professional
- Use specific names, types, and servers from the data above
- When asked about datasources, list actual names from DATASOURCE DETAILS
- When asked about connections, use CONNECTION DETAILS with server names and types
- Provide context and recommendations
- Don't be overly verbose - get to the point

EXAMPLE RESPONSES:
- If asked "how many workbooks?", say "You have {workbooks} workbooks in your Tableau environment."
- If asked "list datasource names", list the actual names from DATASOURCE DETAILS
- If asked "what connections?", list the actual connection types and servers from CONNECTION DETAILS

Remember: You have DETAILED data about their Tableau environment including specific names and connections. USE IT!"""

    return system_prompt

def query_ollama(messages, model=None):
    """Send a query to Ollama and get response"""
    model = model or AI_CONFIG['model']

    try:
        response = requests.post(
            f"{AI_CONFIG['ollama_url']}/api/chat",
            json={
                'model': model,
                'messages': messages,
                'stream': False
            },
            timeout=120
        )

        if response.status_code == 200:
            result = response.json()
            return {
                'success': True,
                'response': result.get('message', {}).get('content', ''),
                'model': model
            }
        else:
            return {
                'success': False,
                'error': f"Ollama returned status {response.status_code}"
            }
    except requests.exceptions.Timeout:
        return {'success': False, 'error': 'Request timed out. The model may be loading.'}
    except requests.exceptions.ConnectionError:
        return {'success': False, 'error': 'Cannot connect to Ollama. Make sure Ollama is running (ollama serve).'}
    except Exception as e:
        return {'success': False, 'error': str(e)}

def get_data_for_ai_query(query_type, filters=None):
    """Get relevant data based on query type for AI analysis"""
    data = {}

    # Find the most recent job result with data
    for job_id, result in job_results.items():
        if isinstance(result, dict):
            if query_type == 'workbooks' and 'workbook_analytics' in result:
                data['workbooks'] = result['workbook_analytics'][:50]  # Limit for context
            elif query_type == 'users' and 'user_analytics' in result:
                data['users'] = result['user_analytics'][:50]
            elif query_type == 'datasources' and 'datasource_analytics' in result:
                data['datasources'] = result['datasource_analytics'][:50]
            elif query_type == 'summary' and 'kpi_summary' in result:
                data['kpi_summary'] = result['kpi_summary']
            elif 'inventory' in result:
                data['inventory'] = result['inventory'][:50]

    return data


@app.route('/api/ai/status', methods=['GET'])
def get_ai_status():
    """Check AI availability and status - supports multiple providers"""
    # Check Ollama status
    ollama_status = check_ollama_status()

    # Determine overall availability
    available = (
        ollama_status.get('available', False) or
        bool(AI_CONFIG['claude_api_key']) or
        bool(AI_CONFIG['openai_api_key'])
    )

    # Determine current active model based on provider
    if AI_CONFIG['provider'] == 'claude' and AI_CONFIG['claude_api_key']:
        current_model = AI_CONFIG['claude_model']
    elif AI_CONFIG['provider'] == 'openai' and AI_CONFIG['openai_api_key']:
        current_model = AI_CONFIG['openai_model']
    else:
        current_model = AI_CONFIG['ollama_model']

    status = {
        'available': available,
        'current_model': current_model,
        'provider': AI_CONFIG['provider'],
        'ollama_available': ollama_status.get('available', False),
        'ollama_models': ollama_status.get('models', []),
        'claude_configured': bool(AI_CONFIG['claude_api_key']),
        'openai_configured': bool(AI_CONFIG['openai_api_key']),
        'config': {
            'enabled': AI_CONFIG['enabled'],
            'provider': AI_CONFIG['provider'],
            'model': current_model,
            'ollama_url': AI_CONFIG['ollama_url']
        }
    }
    return jsonify(status)


@app.route('/api/ai/models', methods=['GET'])
def get_ai_models():
    """Get list of available Ollama models"""
    status = check_ollama_status()
    return jsonify({
        'models': status.get('models', []),
        'current': AI_CONFIG['model']
    })


@app.route('/api/ai/config', methods=['POST'])
def update_ai_config():
    """Update AI configuration - supports Claude, OpenAI, and Ollama"""
    data = request.json

    if 'provider' in data:
        AI_CONFIG['provider'] = data['provider']
    if 'model' in data:
        AI_CONFIG['model'] = data['model']
        # Also update provider-specific model
        if AI_CONFIG['provider'] == 'claude':
            AI_CONFIG['claude_model'] = data['model']
        elif AI_CONFIG['provider'] == 'openai':
            AI_CONFIG['openai_model'] = data['model']
        else:
            AI_CONFIG['ollama_model'] = data['model']
    if 'api_key' in data:
        # Route API key to correct provider
        if AI_CONFIG['provider'] == 'claude':
            AI_CONFIG['claude_api_key'] = data['api_key']
        elif AI_CONFIG['provider'] == 'openai':
            AI_CONFIG['openai_api_key'] = data['api_key']
    if 'claude_api_key' in data:
        AI_CONFIG['claude_api_key'] = data['claude_api_key']
    if 'openai_api_key' in data:
        AI_CONFIG['openai_api_key'] = data['openai_api_key']
    if 'ollama_url' in data:
        AI_CONFIG['ollama_url'] = data['ollama_url']
    if 'enabled' in data:
        AI_CONFIG['enabled'] = data['enabled']

    # Return sanitized config (hide API keys)
    safe_config = {
        'provider': AI_CONFIG['provider'],
        'model': AI_CONFIG['model'],
        'enabled': AI_CONFIG['enabled'],
        'ollama_url': AI_CONFIG['ollama_url'],
        'has_claude_key': bool(AI_CONFIG['claude_api_key']),
        'has_openai_key': bool(AI_CONFIG['openai_api_key'])
    }
    return jsonify({'success': True, 'config': safe_config})


@app.route('/api/ai/chat', methods=['POST'])
def ai_chat():
    """Chat with AI assistant"""
    if not AI_CONFIG['enabled']:
        return jsonify({'success': False, 'error': 'AI is disabled'})

    data = request.json
    user_message = data.get('message', '').strip()
    conversation_id = data.get('conversation_id', str(uuid.uuid4()))
    frontend_context = data.get('context', '')  # Context sent from frontend

    if not user_message:
        return jsonify({'success': False, 'error': 'No message provided'})

    # Get or create conversation history
    if conversation_id not in ai_conversations:
        ai_conversations[conversation_id] = []

    # Build context - combine backend and frontend context
    context = get_context_summary()

    # Parse frontend context if provided
    if frontend_context:
        try:
            fe_ctx = json.loads(frontend_context) if isinstance(frontend_context, str) else frontend_context
            if isinstance(fe_ctx, dict):
                context['frontend_stats'] = fe_ctx
                context['workbook_count'] = fe_ctx.get('workbooks', context.get('workbook_count', 0))
                context['datasource_count'] = fe_ctx.get('datasources', context.get('datasource_count', 0))
                context['project_count'] = fe_ctx.get('projects', 0)
                context['user_count'] = fe_ctx.get('users', context.get('user_count', 0))
                context['view_count'] = fe_ctx.get('views', 0)
                context['has_extraction'] = True  # We have data from frontend
        except (json.JSONDecodeError, TypeError):
            pass

    system_prompt = build_system_prompt_v2(context)

    # Build messages array
    messages = [{'role': 'system', 'content': system_prompt}]

    # Add conversation history (last 10 messages for context)
    messages.extend(ai_conversations[conversation_id][-10:])

    # Add current user message
    messages.append({'role': 'user', 'content': user_message})

    # Query AI (routes to Claude, OpenAI, or Ollama based on config)
    result = query_ai(messages)

    if result['success']:
        # Store in conversation history
        ai_conversations[conversation_id].append({'role': 'user', 'content': user_message})
        ai_conversations[conversation_id].append({'role': 'assistant', 'content': result['response']})

        return jsonify({
            'success': True,
            'response': result['response'],
            'conversation_id': conversation_id,
            'model': result.get('model')
        })
    else:
        return jsonify({
            'success': False,
            'error': result.get('error', 'Unknown error')
        })


@app.route('/api/ai/chat/clear', methods=['POST'])
def clear_ai_chat():
    """Clear conversation history"""
    data = request.json
    conversation_id = data.get('conversation_id')

    if conversation_id and conversation_id in ai_conversations:
        del ai_conversations[conversation_id]

    return jsonify({'success': True})


@app.route('/api/ai/insights', methods=['POST'])
def get_ai_insights():
    """Generate AI-powered governance insights"""
    if not AI_CONFIG['enabled']:
        return jsonify({'success': False, 'error': 'AI is disabled'})

    data = request.json
    insight_type = data.get('type', 'general')  # general, sunset, license, security

    # Get relevant data
    context = get_context_summary()

    if not context.get('has_extraction') and not context.get('has_analytics'):
        return jsonify({
            'success': False,
            'error': 'No data available. Please run Deep Extraction or upload Admin Insights first.'
        })

    # Build insight-specific prompt
    prompts = {
        'general': """Analyze the Tableau environment and provide 5 key governance insights covering:
1. Content health overview
2. Usage patterns
3. Potential issues or risks
4. Quick wins for optimization
5. Recommended next steps

Be specific and actionable.""",

        'sunset': """Identify content that should be considered for retirement (sunset candidates). Look for:
1. Content not accessed in 90+ days
2. Low view counts
3. No recent updates
4. Orphaned content (owner inactive)

Provide specific recommendations with reasoning.""",

        'license': """Analyze license utilization and provide recommendations:
1. Identify potential license waste (inactive Creator/Explorer users)
2. Suggest license type optimizations
3. Flag users who could be downgraded
4. Calculate potential cost savings

Be specific with numbers when available.""",

        'security': """Perform a security-focused analysis:
1. Identify admin accounts that haven't logged in recently
2. Flag content with potentially sensitive data (look for patterns in names)
3. Check for overly broad permissions
4. Identify external data connections

Highlight risks and remediation steps."""
    }

    prompt = prompts.get(insight_type, prompts['general'])

    # Add data context to prompt
    analytics_data = get_data_for_ai_query('summary')
    if analytics_data:
        prompt += f"\n\nAvailable metrics: {json.dumps(analytics_data, default=str)[:2000]}"

    messages = [
        {'role': 'system', 'content': build_system_prompt(context)},
        {'role': 'user', 'content': prompt}
    ]

    result = query_ollama(messages)

    if result['success']:
        return jsonify({
            'success': True,
            'insights': result['response'],
            'type': insight_type,
            'model': result.get('model')
        })
    else:
        return jsonify({
            'success': False,
            'error': result.get('error')
        })


@app.route('/api/ai/search', methods=['POST'])
def ai_natural_language_search():
    """Convert natural language to search filters"""
    if not AI_CONFIG['enabled']:
        return jsonify({'success': False, 'error': 'AI is disabled'})

    data = request.json
    query = data.get('query', '').strip()

    if not query:
        return jsonify({'success': False, 'error': 'No query provided'})

    # Prompt to convert natural language to filters
    prompt = f"""Convert this natural language search into structured filters for a Tableau metadata search.

User query: "{query}"

Return a JSON object with these possible filter fields:
- content_type: "Workbook", "Data Source", "Flow", "Project", or null for all
- owner: owner name/email pattern or null
- project: project name pattern or null
- search_text: text to search in names
- days_inactive: number of days since last access (for dormant content)
- tier: 1, 2, or 3 for criticality tier
- has_custom_sql: true/false
- has_calculations: true/false

Only include fields that are relevant to the query. Return ONLY valid JSON, no explanation.

Example:
Query: "Show me workbooks not used in 90 days owned by John"
Response: {{"content_type": "Workbook", "days_inactive": 90, "owner": "John"}}"""

    messages = [
        {'role': 'system', 'content': 'You are a query parser. Return only valid JSON, no markdown or explanation.'},
        {'role': 'user', 'content': prompt}
    ]

    result = query_ollama(messages)

    if result['success']:
        try:
            # Try to parse the response as JSON
            response_text = result['response'].strip()
            # Remove markdown code blocks if present
            if response_text.startswith('```'):
                response_text = response_text.split('```')[1]
                if response_text.startswith('json'):
                    response_text = response_text[4:]

            filters = json.loads(response_text)

            return jsonify({
                'success': True,
                'filters': filters,
                'original_query': query,
                'model': result.get('model')
            })
        except json.JSONDecodeError:
            # If parsing fails, return the raw response
            return jsonify({
                'success': True,
                'filters': {'search_text': query},
                'original_query': query,
                'parse_error': 'Could not parse AI response as filters',
                'raw_response': result['response']
            })
    else:
        return jsonify({
            'success': False,
            'error': result.get('error')
        })


@app.route('/api/ai/explain', methods=['POST'])
def ai_explain_content():
    """Get AI explanation for specific content"""
    if not AI_CONFIG['enabled']:
        return jsonify({'success': False, 'error': 'AI is disabled'})

    data = request.json
    content_type = data.get('type', 'workbook')
    content_data = data.get('data', {})

    if not content_data:
        return jsonify({'success': False, 'error': 'No content data provided'})

    prompt = f"""Analyze this Tableau {content_type} and provide insights:

{json.dumps(content_data, indent=2, default=str)[:3000]}

Provide:
1. Brief summary of what this content appears to be
2. Usage assessment (active, at-risk, dormant)
3. Governance recommendations
4. Any potential issues or concerns

Keep response concise and actionable."""

    context = get_context_summary()
    messages = [
        {'role': 'system', 'content': build_system_prompt(context)},
        {'role': 'user', 'content': prompt}
    ]

    result = query_ollama(messages)

    if result['success']:
        return jsonify({
            'success': True,
            'explanation': result['response'],
            'content_type': content_type,
            'model': result.get('model')
        })
    else:
        return jsonify({
            'success': False,
            'error': result.get('error')
        })


if __name__ == '__main__':
    app.run(debug=True, port=5001)
