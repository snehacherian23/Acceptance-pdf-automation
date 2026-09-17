"""
app/services/excel_service.py
-------------------------------
Reads the uploaded Excel sheet(s) and turns every row into a
normalized dict of Acceptance Test field values, ready to hand to
docx_template_service to generate one Word document per row.

The Excel sheet is now the SINGLE SOURCE OF TRUTH for the whole
workflow (no more extracting values from PDFs). Since the exact
column headers in a real production sheet can vary ("Account Number"
vs "A/C No" vs "Parent MPI ID", etc.), each target field is matched
against a list of acceptable keyword variants rather than requiring
an exact header string -- this is the same resilience approach used
throughout the project for real-world spreadsheet inconsistency.
"""

import re
import pandas as pd
from openpyxl import load_workbook


class ExcelLookupError(Exception):
    """Raised when an Excel file can't be read at all."""
    pass


def normalize_id(value):
    """
    Normalize an ID-like value (Account Number, Property ID, etc.) so
    Excel's mixed int/float/string cell storage doesn't leak formatting
    artifacts into the generated document.

    Handles:
      - Excel storing numeric IDs as floats (916024.0 -> "916024")
      - Leading/trailing whitespace or stray newlines
      - Thousands-separator commas (916,024 -> 916024)
      - A leading apostrophe some spreadsheets use to force text
        formatting on a numeric-looking cell ('916024 -> 916024)
      - Non-breaking spaces / other invisible whitespace variants
    """
    if value is None:
        return ""
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""

    text = text.replace("\xa0", " ").strip()

    if text.startswith("'"):
        text = text[1:]

    text = text.replace(",", "")

    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".")[0]

    return text.strip()


# --- Field mapping: target field -> acceptable header keyword variants ----
# Each target field is matched against the FIRST Excel column whose header
# contains ANY of these keywords (case-insensitive substring match).
# Order matters within each list only as a tie-breaker.
FIELD_KEYWORDS = {
    "customer_name":        ["customer name", "customer"],
    "account_number":       ["account number", "account no", "a/c no", "ac no", "parent mpi id", "mpi id"],
    "site_name":            ["site name", "location name"],
    "customer_contact_no":  ["customer contact", "contact no", "contact number"],
    "account_manager":      ["account manager"],
    "work_order_type":      ["work order type", "work order"],
    "project_manager":      ["project manager"],
    "bandwidth":            ["bandwidth", "b/w"],
    "installation_date":    ["installation date", "install date"],
    "solution":             ["solution"],
    "service_type":         ["service type", "service"],
    "barcode":               ["barcode"],
    "exchange_port_details": ["exchange/port", "exchange port", "port details", "exchange"],
    "cpe_serial_number":    ["cpe serial", "serial number", "serial no"],
    "router_type":           ["router type"],
    "property_id":           ["property id", "property id/code", "property code"],
    "technician_name":      ["technician"],
}

# Human-readable labels, used in the report and in logs.
FIELD_LABELS = {
    "customer_name": "Customer Name",
    "account_number": "Account Number",
    "site_name": "Site Name",
    "customer_contact_no": "Customer Contact NO",
    "account_manager": "Account Manager",
    "work_order_type": "Work Order Type",
    "project_manager": "Project Manager",
    "bandwidth": "Bandwidth",
    "installation_date": "Installation Date",
    "solution": "Solution",
    "service_type": "Service Type",
    "barcode": "Barcode",
    "exchange_port_details": "Exchange/Port Details",
    "cpe_serial_number": "CPE Serial Number",
    "router_type": "Router Type",
    "property_id": "Property ID",
    "technician_name": "Technician Name",
}


def _detect_header_row(excel_path, sheet_name=0, max_scan_rows=20):
    """
    Find the real header row instead of assuming it's the first row.

    Real-world sheets sometimes have a row of generic placeholders
    ("Column1", "Column2", ...) above the actual headers ("SERVICE",
    "CUSTOMER NAME", "AC No", ...). Scanning blind and always taking
    row 0 would silently misread the whole sheet in that case.

    Approach: read the first `max_scan_rows` rows with no header
    assumption, then score each row by how many of its cells contain
    one of our known field keywords (the same keyword list used for
    column mapping). The row with the highest score is almost
    certainly the real header row -- a placeholder row like "Column1,
    Column2..." will score 0, while a real header row will match
    several known fields at once.

    Returns:
        0-indexed row number (within the sheet) to use as the header.
        Falls back to 0 (the first row) if no row scores well enough
        to be confident -- preserves the original behavior for sheets
        that genuinely do have headers in row 1.
    """
    try:
        raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, nrows=max_scan_rows)
    except Exception:
        return 0

    all_keywords = [kw for kws in FIELD_KEYWORDS.values() for kw in kws]

    best_row = 0
    best_score = 0
    for i in range(len(raw)):
        cells = [str(v).strip().lower() for v in raw.iloc[i].tolist() if pd.notna(v) and str(v).strip()]
        score = sum(1 for cell in cells if any(kw in cell for kw in all_keywords))
        if score > best_score:
            best_score = score
            best_row = i

    # Require at least 2 recognized fields to be confident this is a
    # real header row, not a coincidental partial match in a data row.
    return best_row if best_score >= 2 else 0


def _map_columns(columns):
    """
    Build {target_field: actual_excel_column_name_or_None} for one
    sheet's header row.
    """
    mapping = {}
    for field, keywords in FIELD_KEYWORDS.items():
        found = None
        for col in columns:
            col_lower = str(col).lower()
            if any(kw in col_lower for kw in keywords):
                found = col
                break
        mapping[field] = found
    return mapping


def find_bordered_region(excel_path, sheet_name=0):
    """
    Detect a thick-bordered block of cells in the sheet, used to
    distinguish "new records" from historical ones when a single
    sheet contains both (a common real-world pattern: new rows are
    boxed with a thick black border to call them out visually).

    Uses openpyxl (not pandas) since border/style info isn't exposed
    through pandas' DataFrame -- only openpyxl reads cell-level
    formatting.

    Returns:
        (min_row, max_row) as 1-indexed Excel row numbers (inclusive,
        matching the same numbering used elsewhere: row_number =
        pandas_index + 2), or None if no thick border is found
        anywhere in the sheet.
    """
    wb = load_workbook(excel_path, data_only=True)
    ws = wb[sheet_name] if isinstance(sheet_name, str) else wb.worksheets[sheet_name]

    thick_rows = set()
    for row in ws.iter_rows():
        for cell in row:
            border = cell.border
            sides = (border.top, border.bottom, border.left, border.right)
            if any(side is not None and side.style == "thick" for side in sides):
                thick_rows.add(cell.row)

    if not thick_rows:
        return None
    return (min(thick_rows), max(thick_rows))


def read_rows(excel_path, filename=None, sheet_name=0):
    """
    Read every row of one Excel file into a normalized field dict.

    If the sheet contains a thick-bordered block (see
    find_bordered_region), ONLY the rows inside that block are
    returned -- this is how a single sheet mixing historical and new
    records is handled: new records are boxed, and only those get
    processed. If no bordered block is found, every row is returned
    (the common case: the sheet contains only current/new records).

    Rows are never marked as failed for having blank optional fields
    (Property ID, Bandwidth, Installation Date, Router Type, etc.) --
    those are simply left blank in the generated document. A row is
    only skipped if it's entirely empty (no data in any mapped field
    at all -- a stray spacer row, not a real record).

    Returns:
        dict with:
            "rows": [ {field_dicts + "_source_file", "_row_number", "_blank_fields"} ... ]
            "column_mapping": {field: excel_column_or_None}
            "row_count": int
            "logs": [str, ...]
    """
    filename = filename or excel_path
    logs = []

    header_row = _detect_header_row(excel_path, sheet_name)
    if header_row > 0:
        logs.append(f"'{filename}': detected the real header row at Excel row {header_row + 1} (row(s) above it appear to be placeholders/titles and were skipped).")

    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header_row)
    except Exception as exc:
        raise ExcelLookupError(f"Could not read '{filename}': {exc}")

    mapping = _map_columns(df.columns)
    matched = [f for f, c in mapping.items() if c is not None]
    unmatched = [f for f, c in mapping.items() if c is None]

    logs.append(f"'{filename}': matched {len(matched)}/{len(FIELD_KEYWORDS)} known fields to columns.")
    if unmatched:
        readable = ", ".join(FIELD_LABELS[f] for f in unmatched)
        logs.append(f"'{filename}': no column found for: {readable} -- these will be left blank.")

    # --- Detect a thick-bordered "new records" block, if any ------------
    try:
        region = find_bordered_region(excel_path, sheet_name)
    except Exception as exc:
        region = None
        logs.append(f"'{filename}': could not check for a bordered region ({exc}); processing all rows.")

    if region:
        logs.append(f"'{filename}': detected a thick-bordered block at Excel rows {region[0]}-{region[1]} -- only these rows will be treated as new records.")
    else:
        logs.append(f"'{filename}': no bordered region found -- treating all rows as current/new records.")

    rows = []
    skipped_outside_region = 0
    skipped_blank = 0

    for idx, row in df.iterrows():
        row_number = idx + header_row + 2  # 1-indexed Excel row, relative to the detected header's real position

        if region and not (region[0] <= row_number <= region[1]):
            skipped_outside_region += 1
            continue

        values = {}
        for field, col in mapping.items():
            if col is None:
                values[field] = ""
                continue
            raw = row[col]
            if pd.isna(raw):
                values[field] = ""
            elif field in ("account_number", "property_id"):
                values[field] = normalize_id(raw)
            elif field == "installation_date":
                # Excel dates come through as Timestamps -- render as a plain date string
                values[field] = raw.strftime("%d/%m/%Y") if hasattr(raw, "strftime") else str(raw).strip()
            else:
                values[field] = str(raw).strip()

        if not any(values.values()):
            skipped_blank += 1
            continue  # entirely empty row -- not a real record

        blank_fields = [f for f, v in values.items() if not v]

        rows.append({
            **values,
            "_source_file": filename,
            "_row_number": row_number,
            "_blank_fields": blank_fields,
        })

    logs.append(f"'{filename}': read {len(rows)} data row(s).")
    if skipped_outside_region:
        logs.append(f"'{filename}': skipped {skipped_outside_region} historical row(s) outside the bordered region.")
    if skipped_blank:
        logs.append(f"'{filename}': skipped {skipped_blank} entirely blank row(s).")

    return {
        "rows": rows,
        "column_mapping": mapping,
        "row_count": len(rows),
        "logs": logs,
    }


def read_all_rows(excel_entries):
    """
    Read and combine rows from MULTIPLE Excel files (the dashboard
    supports uploading several sheets for one batch). Rows keep track
    of which file and row number they came from, for the report and
    for error messages.

    Args:
        excel_entries: list of (filename, path) tuples, in upload order.

    Returns:
        dict with "rows" (combined list across all files), "row_count", "logs".
    """
    all_rows = []
    logs = [f"Reading {len(excel_entries)} Excel sheet(s)."]

    for filename, path in excel_entries:
        result = read_rows(path, filename=filename)
        logs.extend(result["logs"])
        all_rows.extend(result["rows"])

    logs.append(f"Combined total: {len(all_rows)} row(s) across {len(excel_entries)} file(s).")

    return {
        "rows": all_rows,
        "row_count": len(all_rows),
        "logs": logs,
    }
