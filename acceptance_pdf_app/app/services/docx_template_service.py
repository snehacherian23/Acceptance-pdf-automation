"""
app/services/docx_template_service.py
----------------------------------------
Generates one Acceptance Test Word document per Excel row, reproducing
the official Omantel Acceptance Test Procedure template as closely as
possible with python-docx: ONE continuous master table (banner through
signatures), matching section header styling, checkbox formatting,
the banded Customer Experience column (with real icon images), and a
uniform outer border.

OUTER BORDER: set via the table's tblBorders (table-level), the
correct OOXML mechanism for a guaranteed-uniform frame around a table
regardless of internal merges -- NOT approximated by overriding
individual cell borders, which can render inconsistently (confirmed
by testing: this was the actual bug in the previous version). Internal
cells explicitly set only their INNER-facing sides; perimeter-facing
sides are left unset so the table-level border shows through
consistently on all four sides.

CUSTOMER EXPERIENCE ICONS: the official template uses small icon
graphics (smiley, thumbs-up, neutral face), not Unicode emoji. This
version embeds the actual PNG assets (app/assets/icons/), background
removed, inline with the text via Run.add_picture().

Calibrated directly against the official reference PDF -- this is a
reproduction, not a redesign.

DATA vs. TEMPLATE DEFAULTS: the field grid (Customer Name, Account
Number, Bandwidth, etc.) is populated dynamically from the Excel row.
The Test & Result / Compliance / Customer Experience checkmarks are
NOT supplied by Excel -- they're pre-filled with the same default
values the official template itself ships with (static per the
reference, not per-row).
"""

import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# --- Brand / template colors (calibrated against the reference PDF) ---
BANNER_BLUE_HEX = "2E4A9E"
SECTION_HEADER_GREY_HEX = "BFBFBF"   # light grey bars: "Test and Result", etc.
BORDER_GREY_HEX = "808080"           # thin internal grid lines
BORDER_BLACK_HEX = "000000"          # thick outer perimeter border

INNER_BORDER_SIZE = 4    # 0.5pt -- internal grid lines
OUTER_BORDER_SIZE = 16   # 2pt -- outer perimeter, uniform on all 4 sides, thicker than internal grid but not heavy

# --- Checkbox glyphs used for the pre-filled test/compliance defaults ---
CHECKED = "\u2611"    # ☑
UNCHECKED = "\u2610"  # ☐

# --- Column widths for the shared 4-column grid (total ~17cm, fits A4 with margins) ---
COL_WIDTHS = [Cm(3.4), Cm(5.2), Cm(3.6), Cm(4.8)]

# --- Customer Experience icon assets ---
ICONS_DIR = os.path.join(os.path.dirname(__file__), "..", "assets", "icons")
ICON_HEIGHT = Cm(0.4)  # small, inline with text -- matches the reference's icon scale


# ---------------------------------------------------------------------
# Low-level table styling helpers
# ---------------------------------------------------------------------
def _set_table_outer_borders(table, color=BORDER_BLACK_HEX, size=OUTER_BORDER_SIZE,
                              inside_color=BORDER_GREY_HEX, inside_size=INNER_BORDER_SIZE):
    """
    Set the table's outer frame via tblBorders (table-level) -- the
    correct OOXML mechanism for a guaranteed-uniform border on all
    four sides regardless of internal cell merges. insideH/insideV
    are also set here as the table's default internal grid, though
    individual cells still set their own explicit inner borders too.
    """
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    for edge in ("insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(inside_size))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), inside_color)
        borders.append(el)
    tblPr.append(borders)


def _set_cell_borders(cell, color=BORDER_GREY_HEX, size=INNER_BORDER_SIZE, skip=()):
    """
    Set thin internal borders on a cell, explicitly SKIPPING any side
    listed in `skip` (typically because that side lies on the table's
    outer perimeter and should inherit the table-level border instead
    of being locally overridden -- leaving it unset, rather than
    setting it to nil or duplicating the outer spec, is what keeps the
    perimeter uniform).
    """
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        if edge in skip:
            continue
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tcPr.append(borders)


def _set_spacer_borders(cell, hide_top=True, hide_bottom=True):
    """
    Spacer rows: explicitly hide the internal top/bottom divider (nil,
    overriding the table's default thin grid line) to create genuine
    blank breathing room -- but left/right are left UNSET so the
    table's outer border still renders correctly on the true perimeter
    for that row.
    """
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    if hide_top:
        el = OxmlElement("w:top")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    if hide_bottom:
        el = OxmlElement("w:bottom")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tcPr.append(borders)


def _set_cell_background(cell, hex_color):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


def _set_cell_margins(cell, top=15, bottom=15, left=80, right=80):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for edge, value in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:w"), str(value))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tcPr.append(mar)


def _set_cell_text(cell, text, bold=False, size=9, color=None, align=None, italic=False):
    cell.text = ""
    para = cell.paragraphs[0]
    if align:
        para.alignment = align
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    run = para.add_run(str(text) if text is not None else "")
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    return run


def _add_icon(paragraph, icon_name, height=ICON_HEIGHT):
    """Insert a small inline icon image (background already removed)
    into the given paragraph, matching the official template's use of
    real icon graphics rather than Unicode emoji."""
    icon_path = os.path.join(ICONS_DIR, f"{icon_name}.png")
    run = paragraph.add_run(" ")
    run.add_picture(icon_path, height=height)


def _row_cells(table, row_idx):
    return table.rows[row_idx].cells


def _hmerge(table, row_idx, col_start, col_end):
    cells = _row_cells(table, row_idx)
    merged = cells[col_start]
    for c in range(col_start + 1, col_end + 1):
        merged = merged.merge(cells[c])
    return merged


def _vmerge(table, row_start, row_end, col_idx):
    merged = _row_cells(table, row_start)[col_idx]
    for r in range(row_start + 1, row_end + 1):
        merged = merged.merge(_row_cells(table, r)[col_idx])
    return merged


def _row_border_skip(row_idx, total_rows, is_leftmost, is_rightmost):
    """Which sides of a cell should be skipped (left unset, inheriting
    the table-level outer border) because they lie on the true
    perimeter of the whole table."""
    skip = set()
    if row_idx == 0:
        skip.add("top")
    if row_idx == total_rows - 1:
        skip.add("bottom")
    if is_leftmost:
        skip.add("left")
    if is_rightmost:
        skip.add("right")
    return skip


def _full_width_row(table, row_idx, total_rows, text, bg_hex=None, bold=True, size=10,
                     color=None, align=WD_ALIGN_PARAGRAPH.LEFT, italic=False,
                     spacer=False):
    cell = _hmerge(table, row_idx, 0, len(COL_WIDTHS) - 1)
    if spacer:
        _set_spacer_borders(cell)
        table.rows[row_idx].height = Cm(0.15)
        table.rows[row_idx].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    else:
        skip = _row_border_skip(row_idx, total_rows, is_leftmost=True, is_rightmost=True)
        _set_cell_borders(cell, skip=skip)
    if bg_hex:
        _set_cell_background(cell, bg_hex)
    _set_cell_margins(cell)
    # IMPORTANT: always explicitly format the cell's paragraph, even when
    # there's no text (spacer rows). Leaving a cell's default paragraph
    # untouched means it silently inherits Word's document-wide default
    # (11pt font, 10pt space-after, 1.15x line height -- roughly 0.8cm of
    # unwanted height per row). Confirmed by inspecting a blank document's
    # styles.xml docDefaults. This was adding ~3.2cm across the 4 spacer
    # rows alone, a real and significant contributor to page overflow.
    _set_cell_text(cell, text or "", bold=bold, size=(3 if spacer else size),
                    color=color, align=align, italic=italic)
    return cell


def _label_value_row(table, row_idx, total_rows, label1, value1, label2, value2, size=9):
    """The recurring 'label | value | label | value' row -- labels are
    bold text on plain white (the reference template does not shade
    these cells; shading is reserved for the section header bars)."""
    cells = _row_cells(table, row_idx)
    _set_cell_text(cells[0], label1, bold=True, size=size)
    _set_cell_text(cells[1], value1, size=size)
    _set_cell_text(cells[2], label2, bold=True, size=size)
    _set_cell_text(cells[3], value2, size=size)
    n = len(cells)
    for i, c in enumerate(cells):
        skip = _row_border_skip(row_idx, total_rows, is_leftmost=(i == 0), is_rightmost=(i == n - 1))
        _set_cell_borders(c, skip=skip)
        _set_cell_margins(c)


# ---------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------
def generate_acceptance_document(row_data, output_path):
    """
    Build one Acceptance Test .docx from a single row's field data
    (as produced by excel_service.read_rows) and save it to output_path.

    Returns output_path.
    """
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.3)
    section.right_margin = Cm(1.3)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.0)

    # --- Row plan (ONE continuous table, matching the official layout) ---
    # 0        : banner
    # 1-8      : field grid (8 rows)
    # 9        : spacer
    # 10       : "Test and Result" heading
    # 11-15    : test/result rows (5)
    # 16       : spacer
    # 17       : compliance/experience split heading
    # 18-26    : compliance rows (9), Customer Experience vertically merged
    # 27       : spacer
    # 28       : "Customer Service Approval" heading
    # 29       : Customer Name: / Technician Name:
    # 30       : Date:
    # 31       : Signature / Signature  (LAST ROW of the table)
    # Remarks is a separate paragraph AFTER the table -- the reference
    # template shows it centered below the bordered box, not inside it.
    TOTAL_ROWS = 32
    LAST_ROW = TOTAL_ROWS - 1

    table = doc.add_table(rows=TOTAL_ROWS, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for col_idx, width in enumerate(COL_WIDTHS):
        for row in table.rows:
            row.cells[col_idx].width = width
    for row in table.rows:
        row.height = Cm(0.4)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

    _set_table_outer_borders(table)

    # --- Banner (row 0) ---
    banner_cell = _full_width_row(table, 0, TOTAL_ROWS, None, bg_hex=BANNER_BLUE_HEX)
    _set_cell_margins(banner_cell, top=50, bottom=50, left=80, right=80)
    banner_cell.text = ""
    p1 = banner_cell.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_before = Pt(0)
    p1.paragraph_format.space_after = Pt(2)
    r1 = p1.add_run("Omantel")
    r1.bold = True
    r1.font.size = Pt(24)
    r1.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p2 = banner_cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    r2 = p2.add_run("Acceptance Test Procedure")
    r2.bold = True
    r2.font.size = Pt(12)
    r2.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # --- Field grid (rows 1-8) -- populated dynamically from the Excel row ---
    grid_rows = [
        ("Customer Name", "customer_name", "Account Number:", "account_number"),
        ("Site Name", "site_name", "Customer Contact NO:", "customer_contact_no"),
        ("Account Manager", "account_manager", "Work Order Type:", "work_order_type"),
        ("Project Manager", "project_manager", "Bandwidth:", "bandwidth"),
        ("Installation Date", "installation_date", "Solution:", "solution"),
        ("Service Type", "service_type", "Barcode:", "barcode"),
        ("Exchange/Port Details", "exchange_port_details", "CPE Serial Number:", "cpe_serial_number"),
        ("Router Type", "router_type", "Property ID:", "property_id"),
    ]
    for i, (label1, key1, label2, key2) in enumerate(grid_rows):
        _label_value_row(table, 1 + i, TOTAL_ROWS, label1, row_data.get(key1, ""), label2, row_data.get(key2, ""))

    # --- Spacer (row 9) ---
    _full_width_row(table, 9, TOTAL_ROWS, None, spacer=True)

    # --- Test and Result heading (row 10) ---
    _full_width_row(table, 10, TOTAL_ROWS, "Test and Result", bg_hex=SECTION_HEADER_GREY_HEX,
                     color=RGBColor(0x00, 0x00, 0x00), size=10)

    # --- Test/Result rows (11-15) -- 5 rows, matching the official layout ---
    test_rows = [
        (f"Head office Overall Status ({CHECKED}Pass/{UNCHECKED}Fail/{UNCHECKED}NA)",
         f"Actual Result ({CHECKED}Reachable / {UNCHECKED}Not Reachable)"),
        (f"Branch office Overall Status ({CHECKED}Pass/{UNCHECKED}Fail/{UNCHECKED}NA)",
         f"Actual Result ({CHECKED}Reachable / {UNCHECKED}Not Reachable)"),
        (f"Perform Web browsing- VPN ({CHECKED}Pass/{UNCHECKED}Fail/{UNCHECKED}NA)",
         f"Ping test to IP outside the sites-VPN ({CHECKED}Pass /{UNCHECKED}Fail)"),
        ("Connectivity and performance Test",
         "Alternate DNS Server test"),
        (f"({UNCHECKED}Pass{CHECKED}/{UNCHECKED}Fail/ {UNCHECKED}NA)",
         f"({CHECKED}Pass/{UNCHECKED}Fail/ {UNCHECKED}NA)"),
    ]
    for i, (left, right) in enumerate(test_rows):
        row_idx = 11 + i
        left_cell = _hmerge(table, row_idx, 0, 1)
        right_cell = _hmerge(table, row_idx, 2, 3)
        _set_cell_text(left_cell, left, size=8)
        _set_cell_text(right_cell, right, size=8)
        skip_l = _row_border_skip(row_idx, TOTAL_ROWS, is_leftmost=True, is_rightmost=False)
        skip_r = _row_border_skip(row_idx, TOTAL_ROWS, is_leftmost=False, is_rightmost=True)
        _set_cell_borders(left_cell, skip=skip_l)
        _set_cell_borders(right_cell, skip=skip_r)
        _set_cell_margins(left_cell)
        _set_cell_margins(right_cell)

    # --- Spacer (row 16) ---
    _full_width_row(table, 16, TOTAL_ROWS, None, spacer=True)

    # --- Compliance / Customer Experience split heading (row 17) ---
    compliance_header = _hmerge(table, 17, 0, 2)
    experience_header = _row_cells(table, 17)[3]
    for cell, text, leftmost, rightmost in (
        (compliance_header, "Customer Requirement /Compliance", True, False),
        (experience_header, "Customer Experience", False, True),
    ):
        skip = _row_border_skip(17, TOTAL_ROWS, is_leftmost=leftmost, is_rightmost=rightmost)
        _set_cell_background(cell, SECTION_HEADER_GREY_HEX)
        _set_cell_borders(cell, skip=skip)
        _set_cell_margins(cell)
        _set_cell_text(cell, text, bold=True, size=10, color=RGBColor(0x00, 0x00, 0x00))

    # --- Compliance rows (18-26) -- pre-filled Yes defaults + banded Customer Experience ---
    compliance_items = [
        "Is The Rack Condition Good?",
        "Is The Router Connected to Ups Power?",
        "The service was activated as scheduled.",
        "Does the Air Condition available in IT Room",
        "The service has been reliable since installation.",
        "The overall quality of the service meets my expectations.",
        "The Technician acted professionally.",
        "The Technician provided a clear explanation for the service.",
        "I am satisfied with the Behavior of the technician.",
    ]
    yes_no_default = f"{CHECKED} Yes      {UNCHECKED} NO"

    for i, item in enumerate(compliance_items):
        row_idx = 18 + i
        item_cell = _hmerge(table, row_idx, 0, 1)
        checkbox_cell = _row_cells(table, row_idx)[2]
        _set_cell_text(item_cell, item, size=8)
        _set_cell_text(checkbox_cell, yes_no_default, size=8)
        skip_item = _row_border_skip(row_idx, TOTAL_ROWS, is_leftmost=True, is_rightmost=False)
        skip_box = _row_border_skip(row_idx, TOTAL_ROWS, is_leftmost=False, is_rightmost=False)
        _set_cell_borders(item_cell, skip=skip_item)
        _set_cell_borders(checkbox_cell, skip=skip_box)
        _set_cell_margins(item_cell)
        _set_cell_margins(checkbox_cell)

    # Customer Experience column: banded exactly like the reference,
    # using real icon images (not Unicode emoji) inline with the text.
    experience_bands = [
        (18, 19, "Excellent", ["excellent_smiley", "excellent_thumbsup"]),
        (20, 21, "Satisfied", ["satisfied_smiley"]),
        (22, 23, "Fair", ["fair_smiley"]),
        (24, 26, "UNSATISFIED", []),
    ]
    for start, end, label, icons in experience_bands:
        cell = _vmerge(table, start, end, 3)
        skip = _row_border_skip(start, TOTAL_ROWS, is_leftmost=False, is_rightmost=True)
        _set_cell_borders(cell, skip=skip)
        _set_cell_margins(cell)
        cell.text = ""
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        run = para.add_run(f"\u2022{label}")
        run.bold = True
        run.font.size = Pt(8)
        for icon_name in icons:
            _add_icon(para, icon_name)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # --- Spacer (row 27) ---
    _full_width_row(table, 27, TOTAL_ROWS, None, spacer=True)

    # --- Customer Service Approval heading (row 28) ---
    _full_width_row(table, 28, TOTAL_ROWS, "Customer Service Approval", bg_hex=SECTION_HEADER_GREY_HEX,
                     color=RGBColor(0x00, 0x00, 0x00), size=10)

    # --- Approval fields (rows 29-30) ---
    cells29 = _row_cells(table, 29)
    _set_cell_text(cells29[0], "Customer Name:", bold=True, size=9)
    _set_cell_text(cells29[1], "", size=9)
    _set_cell_text(cells29[2], "Technician Name:", bold=True, size=9)
    _set_cell_text(cells29[3], row_data.get("technician_name", ""), size=9)
    n = len(cells29)
    for i, c in enumerate(cells29):
        skip = _row_border_skip(29, TOTAL_ROWS, is_leftmost=(i == 0), is_rightmost=(i == n - 1))
        _set_cell_borders(c, skip=skip)
        _set_cell_margins(c)

    cells30 = _row_cells(table, 30)
    _set_cell_text(cells30[0], "Date:", bold=True, size=9)
    date_value_cell = _hmerge(table, 30, 1, 3)
    _set_cell_text(date_value_cell, "", size=9)
    skip_label = _row_border_skip(30, TOTAL_ROWS, is_leftmost=True, is_rightmost=False)
    skip_value = _row_border_skip(30, TOTAL_ROWS, is_leftmost=False, is_rightmost=True)
    _set_cell_borders(cells30[0], skip=skip_label)
    _set_cell_borders(date_value_cell, skip=skip_value)
    _set_cell_margins(cells30[0])
    _set_cell_margins(date_value_cell)

    # --- Signature row (row 31 -- LAST ROW of the table) ---
    sig_left = _hmerge(table, 31, 0, 1)
    sig_right = _hmerge(table, 31, 2, 3)
    skip_sig_l = _row_border_skip(31, TOTAL_ROWS, is_leftmost=True, is_rightmost=False)
    skip_sig_r = _row_border_skip(31, TOTAL_ROWS, is_leftmost=False, is_rightmost=True)
    for cell, skip in ((sig_left, skip_sig_l), (sig_right, skip_sig_r)):
        _set_cell_text(cell, "Signature:", bold=True, size=9)
        cell.paragraphs[0].paragraph_format.space_before = Pt(10)
        _set_cell_borders(cell, skip=skip)
        _set_cell_margins(cell)

    assert LAST_ROW == 31, "row plan drifted -- update TOTAL_ROWS/comments together"

    # --- Remarks: a plain paragraph BELOW the table, matching the
    # reference (the disclaimer sits outside the bordered box) ---
    remarks = doc.add_paragraph()
    remarks.paragraph_format.space_before = Pt(6)
    remarks.paragraph_format.space_after = Pt(0)
    remarks_run = remarks.add_run(
        "*REMARKS: if customer not fulfilling above requirements / conditions, "
        "they will be responsible for CPE damage & outages Fault."
    )
    remarks_run.bold = True
    remarks_run.italic = True
    remarks_run.font.size = Pt(10)
    remarks_run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    remarks.alignment = WD_ALIGN_PARAGRAPH.CENTER

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    return output_path


def build_output_filename(row_data):
    """
    Filename convention: <AccountNumber>.docx, falling back to a row-
    number-based name if Account Number is missing (still generated,
    just flagged as missing that field in the report).
    """
    account_number = row_data.get("account_number", "")
    safe = "".join(ch for ch in str(account_number) if ch.isalnum())
    if safe:
        return f"{safe}.docx"
    return f"row_{row_data.get('_row_number', 'unknown')}.docx"
