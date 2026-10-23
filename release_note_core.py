"""
release_note_core.py

Core logic for turning a excel file export (CSV or XLSX) into a
formatted .docx release note. Supports loading a pre-existing .docx template
file and substituting unique placeholders across paragraphs, tables, headers,
and footers:

  - {{VERSION}}
  - {{RELEASE_DATE}}
  - {{PREPARED_BY}}
  - {{REVIEWED_BY}}
  - {{DOCUMENT_ID}}
  - {{PROJECT_ID}}
  - {{PROJECT_NAME}}
  - {{CLIENT_NAME}}
  - {{REVISION_NO}}
  - {{LIST_OF_CHANGES}}

Screenshot images are left as a "[ Paste screenshot here ]" placeholder
under each ticket's detail table.
"""

import os
import re
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_BREAK
from docx.table import Table as DocxTable

REQUIRED_COLS = [
    "Ticket No", "Facility", "Module", "Category", "Sub Category",
    "Reported Date", "Title", "Description", "Solution",
    "Root Cause/Analysis", "Resolved Date", "Verify Status",
]

SUMMARY_HEADERS = [
    "No", "Ticket No", "Facility", "Module", "Reported Date",
    "Title", "Description", "Root Cause/Analysis", "Resolved Date", "Verify Status",
]

DETAIL_HEADERS = ["Ticket No", "Reporting Date", "Description", "Date Solved", "Remarks"]
SUMMARY_COLUMN_WIDTHS = [0.7, 0.9, 0.8, 0.8, 1.2, 1.5, 2.4, 1.4, 0.9, 0.8]


# ---------------------------------------------------------------------------
# Data loading & filtering
# ---------------------------------------------------------------------------

def load_export(path: str) -> pd.DataFrame:
    if path.lower().endswith((".xlsx", ".xlsm", ".xls")):
        df = pd.read_excel(path)
    else:
        df = pd.read_csv(path)
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Export is missing expected columns: {missing}")
    return df


def clean(val) -> str:
    if pd.isna(val):
        return ""
    text = str(val).strip()
    text = re.sub(r"\n\s*\n+", "\n", text)
    return text


def filter_tickets(df: pd.DataFrame, include_rejected: bool) -> pd.DataFrame:
    if not include_rejected and "Verify Status" in df.columns:
        df = df[df["Verify Status"].astype(str).str.strip().str.lower() != "qc rejected"]
    return df


def split_by_category(df: pd.DataFrame):
    categories = df["Category"].astype(str).str.strip().str.lower()
    enhancements = df[categories.str.contains(r"\bcr\b", case=False, na=False)]
    bugs = df[categories.str.contains(r"\bbug\b", case=False, na=False)]
    return enhancements, bugs


# ---------------------------------------------------------------------------
# Placeholder replacement logic
# ---------------------------------------------------------------------------

def replace_placeholders_in_paragraph(paragraph, replacements: dict):
    p_text = paragraph.text
    if not p_text or not any(k in p_text for k in replacements):
        return

    # Try replacing within individual runs first to retain font/style formatting
    for run in paragraph.runs:
        for k, v in replacements.items():
            if k in run.text:
                run.text = run.text.replace(k, str(v))

    # Fallback if a placeholder was split across multiple run objects by Word
    if any(k in paragraph.text for k in replacements):
        new_text = paragraph.text
        for k, v in replacements.items():
            new_text = new_text.replace(k, str(v))
        if paragraph.runs:
            paragraph.runs[0].text = new_text
            for run in paragraph.runs[1:]:
                run.text = ""
        else:
            paragraph.text = new_text


def replace_placeholders_in_table(table, replacements: dict):
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                replace_placeholders_in_paragraph(p, replacements)
            for inner_table in cell.tables:
                replace_placeholders_in_table(inner_table, replacements)


def replace_placeholders_in_doc(document: Document, replacements: dict):
    # Paragraphs in body
    for p in document.paragraphs:
        replace_placeholders_in_paragraph(p, replacements)

    # Tables in body
    for table in document.tables:
        replace_placeholders_in_table(table, replacements)

    # Headers & Footers across all sections
    for section in document.sections:
        hf_list = [
            getattr(section, 'header', None),
            getattr(section, 'first_page_header', None),
            getattr(section, 'even_page_header', None),
            getattr(section, 'footer', None),
            getattr(section, 'first_page_footer', None),
            getattr(section, 'even_page_footer', None),
        ]
        for hf in hf_list:
            if hf is not None:
                for p in hf.paragraphs:
                    replace_placeholders_in_paragraph(p, replacements)
                for table in hf.tables:
                    replace_placeholders_in_table(table, replacements)

    # Text boxes (shapes with text)
    try:
        from docx.oxml import parse_xml
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
              'v': 'urn:schemas-microsoft-com:vml',
              'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'}
        
        # Find all text boxes in the document body
        for shape_elem in document.element.body.xpath('.//v:textpath', namespaces=ns):
            if shape_elem.text:
                for k, v in replacements.items():
                    if k in shape_elem.text:
                        shape_elem.text = shape_elem.text.replace(k, str(v))
        
        # Find text in shape runs (alternative text box representation)
        for t_elem in document.element.body.xpath('.//w:t', namespaces=ns):
            if t_elem.text:
                for k, v in replacements.items():
                    if k in t_elem.text:
                        t_elem.text = t_elem.text.replace(k, str(v))
    except Exception:
        # If text box processing fails, continue without it
        pass


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------

def _bold_centered(document, text, size=16):
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(size)
    return p


def _page_break(document):
    p = document.add_paragraph()
    run = p.add_run()
    run.add_break(WD_BREAK.PAGE)


def _add_toc(document):
    paragraph = document.add_paragraph()
    run = paragraph.add_run()
    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')

    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = 'TOC \\o "1-3" \\h \\z \\u'

    fld_sep = OxmlElement('w:fldChar')
    fld_sep.set(qn('w:fldCharType'), 'separate')

    placeholder = OxmlElement('w:t')
    placeholder.text = "Right-click here and choose \"Update Field\" in Word to build the table of contents."

    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')

    r = run._r
    r.append(fld_begin)
    r.append(instr)
    r.append(fld_sep)
    r.append(placeholder)
    r.append(fld_end)


def _set_cell_text(cell, text, bold=False):
    cell.text = ""
    lines = (text or "").replace("\r\n", "\n").split("\n")
    p = cell.paragraphs[0]
    for i, line in enumerate(lines):
        if i > 0:
            p = cell.add_paragraph()
        run = p.add_run(line)
        run.bold = bold


def _set_table_column_width(table, widths):
    """Set column widths for a table. Widths is a list of widths in inches."""
    for i, width in enumerate(widths):
        for cell in table.columns[i].cells:
            cell.width = Inches(width)


def _mark_row_as_table_header(row):
    """Mark a row as a repeating header row in Word."""
    tr = row._tr
    tr_pr = tr.get_or_add_trPr()
    tbl_header = OxmlElement('w:tblHeader')
    tr_pr.append(tbl_header)


def _get_first_xml_child(element, path):
    values = element.xpath(path)
    return values[0] if values else None


def _get_section_available_width(section):
    """Return available page width in inches for the section margins."""
    try:
        page_width = section.page_width
        left_margin = section.left_margin
        right_margin = section.right_margin
        available = page_width - left_margin - right_margin
        return available.inches
    except Exception:
        sectPr = section._sectPr
        pgSz = _get_first_xml_child(sectPr, './w:pgSz')
        if pgSz is None:
            return 6.5
        pgW = int(pgSz.get(qn('w:w'), 12240))
        pgMar = _get_first_xml_child(sectPr, './w:pgMar')
        if pgMar is None:
            return (pgW / 1440.0) - 1.0
        left_child = _get_first_xml_child(pgMar, './w:left')
        right_child = _get_first_xml_child(pgMar, './w:right')
        left = int(left_child.get(qn('w:w'), 1440)) if left_child is not None else 1440
        right = int(right_child.get(qn('w:w'), 1440)) if right_child is not None else 1440
        # Word sizes are in twentieths of a point (twips): 1440 twips = 1 inch
        return max((pgW - left - right) / 1440.0, 4.0)


def _scale_column_widths(section, widths):
    """Scale column widths to fit the section's available width."""
    available = _get_section_available_width(section)
    total = sum(widths)
    if total <= 0:
        return widths
    factor = available / total
    return [w * factor for w in widths]


def _shade_cell(cell, hex_color="D9D9D9"):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), hex_color)
    tc_pr.append(shd)


def _add_table(document, headers, rows, shade_header=True, column_widths=None):
    table = document.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        _set_cell_text(hdr_cells[i], h, bold=True)
        if shade_header:
            _shade_cell(hdr_cells[i])
    _mark_row_as_table_header(table.rows[0])
    for row_values in rows:
        cells = table.add_row().cells
        for i, val in enumerate(row_values):
            _set_cell_text(cells[i], val)
    if column_widths:
        section = document.sections[-1]
        scaled_widths = _scale_column_widths(section, column_widths)
        _set_table_column_width(table, scaled_widths)
        table.autofit = False
    return table


def _find_and_replace_placeholder_with_table(document, placeholder, headers, rows, shade_header=True):
    for p in list(document.paragraphs):
        if placeholder in p.text:
            tbl_elem = document.add_table(rows=1, cols=len(headers))._element
            p._element.addprevious(tbl_elem)
            table = DocxTable(tbl_elem, document)
            table.style = "Table Grid"
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            hdr_cells = table.rows[0].cells
            for i, h in enumerate(headers):
                _set_cell_text(hdr_cells[i], h, bold=True)
                if shade_header:
                    _shade_cell(hdr_cells[i])
            _mark_row_as_table_header(table.rows[0])
            for row_values in rows:
                cells = table.add_row().cells
                for i, val in enumerate(row_values):
                    _set_cell_text(cells[i], val)
            p._element.getparent().remove(p._element)
            return True
    return False


def _screenshot_placeholder(document):
    p = document.add_paragraph()
    run = p.add_run("[ Paste screenshot here ]")
    run.italic = True
    run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)


def summary_rows(rows_df):
    rows = []
    for i, (_, r) in enumerate(rows_df.iterrows(), start=1):
        rows.append([
            str(i), clean(r.get("Ticket No")), clean(r.get("Facility")), clean(r.get("Module")),
            clean(r.get("Reported Date")), clean(r.get("Title")), clean(r.get("Description")),
            clean(r.get("Root Cause/Analysis")), clean(r.get("Resolved Date")), clean(r.get("Verify Status")),
        ])
    return rows


def detail_sections(document, rows_df):
    # Column widths for Resolution - Result tables (in inches)
    detail_column_widths = [1.0, 1.2, 3.5, 1.2, 0.8]
    
    for _, r in rows_df.iterrows():
        document.add_heading(clean(r.get("Title")), level=3)
        _add_table(document, DETAIL_HEADERS, [[
            clean(r.get("Ticket No")), clean(r.get("Reported Date")),
            clean(r.get("Description")), clean(r.get("Resolved Date")), "",
        ]], column_widths=detail_column_widths)
        _screenshot_placeholder(document)
        _page_break(document)


def _set_section_to_landscape(section):
    """Set a section to landscape orientation."""
    sectPr = section._sectPr
    pgSz = sectPr.xpath('./w:pgSz')[0]
    
    # Swap width and height for landscape (typically 11" x 8.5" -> 11" x 8.5")
    width = pgSz.get(qn('w:w'))
    height = pgSz.get(qn('w:h'))
    
    if width and height:
        pgSz.set(qn('w:w'), height)
        pgSz.set(qn('w:h'), width)
    
    # Set orientation to landscape
    pgSz.set(qn('w:orient'), 'landscape')


def _set_section_to_portrait(section):
    """Set a section to portrait orientation."""
    sectPr = section._sectPr
    pgSz = sectPr.xpath('./w:pgSz')[0]
    
    # Swap width and height for portrait
    width = pgSz.get(qn('w:w'))
    height = pgSz.get(qn('w:h'))
    
    if width and height:
        pgSz.set(qn('w:w'), height)
        pgSz.set(qn('w:h'), width)
    
    # Set orientation to portrait
    pgSz.set(qn('w:orient'), 'portrait')


def _set_footer_text(document, text):
    """Set footer text in all sections of the document."""
    for section in document.sections:
        footer = section.footer
        if footer.paragraphs:
            footer.paragraphs[0].text = text
        else:
            footer.add_paragraph(text)


def _set_footer_in_section(section, text):
    """Set footer text in a specific section."""
    footer = section.footer
    if footer.paragraphs:
        footer.paragraphs[0].text = text
    else:
        footer.add_paragraph(text)


def _append_ticket_sections(document, enhancements, bugs, header):
    # Summary: Refinement / Technical Improvement
    document.add_heading("Summary \u2013 Refinement / Technical Improvement", level=1)
    document.add_heading("Refinement List", level=2)
    document.add_paragraph(
        f"The following are the summary of resolved refinement and technical improvement "
        f"list that included in {header.get('version_no','')} version release."
    )
    
    # Set current section to landscape for the refinement summary table
    current_section = document.sections[-1]
    _set_section_to_landscape(current_section)
    
    _add_table(document, SUMMARY_HEADERS, summary_rows(enhancements), column_widths=SUMMARY_COLUMN_WIDTHS)
    
    # Add new section and set back to portrait for Resolution - Result
    new_section = document.add_section()
    _set_section_to_portrait(new_section)
    _set_footer_in_section(new_section, header.get("document_id", ""))

    document.add_heading("Resolution - Result", level=1)
    document.add_heading("Resolution \u2013 Result", level=2)
    detail_sections(document, enhancements)

    # Add new section and set to landscape for Summary - Defect
    new_section = document.add_section()
    _set_section_to_landscape(new_section)
    _set_footer_in_section(new_section, header.get("document_id", ""))
    
    # Summary: Defect
    document.add_heading("Summary - Defect", level=1)
    document.add_heading("Defect List", level=2)
    document.add_paragraph(
        f"The following are the summary of resolved defect list that included in "
        f"{header.get('version_no','')} version release."
    )
    
    _add_table(document, SUMMARY_HEADERS, summary_rows(bugs), column_widths=SUMMARY_COLUMN_WIDTHS)

    # Add new section and set back to portrait for Resolution - Result
    new_section = document.add_section()
    _set_section_to_portrait(new_section)
    _set_footer_in_section(new_section, header.get("document_id", ""))
    
    document.add_heading("Resolution - Result", level=1)
    document.add_heading("Resolution \u2013 Result", level=2)
    detail_sections(document, bugs)


# ---------------------------------------------------------------------------
# Main generation
# ---------------------------------------------------------------------------

def generate_release_note(
    input_path: str,
    output_path: str,
    header: dict,
    include_rejected: bool = False,
    template_path: str = None,
):
    df = load_export(input_path)
    df = filter_tickets(df, include_rejected)
    enhancements, bugs = split_by_category(df)

    replacements = {
        "{{VERSION}}": header.get("version_no", ""),
        "{{VERSION_NO}}": header.get("version_no", ""),
        "{{RELEASE_DATE}}": header.get("release_date", ""),
        "{{PREPARED_BY}}": header.get("prepared_by", ""),
        "{{REVIEWED_BY}}": header.get("reviewed_by", ""),
        "{{DOCUMENT_ID}}": header.get("document_id", ""),
        "{{PROJECT_ID}}": header.get("project_id", ""),
        "{{PROJECT_NAME}}": header.get("project_name", ""),
        "{{CLIENT_NAME}}": header.get("client_name", ""),
        "{{REVISION_NO}}": header.get("revision_no", ""),
        "{{LIST_OF_CHANGES}}": header.get("list_of_changes", ""),
    }

    if template_path and os.path.exists(template_path):
        document = Document(template_path)
        replace_placeholders_in_doc(document, replacements)

        ref_replaced = _find_and_replace_placeholder_with_table(
            document, "{{REFINEMENT_SUMMARY_TABLE}}", SUMMARY_HEADERS, summary_rows(enhancements)
        )
        def_replaced = _find_and_replace_placeholder_with_table(
            document, "{{DEFECT_SUMMARY_TABLE}}", SUMMARY_HEADERS, summary_rows(bugs)
        )

        if not ref_replaced and not def_replaced:
            _append_ticket_sections(document, enhancements, bugs, header)
    else:
        document = Document()

        # Title page
        _bold_centered(document, "COMPANY NAME", 24)
        _bold_centered(document, f"FOR {header.get('client_name','')}", 16)
        _bold_centered(document, "RELEASE NOTE", 16)
        _bold_centered(document, "FOR", 16)
        _bold_centered(document, f"VERSION {header.get('version_no','')}", 16)
        document.add_paragraph()

        document.add_heading("Revision History", level=2)
        _add_table(
            document,
            ["Version No", "Date of Release", "Prepared by", "Reviewed / Verified by", "List of changes from Previous Version"],
            [[
                header.get("revision_no", "1.0"),
                header.get("release_date", ""),
                header.get("prepared_by", ""),
                header.get("reviewed_by", ""),
                header.get("list_of_changes", "Document release"),
            ]],
        )
        _page_break(document)

        # Table of Contents
        document.add_heading("Table of Contents", level=1)
        _add_toc(document)
        _page_break(document)

        # Introduction
        document.add_heading("Introduction", level=1)
        document.add_heading("Project Information", level=2)
        _add_table(
            document, ["Field", "Value"],
            [
                ["Project ID :", header.get("project_id", "")],
                ["Project Name :", header.get("project_name", "")],
                ["Version No :", header.get("version_no", "")],
                ["Client Name :", header.get("client_name", "")],
                ["Release Date :", header.get("release_date", "")],
            ],
            shade_header=False,
        )
        document.add_heading("Objectives", level=2)
        _add_table(
            document, ["No.", "Objective", "Reference"],
            [["1.", "Release Notes Verification", "Refer Section 3 \u2013 Refinement and Defect list"]],
        )

        # Purpose
        document.add_heading("Purpose", level=1)
        document.add_paragraph(
            "The purpose of the document is to update the end-user of the changes and "
            "defects have been resolved to be included in the new version of application "
            "and will be deployed to production server."
        )
        document.add_paragraph(
            "This document outlines the list of request and issue to be delivered in the "
            "newer version of application and the guidelines on how to deploy the new "
            "version of application in production."
        )

        _append_ticket_sections(document, enhancements, bugs, header)
        replace_placeholders_in_doc(document, replacements)

    document.save(output_path)
    return {
        "enhancements_count": len(enhancements),
        "bugs_count": len(bugs),
        "total_included": len(df),
    }