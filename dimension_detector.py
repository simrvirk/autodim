"""
dimension_detector.py
=====================
Engineering drawing dimension detector and measurement template parser.

Functions
---------
detect_dimensions(pdf_path, page_num) -> list[dict]
    Extract text blocks from a drawing PDF and classify which ones are
    engineering dimensions using the Claude API.

parse_spreadsheet_template(template_path) -> dict
    Fully parse an Excel measurement record template, preserving its
    structure, formulas, formatting, and conditional-formatting rules
    so the template can later be populated with real dimension data.
"""

from __future__ import annotations

import json
import os
from typing import Any

import re

import fitz  # PyMuPDF
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Regex-based dimension classifier
# ---------------------------------------------------------------------------

def _classify_dimension(text: str) -> tuple[Optional[str], float]:
    """Return (category, confidence) if text looks like an engineering dimension.

    Returns (None, 0.0) for text that is not a dimension.
    Confidence is 0.0–1.0; only items with confidence > 0 are returned.
    """
    t = text.strip()

    # Reference dimensions in parentheses are for information only — skip them
    if re.match(r'^\(.*\)$', t):
        return None, 0.0

    # Reject obviously non-dimension text
    if not t or len(t) > 80:
        return None, 0.0
    word_count = len(t.split())
    if word_count > 8:
        return None, 0.0
    # Skip plain revision letters / single uppercase words (not datum context)
    if re.match(r'^[A-Z]{2,}$', t) and not re.search(r'\d', t):
        return None, 0.0

    # --- Angular ---
    if re.search(r'\d+\.?\d*\s*°', t):
        return "angular", 0.92

    # --- Diameter ---
    if re.search(r'[Øø⌀∅]\s*\d+', t):
        return "diameter", 0.95
    if re.search(r'\bDIA\.?\s*\d+', t, re.IGNORECASE):
        return "diameter", 0.90

    # --- Radius ---
    if re.match(r'^[Rr]\s*\d+\.?\d*', t):
        return "radius", 0.92
    if re.search(r'\bCR\s*\d+\.?\d*', t, re.IGNORECASE):
        return "radius", 0.88

    # --- Thread / screw-hole callouts — intentionally ignored ---
    if re.match(r'^[Mm]\d+[xX×]\d+', t):
        return None, 0.0
    if re.search(r'\d+-\d+\s*(UNC|UNF|UNEF|NPT|NPS)\b', t, re.IGNORECASE):
        return None, 0.0
    if re.match(r'^#\d+-\d+', t):
        return None, 0.0

    # --- Surface finish ---
    if re.search(r'\b[Rr][aAzZqQwW]\s*\d+', t):
        return "finish", 0.92
    if re.search(r'\d+\.?\d*\s*RMS\b', t, re.IGNORECASE):
        return "finish", 0.90
    if re.match(r'^N\d{1,2}$', t):
        return "finish", 0.78

    # --- GD&T symbols and keywords — intentionally ignored ---
    if re.search(r'[⊕⊙⌖⊞⊟◎⌯⌦⌧]', t):
        return None, 0.0
    if re.search(
        r'\b(TRUE\s*POS|T\.P\.|FLATNESS|PERPENDICULARITY|CONCENTRICITY'
        r'|RUNOUT|CIRCULARITY|CYLINDRICITY|PARALLELISM|STRAIGHTNESS'
        r'|PROFILE|SYMMETRY)\b', t, re.IGNORECASE
    ):
        return None, 0.0

    # --- Hole / feature callouts — intentionally ignored ---
    if re.search(r'\b(C\'?BORE|CBORE|CSINK|SPOTFACE|THRU)\b', t, re.IGNORECASE):
        return None, 0.0

    # --- Chamfer / taper ---
    if re.search(r'\d+\s*[Xx×]\s*\d+\.?\d*\s*°', t):
        return "angular", 0.88
    if re.search(r'\bTAPER\b', t, re.IGNORECASE):
        return "linear", 0.80

    # --- Tolerance-only blocks (e.g. "±0.005") ---
    if re.match(r'^[±]\s*\d+\.?\d*$', t):
        return "linear", 0.88

    # --- Linear with explicit tolerance ---
    if re.search(r'\d+\.?\d*\s*[±]\s*\d+\.?\d*', t):
        return "linear", 0.92
    if re.search(r'\d+\.?\d*\s*\+\d+\.?\d*\s*/?\s*-\d+\.?\d*', t):
        return "linear", 0.92

    # --- Plain decimal number (likely a dimension if it has a decimal point) ---
    m = re.match(r'^([+-]?\d{1,5}\.\d{1,5})\s*$', t)
    if m:
        val = abs(float(m.group(1)))
        if 0.0001 <= val <= 99999:
            return "linear", 0.72

    return None, 0.0


# ---------------------------------------------------------------------------
# Public API — PDF dimension detection
# ---------------------------------------------------------------------------

def detect_dimensions(pdf_path: str, page_num: int = 0) -> list[dict]:
    """Detect engineering dimensions in a single page of a drawing PDF.

    Parameters
    ----------
    pdf_path : str
        Absolute or relative path to the PDF file.
    page_num : int
        Zero-based page index. Defaults to 0 (first page).

    Returns
    -------
    list[dict]
        One dict per detected dimension:
        {
            "index"     : int,    # position in the extracted block list
            "text"      : str,    # raw text string from the PDF
            "x"         : float,  # left edge of bounding box (points)
            "y"         : float,  # top edge of bounding box (points)
            "width"     : float,  # bounding box width (points)
            "height"    : float,  # bounding box height (points)
            "confidence": float,  # 0–1 from Claude
            "category"  : str,    # dimension category label
            "included"  : True    # always True for returned items
        }

    Raises
    ------
    FileNotFoundError
        If pdf_path does not point to an existing file.
    ValueError
        If page_num is out of range for the document.
    RuntimeError
        If the page contains no extractable text (scanned PDF) or if the
        Claude API returns an unexpected payload.
    """
    # ------------------------------------------------------------------
    # 1. Validate inputs
    # ------------------------------------------------------------------
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"PDF not found: {pdf_path!r}")

    doc = fitz.open(pdf_path)
    total_pages = len(doc)

    if page_num < 0 or page_num >= total_pages:
        doc.close()
        raise ValueError(
            f"page_num {page_num} is out of range for a {total_pages}-page document."
        )

    # ------------------------------------------------------------------
    # 2. Extract individual text lines with bounding boxes via PyMuPDF
    # ------------------------------------------------------------------
    page = doc[page_num]

    # Use get_text("dict") instead of get_text("blocks") so that each *line*
    # gets its own bounding box.  The "blocks" API merges adjacent lines into
    # one block, causing dimension strings like "11.00 +/- 0.2", "2X", "6.00"
    # to be concatenated into a single entry and mis-classified.
    raw_dict = page.get_text("dict")
    doc.close()

    text_blocks: list[dict] = []
    for block in raw_dict.get("blocks", []):
        if block.get("type") != 0:
            continue  # skip image blocks
        for line in block.get("lines", []):
            # Concatenate all spans on the line into one string
            line_text = "".join(
                span["text"] for span in line.get("spans", [])
            ).strip()
            if not line_text:
                continue
            bbox = line["bbox"]   # (x0, y0, x1, y1)
            text_blocks.append({
                "index" : len(text_blocks),
                "text"  : line_text,
                "x"     : round(bbox[0], 3),
                "y"     : round(bbox[1], 3),
                "width" : round(bbox[2] - bbox[0], 3),
                "height": round(bbox[3] - bbox[1], 3),
            })

    if not text_blocks:
        raise RuntimeError(
            "No text could be extracted from page "
            f"{page_num} of '{pdf_path}'. "
            "This PDF may be a scanned image. Run OCR (e.g. ocrmypdf) on it "
            "first, or use a vector-exported drawing PDF."
        )

    # ------------------------------------------------------------------
    # 3. Classify each text block with regex patterns
    # ------------------------------------------------------------------
    results: list[dict] = []
    for block in text_blocks:
        category, confidence = _classify_dimension(block["text"])
        if category is None:
            continue
        results.append({
            "index"     : block["index"],
            "text"      : block["text"],
            "x"         : block["x"],
            "y"         : block["y"],
            "width"     : block["width"],
            "height"    : block["height"],
            "confidence": confidence,
            "category"  : category,
            "included"  : True,
        })

    # Return in document order (ascending y then x — top-to-bottom, left-to-right)
    results.sort(key=lambda d: (d["y"], d["x"]))
    return results


# ---------------------------------------------------------------------------
# Public API — Excel template parser
# ---------------------------------------------------------------------------

def parse_spreadsheet_template(template_path: str) -> dict:
    """Fully parse an Excel measurement record template (.xlsx).

    Reads the workbook without evaluating formulas (so formula strings are
    preserved) and extracts every structural and formatting detail needed to
    later populate a copy of the template with real dimension data.

    Parameters
    ----------
    template_path : str
        Path to the .xlsx template file.

    Returns
    -------
    dict
        Top-level keys:
        ├─ "file_path"        : str
        ├─ "sheet_names"      : list[str]
        └─ "sheets"           : dict[sheet_name -> sheet_info]

        Each sheet_info dict contains:
        ├─ "dimensions"          : str   (e.g. "A1:Z100")
        ├─ "max_row"             : int
        ├─ "max_col"             : int
        ├─ "freeze_panes"        : str | None
        ├─ "column_headers"      : list[{"column_letter", "column_index",
        │                                "row", "value"}]
        ├─ "formulas"            : list[{"cell", "formula"}]
        ├─ "merged_cells"        : list[{"range", "min_row", "max_row",
        │                                "min_col", "max_col"}]
        ├─ "conditional_formats" : list[{"range", "rules"}]
        ├─ "column_widths"       : dict[column_letter -> width]
        ├─ "row_heights"         : dict[row_number -> height]
        ├─ "cell_styles"         : dict[cell_address -> style_dict]
        └─ "data_entry_rows"     : list[int]   (rows that appear to be for data)

    Raises
    ------
    FileNotFoundError
        If template_path does not exist.
    ValueError
        If the file is not a recognisable .xlsx workbook.
    """
    if not os.path.isfile(template_path):
        raise FileNotFoundError(f"Template not found: {template_path!r}")
    if not template_path.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        raise ValueError(
            f"Expected an Excel workbook (.xlsx/.xlsm), got: {template_path!r}"
        )

    # Load twice:
    #   • formula_wb  — formulas as strings (data_only=False, default)
    #   • value_wb    — computed cell values (data_only=True) for style inspection
    #   We use formula_wb as the primary source; value_wb only for reading
    #   cached computed values where useful (not saved over, so no data loss).
    try:
        formula_wb = load_workbook(template_path, data_only=False)
    except Exception as exc:
        raise ValueError(f"Could not open workbook: {exc}") from exc

    result: dict[str, Any] = {
        "file_path"  : os.path.abspath(template_path),
        "sheet_names": formula_wb.sheetnames,
        "sheets"     : {},
    }

    for sheet_name in formula_wb.sheetnames:
        ws = formula_wb[sheet_name]
        sheet_info: dict[str, Any] = {}

        # ------------------------------------------------------------------
        # Basic dimensions
        # ------------------------------------------------------------------
        sheet_info["dimensions"]   = ws.dimensions
        sheet_info["max_row"]      = ws.max_row
        sheet_info["max_col"]      = ws.max_column
        sheet_info["freeze_panes"] = (
            str(ws.freeze_panes) if ws.freeze_panes else None
        )

        # ------------------------------------------------------------------
        # Column headers
        # Heuristic: scan the first 5 rows for cells whose row is a header.
        # We identify a header row as one where most occupied cells are
        # non-numeric strings.  We collect every cell that has a string value
        # in that row.
        # ------------------------------------------------------------------
        headers: list[dict] = []
        header_rows_found: set[int] = set()

        for row_idx in range(1, min(6, (ws.max_row or 0) + 1)):
            row_cells = list(ws.iter_rows(min_row=row_idx, max_row=row_idx,
                                          values_only=False))[0]
            str_count   = sum(1 for c in row_cells
                              if c.value is not None
                              and isinstance(c.value, str)
                              and not str(c.value).startswith("="))
            total_count = sum(1 for c in row_cells if c.value is not None)
            if total_count > 0 and str_count / total_count >= 0.5:
                header_rows_found.add(row_idx)
                for cell in row_cells:
                    if cell.value is not None:
                        headers.append({
                            "column_letter": get_column_letter(cell.column),
                            "column_index" : cell.column,
                            "row"          : cell.row,
                            "value"        : cell.value,
                        })

        sheet_info["column_headers"] = headers

        # ------------------------------------------------------------------
        # Formulas — every cell whose value starts with "="
        # ------------------------------------------------------------------
        formulas: list[dict] = []
        for row in ws.iter_rows():
            for cell in row:
                val = cell.value
                if isinstance(val, str) and val.startswith("="):
                    formulas.append({
                        "cell"   : cell.coordinate,
                        "formula": val,
                    })
        sheet_info["formulas"] = formulas

        # ------------------------------------------------------------------
        # Merged cells
        # ------------------------------------------------------------------
        merged: list[dict] = []
        for rng in ws.merged_cells.ranges:
            merged.append({
                "range"  : str(rng),
                "min_row": rng.min_row,
                "max_row": rng.max_row,
                "min_col": rng.min_col,
                "max_col": rng.max_col,
            })
        sheet_info["merged_cells"] = merged

        # ------------------------------------------------------------------
        # Conditional formatting rules
        # openpyxl stores them as ConditionalFormattingList.
        # Each entry: (sqref, [Rule, ...])
        # ------------------------------------------------------------------
        cf_list: list[dict] = []
        for sqref, rules in ws.conditional_formatting._cf_rules.items():
            serialised_rules = []
            for rule in rules:
                rule_dict: dict[str, Any] = {
                    "type"    : rule.type,
                    "priority": rule.priority,
                    "operator": getattr(rule, "operator", None),
                    "formula" : (list(rule.formula)
                                 if getattr(rule, "formula", None)
                                 else None),
                }
                # Differential style (fill/font/border applied when rule fires)
                dxf = getattr(rule, "dxf", None)
                if dxf is not None:
                    dxf_info: dict[str, Any] = {}
                    if dxf.fill and dxf.fill.fgColor:
                        dxf_info["fill_fg_color"] = dxf.fill.fgColor.rgb
                    if dxf.font:
                        dxf_info["font_bold"]  = dxf.font.bold
                        dxf_info["font_color"] = (
                            dxf.font.color.rgb if dxf.font.color else None
                        )
                    if dxf.border:
                        dxf_info["has_border"] = True
                    rule_dict["dxf"] = dxf_info

                # Color scale / data bar / icon set specifics
                for attr in ("colorScale", "dataBar", "iconSet"):
                    if getattr(rule, attr, None) is not None:
                        rule_dict[attr] = str(getattr(rule, attr))

                serialised_rules.append(rule_dict)

            cf_list.append({
                "range": str(sqref),
                "rules": serialised_rules,
            })
        sheet_info["conditional_formats"] = cf_list

        # ------------------------------------------------------------------
        # Column widths
        # ------------------------------------------------------------------
        col_widths: dict[str, float | None] = {}
        for col_letter, col_dim in ws.column_dimensions.items():
            col_widths[col_letter] = col_dim.width
        sheet_info["column_widths"] = col_widths

        # ------------------------------------------------------------------
        # Row heights
        # ------------------------------------------------------------------
        row_heights: dict[int, float | None] = {}
        for row_idx, row_dim in ws.row_dimensions.items():
            row_heights[row_idx] = row_dim.height
        sheet_info["row_heights"] = row_heights

        # ------------------------------------------------------------------
        # Cell styles — serialise only non-default cells to keep the dict lean
        # ------------------------------------------------------------------
        cell_styles: dict[str, dict] = {}
        for row in ws.iter_rows():
            for cell in row:
                style = _serialise_cell_style(cell)
                if style:
                    cell_styles[cell.coordinate] = style
        sheet_info["cell_styles"] = cell_styles

        # ------------------------------------------------------------------
        # Data-entry rows heuristic
        # A row is a "data entry" row if:
        #   • it is below all detected header rows
        #   • it has at least one empty cell in a column that has a header
        #   • it is not entirely empty (it has some structure/formula)
        # ------------------------------------------------------------------
        header_row_max = max(header_rows_found) if header_rows_found else 1
        header_col_indices = {h["column_index"] for h in headers}

        data_entry_rows: list[int] = []
        for row_idx in range(header_row_max + 1, (ws.max_row or 0) + 1):
            row_cells = {
                c.column: c
                for c in ws[row_idx]
                if c.column in header_col_indices
            }
            has_formula = any(
                isinstance(c.value, str) and c.value.startswith("=")
                for c in row_cells.values()
            )
            has_empty = any(
                c.value is None for c in row_cells.values()
            )
            # Rows with a mix of formulas and empty cells are data-entry rows
            if has_empty and (has_formula or _row_has_style(ws, row_idx)):
                data_entry_rows.append(row_idx)

        sheet_info["data_entry_rows"] = data_entry_rows

        result["sheets"][sheet_name] = sheet_info

    formula_wb.close()
    return result


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

def _serialise_cell_style(cell) -> dict:
    """Return a dict of non-default style attributes for a cell.

    Returns an empty dict if the cell has no meaningful styling, so callers
    can skip storing it.
    """
    style: dict[str, Any] = {}

    # Font
    f = cell.font
    if f:
        font_info: dict[str, Any] = {}
        if f.bold:
            font_info["bold"] = True
        if f.italic:
            font_info["italic"] = True
        if f.underline:
            font_info["underline"] = f.underline
        if f.size and f.size != 11:
            font_info["size"] = f.size
        if f.name and f.name not in ("Calibri", "Arial", ""):
            font_info["name"] = f.name
        if f.color and f.color.type == "rgb" and f.color.rgb not in (
            "00000000", "FF000000"
        ):
            font_info["color"] = f.color.rgb
        if font_info:
            style["font"] = font_info

    # Fill
    fl = cell.fill
    if fl and fl.fill_type and fl.fill_type != "none":
        fill_info: dict[str, Any] = {"fill_type": fl.fill_type}
        try:
            if fl.fgColor and fl.fgColor.type == "rgb":
                fill_info["fg_color"] = fl.fgColor.rgb
            if fl.bgColor and fl.bgColor.type == "rgb":
                fill_info["bg_color"] = fl.bgColor.rgb
        except Exception:
            pass
        style["fill"] = fill_info

    # Alignment
    al = cell.alignment
    if al:
        align_info: dict[str, Any] = {}
        if al.horizontal and al.horizontal != "general":
            align_info["horizontal"] = al.horizontal
        if al.vertical and al.vertical != "bottom":
            align_info["vertical"] = al.vertical
        if al.wrap_text:
            align_info["wrap_text"] = True
        if align_info:
            style["alignment"] = align_info

    # Number format
    nf = cell.number_format
    if nf and nf not in ("General", "@", ""):
        style["number_format"] = nf

    # Border — just flag its presence; full border serialisation is verbose
    bd = cell.border
    if bd:
        sides = [bd.left, bd.right, bd.top, bd.bottom]
        if any(s and s.border_style for s in sides):
            style["has_border"] = True

    return style


def _row_has_style(ws, row_idx: int) -> bool:
    """Return True if any cell in the row has non-default styling."""
    for cell in ws[row_idx]:
        if _serialise_cell_style(cell):
            return True
    return False
