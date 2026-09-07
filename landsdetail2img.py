#!/usr/bin/env python
"""
landsdetail2img.py
===================

Converts long landscape-format tables inside a Word (.docx) report into a
series of PNG screenshots, so they can be pasted into other documents
(e.g. PowerPoint slides) without manually taking screenshots in Word.

Each ".docx" is expected to contain one or more sections, where a section
looks like:

    <Heading> 9.x Some Section Title </Heading>
    <Normal>  Some short description text ...            </Normal>
    <Table>   #, Findings, Affected, Risk Description, ... (very long)</Table>

The script walks the document top to bottom, and for every table it finds,
it uses the *closest preceding Heading-styled paragraph* as the "section
name" used in the output file name.

Because the table can be arbitrarily long, each table is split into one or
more images:
    * every image always starts with the table's own header row(s) (the
      first row(s) marked as a repeating header in the .docx),
    * every image contains at least one *complete* data row (a data row / a
      group of vertically-merged rows is NEVER split across two images),
    * as many additional whole data rows as possible are packed into the
      same image, as long as the total VISIBLE ROW COUNT of those data rows
      -- i.e. the number of rendered lines, counting both text wrapping and
      explicit line breaks, NOT the number of logical <w:tr> table rows --
      does not exceed --max-rows (default: 40). A single data row containing
      a lot of text can by itself take up many "visible rows"; a table
      section with long cells will therefore typically only fit 1-3 logical
      data rows per image, while a table with short cells may fit dozens.
    * Word's own PAGE headers/footers (running titles, logos, page numbers,
      etc.) are always stripped out -- only the table itself is rendered.

Output file naming:
    <section-name>-<1-based index>.png

Rendering engines
------------------
Two rendering engines are available via --engine:

  "libreoffice" (DEFAULT, recommended)
      For each output image, this builds a temporary .docx that is a copy
      of your ORIGINAL document with everything removed except the target
      table (pruned down to just the rows for that image). It then shells
      out to a locally-installed LibreOffice (`soffice --headless
      --convert-to pdf`) to render that temporary document to PDF, and
      rasterizes + crops the result to a PNG.
      Because the actual table markup (styles, direct formatting, theme
      fonts, etc.) is reused verbatim and rendered by a real, mature
      word-processor layout engine, this reproduces your original
      document's fonts, spacing, and line-wrapping far more faithfully
      than a hand-rolled renderer ever could. Requires LibreOffice to be
      installed on the machine running the script (free, cross-platform:
      https://www.libreoffice.org/download/).

  "pillow" (fallback, no external dependency other than pip packages)
      Parses the table's XML directly and draws it with Pillow, doing our
      own text-wrapping/layout. Useful if LibreOffice cannot be installed
      in your environment, but is inherently an approximation of Word's
      real layout/font rendering.

Note on how "visible rows" (wrapped lines) are estimated
----------------------------------------------------------
Deciding how many logical data rows fit within a --max-rows budget of
*wrapped lines* requires knowing, in advance, how many lines each cell's
text will wrap to -- which depends on the real column widths and fonts.
Both engines therefore share a single lightweight text-wrapping estimator
(using Pillow + metrically-compatible fonts) purely to COUNT how many lines
each row will take, regardless of which engine is used to do the actual,
final pixel rendering of the image. This keeps the row-bucketing decision
fast (no need to invoke LibreOffice repeatedly just to measure row counts),
at the cost of the estimate occasionally being off by about one line versus
Word's own exact line-breaking in edge cases.

Usage
-----
    python landsdetail2img.py "QC FUP v1.2.docx" -o out_dir
    python landsdetail2img.py report1.docx report2.docx -o out_dir --dpi 200 --max-rows 30
    python landsdetail2img.py report.docx -o out_dir --engine pillow
    python landsdetail2img.py report.docx -o out_dir --soffice-path "C:\\Program Files\\LibreOffice\\program\\soffice.exe"

Run `python landsdetail2img.py -h` for the full list of options.
"""

from __future__ import annotations

import argparse
import glob
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from lxml import etree

# ---------------------------------------------------------------------------
# Configuration defaults (all overridable via CLI flags, see --help)
# ---------------------------------------------------------------------------
DEFAULT_DPI = 150
DEFAULT_MAX_ROWS = 40
TWIPS_PER_INCH = 1440.0

# Fixed, nominal DPI used ONLY for estimating how many lines each cell's
# text will wrap to (for row-bucketing decisions). This is deliberately
# decoupled from the user's requested output --dpi: the wrapping decision
# (how many words fit on a line) is scale-invariant, since both the
# available column width and the glyph advance widths scale together with
# DPI, so any fixed reference value works equally well here.
LINE_ESTIMATION_DPI = 150

# How many temporary per-chunk .docx files to hand to a single `soffice`
# invocation at once. Batching multiple files into one soffice call avoids
# paying LibreOffice's ~2-3s cold-start cost per file.
SOFFICE_BATCH_SIZE = 40

# The temporary single-table documents rendered via LibreOffice use a
# generous, but NOT absurdly huge, custom page size:
#   - width  = the table's own real width (from its column grid) + a small
#              safety buffer, so no column is ever clipped by the page edge.
#   - height = estimated per-chunk from the ESTIMATED VISIBLE-LINE COUNT it
#              contains, with a generous per-line allowance, rather than one
#              fixed huge constant -- this keeps the intermediate PDF/PNG
#              small and fast to rasterize, and (combined with stripping
#              page headers/footers) ensures the final cropped image
#              contains ONLY the table, with no large blank gap.
CHUNK_PAGE_MARGIN_TWIPS = 200  # ~0.14 inch on each side, auto-cropped anyway
CHUNK_WIDTH_SAFETY_BUFFER_TWIPS = 720  # +0.5 inch safety margin on width
LINE_HEIGHT_ESTIMATE_TWIPS = 400  # ~0.28 inch per wrapped line, generous upper bound
ROW_CHROME_ESTIMATE_TWIPS = 200  # extra per-row allowance (cell padding, borders)
MIN_CHUNK_PAGE_HEIGHT_TWIPS = 6_000
MAX_CHUNK_PAGE_HEIGHT_TWIPS = 900_000

# Canonical child-element order for CT_TblPrBase (ECMA-376 §17.4), needed
# whenever we insert a new element into an existing <w:tblPr> so the result
# stays schema-valid (Word/LibreOffice are lenient, but let's not rely on
# that).
TBLPR_CHILD_ORDER = [
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
    "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
    "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
    "tblCaption", "tblDescription",
]

# ---------------------------------------------------------------------------
# Pillow-engine-only constants (used both by the Pillow rendering engine AND
# by the shared line-count estimator used for row-bucketing in both engines)
# ---------------------------------------------------------------------------
BORDER_COLOR = (0, 0, 0)
BORDER_WIDTH_PT = 0.5

DEFAULT_TEXT_COLOR = (0, 0, 0)
DEFAULT_FONT_NAME = "Calibri"
DEFAULT_FONT_SIZE_PT = 11.0

CELL_LEFT_RIGHT_PAD_TWIPS = 108
CELL_TOP_BOTTOM_PAD_TWIPS = 40
LINE_SPACING_FACTOR = 1.20


# ---------------------------------------------------------------------------
# Data structures (shared by both engines)
# ---------------------------------------------------------------------------
@dataclass
class Segment:
    text: str
    font_name: str
    size_pt: float
    bold: bool
    italic: bool
    color: tuple


@dataclass
class CellData:
    row0: int
    col0: int
    rowspan: int
    colspan: int
    shading: Optional[tuple]
    raw_lines: list  # list[list[Segment]] - explicit (pre word-wrap) lines


# ---------------------------------------------------------------------------
# XML parsing helpers (shared by both engines)
# ---------------------------------------------------------------------------
def _get_bool_prop(pr_el, tag: str) -> bool:
    """OOXML boolean run/paragraph property, e.g. <w:b/> or <w:b w:val="0"/>."""
    if pr_el is None:
        return False
    el = pr_el.find(qn(f"w:{tag}"))
    if el is None:
        return False
    val = el.get(qn("w:val"))
    if val is None:
        return True
    return val.lower() not in ("0", "false", "off", "none")


def _hex_to_rgb(hex_str: str) -> tuple:
    hex_str = hex_str.lstrip("#")
    return tuple(int(hex_str[i : i + 2], 16) for i in (0, 2, 4))


def extract_run_format(rPr) -> dict:
    font_name = DEFAULT_FONT_NAME
    size_pt = DEFAULT_FONT_SIZE_PT
    color = DEFAULT_TEXT_COLOR
    bold = False
    italic = False

    if rPr is not None:
        rfonts = rPr.find(qn("w:rFonts"))
        if rfonts is not None:
            ascii_name = rfonts.get(qn("w:ascii")) or rfonts.get(qn("w:hAnsi"))
            if ascii_name:
                font_name = ascii_name

        sz = rPr.find(qn("w:sz"))
        if sz is not None and sz.get(qn("w:val")):
            try:
                size_pt = int(sz.get(qn("w:val"))) / 2.0
            except ValueError:
                pass

        color_el = rPr.find(qn("w:color"))
        if color_el is not None:
            val = color_el.get(qn("w:val"))
            if val and val.lower() != "auto":
                try:
                    color = _hex_to_rgb(val)
                except ValueError:
                    pass

        bold = _get_bool_prop(rPr, "b")
        italic = _get_bool_prop(rPr, "i")

    return dict(font_name=font_name, size_pt=size_pt, bold=bold, italic=italic, color=color)


def cell_raw_lines(tc) -> list:
    """Return list[list[Segment]]: one inner list per explicit line, where a
    new explicit line is started by a paragraph boundary or a <w:br/>."""
    lines: list = []
    for p in tc.findall(qn("w:p")):
        current: list = []
        for r in p.findall(qn("w:r")):
            rPr = r.find(qn("w:rPr"))
            fmt = extract_run_format(rPr)
            for child in r:
                tag = etree.QName(child).localname
                if tag == "t":
                    text = child.text or ""
                    if text:
                        current.append(Segment(text=text, **fmt))
                elif tag == "br":
                    lines.append(current)
                    current = []
                elif tag == "tab":
                    current.append(Segment(text="\t", **fmt))
        lines.append(current)
    if not lines:
        lines = [[]]
    return lines


def cell_shading(tcPr) -> Optional[tuple]:
    if tcPr is None:
        return None
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        return None
    fill = shd.get(qn("w:fill"))
    if not fill or fill.lower() == "auto":
        return None
    try:
        return _hex_to_rgb(fill)
    except ValueError:
        return None


def parse_table_grid(table: Table):
    """Parse a python-docx Table into (col_widths_twips, grid, header_rows).

    grid is a 2D list [n_rows][n_cols] of CellData objects. Cells that are
    covered by a horizontal (gridSpan) or vertical (vMerge) merge point to the
    SAME CellData instance as their "owner" cell (top-left of the merge).
    """
    tbl = table._tbl
    tblGrid = tbl.find(qn("w:tblGrid"))
    col_widths_twips = []
    if tblGrid is not None:
        for gridCol in tblGrid.findall(qn("w:gridCol")):
            w = gridCol.get(qn("w:w"))
            col_widths_twips.append(float(w) if w else 1000.0)
    n_cols = len(col_widths_twips)

    trs = tbl.findall(qn("w:tr"))
    n_rows = len(trs)
    grid = [[None] * n_cols for _ in range(n_rows)]
    header_rows = []
    open_vmerge: dict = {}  # col_index -> CellData currently open

    for r, tr in enumerate(trs):
        trPr = tr.find(qn("w:trPr"))
        if trPr is not None and trPr.find(qn("w:tblHeader")) is not None:
            val = trPr.find(qn("w:tblHeader")).get(qn("w:val"))
            if val is None or val.lower() not in ("0", "false", "off"):
                header_rows.append(r)

        col_cursor = 0
        for tc in tr.findall(qn("w:tc")):
            tcPr = tc.find(qn("w:tcPr"))
            gridSpan_el = tcPr.find(qn("w:gridSpan")) if tcPr is not None else None
            gridspan = int(gridSpan_el.get(qn("w:val"))) if gridSpan_el is not None else 1

            vMerge_el = tcPr.find(qn("w:vMerge")) if tcPr is not None else None
            is_continue = False
            is_restart = False
            if vMerge_el is not None:
                val = vMerge_el.get(qn("w:val"))
                if val is None or val.lower() == "continue":
                    is_continue = True
                else:
                    is_restart = True

            if is_continue and col_cursor in open_vmerge:
                owner = open_vmerge[col_cursor]
                owner.rowspan += 1
                for cc in range(col_cursor, min(col_cursor + gridspan, n_cols)):
                    grid[r][cc] = owner
            else:
                cell = CellData(
                    row0=r,
                    col0=col_cursor,
                    rowspan=1,
                    colspan=gridspan,
                    shading=cell_shading(tcPr),
                    raw_lines=cell_raw_lines(tc),
                )
                for cc in range(col_cursor, min(col_cursor + gridspan, n_cols)):
                    grid[r][cc] = cell
                if is_restart:
                    open_vmerge[col_cursor] = cell
                else:
                    open_vmerge.pop(col_cursor, None)

            col_cursor += gridspan

    if not header_rows:
        header_rows = [0] if n_rows else []

    return col_widths_twips, grid, header_rows, n_rows, n_cols


# ---------------------------------------------------------------------------
# Section / heading discovery (shared by both engines)
# ---------------------------------------------------------------------------
def sanitize_filename_part(text: str) -> str:
    text = text.strip()
    text = re.sub(r"[\\/:*?\"<>|]", "_", text)
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_.")
    return text or "Section"


def discover_sections(document: Document):
    """Walk the document body top-to-bottom and pair every top-level table
    with the closest preceding Heading-styled paragraph text. Returns a list
    of (heading_text, table_index, table) where table_index is the table's
    0-based position among top-level <w:tbl> children of the body (stable
    across independent reloads of the same source file, which the
    LibreOffice engine relies on)."""
    body = document.element.body
    sections = []
    current_heading = None
    table_index = 0
    for child in body.iterchildren():
        tag = etree.QName(child).localname
        if tag == "p":
            para = Paragraph(child, document)
            style_name = para.style.name if para.style is not None else ""
            if style_name and style_name.lower().startswith("heading") and para.text.strip():
                current_heading = para.text.strip()
        elif tag == "tbl":
            table = Table(child, document)
            sections.append((current_heading, table_index, table))
            table_index += 1
    return sections


# ---------------------------------------------------------------------------
# Text-wrapping estimator (shared: used for BOTH the Pillow rendering engine
# AND, regardless of engine, to estimate how many visible/wrapped lines each
# row will take, for row-bucketing purposes).
# ---------------------------------------------------------------------------
_system = platform.system()
if _system == "Windows":
    _win_dir = os.environ.get("WINDIR", r"C:\Windows")
    FONT_SEARCH_DIRS = [
        os.path.join(_win_dir, "Fonts"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"),
    ]
elif _system == "Darwin":
    FONT_SEARCH_DIRS = [
        "/Library/Fonts",
        "/System/Library/Fonts",
        "/System/Library/Fonts/Supplemental",
        os.path.expanduser("~/Library/Fonts"),
    ]
else:
    FONT_SEARCH_DIRS = [
        "/usr/share/fonts",
        "/usr/local/share/fonts",
        os.path.expanduser("~/.fonts"),
        os.path.expanduser("~/.local/share/fonts"),
    ]

FONT_CANDIDATES = {
    "times new roman": [
        ("times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"),
        ("Times New Roman.ttf", "Times New Roman Bold.ttf", "Times New Roman Italic.ttf", "Times New Roman Bold Italic.ttf"),
        ("LiberationSerif-Regular.ttf", "LiberationSerif-Bold.ttf", "LiberationSerif-Italic.ttf", "LiberationSerif-BoldItalic.ttf"),
        ("DejaVuSerif.ttf", "DejaVuSerif-Bold.ttf", "DejaVuSerif-Italic.ttf", "DejaVuSerif-BoldItalic.ttf"),
    ],
    "calibri": [
        ("calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf"),
        ("Calibri.ttf", "Calibri Bold.ttf", "Calibri Italic.ttf", "Calibri Bold Italic.ttf"),
        ("LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf", "LiberationSans-Italic.ttf", "LiberationSans-BoldItalic.ttf"),
        ("DejaVuSans.ttf", "DejaVuSans-Bold.ttf", "DejaVuSans-Oblique.ttf", "DejaVuSans-BoldOblique.ttf"),
    ],
    "arial": [
        ("arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"),
        ("Arial.ttf", "Arial Bold.ttf", "Arial Italic.ttf", "Arial Bold Italic.ttf"),
        ("LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf", "LiberationSans-Italic.ttf", "LiberationSans-BoldItalic.ttf"),
        ("DejaVuSans.ttf", "DejaVuSans-Bold.ttf", "DejaVuSans-Oblique.ttf", "DejaVuSans-BoldOblique.ttf"),
    ],
}
GENERIC_FALLBACK_GROUPS = [
    ("arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"),
    ("calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf"),
    ("LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf", "LiberationSans-Italic.ttf", "LiberationSans-BoldItalic.ttf"),
    ("DejaVuSans.ttf", "DejaVuSans-Bold.ttf", "DejaVuSans-Oblique.ttf", "DejaVuSans-BoldOblique.ttf"),
]

_FONT_INDEX: Optional[dict] = None
_FONT_CACHE: dict = {}


def _build_font_index() -> dict:
    index: dict = {}
    for base_dir in FONT_SEARCH_DIRS:
        if not base_dir or not os.path.isdir(base_dir):
            continue
        for root, _dirs, files in os.walk(base_dir):
            for fname in files:
                if fname.lower().endswith((".ttf", ".ttc", ".otf")):
                    index.setdefault(fname.lower(), os.path.join(root, fname))
    return index


def _get_font_index() -> dict:
    global _FONT_INDEX
    if _FONT_INDEX is None:
        _FONT_INDEX = _build_font_index()
    return _FONT_INDEX


def _resolve_font_path(family_key: str, style_idx: int) -> Optional[str]:
    index = _get_font_index()
    for group in FONT_CANDIDATES.get(family_key, []):
        fname = group[style_idx].lower()
        if fname in index:
            return index[fname]
    for group in GENERIC_FALLBACK_GROUPS:
        fname = group[style_idx].lower()
        if fname in index:
            return index[fname]
    if index:
        return next(iter(index.values()))
    return None


def twips_to_px(twips: float, dpi: int) -> float:
    return twips / TWIPS_PER_INCH * dpi


def pt_to_px(pt: float, dpi: int) -> float:
    return pt / 72.0 * dpi


def get_font(name: str, size_pt: float, bold: bool, italic: bool, dpi: int):
    from PIL import ImageFont

    key = (name.lower() if name else "", round(size_pt, 1), bold, italic, dpi)
    if key in _FONT_CACHE:
        return _FONT_CACHE[key]

    style_idx = (1 if bold else 0) + (2 if italic else 0)
    family_key = (name or "").lower()
    size_px = max(1, round(pt_to_px(size_pt, dpi)))

    path = _resolve_font_path(family_key, style_idx)
    font = None
    if path:
        try:
            font = ImageFont.truetype(path, size_px)
        except OSError:
            font = None
    if font is None:
        try:
            font = ImageFont.load_default(size=size_px)
        except TypeError:
            font = ImageFont.load_default()

    _FONT_CACHE[key] = font
    return font


def wrap_cell_lines(cell: "CellData", max_width_px: float, dpi: int) -> list:
    """Return list of visual sub-lines for the whole cell (all explicit
    lines, tokenized + word-wrapped). Each sub-line is a list of
    (text, font, color) tuples ready to be drawn left-to-right."""
    all_subs = []
    for line_segments in cell.raw_lines:
        if not line_segments:
            all_subs.append([])
            continue
        all_subs.extend(_wrap_one_line(line_segments, max_width_px, dpi))
    return all_subs


def _split_long_chunk(text: str, font, max_width_px: float) -> list:
    """Break a too-long, unbreakable chunk (e.g. a URL) into pieces that
    each individually fit within max_width_px."""
    if font.getlength(text) <= max_width_px or len(text) <= 1:
        return [text]
    pieces = []
    current = ""
    for ch in text:
        trial = current + ch
        if current and font.getlength(trial) > max_width_px:
            pieces.append(current)
            current = ch
        else:
            current = trial
    if current:
        pieces.append(current)
    return pieces


def _wrap_one_line(segments: list, max_width_px: float, dpi: int) -> list:
    """Tokenize one explicit line into (space | word) tokens -- where a
    'word' may legitimately span multiple original runs when there is no
    actual space character between them -- then greedily word-wrap those
    tokens into visual sub-lines that each fit max_width_px."""
    tokens = []  # list of ("space", font) | ("word", [(text, font, color), ...])
    current_word: list = []
    for seg in segments:
        font = get_font(seg.font_name, seg.size_pt, seg.bold, seg.italic, dpi)
        text = seg.text
        i, n = 0, len(text)
        while i < n:
            if text[i] == " ":
                if current_word:
                    tokens.append(("word", current_word))
                    current_word = []
                tokens.append(("space", font))
                i += 1
            else:
                j = i
                while j < n and text[j] != " ":
                    j += 1
                current_word.append((text[i:j], font, seg.color))
                i = j
    if current_word:
        tokens.append(("word", current_word))

    sub_lines: list = []
    current_line: list = []
    current_width = 0.0
    pending_space: Optional[tuple] = None

    def word_width(parts) -> float:
        return sum(font.getlength(text) for text, font, _ in parts)

    def flush_line():
        nonlocal current_line, current_width, pending_space
        sub_lines.append(current_line)
        current_line = []
        current_width = 0.0
        pending_space = None

    for kind, payload in tokens:
        if kind == "space":
            font = payload
            if current_line:
                pending_space = (" ", font, DEFAULT_TEXT_COLOR)
            continue

        parts = payload
        w = word_width(parts)
        space_w = pending_space[1].getlength(" ") if pending_space else 0.0

        if w > max_width_px and not current_line and pending_space is None:
            for text, font, color in parts:
                for piece in _split_long_chunk(text, font, max_width_px):
                    piece_w = font.getlength(piece)
                    if current_line and current_width + piece_w > max_width_px:
                        flush_line()
                    current_line.append((piece, font, color))
                    current_width += piece_w
            continue

        if current_line and current_width + space_w + w > max_width_px:
            flush_line()
            if current_line and pending_space:
                current_line.append(pending_space)
                current_width += pending_space[1].getlength(" ")
        elif pending_space:
            current_line.append(pending_space)
            current_width += space_w

        pending_space = None
        current_line.extend(parts)
        current_width += w

    if current_line or not sub_lines:
        sub_lines.append(current_line)
    return sub_lines


def line_height_px(sub_line: list, dpi: int) -> float:
    if not sub_line:
        default_font = get_font(DEFAULT_FONT_NAME, DEFAULT_FONT_SIZE_PT, False, False, dpi)
        return default_font.size * LINE_SPACING_FACTOR
    max_size = max(font.size for _, font, _ in sub_line)
    return max_size * LINE_SPACING_FACTOR


def estimate_row_line_counts(grid, col_widths_twips, n_rows, n_cols, dpi=LINE_ESTIMATION_DPI):
    """For every row, estimate the number of VISIBLE (wrapped) lines its
    tallest cell will require -- i.e. the same quantity a viewer would count
    if they looked at the rendered table and counted text lines within that
    row, including both explicit line breaks and word-wrap. Cells that span
    multiple rows (vMerge) have their line-count requirement distributed
    across their row span the same way real row-height distribution works:
    any excess beyond what the other cells already imply is attributed to
    the LAST row of the span.

    Returns a list[int] of length n_rows (min 1 per row)."""
    col_widths_px = [twips_to_px(w, dpi) for w in col_widths_twips]
    pad_lr_px = twips_to_px(CELL_LEFT_RIGHT_PAD_TWIPS, dpi)

    min_lines = [0] * n_rows
    seen_ids = set()
    for r in range(n_rows):
        for c in range(n_cols):
            cell = grid[r][c]
            if cell is None or cell.row0 != r or cell.col0 != c:
                continue
            if id(cell) in seen_ids:
                continue
            seen_ids.add(id(cell))

            width_px = sum(col_widths_px[cell.col0 : cell.col0 + cell.colspan]) - 2 * pad_lr_px
            width_px = max(1.0, width_px)
            sub_lines = wrap_cell_lines(cell, width_px, dpi)
            n_lines = max(1, len(sub_lines))

            if cell.rowspan == 1:
                min_lines[r] = max(min_lines[r], n_lines)
            else:
                span = range(cell.row0, cell.row0 + cell.rowspan)
                current_sum = sum(min_lines[rr] for rr in span)
                if n_lines > current_sum:
                    min_lines[cell.row0 + cell.rowspan - 1] += n_lines - current_sum

    for r in range(n_rows):
        if min_lines[r] <= 0:
            min_lines[r] = 1
    return min_lines


# ---------------------------------------------------------------------------
# Row grouping (bands that must stay together) + chunking (<= max_rows
# VISIBLE/WRAPPED LINES, not logical row count) -- shared by both engines.
# ---------------------------------------------------------------------------
def group_into_bands(grid, data_row_indices, n_cols):
    bands = []
    current = []
    for r in data_row_indices:
        shares_merge = False
        if current:
            prev_r = current[-1]
            for c in range(n_cols):
                a, b = grid[r][c], grid[prev_r][c]
                if a is not None and a is b:
                    shares_merge = True
                    break
        if shares_merge:
            current.append(r)
        else:
            if current:
                bands.append(current)
            current = [r]
    if current:
        bands.append(current)
    return bands


def build_chunks(bands, row_weights, max_rows):
    """Greedily pack bands (groups of rows that must stay together) into
    chunks, where a chunk's "size" is the SUM of row_weights over all rows
    in it (row_weights[r] = estimated visible/wrapped line count of row r,
    NOT simply 1-per-row) -- so a chunk never exceeds max_rows worth of
    visible lines. A single band that by itself already exceeds max_rows is
    still kept whole (never split), per spec: a data row must never be split
    across two images."""
    chunks = []
    current_rows: list = []
    current_weight = 0
    for band in bands:
        band_weight = sum(row_weights[r] for r in band)
        if current_rows and current_weight + band_weight > max_rows:
            chunks.append(current_rows)
            current_rows = []
            current_weight = 0
        current_rows.extend(band)
        current_weight += band_weight
    if current_rows:
        chunks.append(current_rows)
    return chunks


def compute_row_chunks(table: Table, max_rows: int):
    """Returns (header_rows, chunks, col_widths_twips, grid, n_rows, n_cols,
    row_weights)."""
    col_widths_twips, grid, header_rows, n_rows, n_cols = parse_table_grid(table)
    row_weights = estimate_row_line_counts(grid, col_widths_twips, n_rows, n_cols)
    data_row_indices = [r for r in range(n_rows) if r not in header_rows]
    bands = group_into_bands(grid, data_row_indices, n_cols)
    chunks = build_chunks(bands, row_weights, max_rows)
    return header_rows, chunks, col_widths_twips, grid, n_rows, n_cols, row_weights


# ===========================================================================
# ENGINE 1 (default): LibreOffice-backed rendering
# ===========================================================================
def find_soffice(explicit_path: Optional[str]) -> Optional[str]:
    if explicit_path:
        return explicit_path if os.path.exists(explicit_path) else None

    for exe_name in ("soffice", "soffice.exe", "libreoffice"):
        found = shutil.which(exe_name)
        if found:
            return found

    system = platform.system()
    candidates = []
    if system == "Windows":
        for base in (
            os.environ.get("PROGRAMFILES", r"C:\Program Files"),
            os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"),
        ):
            if base:
                candidates.append(os.path.join(base, "LibreOffice", "program", "soffice.exe"))
    elif system == "Darwin":
        candidates.append("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    else:
        candidates.extend(["/usr/bin/soffice", "/usr/local/bin/soffice", "/opt/libreoffice/program/soffice"])
        candidates.extend(glob.glob("/opt/libreoffice*/program/soffice"))

    for c in candidates:
        if c and os.path.exists(c):
            return c
    return None


def _set_tblpr_child(tblPr, tag_localname: str, attrib: dict):
    """Get-or-create a direct child of <w:tblPr> named w:{tag_localname},
    inserting it at the schema-correct position if newly created, then set
    the given attributes on it. Returns the element."""
    tag_q = qn(f"w:{tag_localname}")
    el = tblPr.find(tag_q)
    if el is None:
        el = etree.Element(tag_q)
        try:
            new_idx = TBLPR_CHILD_ORDER.index(tag_localname)
        except ValueError:
            new_idx = len(TBLPR_CHILD_ORDER)
        insert_at = len(tblPr)
        for i, child in enumerate(tblPr):
            child_tag = etree.QName(child).localname
            try:
                child_idx = TBLPR_CHILD_ORDER.index(child_tag)
            except ValueError:
                child_idx = len(TBLPR_CHILD_ORDER)
            if child_idx > new_idx:
                insert_at = i
                break
        tblPr.insert(insert_at, el)
    for k, v in attrib.items():
        el.set(qn(f"w:{k}"), v)
    return el


def estimate_page_height_twips(rows_to_keep: list, row_weights: list) -> int:
    """Estimate a safe page height (in twips) for a chunk containing the
    given rows, based on the ESTIMATED VISIBLE LINE COUNT of those rows
    (not just their count), so the temporary single-table document's custom
    page is large enough to fit the whole chunk on one page (avoiding
    unwanted pagination) without being wastefully huge."""
    total_lines = sum(row_weights[r] for r in rows_to_keep)
    n_rows = len(rows_to_keep)
    est = total_lines * LINE_HEIGHT_ESTIMATE_TWIPS + n_rows * ROW_CHROME_ESTIMATE_TWIPS
    return int(min(MAX_CHUNK_PAGE_HEIGHT_TWIPS, max(MIN_CHUNK_PAGE_HEIGHT_TWIPS, est)))


def prune_document_to_single_table(docx_path: Path, table_index: int, rows_to_keep: list, page_height_twips: int):
    """Reload docx_path fresh and strip its body down to ONLY the table at
    `table_index` (matched by position among top-level <w:tbl> children),
    keeping only the <w:tr> rows whose indices are in rows_to_keep (in that
    same relative order).

    Also, to guarantee the final render contains NOTHING but the table
    itself (no page header/footer, no clipped columns):
      * strips any headerReference/footerReference from the section
        properties, so Word's running page header/footer (title, logo,
        page numbers, etc.) never gets drawn,
      * forces the table's layout to "fixed" and its indentation to 0 and
        its declared width to an absolute (dxa) value matching the sum of
        its own column widths, so the table can never be auto-refitted,
        indented, or stretched/shrunk to some other width by the layout
        engine,
      * sets a custom page size sized (with a safety buffer) to exactly
        fit the table's real width and the given estimated page height.

    Returns (document, table_width_twips).
    """
    document = Document(str(docx_path))
    body = document.element.body

    tables = body.findall(qn("w:tbl"))
    target = tables[table_index]

    sectPr = body.find(qn("w:sectPr"))
    for child in list(body):
        if child is not target and child is not sectPr:
            body.remove(child)

    trs = target.findall(qn("w:tr"))
    keep_set = set(rows_to_keep)
    for i, tr in enumerate(trs):
        if i not in keep_set:
            target.remove(tr)

    # --- Strip page headers/footers so ONLY the table renders -------------
    if sectPr is not None:
        for tag in ("headerReference", "footerReference"):
            for el in sectPr.findall(qn(f"w:{tag}")):
                sectPr.remove(el)
        titlePg = sectPr.find(qn("w:titlePg"))
        if titlePg is not None:
            sectPr.remove(titlePg)

    # --- Compute the table's true width from its own column grid ----------
    tblGrid = target.find(qn("w:tblGrid"))
    table_width_twips = 0.0
    if tblGrid is not None:
        for gridCol in tblGrid.findall(qn("w:gridCol")):
            w = gridCol.get(qn("w:w"))
            table_width_twips += float(w) if w else 1000.0
    if table_width_twips <= 0:
        table_width_twips = 12000.0

    # --- Lock the table's own layout/indent/width so nothing else (page
    #     width, autofit, pct-based width, inherited indent) can cause a
    #     column to be resized, shifted, or clipped ------------------------
    tblPr = target.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = etree.Element(qn("w:tblPr"))
        target.insert(0, tblPr)

    _set_tblpr_child(tblPr, "tblLayout", {"type": "fixed"})
    _set_tblpr_child(tblPr, "tblInd", {"w": "0", "type": "dxa"})
    _set_tblpr_child(tblPr, "tblW", {"w": str(int(table_width_twips)), "type": "dxa"})

    page_width_twips = int(
        table_width_twips + 2 * CHUNK_PAGE_MARGIN_TWIPS + CHUNK_WIDTH_SAFETY_BUFFER_TWIPS
    )

    if sectPr is None:
        sectPr = etree.SubElement(body, qn("w:sectPr"))

    for tag in ("w:pgSz", "w:pgMar"):
        el = sectPr.find(qn(tag))
        if el is not None:
            sectPr.remove(el)

    pgSz = etree.SubElement(sectPr, qn("w:pgSz"))
    pgSz.set(qn("w:w"), str(page_width_twips))
    pgSz.set(qn("w:h"), str(page_height_twips))
    pgSz.set(qn("w:orient"), "landscape")

    pgMar = etree.SubElement(sectPr, qn("w:pgMar"))
    m = str(int(CHUNK_PAGE_MARGIN_TWIPS))
    for side in ("top", "right", "bottom", "left"):
        pgMar.set(qn(f"w:{side}"), m)
    for side in ("header", "footer", "gutter"):
        pgMar.set(qn(f"w:{side}"), "0")

    return document, table_width_twips


def convert_docx_batch_to_pdf(soffice_path: str, docx_paths: list, out_dir: Path):
    for i in range(0, len(docx_paths), SOFFICE_BATCH_SIZE):
        batch = docx_paths[i : i + SOFFICE_BATCH_SIZE]
        cmd = [
            soffice_path,
            "--headless",
            "--norestore",
            "--convert-to",
            "pdf",
            "--outdir",
            str(out_dir),
        ] + [str(p) for p in batch]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice conversion failed (exit {result.returncode}):\n"
                f"stdout: {result.stdout}\nstderr: {result.stderr}"
            )


def render_pdf_to_cropped_png(
    pdf_path: Path,
    dpi: int,
    out_path: Path,
    table_width_twips: float,
    pad_px: int = 4,
):
    """Rasterize page 1 of the (single-table) PDF and crop it tightly to the
    table's content.

    The horizontal crop bounds are computed DETERMINISTICALLY from the
    table's own known width (margin -> margin + table_width), rather than
    "auto-detected" from whichever pixels happen to be non-white -- this
    guarantees no column is ever clipped away by a faulty whitespace
    heuristic. Only the BOTTOM edge (where the table's content actually
    ends) is auto-detected, since real per-row rendered height depends on
    Word/LibreOffice's own text wrapping and isn't known in advance.
    """
    import fitz  # PyMuPDF
    from PIL import Image, ImageChops

    pdf_doc = fitz.open(str(pdf_path))
    n_pages = len(pdf_doc)
    if n_pages == 0:
        raise RuntimeError(f"LibreOffice produced an empty PDF for {pdf_path.name}")
    if n_pages > 1:
        print(
            f"  [!] Warning: '{pdf_path.stem}' overflowed onto {n_pages} pages "
            f"(table content taller than the estimated page height); only the "
            f"first page was used. Consider lowering --max-rows."
        )

    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    page = pdf_doc[0]
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

    bg = Image.new("RGB", img.size, (255, 255, 255))
    diff = ImageChops.difference(img, bg)
    bbox = diff.getbbox()
    top = bbox[1] if bbox else 0
    bottom = bbox[3] if bbox else img.height

    margin_px = CHUNK_PAGE_MARGIN_TWIPS / TWIPS_PER_INCH * dpi
    table_width_px = table_width_twips / TWIPS_PER_INCH * dpi
    left = margin_px
    right = margin_px + table_width_px

    l = max(0, int(round(left)) - pad_px)
    r = min(img.width, int(round(right)) + pad_px)
    t = max(0, top - pad_px)
    b = min(img.height, bottom + pad_px)

    if r <= l:
        l, r = 0, img.width
    if b <= t:
        t, b = 0, img.height

    img = img.crop((l, t, r, b))

    out_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(out_path, "PNG")
    pdf_doc.close()


def process_docx_libreoffice(docx_path: Path, out_dir: Path, dpi: int, max_rows: int, soffice_path: str, keep_temp: bool):
    document = Document(str(docx_path))
    sections = discover_sections(document)
    if not sections:
        print(f"  [!] No tables found in {docx_path.name}")
        return []

    used_names: dict = {}
    # plan[i] = (out_path, table_index, rows_to_keep, n_data_rows, n_visible_lines)
    plan = []
    for heading_text, table_index, table in sections:
        base_name = sanitize_filename_part(heading_text or "Section")
        if base_name in used_names:
            used_names[base_name] += 1
            section_name = f"{base_name}_{used_names[base_name]}"
        else:
            used_names[base_name] = 1
            section_name = base_name

        header_rows, chunks, _cw, _grid, n_rows, _n_cols, row_weights = compute_row_chunks(table, max_rows)
        if n_rows == 0:
            continue
        if not chunks:
            print(f"  [!] Section '{section_name}': table has no data rows, skipped")
            continue

        for idx, chunk_rows in enumerate(chunks, start=1):
            rows_to_keep = sorted(set(header_rows) | set(chunk_rows))
            n_visible_lines = sum(row_weights[r] for r in chunk_rows)
            out_path = out_dir / f"{section_name}-{idx}.png"
            plan.append((out_path, table_index, rows_to_keep, len(chunk_rows), n_visible_lines, chunk_rows, row_weights))

    if not plan:
        return []

    tmp_root = Path(tempfile.mkdtemp(prefix="landsdetail2img_"))
    try:
        docx_tmp_dir = tmp_root / "docx"
        pdf_tmp_dir = tmp_root / "pdf"
        docx_tmp_dir.mkdir(parents=True, exist_ok=True)
        pdf_tmp_dir.mkdir(parents=True, exist_ok=True)

        temp_docx_paths = []
        table_widths = []
        for i, (out_path, table_index, rows_to_keep, _n, _nl, _chunk_rows, row_weights) in enumerate(plan):
            page_h = estimate_page_height_twips(rows_to_keep, row_weights)
            pruned_doc, table_width_twips = prune_document_to_single_table(
                docx_path, table_index, rows_to_keep, page_h
            )
            temp_docx_path = docx_tmp_dir / f"chunk_{i:05d}.docx"
            pruned_doc.save(str(temp_docx_path))
            temp_docx_paths.append(temp_docx_path)
            table_widths.append(table_width_twips)

        print(f"  Rendering {len(temp_docx_paths)} chunk(s) via LibreOffice ...")
        convert_docx_batch_to_pdf(soffice_path, temp_docx_paths, pdf_tmp_dir)

        generated_files = []
        for i, (out_path, table_index, rows_to_keep, n_data_rows, n_visible_lines, _chunk_rows, _rw) in enumerate(plan):
            pdf_path = pdf_tmp_dir / f"chunk_{i:05d}.pdf"
            if not pdf_path.exists():
                print(f"  [!] Missing expected PDF for {out_path.name}, skipped")
                continue
            render_pdf_to_cropped_png(pdf_path, dpi, out_path, table_widths[i])
            generated_files.append(out_path)
            print(f"  -> {out_path}  ({n_data_rows} data row(s), ~{n_visible_lines} visible line(s))")

        return generated_files
    finally:
        if keep_temp:
            print(f"  [i] Temp files kept at: {tmp_root}")
        else:
            shutil.rmtree(tmp_root, ignore_errors=True)


# ===========================================================================
# ENGINE 2 (fallback): Pillow-backed rendering
# ===========================================================================
def compute_row_heights(grid, col_widths_px, n_rows, n_cols, dpi):
    pad_tb_px = twips_to_px(CELL_TOP_BOTTOM_PAD_TWIPS, dpi)
    min_height = [0.0] * n_rows
    seen_ids = set()

    for r in range(n_rows):
        for c in range(n_cols):
            cell = grid[r][c]
            if cell is None or cell.row0 != r or cell.col0 != c:
                continue
            if id(cell) in seen_ids:
                continue
            seen_ids.add(id(cell))

            width_px = sum(col_widths_px[cell.col0 : cell.col0 + cell.colspan])
            width_px = max(1.0, width_px - 2 * twips_to_px(CELL_LEFT_RIGHT_PAD_TWIPS, dpi))
            sub_lines = wrap_cell_lines(cell, width_px, dpi)
            content_h = sum(line_height_px(sl, dpi) for sl in sub_lines) + 2 * pad_tb_px

            if cell.rowspan == 1:
                min_height[r] = max(min_height[r], content_h)
            else:
                span = range(cell.row0, cell.row0 + cell.rowspan)
                current_sum = sum(min_height[rr] for rr in span)
                if content_h > current_sum:
                    min_height[cell.row0 + cell.rowspan - 1] += content_h - current_sum

    default_font = get_font(DEFAULT_FONT_NAME, DEFAULT_FONT_SIZE_PT, False, False, dpi)
    fallback_h = default_font.size * LINE_SPACING_FACTOR + 2 * pad_tb_px
    for r in range(n_rows):
        if min_height[r] <= 0:
            min_height[r] = fallback_h
    return min_height


def render_rows_to_image_pillow(rows, grid, col_widths_px, row_heights, n_cols, dpi, out_path: Path):
    from PIL import Image, ImageDraw

    border_w = max(1, round(pt_to_px(BORDER_WIDTH_PT, dpi)))
    pad_lr_px = twips_to_px(CELL_LEFT_RIGHT_PAD_TWIPS, dpi)
    pad_tb_px = twips_to_px(CELL_TOP_BOTTOM_PAD_TWIPS, dpi)

    width_px = int(round(sum(col_widths_px))) + border_w
    height_px = int(round(sum(row_heights[r] for r in rows))) + border_w

    img = Image.new("RGB", (max(1, width_px), max(1, height_px)), "white")
    draw = ImageDraw.Draw(img)

    row_index_of = {r: i for i, r in enumerate(rows)}
    y_cursor = 0.0
    for r in rows:
        x_cursor = 0.0
        row_h = row_heights[r]
        for c in range(n_cols):
            cell = grid[r][c]
            col_w = col_widths_px[c]
            if cell is not None and cell.row0 == r and cell.col0 == c:
                cell_w = sum(col_widths_px[cell.col0 : cell.col0 + cell.colspan])
                span_rows = [rr for rr in range(cell.row0, cell.row0 + cell.rowspan) if rr in row_index_of]
                cell_h = sum(row_heights[rr] for rr in span_rows) if span_rows else row_h

                x0, y0 = x_cursor, y_cursor
                x1, y1 = x0 + cell_w, y0 + cell_h

                if cell.shading:
                    draw.rectangle([x0, y0, x1, y1], fill=cell.shading)
                draw.rectangle([x0, y0, x1, y1], outline=BORDER_COLOR, width=border_w)

                max_text_w = max(1.0, cell_w - 2 * pad_lr_px)
                sub_lines = wrap_cell_lines(cell, max_text_w, dpi)
                ty = y0 + pad_tb_px
                for sl in sub_lines:
                    lh = line_height_px(sl, dpi)
                    tx = x0 + pad_lr_px
                    for text, font, color in sl:
                        if text:
                            draw.text((tx, ty), text, font=font, fill=color)
                            tx += font.getlength(text)
                    ty += lh
            x_cursor += col_w
        y_cursor += row_h

    out_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(out_path, "PNG")


def process_docx_pillow(docx_path: Path, out_dir: Path, dpi: int, max_rows: int):
    document = Document(str(docx_path))
    sections = discover_sections(document)
    if not sections:
        print(f"  [!] No tables found in {docx_path.name}")
        return []

    used_names: dict = {}
    generated_files = []

    for heading_text, _table_index, table in sections:
        base_name = sanitize_filename_part(heading_text or "Section")
        if base_name in used_names:
            used_names[base_name] += 1
            section_name = f"{base_name}_{used_names[base_name]}"
        else:
            used_names[base_name] = 1
            section_name = base_name

        header_rows, chunks, col_widths_twips, grid, n_rows, n_cols, row_weights = compute_row_chunks(table, max_rows)
        if n_rows == 0 or n_cols == 0:
            continue
        if not chunks:
            print(f"  [!] Section '{section_name}': table has no data rows, skipped")
            continue

        col_widths_px = [twips_to_px(w, dpi) for w in col_widths_twips]
        row_heights = compute_row_heights(grid, col_widths_px, n_rows, n_cols, dpi)

        for idx, chunk_rows in enumerate(chunks, start=1):
            rows = header_rows + chunk_rows
            n_visible_lines = sum(row_weights[r] for r in chunk_rows)
            out_path = out_dir / f"{section_name}-{idx}.png"
            render_rows_to_image_pillow(rows, grid, col_widths_px, row_heights, n_cols, dpi, out_path)
            generated_files.append(out_path)
            print(f"  -> {out_path}  ({len(chunk_rows)} data row(s), ~{n_visible_lines} visible line(s))")

    return generated_files


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Convert long Word (.docx) table sections into PNG screenshots."
    )
    parser.add_argument("docx_files", nargs="+", help="One or more source .docx files")
    parser.add_argument(
        "-o", "--output-dir", default="output", help="Directory to write PNG files into (default: ./output)"
    )
    parser.add_argument(
        "--dpi", type=int, default=DEFAULT_DPI, help=f"Render resolution in DPI (default: {DEFAULT_DPI})"
    )
    parser.add_argument(
        "--max-rows",
        type=int,
        default=DEFAULT_MAX_ROWS,
        help=(
            f"Max VISIBLE ROWS (rendered lines, counting both text-wrapping and "
            f"explicit line breaks -- NOT logical table rows) packed into a single "
            f"image (default: {DEFAULT_MAX_ROWS}). A row with a lot of text may by "
            f"itself take up many visible rows; long-celled tables will therefore "
            f"typically pack far fewer than {DEFAULT_MAX_ROWS} logical rows per image."
        ),
    )
    parser.add_argument(
        "--engine",
        choices=["libreoffice", "pillow"],
        default="libreoffice",
        help=(
            "Rendering backend. 'libreoffice' (default) reuses your original "
            "table's real formatting and renders it with a real word-processor "
            "layout engine for maximum fidelity, but requires LibreOffice to be "
            "installed. 'pillow' is a pure-Python, dependency-light approximation."
        ),
    )
    parser.add_argument(
        "--soffice-path",
        default=None,
        help="Explicit path to the LibreOffice 'soffice' executable, if it isn't auto-detected.",
    )
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        help="Keep the temporary per-chunk .docx/.pdf files (for debugging the libreoffice engine).",
    )
    args = parser.parse_args()

    out_dir = Path(args.output_dir)

    soffice_path = None
    if args.engine == "libreoffice":
        soffice_path = find_soffice(args.soffice_path)
        if not soffice_path:
            print(
                "[!] Could not find a LibreOffice installation (the 'soffice' executable).\n"
                "    Options:\n"
                "      1. Install LibreOffice (free): https://www.libreoffice.org/download/\n"
                "         then re-run this command.\n"
                "      2. If it's already installed somewhere non-standard, pass its path via:\n"
                "         --soffice-path \"C:\\Path\\To\\soffice.exe\"\n"
                "      3. Or use the lower-fidelity, dependency-free fallback engine:\n"
                "         --engine pillow\n",
                file=sys.stderr,
            )
            sys.exit(1)

    all_generated = []
    for docx_file in args.docx_files:
        docx_path = Path(docx_file)
        if not docx_path.exists():
            print(f"[!] File not found: {docx_path}", file=sys.stderr)
            continue
        print(f"Processing {docx_path.name} ...")
        if args.engine == "libreoffice":
            files = process_docx_libreoffice(
                docx_path, out_dir, args.dpi, args.max_rows, soffice_path, args.keep_temp
            )
        else:
            files = process_docx_pillow(docx_path, out_dir, args.dpi, args.max_rows)
        if files:
            all_generated.extend(files)

    print(f"\nDone. Generated {len(all_generated)} image(s) in '{out_dir}/'.")


if __name__ == "__main__":
    main()
