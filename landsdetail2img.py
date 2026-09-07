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
    * every image always starts with the header row(s) of the table,
    * every image contains at least one *complete* data row (a data row / a
      group of vertically-merged rows is NEVER split across two images),
    * as many additional whole data rows as possible are packed into the
      same image, as long as the total number of data rows in that image
      does not exceed --max-rows (default: 40).

Output file naming:
    <section-name>-<1-based index>.png

Usage
-----
    python landsdetail2img.py "QC FUP v1.2.docx" -o out_dir
    python landsdetail2img.py report1.docx report2.docx -o out_dir --dpi 200 --max-rows 30

Run `python landsdetail2img.py -h` for the full list of options.
"""

from __future__ import annotations

import argparse
import os
import platform
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from lxml import etree
from PIL import Image, ImageDraw, ImageFont

# ---------------------------------------------------------------------------
# Configuration defaults (all overridable via CLI flags, see --help)
# ---------------------------------------------------------------------------
DEFAULT_DPI = 150
DEFAULT_MAX_ROWS = 40
TWIPS_PER_INCH = 1440.0

BORDER_COLOR = (0, 0, 0)
BORDER_WIDTH_PT = 0.5  # matches the thin "Table Grid" borders used in Word

DEFAULT_TEXT_COLOR = (0, 0, 0)
DEFAULT_FONT_NAME = "Calibri"
DEFAULT_FONT_SIZE_PT = 11.0

# Left/right cell padding mirrors Word's default table cell margins (108 twips
# ~= 0.19cm). Top/bottom padding is a small cosmetic inset since Word's own
# "Table Grid" style uses 0 twips there, which looks too cramped as an image.
CELL_LEFT_RIGHT_PAD_TWIPS = 108
CELL_TOP_BOTTOM_PAD_TWIPS = 40
LINE_SPACING_FACTOR = 1.20

# ---------------------------------------------------------------------------
# Cross-platform font resolution
# ---------------------------------------------------------------------------
# Word documents reference fonts by *family name* (e.g. "Times New Roman"),
# but Pillow's ImageFont.truetype() needs an actual *file path* on disk.
# Font file locations/names differ wildly between Windows, macOS and Linux,
# so instead of hardcoding one platform's paths, we build an index of every
# .ttf/.ttc/.otf file we can find on the current machine (keyed by lowercase
# filename), and then look up the best available match for each requested
# family/style. If nothing suitable is found anywhere, we fall back to
# Pillow's bundled default font so the script never crashes for a missing
# font file.
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

# For each Word font family (lower-cased) we care about, list candidate
# (regular, bold, italic, bold-italic) filenames across every platform. Only
# the ones that actually exist in the discovered font index will be used.
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

# Used when the requested family isn't in FONT_CANDIDATES at all, or none of
# its candidate files can be found on this machine.
GENERIC_FALLBACK_GROUPS = [
    ("arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"),
    ("calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf"),
    ("LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf", "LiberationSans-Italic.ttf", "LiberationSans-BoldItalic.ttf"),
    ("DejaVuSans.ttf", "DejaVuSans-Bold.ttf", "DejaVuSans-Oblique.ttf", "DejaVuSans-BoldOblique.ttf"),
]

_FONT_INDEX: Optional[dict] = None


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

    # Absolute last resort: grab literally any installed font so we still
    # render *something* instead of crashing.
    if index:
        return next(iter(index.values()))
    return None


_FONT_CACHE: dict = {}


def twips_to_px(twips: float, dpi: int) -> float:
    return twips / TWIPS_PER_INCH * dpi


def pt_to_px(pt: float, dpi: int) -> float:
    return pt / 72.0 * dpi


def get_font(name: str, size_pt: float, bold: bool, italic: bool, dpi: int):
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
        # No usable TrueType font found anywhere on this machine; fall back
        # to Pillow's bundled default font rather than crashing.
        try:
            font = ImageFont.load_default(size=size_px)
        except TypeError:
            # Older Pillow versions don't accept a `size` kwarg here.
            font = ImageFont.load_default()

    _FONT_CACHE[key] = font
    return font


# ---------------------------------------------------------------------------
# Data structures
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
# XML parsing helpers
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
# Text wrapping
# ---------------------------------------------------------------------------
def _split_long_word(word: str, font, max_width_px: float) -> list:
    """Break a single very long 'word' (e.g. a URL) into pieces that each fit
    within max_width_px, so it doesn't overflow the cell horizontally."""
    if font.getlength(word) <= max_width_px or len(word) <= 1:
        return [word]
    pieces = []
    current = ""
    for ch in word:
        trial = current + ch
        if current and font.getlength(trial) > max_width_px:
            pieces.append(current)
            current = ch
        else:
            current = trial
    if current:
        pieces.append(current)
    return pieces


def wrap_line(segments: list, max_width_px: float, dpi: int) -> list:
    """Word-wrap one explicit line (list[Segment]) to fit max_width_px.

    Returns a list of "visual sub-lines", each a list of (text, font, color)
    tuples ready to be drawn left-to-right.
    """
    subs: list = []
    current: list = []
    current_width = 0.0

    def flush():
        nonlocal current, current_width
        subs.append(current)
        current = []
        current_width = 0.0

    for seg in segments:
        font = get_font(seg.font_name, seg.size_pt, seg.bold, seg.italic, dpi)
        space_w = font.getlength(" ")
        # Preserve simple whitespace splitting; consecutive spaces collapse
        # into single separators (visually indistinguishable in practice).
        words = [w for w in seg.text.split(" ")]
        for i, raw_word in enumerate(words):
            if raw_word == "":
                if i != 0 and current:
                    if current_width + space_w <= max_width_px:
                        current.append((" ", font, seg.color))
                        current_width += space_w
                continue
            for word in _split_long_word(raw_word, font, max_width_px):
                word_w = font.getlength(word)
                sep_w = space_w if current else 0.0
                if current and current_width + sep_w + word_w > max_width_px:
                    flush()
                    current.append((word, font, seg.color))
                    current_width = word_w
                else:
                    if current:
                        current.append((" ", font, seg.color))
                        current_width += sep_w
                    current.append((word, font, seg.color))
                    current_width += word_w
    if current or not subs:
        subs.append(current)
    return subs


def wrap_cell(cell: CellData, max_width_px: float, dpi: int) -> list:
    """Return list of visual sub-lines for the whole cell (all explicit
    lines, word-wrapped)."""
    all_subs = []
    for line_segments in cell.raw_lines:
        if not line_segments:
            all_subs.append([])  # blank explicit line
        else:
            all_subs.extend(wrap_line(line_segments, max_width_px, dpi))
    return all_subs


def line_height_px(sub_line: list, dpi: int) -> float:
    if not sub_line:
        default_font = get_font(DEFAULT_FONT_NAME, DEFAULT_FONT_SIZE_PT, False, False, dpi)
        return default_font.size * LINE_SPACING_FACTOR
    max_size = max(font.size for _, font, _ in sub_line)
    return max_size * LINE_SPACING_FACTOR


# ---------------------------------------------------------------------------
# Section / heading discovery
# ---------------------------------------------------------------------------
def sanitize_filename_part(text: str) -> str:
    text = text.strip()
    text = re.sub(r"[\\/:*?\"<>|]", "_", text)
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_.")
    return text or "Section"


def discover_sections(document: Document):
    """Walk the document body top-to-bottom and pair every table with the
    closest preceding Heading-styled paragraph text."""
    body = document.element.body
    sections = []
    current_heading = None
    for child in body.iterchildren():
        tag = etree.QName(child).localname
        if tag == "p":
            para = Paragraph(child, document)
            style_name = para.style.name if para.style is not None else ""
            if style_name and style_name.lower().startswith("heading") and para.text.strip():
                current_heading = para.text.strip()
        elif tag == "tbl":
            table = Table(child, document)
            sections.append((current_heading, table))
    return sections


# ---------------------------------------------------------------------------
# Row height computation
# ---------------------------------------------------------------------------
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
            sub_lines = wrap_cell(cell, width_px, dpi)
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


# ---------------------------------------------------------------------------
# Row grouping (bands that must stay together) + chunking (<= max_rows)
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


def build_chunks(bands, max_rows):
    chunks = []
    current_rows: list = []
    current_count = 0
    for band in bands:
        band_len = len(band)
        if current_rows and current_count + band_len > max_rows:
            chunks.append(current_rows)
            current_rows = []
            current_count = 0
        current_rows.extend(band)
        current_count += band_len
    if current_rows:
        chunks.append(current_rows)
    return chunks


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------
def render_rows_to_image(rows, grid, col_widths_px, row_heights, n_cols, dpi, out_path: Path):
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
                # height may span rows not all present in this image (shouldn't
                # happen because bands keep merges together, but guard anyway)
                span_rows = [rr for rr in range(cell.row0, cell.row0 + cell.rowspan) if rr in row_index_of]
                cell_h = sum(row_heights[rr] for rr in span_rows) if span_rows else row_h

                x0, y0 = x_cursor, y_cursor
                x1, y1 = x0 + cell_w, y0 + cell_h

                if cell.shading:
                    draw.rectangle([x0, y0, x1, y1], fill=cell.shading)
                draw.rectangle([x0, y0, x1, y1], outline=BORDER_COLOR, width=border_w)

                max_text_w = max(1.0, cell_w - 2 * pad_lr_px)
                sub_lines = wrap_cell(cell, max_text_w, dpi)
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


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------
def process_docx(docx_path: Path, out_dir: Path, dpi: int, max_rows: int):
    document = Document(str(docx_path))
    sections = discover_sections(document)

    if not sections:
        print(f"  [!] No tables found in {docx_path.name}")
        return

    used_names: dict = {}
    generated_files = []

    for heading_text, table in sections:
        base_name = sanitize_filename_part(heading_text or "Section")
        if base_name in used_names:
            used_names[base_name] += 1
            section_name = f"{base_name}_{used_names[base_name]}"
        else:
            used_names[base_name] = 1
            section_name = base_name

        col_widths_twips, grid, header_rows, n_rows, n_cols = parse_table_grid(table)
        if n_rows == 0 or n_cols == 0:
            continue
        col_widths_px = [twips_to_px(w, dpi) for w in col_widths_twips]
        row_heights = compute_row_heights(grid, col_widths_px, n_rows, n_cols, dpi)

        data_row_indices = [r for r in range(n_rows) if r not in header_rows]
        bands = group_into_bands(grid, data_row_indices, n_cols)
        chunks = build_chunks(bands, max_rows)

        if not chunks:
            print(f"  [!] Section '{section_name}': table has no data rows, skipped")
            continue

        for idx, chunk_rows in enumerate(chunks, start=1):
            rows = header_rows + chunk_rows
            out_path = out_dir / f"{section_name}-{idx}.png"
            render_rows_to_image(rows, grid, col_widths_px, row_heights, n_cols, dpi, out_path)
            generated_files.append(out_path)
            print(f"  -> {out_path}  ({len(chunk_rows)} data row(s))")

    return generated_files


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
        help=f"Max data rows packed into a single image (default: {DEFAULT_MAX_ROWS})",
    )
    args = parser.parse_args()

    out_dir = Path(args.output_dir)
    all_generated = []
    for docx_file in args.docx_files:
        docx_path = Path(docx_file)
        if not docx_path.exists():
            print(f"[!] File not found: {docx_path}", file=sys.stderr)
            continue
        print(f"Processing {docx_path.name} ...")
        files = process_docx(docx_path, out_dir, args.dpi, args.max_rows)
        if files:
            all_generated.extend(files)

    print(f"\nDone. Generated {len(all_generated)} image(s) in '{out_dir}/'.")


if __name__ == "__main__":
    main()
