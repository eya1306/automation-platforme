#!/usr/bin/env python3
"""
excel_diff.py - Excel Change Review

Compares two Excel workbooks and writes out two highlighted copies:

    <name>_OLD.xlsx   changed / deleted cells filled RED
    <name>_NEW.xlsx   changed / added cells filled GREEN,
                      with a note saying what was removed

Every sheet is compared, and every kind of difference is caught: modified
values, changed formulas, added cells, deleted cells, added sheets and deleted
sheets. Nothing to configure. The two originals are never modified.

Rows are ALIGNED before they are compared, so inserting or deleting a row does
not make every row beneath it look modified -- an inserted row is reported as
one insertion and the rest still line up against their real counterparts. Pass
the blank template with --template and the template's own headings and labels
are left out of the comparison entirely.

Any spreadsheet LibreOffice can open is accepted: .xlsx, .xlsm, .xltx, .xltm,
.xls, .xlt, .xlsb, .ods, .csv and the rest. The highlighted copies always come
out as .xlsx (or .xlsm), because the old .xls container cannot carry the fills
and notes this tool adds.

Both copies keep the template exactly as it was authored -- the same layout,
cell styles, borders, merged ranges, column widths, images, formulas and sheet
tab colours. The only thing added is a fill on the cells that actually differ,
plus the note explaining each one.

Run it
------
    python3 excel_diff.py                     -> opens the window
    python3 excel_diff.py old.xlsx new.xlsx   -> command line
"""

from __future__ import annotations

import argparse
import copy
import datetime as _dt
import difflib
import os
import re
import shutil
import subprocess
import tempfile
import sys
import threading
import queue
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter, column_index_from_string
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is required:  pip install openpyxl")


# ==========================================================================
# Configuration
# ==========================================================================

FILL_RED = "FFFFC7CE"   # old side: changed or deleted   /  new side: deleted
FILL_GREEN = "FFC6EFCE" # new side: changed or added

# There is deliberately no tab colour here. The workbook template is left
# exactly as it was authored -- sheet tab colours, cell styles, borders, merges,
# column widths, images and formulas all survive the round trip untouched. The
# only thing this tool paints is the fill of a cell that actually differs.

FLOAT_TOLERANCE = 1e-9  # relative tolerance for numeric comparison
NUMERIC_DECIMALS = 10   # numbers are rounded here before they are compared
COMMENT_AUTHOR = "Excel Change Review"
MAX_NOTE_LEN = 250      # long cell values are trimmed inside the note

EXCEL_TYPES = [("Excel files", "*.xlsx *.xlsm *.xlsb *.xltx *.xltm *.xls *.xlt *.ods *.csv"),
               ("Excel 2007+", "*.xlsx *.xlsm *.xlsb *.xltx *.xltm"),
               ("Excel 97-2003", "*.xls *.xlt"),
               ("OpenDocument / text", "*.ods *.fods *.csv *.tsv"),
               ("All files", "*.*")]

APP_TITLE = "Excel Change Review"
__version__ = "2.2 (row-aligned, namespace-safe, scrollable)"
APP_SUBTITLE = ("Outputs TWO highlighted files: changes are RED in the original and GREEN "
                "in the modified file.\nDeleted cells carry a note recording what was removed.")


@dataclass
class Change:
    sheet: str
    cell: str
    kind: str          # modified | added | deleted | sheet-added | sheet-deleted
    old_value: object = None
    new_value: object = None
    whole_row: bool = False
    old_cell: str = None   # coordinate in the OLD file (rows can shift)
    new_cell: str = None   # coordinate in the NEW file


@dataclass
class DiffResult:
    changes: list = field(default_factory=list)
    sheets_only_in_old: list = field(default_factory=list)
    sheets_only_in_new: list = field(default_factory=list)
    out_old: Path = None
    out_new: Path = None
    deleted_rows: list = field(default_factory=list)   # (sheet, row) fully removed
    added_rows: list = field(default_factory=list)     # (sheet, row) newly present

    @property
    def cell_changes(self) -> int:
        return sum(1 for c in self.changes if c.kind in ("modified", "added", "deleted"))

    def count(self, kind: str) -> int:
        return sum(1 for c in self.changes if c.kind == kind)


# ==========================================================================
# Legacy .xls support
# ==========================================================================
#
# openpyxl only reads the modern XML formats. Old .xls workbooks are converted
# to .xlsx first, in a temporary folder -- the user's file is never touched.
# LibreOffice is tried first because it preserves formulas, formatting and
# merged cells. If it is not installed, xlrd is used to salvage the values,
# which is enough to compare but loses the styling.

# What openpyxl can open directly. Everything else is routed through
# LibreOffice first, so the tool accepts any spreadsheet the user throws at it.
NATIVE_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}

# Anything LibreOffice can read, which is every Excel dialect plus the
# OpenDocument and plain-text families.
CONVERTIBLE_SUFFIXES = {
    ".xls", ".xlt", ".xlsb", ".xla", ".xlam", ".xlw", ".xlr",   # Excel dialects
    ".ods", ".ots", ".fods", ".sxc", ".stc",                    # OpenDocument
    ".csv", ".tsv", ".txt", ".dif", ".slk", ".dbf",             # text / interchange
}

# xlrd only understands the old BIFF .xls; it is the fallback when LibreOffice
# is missing, and only for those.
XLRD_SUFFIXES = {".xls", ".xlt"}

SUPPORTED_SUFFIXES = NATIVE_SUFFIXES | CONVERTIBLE_SUFFIXES

LIBREOFFICE_CANDIDATES = (
    "soffice", "libreoffice", "localc",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
)


def find_libreoffice():
    for name in LIBREOFFICE_CANDIDATES:
        found = shutil.which(name)
        if found:
            return found
        if os.path.sep in name and Path(name).exists():
            return name
    return None


def has_vba(path: Path) -> bool:
    """True if this legacy workbook carries a VBA project.

    Directory names inside an OLE compound file are stored as UTF-16, so the
    marker is searched for in that encoding as well as plain bytes.
    """
    try:
        with open(path, "rb") as handle:
            blob = handle.read(4 * 1024 * 1024)
    except OSError:
        return False
    return (b"_VBA_PROJECT" in blob
            or "_VBA_PROJECT".encode("utf-16-le") in blob)


def convert_with_libreoffice(src: Path, dest_dir: Path, keep_macros: bool = False):
    """Convert to .xlsx (or .xlsm) via LibreOffice. Returns the new path, or None."""
    exe = find_libreoffice()
    if not exe:
        return None

    # A workbook with macros has to land in the macro-enabled format, or
    # LibreOffice quietly drops the VBA project on the way out.
    target = 'xlsm:Calc MS Excel 2007 VBA XML' if keep_macros else "xlsx"
    suffix = ".xlsm" if keep_macros else ".xlsx"

    # An isolated profile, so the conversion still works when the user happens
    # to have LibreOffice open (a shared profile makes it exit silently).
    profile = dest_dir / "lo_profile"
    cmd = [exe, f"-env:UserInstallation={profile.as_uri()}",
           "--headless", "--norestore", "--convert-to", target,
           "--outdir", str(dest_dir), str(src)]
    try:
        subprocess.run(cmd, check=True, timeout=240,
                       stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except Exception:
        return None

    produced = dest_dir / (src.stem + suffix)
    if produced.exists():
        return produced
    # Fall back to whatever it actually wrote, if anything.
    for candidate in (dest_dir / (src.stem + ".xlsx"), dest_dir / (src.stem + ".xlsm")):
        if candidate.exists():
            return candidate
    return None


def convert_with_xlrd(src: Path, dest_dir: Path):
    """Fall back to xlrd: rebuild the values as a fresh .xlsx. Returns path or None."""
    try:
        import xlrd
        from openpyxl import Workbook
    except ImportError:
        return None

    try:
        book = xlrd.open_workbook(str(src))
    except Exception:
        return None

    wb = Workbook()
    wb.remove(wb.active)
    for sheet in book.sheets():
        ws = wb.create_sheet((sheet.name or "Sheet")[:31])
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                cell = sheet.cell(r, c)
                if cell.ctype == xlrd.XL_CELL_EMPTY:
                    continue
                value = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        value = xlrd.xldate_as_datetime(value, book.datemode)
                    except Exception:
                        pass
                elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    value = bool(value)
                elif cell.ctype == xlrd.XL_CELL_ERROR:
                    value = "#ERROR"
                ws.cell(row=r + 1, column=c + 1, value=value)

    if not wb.sheetnames:
        wb.create_sheet("Sheet1")
    produced = dest_dir / (src.stem + ".xlsx")
    wb.save(produced)
    return produced


def ensure_modern(path: Path, workdir: Path, label: str, say) -> Path:
    """Return a path openpyxl can read, converting the file first if necessary."""
    suffix = path.suffix.lower()
    if suffix in NATIVE_SUFFIXES:
        return path

    if suffix not in CONVERTIBLE_SUFFIXES:
        # Unknown extension: rather than refuse outright, let LibreOffice try --
        # it sniffs the real format from the bytes, so a mislabelled file still
        # goes through. Only if that fails do we give up.
        say(f"Unrecognised extension '{path.suffix}' on the {label} file - "
            f"attempting to convert it anyway ...")

    # Each file gets its own folder: the two inputs often share a base name.
    dest = workdir / label
    dest.mkdir(parents=True, exist_ok=True)

    say(f"Converting the {label} file ({suffix or 'no extension'}) to .xlsx ...")
    keep_macros = has_vba(path)
    if keep_macros:
        say(f"   the {label} file contains VBA - converting to .xlsm to keep it")
    converted = convert_with_libreoffice(path, dest, keep_macros=keep_macros)
    if converted:
        return converted

    if suffix in XLRD_SUFFIXES:
        say(f"WARNING: LibreOffice is not installed, so the {label} .xls is being read "
            f"with xlrd, which recovers the values only. The output will NOT keep the "
            f"template layout, the styling or the formulas. Install LibreOffice to "
            f"preserve them.")
        converted = convert_with_xlrd(path, dest)
        if converted:
            return converted

    raise ValueError(
        f"The {label} file ('{path.name}') could not be opened.\n\n"
        "Formats openpyxl reads directly: .xlsx, .xlsm, .xltx, .xltm\n"
        "Everything else (.xls, .xlsb, .ods, .csv ...) needs LibreOffice.\n\n"
        "Install one of the following, then try again:\n"
        "  - LibreOffice  (best: handles every format, keeps formatting)\n"
        "  - xlrd:  pip install xlrd   (old .xls only, values only)\n\n"
        "Or open the file in Excel and re-save it as .xlsx."
    )


# ==========================================================================
# Comparison
# ==========================================================================

# Every value is reduced to a canonical, hashable KEY. Two cells are the same
# cell when their keys match. Working in keys rather than in raw values does
# two jobs at once: it makes the comparison immune to the noise a format
# conversion introduces (100 vs 100.0 vs "100", a stray non-breaking space, a
# date that arrives as a datetime), and it makes a whole row hashable, which is
# what the row alignment below needs.

_WHITESPACE = re.compile(r"\s+")
_NUMERIC = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$")


def clean_text(text: str) -> str:
    """Collapse every run of whitespace -- including the non-breaking kind that
    Excel and LibreOffice sprinkle differently -- down to a single space."""
    return _WHITESPACE.sub(" ", str(text).replace("\u00a0", " ")).strip()


def number_key(number):
    value = float(number)
    if value != value or value in (float("inf"), float("-inf")):
        return ("n", repr(value))
    value = round(value, NUMERIC_DECIMALS)
    if value == int(value) and abs(value) < 1e15:
        value = float(int(value))          # 100.0 and 100 land on the same key
    return ("n", value)


def value_key(value):
    """Canonical key for one value. None means 'this cell holds nothing'."""
    if value is None:
        return None
    if isinstance(value, bool):
        return ("b", value)
    if isinstance(value, _dt.datetime):
        if value.time() == _dt.time(0, 0):
            return ("d", value.date().isoformat())
        return ("d", value.replace(microsecond=0).isoformat())
    if isinstance(value, _dt.date):
        return ("d", value.isoformat())
    if isinstance(value, _dt.time):
        return ("t", value.replace(microsecond=0).isoformat())
    if isinstance(value, (int, float)):
        return number_key(value)

    text = clean_text(value)
    if not text:
        return None
    lowered = text.lower()
    if lowered in ("true", "false"):
        return ("b", lowered == "true")
    # A number typed as text in one file and as a number in the other is the
    # same number, not a change.
    compact = text.replace(" ", "")
    if _NUMERIC.match(compact):
        try:
            return number_key(float(compact))
        except ValueError:
            pass
    return ("s", text)


def is_formula(value) -> bool:
    return isinstance(value, str) and value.lstrip().startswith("=")


def formula_key(text: str):
    """Formulas differing only in spacing or case are the same formula."""
    return ("f", _WHITESPACE.sub("", str(text)).upper())


def cell_key(cell):
    """Canonical key for a (raw, cached) pair. None means the cell is empty."""
    raw, cached = cell
    if is_formula(raw):
        return formula_key(raw)
    if raw is None and is_formula(cached):
        return formula_key(cached)
    return value_key(raw if raw is not None else cached)


def values_equal(a, b) -> bool:
    return value_key(a) == value_key(b)


def cells_equal(old, new) -> bool:
    """
    Compare a cell across the two files, covering formulas and values both.

    If either side holds a formula, the formula text decides -- two cells with
    the same formula are unchanged even when only one of the files carries a
    cached result (very common: one file saved by a script, the other by Excel).
    When both cached results are present they must match too, which catches a
    total whose inputs moved. Plain cells compare on value.
    """
    key_old, key_new = cell_key(old), cell_key(new)
    if key_old != key_new:
        return False

    if key_old is not None and key_old[0] == "f":
        cached_old, cached_new = old[1], new[1]
        if (cached_old is not None and cached_new is not None
                and not is_formula(cached_old) and not is_formula(cached_new)):
            return value_key(cached_old) == value_key(cached_new)
    return True


def cell_is_empty(cell) -> bool:
    return cell_key(cell) is None


def extract_grid(ws_raw, ws_val) -> dict:
    """
    {(row, col): (raw, cached)} for every cell that actually holds something.

    Read with iter_rows rather than ws.cell(), which would materialise an empty
    object for every coordinate visited -- on a wide sheet that is millions of
    cells of pure overhead, and it silently stretches the sheet's dimensions.
    """
    collected = {}
    for source, slot in ((ws_raw, 0), (ws_val, 1)):
        if source is None:
            continue
        for row in source.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                collected.setdefault((cell.row, cell.column), [None, None])[slot] = cell.value

    grid = {}
    for position, pair in collected.items():
        cell = (pair[0], pair[1])
        if not cell_is_empty(cell):
            grid[position] = cell
    return grid


def row_signatures(grid: dict, max_row: int) -> list:
    """One hashable fingerprint per row, used to align the two sheets."""
    by_row = {}
    for (row, col), cell in grid.items():
        by_row.setdefault(row, []).append((col, cell_key(cell)))
    return [tuple(sorted(by_row.get(row, ()))) for row in range(1, max_row + 1)]


MIN_ROW_SIMILARITY = 0.2   # below this, two rows are unrelated
MAX_BLOCK_CELLS = 250_000  # guard on the pairing search for huge edited blocks


def row_similarity(signature_a, signature_b) -> float:
    """
    How much two rows look like each other, from 0.0 to 1.0.

    Rows that use the same columns are related; rows that also agree on some of
    the values in them are more related still. This is what tells an edited row
    apart from an unrelated one that merely happens to sit at the same offset.
    """
    if not signature_a or not signature_b:
        return 0.0
    a, b = dict(signature_a), dict(signature_b)
    shared = set(a) & set(b)
    if not shared:
        return 0.0
    identical = sum(1 for column in shared if a[column] == b[column])
    return (len(shared) + identical) / (2.0 * max(len(a), len(b)))


def pair_block(old_rows: list, new_rows: list, old_signatures: list,
               new_signatures: list) -> list:
    """
    Pair the rows inside one edited block, by resemblance rather than by
    position, keeping the original order.

    Position alone is not enough: delete one row and add another a few rows
    further down and the two land in the same block, where pairing by offset
    would marry a deleted row to an unrelated added one and report both as a
    single nonsensical edit.
    """
    count_old, count_new = len(old_rows), len(new_rows)
    if not count_old:
        return [(None, row) for row in new_rows]
    if not count_new:
        return [(row, None) for row in old_rows]

    if count_old * count_new > MAX_BLOCK_CELLS:      # pathological block
        overlap = min(count_old, count_new)
        return ([(old_rows[k], new_rows[k]) for k in range(overlap)]
                + [(row, None) for row in old_rows[overlap:]]
                + [(None, row) for row in new_rows[overlap:]])

    # Best-scoring order-preserving pairing, by dynamic programming.
    best = [[0.0] * (count_new + 1) for _ in range(count_old + 1)]
    move = [[None] * (count_new + 1) for _ in range(count_old + 1)]
    for i in range(1, count_old + 1):
        for j in range(1, count_new + 1):
            score = row_similarity(old_signatures[old_rows[i - 1] - 1],
                                   new_signatures[new_rows[j - 1] - 1])
            options = [(best[i - 1][j], "old"), (best[i][j - 1], "new")]
            if score >= MIN_ROW_SIMILARITY:
                options.append((best[i - 1][j - 1] + score, "pair"))
            best[i][j], move[i][j] = max(options)

    pairs, i, j = [], count_old, count_new
    while i or j:
        if i and j and move[i][j] == "pair":
            pairs.append((old_rows[i - 1], new_rows[j - 1]))
            i, j = i - 1, j - 1
        elif j and (not i or move[i][j] == "new"):
            pairs.append((None, new_rows[j - 1]))
            j -= 1
        else:
            pairs.append((old_rows[i - 1], None))
            i -= 1
    pairs.reverse()
    return pairs


def align_rows(old_signatures: list, new_signatures: list) -> list:
    """
    Pair up the rows of the two sheets, allowing for insertions and deletions.

    Returns [(old_row, new_row), ...] in order, where either side may be None:
    None on the left is a row that was added, None on the right is a row that
    was deleted. Row numbers are 1-based.

    This is the part the position-for-position comparison got wrong. Inserting
    a single row near the top of a sheet shifts every row below it by one, and
    a straight coordinate-to-coordinate diff then reports every one of those
    rows as modified -- including the template's own headings and labels, which
    nobody touched. Aligning first means an inserted row is reported as exactly
    that: one insertion, with everything below it still matched to its real
    counterpart.
    """
    # Only rows that actually hold something take part in the alignment. Blank
    # rows all look identical to each other, so letting them vote lets the
    # matcher slide the whole sheet by a row to buy itself a longer run of
    # "matches" that mean nothing -- and every real row below then lines up
    # against the wrong counterpart. Blank rows have nothing to compare anyway.
    old_index = [i for i, signature in enumerate(old_signatures) if signature]
    new_index = [i for i, signature in enumerate(new_signatures) if signature]

    matcher = difflib.SequenceMatcher(
        a=[old_signatures[i] for i in old_index],
        b=[new_signatures[j] for j in new_index],
        autojunk=False)

    pairs = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            pairs.extend((old_index[i1 + k] + 1, new_index[j1 + k] + 1)
                         for k in range(i2 - i1))
        elif tag == "replace":
            # Inside an edited block, marry each row to the row it most
            # resembles; whatever is left over is a real insertion or deletion.
            pairs.extend(pair_block(
                [old_index[k] + 1 for k in range(i1, i2)],
                [new_index[k] + 1 for k in range(j1, j2)],
                old_signatures, new_signatures))
        elif tag == "delete":
            pairs.extend((old_index[k] + 1, None) for k in range(i1, i2))
        elif tag == "insert":
            pairs.extend((None, new_index[k] + 1) for k in range(j1, j2))

    # A row whose content was simply wiped is not the same event as a row that
    # was pulled out of the sheet. If an unmatched row's own number is still
    # free and blank on the other side, pair the two: the change then reads as
    # "these cells were cleared", and the highlight lands on the same address
    # in both files instead of drifting.
    claimed_old = {o for o, _ in pairs if o is not None}
    claimed_new = {n for _, n in pairs if n is not None}
    blank_old = {i + 1 for i, signature in enumerate(old_signatures) if not signature}
    blank_new = {i + 1 for i, signature in enumerate(new_signatures) if not signature}

    resolved = []
    for row_old, row_new in pairs:
        if row_new is None and row_old in blank_new and row_old not in claimed_new:
            row_new = row_old
            claimed_new.add(row_new)
        elif row_old is None and row_new in blank_old and row_new not in claimed_old:
            row_old = row_new
            claimed_old.add(row_old)
        resolved.append((row_old, row_new))

    resolved.sort(key=lambda pair: (pair[1] if pair[1] is not None else pair[0],
                                    pair[0] if pair[0] is not None else 0))
    return resolved


def display(cell) -> str:
    """Readable rendering of a cell for the notes and the on-screen log."""
    raw, cached = cell
    if is_formula(raw):
        return f"{raw}" if cached is None else f"{raw}  (= {cached})"
    value = raw if raw is not None else cached
    return "" if value is None else str(value)


def shorten(text: str) -> str:
    text = str(text)
    return text if len(text) <= MAX_NOTE_LEN else text[:MAX_NOTE_LEN] + " ..."


# ==========================================================================
# Writing the highlighted copy
# ==========================================================================
#
# The highlighted copies are made by PATCHING the source package, never by
# rebuilding it.
#
# openpyxl cannot be used as the writer. On save it regenerates the workbook's
# style tables from scratch and keeps only the named styles it tracks, while
# leaving every cell's xfId pointing at the ORIGINAL entry number. A workbook
# converted from a legacy .xls carries far more <cellStyleXfs> entries than
# survive that trip -- the SUTC template has 36 and openpyxl writes back 22 --
# so every cell above the cut ends up with a dangling reference and Excel drops
# its fill and its borders. The template renders blank, and the red and green
# highlights vanish with it, because they inherit the same broken reference.
#
# So the converted source .xlsx is copied part for part, and only what has to
# change is rewritten:
#
#   xl/styles.xml   two fills and a few xf records APPENDED -- nothing is ever
#                   renumbered, so every existing reference stays valid
#   the sheets      the `s` attribute of the marked cells, and a legacy drawing
#                   reference where notes are attached
#   new parts       the notes themselves and their VML shapes
#
# Everything else -- layout, fonts, borders, merged ranges, column widths,
# images, formulas and their cached results, sheet tab colours, even VBA in an
# .xlsm -- is passed through byte for byte.

SML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"

COMMENT_REL = REL_NS + "/comments"
VML_REL = REL_NS + "/vmlDrawing"
COMMENT_TYPE = ("application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.comments+xml")
VML_TYPE = "application/vnd.openxmlformats-officedocument.vmlDrawing"

# Serialising a sheet must reproduce the prefixes it arrived with, or the
# r:id references inside it stop resolving.
for _prefix, _uri in (
        ("", SML_NS),
        ("r", REL_NS),
        ("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"),
        ("a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
        ("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"),
        ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")):
    ET.register_namespace(_prefix, _uri)

XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

# ---------------------------------------------------------------------------
# Preserving the namespaces a part arrived with
# ---------------------------------------------------------------------------
#
# ElementTree is a lossy round trip for an Office part, in two ways that both
# make Excel declare the file unreadable and "repair" it -- which strips every
# style in the workbook and takes the highlights down with them:
#
#   1. It only re-declares the namespaces something in the tree actually uses.
#      A file written by Excel carries mc:Ignorable="x14ac" on the root, where
#      "x14ac" is a plain attribute VALUE, not a namespaced name -- so nothing
#      in the tree "uses" it, the xmlns:x14ac declaration is dropped, and the
#      surviving Ignorable now points at a prefix that no longer exists.
#
#   2. Prefixes it has not been told about are renamed. x14ac:dyDescent, which
#      Excel puts on every single <row>, comes back out as ns2:dyDescent.
#
# LibreOffice writes neither of those, which is why a file that has been round
# tripped through it survives and a file straight out of Excel does not.
#
# The cure is to register every prefix a part declares before parsing it, and
# then to put the part's ORIGINAL root tag back verbatim afterwards, so every
# declaration and every mc:Ignorable returns exactly as authored.

_XMLNS_DECL = re.compile(rb'xmlns:([A-Za-z_][\w.-]*)\s*=\s*"([^"]*)"')
_DEFAULT_NS = re.compile(rb'\sxmlns\s*=\s*"([^"]*)"')


def _register_namespaces(data: bytes) -> None:
    """Teach ElementTree every prefix this part declares, before parsing it."""
    for prefix, uri in _XMLNS_DECL.findall(data):
        name = prefix.decode("utf-8", "ignore")
        if name in ("xml", "xmlns"):
            continue
        try:
            ET.register_namespace(name, uri.decode("utf-8", "ignore"))
        except (ValueError, TypeError):
            pass


def _root_tag(data: bytes) -> bytes:
    """The opening tag of the root element, verbatim, declarations and all."""
    position, size = 0, len(data)
    while position < size:
        start = data.find(b"<", position)
        if start < 0:
            return b""
        if data.startswith(b"<?", start):
            position = data.find(b"?>", start) + 2
        elif data.startswith(b"<!--", start):
            position = data.find(b"-->", start) + 3
        elif data.startswith(b"<!", start):
            position = data.find(b">", start) + 1
        else:
            end = data.find(b">", start)
            return data[start:end + 1] if end > 0 else b""
        if position < 2:
            return b""
    return b""


def _restore_root_tag(generated: bytes, original: bytes) -> bytes:
    """Swap ElementTree's root tag for the one the part was authored with."""
    original_tag = _root_tag(original)
    generated_tag = _root_tag(generated)
    if not original_tag or not generated_tag or original_tag == generated_tag:
        return generated

    closes_itself = generated_tag.rstrip().endswith(b"/>")
    body = original_tag.rstrip()
    body = body[:-2] if body.endswith(b"/>") else body[:-1]

    # Anything ElementTree had to declare that the original did not is kept.
    present = {prefix for prefix, _ in _XMLNS_DECL.findall(body)}
    extra = b""
    for prefix, uri in _XMLNS_DECL.findall(generated_tag):
        if prefix not in present:
            extra += b' xmlns:' + prefix + b'="' + uri + b'"'
            present.add(prefix)
    if not _DEFAULT_NS.search(body):
        found = _DEFAULT_NS.search(generated_tag)
        if found:
            extra += b' xmlns="' + found.group(1) + b'"'

    return generated.replace(generated_tag,
                             body + extra + (b"/>" if closes_itself else b">"), 1)


def _parse(data: bytes):
    """Parse a part after registering the prefixes it declares."""
    _register_namespaces(data)
    return ET.fromstring(data)


def _serialise(root, original: bytes = b"") -> bytes:
    generated = (XML_HEADER + ET.tostring(root, encoding="unicode")).encode("utf-8")
    return _restore_root_tag(generated, original) if original else generated

# Where <legacyDrawing> is allowed to sit inside a worksheet. Excel rejects the
# file outright if the children are out of schema order.
SHEET_ORDER = (
    "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
    "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios",
    "autoFilter", "sortState", "dataConsolidate", "customSheetViews",
    "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations",
    "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter",
    "rowBreaks", "colBreaks", "customProperties", "cellWatches",
    "ignoredErrors", "smartTags", "drawing", "legacyDrawing",
    "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems",
    "tableParts", "extLst",
)

CELL_REF = re.compile(r"([A-Z]+)(\d+)")


def _q(tag: str) -> str:
    return f"{{{SML_NS}}}{tag}"


def _sheet_rank(tag: str) -> int:
    """Where a worksheet child sits in schema order; anything unknown sorts last."""
    name = tag.rsplit("}", 1)[-1]
    return SHEET_ORDER.index(name) if name in SHEET_ORDER else len(SHEET_ORDER)


def _split_ref(ref: str):
    """'AJ22' -> (22, 36).  Returns (row, column) as 1-based numbers."""
    match = CELL_REF.fullmatch(ref or "")
    if not match:
        return None
    return int(match.group(2)), column_index_from_string(match.group(1))


class HighlightWriter:
    """
    One source workbook in, one highlighted copy out.

    Call fill() and note() to record what should be marked, then save().
    """

    def __init__(self, src: Path):
        with zipfile.ZipFile(src) as archive:
            # Order matters: some readers want the parts as they were laid out.
            self.order = list(archive.namelist())
            self.parts = {name: archive.read(name) for name in self.order}

        self.styles_source = self.parts["xl/styles.xml"]
        self.styles = _parse(self.styles_source)
        self.sheet_parts = self._locate_sheets()

        self._fill_ids = {}     # rgb -> index into <fills>
        self._style_ids = {}    # (base xf, rgb) -> index into <cellXfs>
        self._fills = {}        # sheet name -> {(row, col): rgb}
        self._notes = {}        # sheet name -> {(row, col): text}

    # -- reading the source layout ----------------------------------------

    def _locate_sheets(self) -> dict:
        targets = {rel.get("Id"): rel.get("Target")
                   for rel in _parse(self.parts["xl/_rels/workbook.xml.rels"])}
        sheets = {}
        book = _parse(self.parts["xl/workbook.xml"])
        for sheet in book.find(_q("sheets")):
            target = (targets.get(sheet.get(f"{{{REL_NS}}}id")) or "").lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            sheets[sheet.get("name")] = target
        return sheets

    def has_sheet(self, name: str) -> bool:
        return name in self.sheet_parts

    # -- recording what to mark -------------------------------------------

    def fill(self, sheet: str, row: int, col: int, rgb: str) -> None:
        self._fills.setdefault(sheet, {})[(row, col)] = rgb

    def note(self, sheet: str, row: int, col: int, text: str) -> None:
        if text:
            self._notes.setdefault(sheet, {})[(row, col)] = shorten(text)

    # -- appending to the style table --------------------------------------

    def _fill_id(self, rgb: str) -> int:
        if rgb not in self._fill_ids:
            fills = self.styles.find(_q("fills"))
            fill = ET.SubElement(fills, _q("fill"))
            pattern = ET.SubElement(fill, _q("patternFill"), {"patternType": "solid"})
            ET.SubElement(pattern, _q("fgColor"), {"rgb": rgb})
            ET.SubElement(pattern, _q("bgColor"), {"rgb": rgb})
            fills.set("count", str(len(fills)))
            self._fill_ids[rgb] = len(fills) - 1
        return self._fill_ids[rgb]

    def _style_id(self, base: int, rgb: str) -> int:
        """An xf identical to the cell's own, but carrying the highlight fill."""
        key = (base, rgb)
        if key not in self._style_ids:
            cell_xfs = self.styles.find(_q("cellXfs"))
            source = cell_xfs[base] if 0 <= base < len(cell_xfs) else None
            xf = copy.deepcopy(source) if source is not None else ET.Element(_q("xf"))
            xf.set("fillId", str(self._fill_id(rgb)))
            xf.set("applyFill", "1")
            cell_xfs.append(xf)
            cell_xfs.set("count", str(len(cell_xfs)))
            self._style_ids[key] = len(cell_xfs) - 1
        return self._style_ids[key]

    # -- painting -----------------------------------------------------------

    def _paint(self, root, fills: dict) -> None:
        data = root.find(_q("sheetData"))
        if data is None:
            return
        rows = {int(r.get("r")): r for r in data.findall(_q("row")) if r.get("r")}

        for (row, col), rgb in sorted(fills.items()):
            row_el = rows.get(row)
            if row_el is None:
                row_el = ET.Element(_q("row"), {"r": str(row)})
                at = next((i for i, r in enumerate(data)
                           if int(r.get("r") or 0) > row), len(data))
                data.insert(at, row_el)
                rows[row] = row_el

            ref = f"{get_column_letter(col)}{row}"
            cell = next((c for c in row_el if c.get("r") == ref), None)
            if cell is None:
                cell = ET.Element(_q("c"), {"r": ref})
                at = next((i for i, c in enumerate(row_el)
                           if (_split_ref(c.get("r")) or (0, 0))[1] > col), len(row_el))
                row_el.insert(at, cell)

            cell.set("s", str(self._style_id(int(cell.get("s") or 0), rgb)))

        self._widen_dimension(root, fills)

    @staticmethod
    def _widen_dimension(root, fills: dict) -> None:
        """Keep <dimension> covering the sheet once cells have been added."""
        dimension = root.find(_q("dimension"))
        if dimension is None or not fills:
            return
        corners = [_split_ref(part) for part in (dimension.get("ref") or "").split(":")]
        corners = [c for c in corners if c]
        corners += list(fills)
        if not corners:
            return
        top, left = min(r for r, _ in corners), min(c for _, c in corners)
        bottom, right = max(r for r, _ in corners), max(c for _, c in corners)
        dimension.set("ref", f"{get_column_letter(left)}{top}:"
                             f"{get_column_letter(right)}{bottom}")

    # -- notes --------------------------------------------------------------

    def _attach_notes(self, sheet_part: str, root, notes: dict, index: int) -> None:
        rels_part = f"xl/worksheets/_rels/{sheet_part.rsplit('/', 1)[-1]}.rels"
        existing = self._existing_targets(rels_part, sheet_part)

        # A sheet that already carries comments must not be given a second
        # comments part -- Excel rejects that outright. Merge into the one
        # that is already there instead, and likewise for its VML shapes.
        comments_part = existing.get("comments")
        vml_part = existing.get("vml")

        if comments_part and comments_part in self.parts:
            self._add_part(comments_part,
                           _merge_comments(self.parts[comments_part], notes))
        else:
            comments_part = self._free_part(f"xl/comments{index}.xml")
            self._add_part(comments_part, _comments_xml(notes))
            self._add_relationship(rels_part, COMMENT_REL,
                                   _relative_target(comments_part))
            self._register_content_type(comments_part)

        if vml_part and vml_part in self.parts:
            self._add_part(vml_part, _merge_vml(self.parts[vml_part], notes))
            self._ensure_vml_content_type()
            return

        vml_part = self._free_part(f"xl/drawings/vmlDrawing{index}.vml")
        self._add_part(vml_part, _vml_drawing(notes))
        vml_id = self._add_relationship(rels_part, VML_REL,
                                        _relative_target(vml_part))
        self._ensure_vml_content_type()

        legacy = root.find(_q("legacyDrawing"))
        if legacy is None:
            legacy = ET.Element(_q("legacyDrawing"))
            rank = SHEET_ORDER.index("legacyDrawing")
            at = next((i for i, child in enumerate(root)
                       if _sheet_rank(child.tag) > rank), len(root))
            root.insert(at, legacy)
        legacy.set(f"{{{REL_NS}}}id", vml_id)

    def _free_part(self, name: str) -> str:
        """A part name that is not already taken, so nothing is overwritten."""
        if name not in self.parts:
            return name
        stem, _, suffix = name.rpartition(".")
        counter = 2
        while f"{stem}_{counter}.{suffix}" in self.parts:
            counter += 1
        return f"{stem}_{counter}.{suffix}"

    def _existing_targets(self, rels_part: str, sheet_part: str) -> dict:
        """Any comments and VML part this sheet is already wired up to."""
        found = {}
        if rels_part not in self.parts:
            return found
        base = sheet_part.rsplit("/", 1)[0]
        try:
            relationships = _parse(self.parts[rels_part])
        except ET.ParseError:
            return found
        for rel in relationships:
            target = (rel.get("Target") or "")
            kind = rel.get("Type") or ""
            if target.startswith("/"):
                resolved = target.lstrip("/")
            else:
                resolved = os.path.normpath(f"{base}/{target}").replace("\\", "/")
            if kind == COMMENT_REL:
                found["comments"] = resolved
            elif kind == VML_REL:
                found["vml"] = resolved
        return found

    def _ensure_vml_content_type(self) -> None:
        name = "[Content_Types].xml"
        xml = self.parts[name].decode("utf-8")
        if 'Extension="vml"' in xml:
            return
        xml = xml.replace("<Override", f'<Default Extension="vml" '
                                       f'ContentType="{VML_TYPE}"/><Override', 1)
        self._add_part(name, xml.encode("utf-8"))

    # -- package plumbing ---------------------------------------------------

    def _add_part(self, name: str, data: bytes) -> None:
        if name not in self.parts:
            self.order.append(name)
        self.parts[name] = data

    def _add_relationship(self, rels_part: str, rel_type: str, target: str) -> str:
        """Append one relationship, textually, so the rest of the part is kept."""
        if rels_part in self.parts:
            xml = self.parts[rels_part].decode("utf-8")
        else:
            xml = (XML_HEADER + '<Relationships xmlns="http://schemas.openxmlformats'
                   '.org/package/2006/relationships"></Relationships>')
        used = {int(n) for n in re.findall(r'Id="rId(\d+)"', xml)}
        rel_id = f"rId{max(used) + 1 if used else 1}"
        entry = f'<Relationship Id="{rel_id}" Type="{rel_type}" Target="{target}"/>'

        if "</Relationships>" in xml:
            xml = xml.replace("</Relationships>", entry + "</Relationships>")
        else:
            # An empty rels part is written self-closed: <Relationships ... />.
            # Opening it up first is the difference between the relationship
            # landing and being silently dropped -- and a dropped r:id leaves a
            # dangling reference that makes Excel condemn the whole workbook.
            xml = re.sub(r"/\s*>\s*$", ">" + entry + "</Relationships>", xml.rstrip(),
                         count=1)
        self._add_part(rels_part, xml.encode("utf-8"))
        return rel_id

    def _register_content_type(self, comments_part: str) -> None:
        name = "[Content_Types].xml"
        xml = self.parts[name].decode("utf-8")
        override = f'<Override PartName="/{comments_part}" ContentType="{COMMENT_TYPE}"/>'
        if override not in xml:
            xml = xml.replace("</Types>", override + "</Types>")
        self._add_part(name, xml.encode("utf-8"))

    # -- output --------------------------------------------------------------

    def save(self, path: Path) -> None:
        index = 0
        for sheet in sorted(set(self._fills) | set(self._notes)):
            part = self.sheet_parts.get(sheet)
            if part is None:
                continue
            source = self.parts[part]
            root = _parse(source)
            self._paint(root, self._fills.get(sheet, {}))
            notes = self._notes.get(sheet)
            if notes:
                index += 1
                self._attach_notes(part, root, notes, index)
            self._add_part(part, _serialise(root, source))

        _register_namespaces(self.styles_source)
        self._add_part("xl/styles.xml", _serialise(self.styles, self.styles_source))

        path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
            for name in self.order:
                archive.writestr(name, self.parts[name])


def _relative_target(part: str) -> str:
    """A worksheet rels Target pointing at `part`, relative to xl/worksheets/."""
    if part.startswith("xl/worksheets/"):
        return part[len("xl/worksheets/"):]
    return "../" + part[len("xl/"):] if part.startswith("xl/") else "/" + part


def _comments_xml(notes: dict) -> bytes:
    root = ET.Element(_q("comments"))
    authors = ET.SubElement(root, _q("authors"))
    ET.SubElement(authors, _q("author")).text = COMMENT_AUTHOR
    listing = ET.SubElement(root, _q("commentList"))
    for (row, col), text in sorted(notes.items()):
        comment = ET.SubElement(listing, _q("comment"),
                                {"ref": f"{get_column_letter(col)}{row}", "authorId": "0"})
        run = ET.SubElement(ET.SubElement(comment, _q("text")), _q("t"))
        run.set(f"{{{XML_NS}}}space", "preserve")
        run.text = text
    return _serialise(root)


_VML_PROLOGUE = (
    '<xml xmlns:v="urn:schemas-microsoft-com:vml"'
    ' xmlns:o="urn:schemas-microsoft-com:office:office"'
    ' xmlns:x="urn:schemas-microsoft-com:office:excel">'
    '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>'
    '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"'
    ' path="m,l,21600r21600,l21600,xe">'
    '<v:stroke joinstyle="miter"/>'
    '<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
)


def _merge_comments(existing: bytes, notes: dict) -> bytes:
    """Add our notes to a comments part the workbook already had."""
    root = _parse(existing)
    authors = root.find(_q("authors"))
    if authors is None:
        authors = ET.Element(_q("authors"))
        root.insert(0, authors)
    names = [a.text or "" for a in authors.findall(_q("author"))]
    if COMMENT_AUTHOR in names:
        author_id = names.index(COMMENT_AUTHOR)
    else:
        ET.SubElement(authors, _q("author")).text = COMMENT_AUTHOR
        author_id = len(names)

    listing = root.find(_q("commentList"))
    if listing is None:
        listing = ET.SubElement(root, _q("commentList"))
    taken = {c.get("ref") for c in listing.findall(_q("comment"))}

    for (row, col), text in sorted(notes.items()):
        ref = f"{get_column_letter(col)}{row}"
        if ref in taken:               # leave the author's own note alone
            continue
        comment = ET.SubElement(listing, _q("comment"),
                                {"ref": ref, "authorId": str(author_id)})
        run = ET.SubElement(ET.SubElement(comment, _q("text")), _q("t"))
        run.set(f"{{{XML_NS}}}space", "preserve")
        run.text = text
    return _serialise(root, existing)


def _merge_vml(existing: bytes, notes: dict) -> bytes:
    """Append our note shapes to the VML drawing the sheet already had."""
    text = existing.decode("utf-8", "replace")
    closing = text.rfind("</xml>")
    if closing < 0:
        return _vml_drawing(notes)
    used = [int(n) for n in re.findall(r'id="_x0000_s(\d+)"', text)]
    shapes = _vml_shapes(notes, start=max(used) + 1 if used else 1025)
    return (text[:closing] + shapes + text[closing:]).encode("utf-8")


def _vml_shapes(notes: dict, start: int = 1025) -> str:
    shapes = []
    for offset, (row, col) in enumerate(sorted(notes)):
        shapes.append(
            f'<v:shape id="_x0000_s{start + offset}" type="#_x0000_t202"'
            ' style="position:absolute;margin-left:70pt;margin-top:5pt;width:220pt;'
            'height:100pt;z-index:1;visibility:hidden" fillcolor="#ffffe1"'
            ' o:insetmode="auto">'
            '<v:fill color2="#ffffe1"/>'
            '<v:shadow on="t" color="black" obscured="t"/>'
            '<v:path o:connecttype="none"/>'
            '<v:textbox style="mso-direction-alt:auto">'
            '<div style="text-align:left"></div></v:textbox>'
            '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>'
            f'<x:Anchor>{col}, 15, {row - 1}, 10, {col + 4}, 15, {row + 4}, 10</x:Anchor>'
            '<x:AutoFill>False</x:AutoFill>'
            f'<x:Row>{row - 1}</x:Row><x:Column>{col - 1}</x:Column>'
            '</x:ClientData></v:shape>'
        )
    return "".join(shapes)


def _vml_drawing(notes: dict) -> bytes:
    """The hidden shapes Excel needs before it will show a note at all."""
    return (_VML_PROLOGUE + _vml_shapes(notes) + "</xml>").encode("utf-8")


# ==========================================================================
# Highlighting
# ==========================================================================

def build_merged_lookup(ws) -> dict:
    lookup = {}
    for rng in ws.merged_cells.ranges:
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                lookup[(r, c)] = rng
    return lookup


def mark(writer: HighlightWriter, sheet: str, row: int, col: int,
         rgb: str, note: str, merged_lookup) -> None:
    """Fill a cell (the whole merged range if it is one) and note what changed."""
    rng = merged_lookup.get((row, col))
    if rng is not None:
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                writer.fill(sheet, r, c, rgb)
        row, col = rng.min_row, rng.min_col
    else:
        writer.fill(sheet, row, col, rgb)

    writer.note(sheet, row, col, note)


# ==========================================================================
# Core
# ==========================================================================

def diff_workbooks(old_path, new_path, out_dir, template=None, progress=None) -> DiffResult:
    """
    Compare the two workbooks and write the highlighted copies.

    `template` is the blank template both files were built from. It is optional
    and never modified -- supplying it simply tells the tool which cells are the
    template's own headings and labels, so those are left alone and only the
    client's content is compared.

    `progress` is an optional callable taking one string; the window uses it to
    stream status lines while the work runs on a background thread.
    """
    old_path, new_path, out_dir = Path(old_path), Path(new_path), Path(out_dir)
    say = progress or (lambda msg: None)

    # Legacy .xls inputs are converted into this scratch folder, which is
    # removed at the end. The originals on disk are never touched.
    workdir = Path(tempfile.mkdtemp(prefix="excel_review_"))
    try:
        src_old = ensure_modern(old_path, workdir, "original", say)
        src_new = ensure_modern(new_path, workdir, "modified", say)

        # Two passes per file: one keeps the formulas (and is what gets written
        # back out), one carries the cached results the comparison needs.
        say(f"Loading {old_path.name} ...")
        wb_old = load_workbook(src_old, data_only=False,
                               keep_vba=src_old.suffix.lower() == ".xlsm")
        wb_old_val = load_workbook(src_old, data_only=True)
        say(f"Loading {new_path.name} ...")
        wb_new = load_workbook(src_new, data_only=False,
                               keep_vba=src_new.suffix.lower() == ".xlsm")
        wb_new_val = load_workbook(src_new, data_only=True)

        # The blank template, if one was given. Read only -- it contributes
        # nothing to the output, it only marks which cells are boilerplate.
        template_keys = {}
        if template:
            template_path = Path(template)
            say(f"Reading the template {template_path.name} ...")
            src_tpl = ensure_modern(template_path, workdir, "template", say)
            wb_tpl = load_workbook(src_tpl, data_only=False)
            wb_tpl_val = load_workbook(src_tpl, data_only=True)
            for sheet_name in wb_tpl.sheetnames:
                grid = extract_grid(
                    wb_tpl[sheet_name],
                    wb_tpl_val[sheet_name] if sheet_name in wb_tpl_val.sheetnames else None)
                template_keys[sheet_name] = {pos: cell_key(cell)
                                             for pos, cell in grid.items()}
            wb_tpl.close()
            wb_tpl_val.close()

        # The highlighted copies are patched straight onto these two packages,
        # so whatever the template looks like is what comes out.
        writer_old = HighlightWriter(src_old)
        writer_new = HighlightWriter(src_new)
    except Exception:
        shutil.rmtree(workdir, ignore_errors=True)
        raise

    result = DiffResult()

    old_sheets, new_sheets = list(wb_old.sheetnames), list(wb_new.sheetnames)
    result.sheets_only_in_old = [s for s in old_sheets if s not in new_sheets]
    result.sheets_only_in_new = [s for s in new_sheets if s not in old_sheets]
    shared = [s for s in old_sheets if s in new_sheets]

    # ---- whole sheets that disappeared or appeared -----------------------
    for name in result.sheets_only_in_old:
        ws = wb_old[name]
        mark(writer_old, name, 1, 1, FILL_RED,
             f"SHEET DELETED\nThis whole sheet is missing from the new file.",
             build_merged_lookup(ws))
        result.changes.append(Change(name, "-", "sheet-deleted"))
        say(f"Sheet deleted: {name}")

    for name in result.sheets_only_in_new:
        ws = wb_new[name]
        mark(writer_new, name, 1, 1, FILL_GREEN,
             f"SHEET ADDED\nThis sheet does not exist in the old file.",
             build_merged_lookup(ws))
        result.changes.append(Change(name, "-", "sheet-added"))
        say(f"Sheet added: {name}")

    # ---- cell by cell, across every shared sheet -------------------------
    sheet_deleted_rows, sheet_added_rows = set(), set()

    for i, name in enumerate(shared, 1):
        say(f"Comparing sheet {i}/{len(shared)}: {name}")
        ws_old, ws_new = wb_old[name], wb_new[name]
        ws_old_val = wb_old_val[name] if name in wb_old_val.sheetnames else None
        ws_new_val = wb_new_val[name] if name in wb_new_val.sheetnames else None

        merged_old = build_merged_lookup(ws_old)
        merged_new = build_merged_lookup(ws_new)

        grid_old = extract_grid(ws_old, ws_old_val)
        grid_new = extract_grid(ws_new, ws_new_val)
        template_grid = template_keys.get(name, {})

        max_row_old = max((r for r, _ in grid_old), default=0)
        max_row_new = max((r for r, _ in grid_new), default=0)

        # Align the rows before looking at a single cell. Everything below
        # then compares like against like, however far the content has shifted.
        pairs = align_rows(row_signatures(grid_old, max_row_old),
                           row_signatures(grid_new, max_row_new))

        columns_old, columns_new = {}, {}
        for (row, col) in grid_old:
            columns_old.setdefault(row, set()).add(col)
        for (row, col) in grid_new:
            columns_new.setdefault(row, set()).add(col)

        sheet_changes = 0
        skipped_template = 0

        for row_old, row_new in pairs:
            whole_row = row_old is None or row_new is None
            columns = set()
            if row_old is not None:
                columns |= columns_old.get(row_old, set())
            if row_new is not None:
                columns |= columns_new.get(row_new, set())

            for col in sorted(columns):
                old_cell = grid_old.get((row_old, col), (None, None)) if row_old else (None, None)
                new_cell = grid_new.get((row_new, col), (None, None)) if row_new else (None, None)

                empty_old, empty_new = cell_is_empty(old_cell), cell_is_empty(new_cell)
                if empty_old and empty_new:
                    continue
                if not whole_row and cells_equal(old_cell, new_cell):
                    continue

                # A cell that still carries the template's own text on BOTH
                # sides is the template speaking, not the client -- it is left
                # alone even when the row it sits on has moved.
                if template_grid:
                    key_old = template_grid.get((row_old, col)) if row_old else None
                    key_new = template_grid.get((row_new, col)) if row_new else None
                    if (key_old is not None and key_new is not None
                            and cell_key(old_cell) == key_old
                            and cell_key(new_cell) == key_new):
                        skipped_template += 1
                        continue

                kind = "added" if empty_old else "deleted" if empty_new else "modified"
                old_text, new_text = display(old_cell), display(new_cell)
                coord_old = f"{get_column_letter(col)}{row_old}" if row_old else None
                coord_new = f"{get_column_letter(col)}{row_new}" if row_new else None

                if kind == "added":
                    note = (f"ROW ADDED\nRow {row_new} is not present in the old file."
                            f"\nValue: {new_text}"
                            if whole_row else
                            "ADDED\nThis cell was empty in the old file.")
                    mark(writer_new, name, row_new, col, FILL_GREEN, note, merged_new)

                elif kind == "deleted":
                    note_old = (f"ROW DELETED\nThe whole of row {row_old} was removed in the "
                                f"new file.\nWas: {old_text}"
                                if whole_row else
                                f"DELETED\nThis cell was cleared in the new file."
                                f"\nWas: {old_text}")
                    mark(writer_old, name, row_old, col, FILL_RED, note_old, merged_old)
                    # Only paint the new file where the cell still exists there.
                    if row_new is not None:
                        mark(writer_new, name, row_new, col, FILL_GREEN,
                             f"DELETED\nThis cell was cleared."
                             f"\nIt used to contain: {old_text}", merged_new)

                else:
                    note = f"MODIFIED\nOld: {old_text}\nNew: {new_text}"
                    mark(writer_old, name, row_old, col, FILL_RED, note, merged_old)
                    mark(writer_new, name, row_new, col, FILL_GREEN, note, merged_new)

                result.changes.append(
                    Change(name, coord_new or coord_old, kind,
                           old_value=old_text or None,
                           new_value=new_text or None,
                           whole_row=whole_row,
                           old_cell=coord_old, new_cell=coord_new)
                )
                sheet_changes += 1

            if whole_row and columns:
                if row_old is not None:
                    sheet_deleted_rows.add((name, row_old))
                else:
                    sheet_added_rows.add((name, row_new))

        if sheet_changes:
            # The sheet tab keeps whatever colour the template gave it -- the
            # changed cells are the only marks this tool leaves behind.
            say(f"   {sheet_changes} change(s) in '{name}'")
        if skipped_template:
            say(f"   {skipped_template} template cell(s) ignored in '{name}'")

    # ---- save -------------------------------------------------------------
    result.deleted_rows = sorted(sheet_deleted_rows)
    result.added_rows = sorted(sheet_added_rows)

    out_dir.mkdir(parents=True, exist_ok=True)
    # Highlighted copies are always written in a modern format: the old .xls
    # container cannot carry the fills and notes this tool adds. The extension
    # follows the package actually being patched, so an .xlsm keeps its macros.
    ext_old = ".xlsm" if src_old.suffix.lower() == ".xlsm" else ".xlsx"
    ext_new = ".xlsm" if src_new.suffix.lower() == ".xlsm" else ".xlsx"
    result.out_old = out_dir / f"{old_path.stem}_OLD{ext_old}"
    result.out_new = out_dir / f"{new_path.stem}_NEW{ext_new}"

    say("Saving highlighted files ...")
    writer_old.save(result.out_old)
    writer_new.save(result.out_new)

    shutil.rmtree(workdir, ignore_errors=True)
    return result


def summary_lines(result: DiffResult, limit: int = 30):
    lines = []
    if not result.changes:
        lines.append("No differences found - the two files match.")
        return lines

    lines.append(f"{result.cell_changes} changed cell(s):  "
                 f"{result.count('modified')} modified, "
                 f"{result.count('added')} added, "
                 f"{result.count('deleted')} deleted.")
    if result.deleted_rows:
        lines.append("Whole rows deleted: " +
                     ", ".join(f"{sheet}!{row}" for sheet, row in result.deleted_rows))
    if result.added_rows:
        lines.append("Whole rows added:   " +
                     ", ".join(f"{sheet}!{row}" for sheet, row in result.added_rows))
    lines.append("")

    by_sheet = {}
    for c in result.changes:
        by_sheet.setdefault(c.sheet, []).append(c)

    for sheet, changes in by_sheet.items():
        lines.append(f"  [{sheet}]  {len(changes)} change(s)")
        for c in changes[:limit]:
            # When rows have shifted the same cell sits at two different
            # coordinates, so both are shown: old -> new.
            if c.old_cell and c.new_cell and c.old_cell != c.new_cell:
                where = f"{c.old_cell}>{c.new_cell}"
            else:
                where = c.old_cell or c.new_cell or c.cell or "-"

            if c.kind.startswith("sheet"):
                lines.append(f"      {c.kind}")
            elif c.kind == "deleted":
                tag = "row deleted" if c.whole_row else "deleted"
                lines.append(f"      {where:<12} {tag:<11} was: {c.old_value!r}")
            elif c.kind == "added":
                tag = "row added" if c.whole_row else "added"
                lines.append(f"      {where:<12} {tag:<11} now: {c.new_value!r}")
            else:
                lines.append(f"      {where:<12} {'modified':<11} "
                             f"{c.old_value!r}  ->  {c.new_value!r}")
        if len(changes) > limit:
            lines.append(f"      ... and {len(changes) - limit} more")
        lines.append("")

    if result.sheets_only_in_old:
        lines.append("Sheets deleted: " + ", ".join(result.sheets_only_in_old))
    if result.sheets_only_in_new:
        lines.append("Sheets added:   " + ", ".join(result.sheets_only_in_new))
    return lines


def open_folder(path) -> None:
    path = str(path)
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
    except Exception:
        pass


# ==========================================================================
# Window
# ==========================================================================

# palette
BG = "#f0f0f0"
CARD = "#ffffff"
BORDER = "#d5d8dc"
TEXT = "#1a1a1a"
MUTED = "#6b7280"
ACCENT = "#1f6feb"
ACCENT_DARK = "#1a5fd0"


def _pick_font(tkfont, *candidates):
    available = set(tkfont.families())
    for name in candidates:
        if name in available:
            return name
    return "TkDefaultFont"


def build_gui(root):
    """Build the whole window on `root`. Returns a handle dict (used by tests)."""
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, font as tkfont

    ui = _pick_font(tkfont, "Segoe UI", "Helvetica Neue", "DejaVu Sans", "Arial")
    mono = _pick_font(tkfont, "Consolas", "Menlo", "DejaVu Sans Mono", "Courier New")

    F_TITLE = (ui, 17)
    F_SUB = (ui, 9)
    F_SECTION = (ui, 11, "bold")
    F_LABEL = (ui, 10)
    F_HINT = (ui, 8)
    F_BTN = (ui, 11)
    F_LOG = (mono, 9)

    root.title(f"{APP_TITLE} {__version__}")
    # Fit the screen rather than assuming one. A fixed 760px tall window runs
    # off the bottom of a laptop display, taking the Compare button with it.
    screen_w, screen_h = root.winfo_screenwidth(), root.winfo_screenheight()
    win_w, win_h = min(820, screen_w - 80), min(780, screen_h - 120)
    root.geometry(f"{win_w}x{win_h}"
                  f"+{max(0, (screen_w - win_w) // 2)}"
                  f"+{max(0, (screen_h - win_h) // 3)}")
    root.minsize(620, 380)
    root.configure(bg=BG)

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("TSeparator", background=BORDER)
    style.configure("Card.TButton", font=(ui, 9), padding=(10, 4))
    style.configure("Review.Horizontal.TProgressbar", troughcolor="#e3e6ea",
                    background=ACCENT, bordercolor=BORDER, lightcolor=ACCENT,
                    darkcolor=ACCENT)

    old_var = tk.StringVar()
    new_var = tk.StringVar()
    tpl_var = tk.StringVar()
    out_var = tk.StringVar(value=str(Path.home() / "Documents"))
    status_var = tk.StringVar(value="Select the two files, then press Compare Files.")

    log_queue = queue.Queue()
    state = {"running": False, "out_dir": None}

    # The action row, the log and the status line are pinned to the bottom of
    # the window and are never scrolled away, so the Compare button is always
    # reachable no matter how short the display is. Only the form above them
    # scrolls.
    tk.Label(root, textvariable=status_var, font=F_SUB, bg=BG, fg=MUTED, anchor="w",
             padx=22).pack(side="bottom", fill="x", pady=(4, 8))

    log_frame = tk.Frame(root, bg=CARD, highlightbackground=BORDER,
                         highlightthickness=1)
    log_frame.pack(side="bottom", fill="both", padx=22, pady=(0, 4))

    actions = tk.Frame(root, bg=BG, padx=22)
    actions.pack(side="bottom", fill="x", pady=(8, 8))

    scroll_area = tk.Frame(root, bg=BG)
    scroll_area.pack(side="top", fill="both", expand=True)

    canvas = tk.Canvas(scroll_area, bg=BG, highlightthickness=0, bd=0)
    form_scroll = tk.Scrollbar(scroll_area, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=form_scroll.set)
    canvas.pack(side="left", fill="both", expand=True)

    outer = tk.Frame(canvas, bg=BG, padx=22, pady=18)
    window_id = canvas.create_window((0, 0), window=outer, anchor="nw")

    def _resize(_event=None):
        canvas.itemconfigure(window_id, width=canvas.winfo_width())
        canvas.configure(scrollregion=canvas.bbox("all"))
        # The scrollbar only appears when there is genuinely something to
        # scroll to, so it does not clutter a window that already fits.
        needed = outer.winfo_reqheight() > canvas.winfo_height()
        if needed and not form_scroll.winfo_ismapped():
            form_scroll.pack(side="right", fill="y")
        elif not needed and form_scroll.winfo_ismapped():
            form_scroll.pack_forget()

    outer.bind("<Configure>", _resize)
    canvas.bind("<Configure>", _resize)

    def _wheel(event):
        if outer.winfo_reqheight() <= canvas.winfo_height():
            return
        if event.num == 4:
            step = -1
        elif event.num == 5:
            step = 1
        else:                                   # Windows and macOS
            step = -1 if event.delta > 0 else 1
        canvas.yview_scroll(step, "units")

    for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
        root.bind_all(sequence, _wheel)

    # ---------------- header ----------------
    tk.Label(outer, text=APP_TITLE, font=F_TITLE, bg=BG, fg=TEXT, anchor="w"
             ).pack(fill="x")
    tk.Label(outer, text=APP_SUBTITLE, font=F_SUB, bg=BG, fg=MUTED, anchor="w",
             justify="left", wraplength=740).pack(fill="x", pady=(2, 14))

    # ---------------- card ----------------
    card = tk.Frame(outer, bg=CARD, highlightbackground=BORDER,
                    highlightthickness=1, bd=0, padx=20, pady=16)
    card.pack(fill="x")
    card.columnconfigure(0, weight=1)

    def file_row(parent, label_text, var, command, row):
        tk.Label(parent, text=label_text, font=F_LABEL, bg=CARD, fg=TEXT, anchor="w"
                 ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 4))
        holder = tk.Frame(parent, bg=CARD)
        holder.grid(row=row + 1, column=0, columnspan=2, sticky="ew", pady=(0, 14))
        holder.columnconfigure(0, weight=1)
        entry = tk.Entry(holder, textvariable=var, font=F_LABEL, relief="solid", bd=1,
                         highlightthickness=0, bg=CARD, fg=TEXT, insertbackground=TEXT)
        entry.grid(row=0, column=0, sticky="ew", ipady=5)
        ttk.Button(holder, text="Select...", style="Card.TButton", command=command
                   ).grid(row=0, column=1, padx=(10, 0))
        return entry

    def pick_old():
        path = filedialog.askopenfilename(title="Select the ORIGINAL (old) Excel file",
                                          filetypes=EXCEL_TYPES)
        if path:
            old_var.set(path)

    def pick_new():
        path = filedialog.askopenfilename(title="Select the MODIFIED (new) Excel file",
                                          filetypes=EXCEL_TYPES)
        if path:
            new_var.set(path)

    def pick_template():
        path = filedialog.askopenfilename(title="Select the blank TEMPLATE (optional)",
                                          filetypes=EXCEL_TYPES)
        if path:
            tpl_var.set(path)

    def pick_out():
        path = filedialog.askdirectory(title="Select the output folder")
        if path:
            out_var.set(path)

    file_row(card, "Original file (old)", old_var, pick_old, 0)
    file_row(card, "Modified file (new)", new_var, pick_new, 2)
    file_row(card, "Blank template (optional - its headings are then ignored)",
             tpl_var, pick_template, 4)

    ttk.Separator(card, orient="horizontal").grid(row=6, column=0, columnspan=2,
                                                  sticky="ew", pady=(2, 14))

    # ---------------- what gets compared ----------------
    tk.Label(card, text="Comparison scope", font=F_SECTION, bg=CARD, fg=TEXT, anchor="w"
             ).grid(row=7, column=0, columnspan=2, sticky="w")
    tk.Label(card,
             text=("Every sheet and every cell is compared, values and formulas alike. "
                   "Hover a highlighted\ncell in Excel to read what changed. The template is "
                   "left untouched -- layout, styles and\nsheet tab colours all survive; only "
                   "the cells that differ are coloured."),
             font=F_HINT, bg=CARD, fg=MUTED, anchor="w", justify="left",
             wraplength=700
             ).grid(row=8, column=0, columnspan=2, sticky="w", pady=(4, 12))

    legend = tk.Frame(card, bg=CARD)
    legend.grid(row=9, column=0, columnspan=2, sticky="w", pady=(0, 14))
    # FILL_* are ARGB for Excel; Tk wants the plain #RRGGBB half.
    for swatch, caption in ((f"#{FILL_RED[2:]}", "original file: every change"),
                            (f"#{FILL_GREEN[2:]}", "modified file: every change")):
        chip = tk.Frame(legend, bg=CARD)
        chip.pack(side="left", padx=(0, 18))
        tk.Frame(chip, bg=swatch, width=15, height=15, highlightbackground=BORDER,
                 highlightthickness=1).pack(side="left")
        tk.Label(chip, text=caption, font=F_HINT, bg=CARD, fg=MUTED).pack(side="left", padx=6)

    ttk.Separator(card, orient="horizontal").grid(row=10, column=0, columnspan=2,
                                                  sticky="ew", pady=(0, 14))

    # ---------------- output ----------------
    tk.Label(card, text="Output", font=F_SECTION, bg=CARD, fg=TEXT, anchor="w"
             ).grid(row=11, column=0, columnspan=2, sticky="w", pady=(0, 10))

    out_holder = tk.Frame(card, bg=CARD)
    out_holder.grid(row=12, column=0, columnspan=2, sticky="ew")
    out_holder.columnconfigure(1, weight=1)
    tk.Label(out_holder, text="Save to:", font=F_LABEL, bg=CARD, fg=TEXT, width=14,
             anchor="w").grid(row=0, column=0, sticky="w")
    tk.Entry(out_holder, textvariable=out_var, font=F_LABEL, relief="solid", bd=1,
             bg=CARD, fg=TEXT, insertbackground=TEXT).grid(row=0, column=1, sticky="ew", ipady=5)
    ttk.Button(out_holder, text="Browse...", style="Card.TButton", command=pick_out
               ).grid(row=0, column=2, padx=(10, 0))

    tk.Label(card, text="The two highlighted copies are named after your files, "
                        "with _OLD and _NEW added.",
             font=F_HINT, bg=CARD, fg=MUTED, anchor="w"
             ).grid(row=13, column=0, columnspan=2, sticky="w", pady=(10, 0))

    # ---------------- action row ----------------
    run_btn = tk.Button(actions, text="Compare Files", font=F_BTN, bg=ACCENT, fg="white",
                        activebackground=ACCENT_DARK, activeforeground="white",
                        relief="flat", bd=0, padx=22, pady=9, cursor="hand2")
    run_btn.pack(side="left")

    open_btn = tk.Button(actions, text="Open output folder", font=(ui, 10), bg=BG, fg=MUTED,
                         relief="flat", bd=0, padx=12, pady=9, state="disabled",
                         command=lambda: state["out_dir"] and open_folder(state["out_dir"]))
    open_btn.pack(side="left", padx=10)

    progress = ttk.Progressbar(actions, mode="indeterminate",
                               style="Review.Horizontal.TProgressbar")
    progress.pack(side="left", fill="x", expand=True, padx=(16, 0))

    # ---------------- log ----------------
    scroll = tk.Scrollbar(log_frame)
    scroll.pack(side="right", fill="y")
    log_lines = 12 if screen_h >= 900 else 8 if screen_h >= 700 else 6
    log = tk.Text(log_frame, height=log_lines, wrap="none", font=F_LOG, bg=CARD, fg=TEXT,
                  relief="flat", bd=0, padx=10, pady=8, state="disabled",
                  yscrollcommand=scroll.set)
    log.pack(fill="both", expand=True)
    scroll.config(command=log.yview)

    def write_log(text=""):
        log.configure(state="normal")
        log.insert("end", str(text) + "\n")
        log.see("end")
        log.configure(state="disabled")

    def clear_log():
        log.configure(state="normal")
        log.delete("1.0", "end")
        log.configure(state="disabled")

    # ---------------- worker ----------------
    def worker(old, new, out, template):
        try:
            result = diff_workbooks(old, new, out, template=template,
                                    progress=lambda m: log_queue.put(("log", m)))
            log_queue.put(("done", result))
        except Exception as exc:            # surfaced in the window, not a console
            log_queue.put(("error", exc))

    def finish(message):
        state["running"] = False
        progress.stop()
        run_btn.configure(state="normal", text="Compare Files")
        status_var.set(message)

    def poll_queue():
        try:
            while True:
                kind, payload = log_queue.get_nowait()

                if kind == "log":
                    write_log(payload)
                    status_var.set(str(payload))

                elif kind == "done":
                    result = payload
                    write_log("")
                    write_log("-" * 64)
                    for line in summary_lines(result):
                        write_log(line)
                    write_log("-" * 64)
                    write_log(f"Written: {result.out_old}")
                    write_log(f"Written: {result.out_new}")
                    finish(f"Done - {result.cell_changes} changed cell(s).")
                    open_btn.configure(state="normal", fg=ACCENT)

                elif kind == "error":
                    write_log(f"ERROR: {payload}")
                    finish("Failed.")
                    messagebox.showerror("Comparison failed", str(payload))
        except queue.Empty:
            pass
        root.after(120, poll_queue)

    def start():
        if state["running"]:
            return
        old, new = old_var.get().strip(), new_var.get().strip()
        out = out_var.get().strip()
        template = tpl_var.get().strip() or None

        if not old or not new:
            messagebox.showwarning("Missing file",
                                   "Please select both the original and the modified file.")
            return
        for label, path in (("original", old), ("modified", new)):
            if not Path(path).is_file():
                messagebox.showerror("File not found",
                                     f"The {label} file does not exist:\n{path}")
                return
        if Path(old).resolve() == Path(new).resolve():
            if not messagebox.askyesno("Same file",
                                       "Both fields point to the same file. Continue anyway?"):
                return
        if template and not Path(template).is_file():
            messagebox.showerror("File not found",
                                 f"The template file does not exist:\n{template}")
            return
        if not out:
            out = str(Path(new).parent / "comparison_output")
            out_var.set(out)

        clear_log()
        open_btn.configure(state="disabled", fg=MUTED)
        state["running"] = True
        state["out_dir"] = out
        run_btn.configure(state="disabled", text="Working...")
        progress.start(12)
        write_log(f"Original: {old}")
        write_log(f"Modified: {new}")
        if template:
            write_log(f"Template: {template}")
        write_log(f"Output:   {out}")
        write_log("")

        threading.Thread(target=worker, args=(old, new, out, template),
                         daemon=True).start()

    run_btn.configure(command=start)
    root.bind("<Return>", lambda _e: start())

    write_log(f"{APP_TITLE} {__version__}")
    write_log("")
    write_log("Select the original file and the modified file, then press Compare Files.")
    write_log("")
    write_log("Your two files are never modified - two highlighted copies are written")
    write_log("to the output folder. Hover a highlighted cell in Excel to read the note")
    write_log("describing the change; deleted cells say what they used to contain.")
    root.after(120, poll_queue)

    return {"root": root, "old_var": old_var, "new_var": new_var, "out_var": out_var,
            "tpl_var": tpl_var,
            "status_var": status_var, "start": start, "log": log, "state": state}


def launch_gui() -> int:
    try:
        import tkinter as tk
    except ImportError:
        sys.exit("tkinter is not installed.\n"
                 "  Debian/Ubuntu: sudo apt install python3-tk\n"
                 "  Fedora:        sudo dnf install python3-tkinter\n"
                 "  Windows/macOS: reinstall Python with the 'tcl/tk' option ticked")
    root = tk.Tk()
    build_gui(root)
    root.mainloop()
    return 0


# ==========================================================================
# CLI
# ==========================================================================

def main(argv=None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if not argv:
        return launch_gui()

    p = argparse.ArgumentParser(description=f"{APP_TITLE} - compare two Excel files "
                                            "and highlight every change.")
    p.add_argument("old", nargs="?", type=Path, help="the original (old) workbook")
    p.add_argument("new", nargs="?", type=Path, help="the modified (new) workbook")
    p.add_argument("-o", "--out-dir", type=Path, default=Path("comparison_output"),
                   help="output folder (default: comparison_output)")
    p.add_argument("-t", "--template", type=Path, default=None,
                   help="the blank template both files were built from; its own "
                        "headings and labels are then ignored (never modified)")
    p.add_argument("--gui", action="store_true", help="open the window instead")
    p.add_argument("--version", action="version",
                   version=f"{APP_TITLE} {__version__}")
    args = p.parse_args(argv)

    if args.gui:
        return launch_gui()
    if not args.old or not args.new:
        p.error("give both files, or no arguments at all to open the window")
    for path in (args.old, args.new):
        if not path.is_file():
            p.error(f"file not found: {path}")
    if args.template and not args.template.is_file():
        p.error(f"template not found: {args.template}")

    print(f"{APP_TITLE} {__version__}")
    result = diff_workbooks(args.old, args.new, args.out_dir, template=args.template)

    print()
    for line in summary_lines(result):
        print(line)
    print(f"\nWritten: {result.out_old}")
    print(f"Written: {result.out_new}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())