#!/usr/bin/env python3
"""Word Document Change Review.

Compares two Word documents (.doc / .docx / .docm) and produces ONE single
comparison report (a .docx) that contains *only the sections that changed*.

For every changed section the report shows, back to back:

    <Section number + title>
    ================ BEFORE ================
    <the original section, copied with its exact formatting>
    ================ AFTER  ================
    <the modified section, copied with its exact formatting>

Only the actual changed words/phrases are highlighted (red on the BEFORE
side, green on the AFTER side); everything else keeps its original look.
Unchanged sections never appear, and no artificial placeholder text (e.g.
"line removed here") is ever inserted -- deleted text simply shows up only
on the BEFORE side and added text only on the AFTER side.

By default the comparison is confined to the "Requirements" chapter: the
first Heading 1 whose title contains the word "Requirements" (e.g. "3
Requirements", "3 System Requirements", "4 Functional Requirements") through
to the *next* Heading 1. Everything outside that chapter -- front matter
before it, and appendices / annexes / revision history after it -- is never
compared and never appears in the report.

To review specific parts instead, switch to "Specific section(s)" mode and
type one section per field (use "+ Add another section" for several). In that
mode the report is laid out one requested section at a time, in the order they
were entered:

    <Report title>                 (Heading 1, e.g. "Impact SDDD CSC CR ...")
        <Section 1 heading>        (Heading 2)
            Before                 (Heading 3)  <- the whole original section
            After                  (Heading 3)  <- the whole modified section
        <Section 2 heading>
            Before
            After
        ...

Formatting fidelity is achieved by deep-copying the underlying Office Open
XML of each section (paragraphs AND tables) straight into the report, then
carrying over the styles, list-numbering definitions and embedded images
they reference, so a copied section looks exactly as it did in its source
document.
"""

from __future__ import annotations

# --------------------------------------------------------------------------- #
# Standard library imports
# --------------------------------------------------------------------------- #
import difflib
import logging
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import zipfile
from copy import deepcopy
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

# Tkinter is only needed for the desktop GUI. Import it lazily so the
# comparison engine can be imported and driven headlessly (e.g. from a
# script or a server) on systems where Tk is not installed.
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    _TK_AVAILABLE = True
except ImportError:  # pragma: no cover - headless environments
    tk = None  # type: ignore
    filedialog = messagebox = ttk = None  # type: ignore
    _TK_AVAILABLE = False

# --------------------------------------------------------------------------- #
# Third-party imports
# --------------------------------------------------------------------------- #
from docx import Document
from docx.document import Document as DocxDocument
from docx.image.image import Image
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.parts.image import ImagePart
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph

# ``rapidfuzz`` gives fast, high-quality fuzzy ratios but is an optional
# third-party dependency. When it is unavailable we transparently fall back
# to the standard library's ``difflib`` so the tool keeps working (only a
# little slower on very large documents). Both expose a ``fuzz.ratio(a, b)``
# returning a 0..100 similarity score.
try:
    from rapidfuzz import fuzz  # type: ignore
except ImportError:  # pragma: no cover - exercised only without rapidfuzz
    class _DifflibFuzz:
        """Minimal drop-in replacement for ``rapidfuzz.fuzz`` using difflib."""

        @staticmethod
        def ratio(a: str, b: str) -> float:
            return difflib.SequenceMatcher(None, a, b, autojunk=False).ratio() * 100.0

    fuzz = _DifflibFuzz()  # type: ignore


# =========================================================================== #
#                                                                             #
#  utils                                                                      #
#  - logging setup, custom exceptions, file validation, .doc conversion,     #
#    .docm repackaging                                                        #
#                                                                             #
# =========================================================================== #

_LOG_FORMAT = "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"


def setup_logging(log_file: Optional[Path] = None, level: int = logging.INFO) -> logging.Logger:
    """Configure and return the application's root logger."""
    logger = logging.getLogger("word_diff_reviewer")
    logger.setLevel(level)
    logger.handlers.clear()

    formatter = logging.Formatter(_LOG_FORMAT)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    if log_file is not None:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    return logger


def get_logger(name: str) -> logging.Logger:
    """Return a child logger under the application's namespace."""
    return logging.getLogger(f"word_diff_reviewer.{name}")


# --------------------------------------------------------------------------- #
# Custom exceptions
# --------------------------------------------------------------------------- #


class DocumentComparisonError(Exception):
    """Base exception for all application-specific errors."""


class InvalidFileError(DocumentComparisonError):
    """Raised when a selected file does not exist or is not readable."""


class UnsupportedFormatError(DocumentComparisonError):
    """Raised when a file extension is not one of .doc/.docx/.docm."""


class CorruptedDocumentError(DocumentComparisonError):
    """Raised when a Word document cannot be parsed because it is corrupted
    or not a valid Office Open XML / binary Word file."""


class ConversionError(DocumentComparisonError):
    """Raised when a legacy .doc file cannot be converted to .docx."""


class EmptyDocumentError(DocumentComparisonError):
    """Raised when a document contains no extractable text."""


class SectionNotFoundError(DocumentComparisonError):
    """Raised when a requested target section/heading cannot be located."""


class PermissionDeniedError(DocumentComparisonError):
    """Raised when the application lacks permission to read or write a file."""


# --------------------------------------------------------------------------- #
# File validation
# --------------------------------------------------------------------------- #

SUPPORTED_EXTENSIONS = {".doc", ".docx", ".docm"}


def validate_input_file(path: str | Path) -> Path:
    """Validate that a path points to a readable, supported Word document."""
    p = Path(path).expanduser().resolve()

    if not p.exists():
        raise InvalidFileError(f"File does not exist: {p}")
    if not p.is_file():
        raise InvalidFileError(f"Path is not a file: {p}")
    if p.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise UnsupportedFormatError(
            f"Unsupported file format '{p.suffix}'. Supported formats: "
            f"{', '.join(sorted(SUPPORTED_EXTENSIONS))}"
        )
    try:
        with open(p, "rb"):
            pass
    except PermissionError as exc:
        raise PermissionDeniedError(f"Permission denied reading file: {p}") from exc

    if p.stat().st_size == 0:
        raise EmptyDocumentError(f"File is empty: {p}")

    return p


def validate_output_folder(path: str | Path) -> Path:
    """Validate (and create if necessary) an output folder."""
    p = Path(path).expanduser().resolve()
    try:
        p.mkdir(parents=True, exist_ok=True)
        probe = p / ".write_test.tmp"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink(missing_ok=True)
    except OSError as exc:
        raise PermissionDeniedError(f"Cannot write to output folder: {p} ({exc})") from exc
    return p


# --------------------------------------------------------------------------- #
# Legacy .doc conversion
# --------------------------------------------------------------------------- #


def _find_soffice() -> Optional[str]:
    """Locate a LibreOffice/OpenOffice executable on PATH, if any."""
    for candidate in ("soffice", "soffice.exe", "libreoffice"):
        found = shutil.which(candidate)
        if found:
            return found
    return None


def convert_doc_to_docx(doc_path: Path, work_dir: Optional[Path] = None) -> Path:
    """Convert a legacy binary .doc file into a modern .docx file.

    Tries, in order:
      1. Microsoft Word COM automation via pywin32 (Windows only).
      2. LibreOffice/OpenOffice headless conversion (cross platform).
    """
    logger = get_logger("utils.convert")
    work_dir = work_dir or Path(tempfile.mkdtemp(prefix="docx_convert_"))
    work_dir.mkdir(parents=True, exist_ok=True)
    target = work_dir / (doc_path.stem + ".docx")

    # --- Attempt 1: Windows COM automation via pywin32 --------------------
    if sys.platform.startswith("win"):
        try:
            import win32com.client  # type: ignore

            logger.info("Converting %s using Microsoft Word COM automation", doc_path)
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            try:
                doc = word.Documents.Open(str(doc_path))
                # wdFormatXMLDocument = 12 (.docx)
                doc.SaveAs(str(target), FileFormat=12)
                doc.Close()
            finally:
                word.Quit()
            if target.exists():
                return target
        except Exception as exc:  # pragma: no cover - Windows only path
            logger.warning("Word COM conversion failed: %s", exc)

    # --- Attempt 2: LibreOffice headless -----------------------------------
    soffice = _find_soffice()
    if soffice:
        try:
            logger.info("Converting %s using LibreOffice headless (%s)", doc_path, soffice)
            subprocess.run(
                [
                    soffice,
                    "--headless",
                    "--norestore",
                    "--convert-to",
                    "docx",
                    "--outdir",
                    str(work_dir),
                    str(doc_path),
                ],
                check=True,
                capture_output=True,
                timeout=120,
            )
            if target.exists():
                return target
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as exc:
            logger.warning("LibreOffice conversion failed: %s", exc)

    raise ConversionError(
        f"Could not convert legacy .doc file '{doc_path.name}' to .docx. "
        "Please install LibreOffice (recommended, cross-platform) or Microsoft "
        "Word (Windows) so the application can convert legacy .doc files, or "
        "save the file as .docx manually before comparing."
    )


# --------------------------------------------------------------------------- #
# .docm (macro-enabled) repackaging
# --------------------------------------------------------------------------- #

_DOCM_MAIN_CT = "application/vnd.ms-word.document.macroEnabled.main+xml"
_DOCX_MAIN_CT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
)


def repackage_docm_as_docx(docm_path: Path, work_dir: Path) -> Path:
    """Repackage a macro-enabled .docm as a plain .docx that python-docx
    can open.

    A .docm is byte-for-byte the same OOXML package as a .docx except for
    (a) the main document part's content type and (b) the embedded VBA
    project parts. python-docx refuses the macro-enabled content type, which
    is why .docm files previously failed with an "unsupported" error even
    though the document content itself is fully readable.

    This function copies the package, rewrites the content type to the
    standard .docx one and strips the VBA parts (macros are irrelevant to a
    text comparison and must not block it). All document content, styles,
    numbering, images, headers/footers etc. are carried over unchanged.
    """
    logger = get_logger("utils.docm")
    work_dir.mkdir(parents=True, exist_ok=True)
    target = work_dir / (docm_path.stem + "_from_docm.docx")

    try:
        with zipfile.ZipFile(docm_path) as zin, zipfile.ZipFile(
            target, "w", zipfile.ZIP_DEFLATED
        ) as zout:
            for name in zin.namelist():
                # Drop the VBA project itself and its relationship part.
                if name.startswith("word/vbaProject") or name.startswith(
                    "word/_rels/vbaProject"
                ) or name == "word/vbaData.xml":
                    continue

                data = zin.read(name)

                if name == "[Content_Types].xml":
                    xml = data.decode("utf-8")
                    xml = xml.replace(_DOCM_MAIN_CT, _DOCX_MAIN_CT)
                    xml = re.sub(r"<Override[^>]*vba(?:Project|Data)[^>]*/>", "", xml)
                    data = xml.encode("utf-8")
                elif name == "word/_rels/document.xml.rels":
                    xml = data.decode("utf-8")
                    xml = re.sub(
                        r"<Relationship[^>]*vba(?:Project|Data)[^>]*/>", "", xml
                    )
                    data = xml.encode("utf-8")

                zout.writestr(name, data)
    except zipfile.BadZipFile as exc:
        raise CorruptedDocumentError(
            f"'{docm_path.name}' is not a valid Word package (corrupted zip)."
        ) from exc

    logger.info("Repackaged %s -> %s", docm_path.name, target.name)
    return target


def is_docx_zip(path: Path) -> bool:
    """Quickly check whether a file is a valid ZIP-based Office document."""
    try:
        return zipfile.is_zipfile(path)
    except OSError:
        return False


# =========================================================================== #
#                                                                             #
#  text extraction helpers                                                    #
#  - single source of truth for paragraph text so that the character         #
#    offsets used by the diff engine always line up exactly with the runs    #
#    walked by the in-place highlighter                                      #
#                                                                             #
# =========================================================================== #

_W_T = None  # populated lazily because qn() needs the docx namespaces


def _tags():
    """Cache frequently used qualified tag names."""
    global _W_T
    if _W_T is None:
        _W_T = {
            "t": qn("w:t"),
            "br": qn("w:br"),
            "cr": qn("w:cr"),
            "tab": qn("w:tab"),
            "r": qn("w:r"),
            "rPr": qn("w:rPr"),
            "p": qn("w:p"),
            "tbl": qn("w:tbl"),
        }
    return _W_T


def run_text(r_elem) -> str:
    """Extract the visible text of a single <w:r>, mapping soft line breaks
    to \\n and tabs to \\t. Non-text content (images, field codes, etc.)
    contributes nothing."""
    tags = _tags()
    parts: List[str] = []
    for child in r_elem:
        tag = child.tag
        if tag == tags["t"]:
            parts.append(child.text or "")
        elif tag in (tags["br"], tags["cr"]):
            parts.append("\n")
        elif tag == tags["tab"]:
            parts.append("\t")
    return "".join(parts)


def paragraph_runs(p_elem) -> List:
    """All <w:r> elements of a paragraph in document order, including runs
    nested inside hyperlinks, smart tags, etc. Word habitually fragments a
    visually contiguous sentence across many runs, so text must always be
    assembled by walking every run."""
    tags = _tags()
    return list(p_elem.iter(tags["r"]))


def paragraph_text(p_elem) -> str:
    """Full visible text of a paragraph, assembled run by run with the same
    rules as :func:`run_text` (this MUST stay consistent with the
    highlighter's run walk)."""
    return "".join(run_text(r) for r in paragraph_runs(p_elem))


# =========================================================================== #
#                                                                             #
#  document_parser                                                            #
#  - loads Word docs and extracts a flat list of DocumentElement objects,    #
#    each holding a live reference to its paragraph for in-place markup      #
#                                                                             #
# =========================================================================== #

logger = get_logger("document_parser")

# Word "Heading 1".."Heading 9" style names map to outline levels 1-9.
_HEADING_STYLE_RE = re.compile(r"^Heading\s*(\d)$", re.IGNORECASE)
_TITLE_STYLES = {"Title", "Subtitle"}


@dataclass
class DocumentElement:
    """A single logical unit of document content (one paragraph, either in
    the body or inside a table cell), with a live reference to the
    underlying python-docx Paragraph so changes can be highlighted in place
    without ever recreating (and thereby reformatting) the content."""

    index: int                     # order of appearance in the document
    text: str                      # visible text (run-walk extraction)
    style_name: str                # underlying Word paragraph style
    is_heading: bool               # True for Heading 1-9 / Title / Subtitle
    heading_level: Optional[int]   # 1-9 for headings, None otherwise
    paragraph: Paragraph           # live reference into the open document
    section_path: List[str] = field(default_factory=list)
    is_table_cell: bool = False

    @property
    def section_label(self) -> str:
        if self.section_path:
            return " > ".join(self.section_path)
        return "Document Start"


class DocumentParser:
    """Parses .doc/.docx/.docm files into a list of ``DocumentElement``."""

    def __init__(self, temp_dir: Optional[Path] = None) -> None:
        self._temp_dir = temp_dir or Path(tempfile.mkdtemp(prefix="word_diff_parse_"))

    # ------------------------------------------------------------------ #
    # Public API
    # ------------------------------------------------------------------ #

    def load(self, path: Path) -> DocxDocument:
        """Open a Word file, converting legacy .doc and repackaging
        macro-enabled .docm first as needed.

        Raises:
            CorruptedDocumentError: If the file cannot be parsed.
            ConversionError: If a .doc file cannot be converted.
        """
        path = Path(path)
        suffix = path.suffix.lower()

        if suffix == ".doc":
            logger.info("Legacy .doc detected, converting: %s", path)
            docx_path = convert_doc_to_docx(path, self._temp_dir)
        elif suffix == ".docm":
            # python-docx rejects the macro-enabled main-part content type,
            # so repackage the (otherwise identical) OOXML container as a
            # plain .docx. Macros never block the comparison this way.
            logger.info("Macro-enabled .docm detected, repackaging: %s", path)
            docx_path = repackage_docm_as_docx(path, self._temp_dir)
        else:
            docx_path = path

        if not is_docx_zip(docx_path):
            raise CorruptedDocumentError(
                f"'{path.name}' does not appear to be a valid Word document "
                "(not a recognizable OOXML package). It may be corrupted, "
                "password-protected, or in an unsupported legacy format."
            )

        try:
            return Document(str(docx_path))
        except Exception as exc:  # python-docx raises varied exception types
            raise CorruptedDocumentError(
                f"Failed to parse '{path.name}': {exc}"
            ) from exc

    def extract_elements(self, document: DocxDocument) -> List[DocumentElement]:
        """Flatten a document's body into an ordered list of paragraph
        elements (body paragraphs plus every paragraph inside every table
        cell, including nested tables), tracking the heading breadcrumb.

        Raises:
            EmptyDocumentError: If no text content is found at all.
        """
        tags = _tags()
        elements: List[DocumentElement] = []
        heading_stack: List[Tuple[int, str]] = []  # (level, text)
        state = {"idx": 0}

        # Pre-index paragraphs/tables by their underlying XML element id so
        # lookups during the body walk are O(1).
        para_by_elem = {id(p._p): p for p in document.paragraphs}
        table_by_elem = {id(t._tbl): t for t in document.tables}

        def add_paragraph(para: Paragraph, in_table: bool) -> None:
            text = paragraph_text(para._p)
            style_name = para.style.name if para.style else "Normal"
            level = self._heading_level(style_name)
            is_heading = level is not None or style_name in _TITLE_STYLES

            if is_heading and text.strip() and not in_table:
                eff_level = level or 0
                while heading_stack and heading_stack[-1][0] >= eff_level:
                    heading_stack.pop()
                heading_stack.append((eff_level, text.strip()))

            elements.append(
                DocumentElement(
                    index=state["idx"],
                    text=text,
                    style_name=style_name,
                    is_heading=is_heading,
                    heading_level=level,
                    paragraph=para,
                    section_path=[h[1] for h in heading_stack],
                    is_table_cell=in_table,
                )
            )
            state["idx"] += 1

        def walk_table(table) -> None:
            seen_cells = set()
            for row in table.rows:
                for cell in row.cells:
                    # Merged cells appear multiple times in row.cells.
                    if id(cell._tc) in seen_cells:
                        continue
                    seen_cells.add(id(cell._tc))
                    for para in cell.paragraphs:
                        add_paragraph(para, in_table=True)
                    for nested in cell.tables:
                        walk_table(nested)

        body = document.element.body
        for child in body.iterchildren():
            if child.tag == tags["p"]:
                para = para_by_elem.get(id(child))
                if para is not None:
                    add_paragraph(para, in_table=False)
            elif child.tag == tags["tbl"]:
                table = table_by_elem.get(id(child))
                if table is not None:
                    walk_table(table)

        if not any(e.text.strip() for e in elements):
            raise EmptyDocumentError("Document contains no extractable text.")

        return elements

    def slice_from_requirements(
        self, elements: List[DocumentElement], keyword: str = "requirements"
    ) -> List[DocumentElement]:
        """Return the elements from the first Heading 1 whose title contains
        ``keyword`` (case-insensitive) through the END of the document.

        Kept for callers that really do want everything from the Requirements
        heading onwards. The comparison itself uses
        :meth:`find_requirements_bounds`, which stops at the next Heading 1 so
        appendices and other chapters stay out of the report.

        Handles the usual variants automatically: "3 Requirements",
        "3. System Requirements", "3 Functional Requirements",
        "4 Requirements", ... -- the section number is irrelevant, only the
        Heading 1 level and the word "Requirements" matter.

        Raises:
            SectionNotFoundError: If no matching Heading 1 exists.
        """
        keyword = keyword.strip().lower()
        for i, el in enumerate(elements):
            if (
                el.is_heading
                and el.heading_level == 1
                and keyword in el.text.strip().lower()
            ):
                logger.info(
                    "Comparison starts at Heading 1: %r (element %d)",
                    el.text.strip(),
                    i,
                )
                return elements[i:]
        raise SectionNotFoundError(
            "Could not find a Heading 1 containing the word 'Requirements'. "
            "Check that the document uses the built-in 'Heading 1' style for "
            "its section titles, or switch to 'Compare entire document'."
        )

    def find_requirements_bounds(
        self, elements: List[DocumentElement], keyword: str = "requirements"
    ) -> Optional[Tuple[int, int]]:
        """Locate the "Requirements" chapter and return ``(start_index,
        end_index)`` into ``elements`` -- the Heading 1 itself through
        everything below it, up to but NOT including the next Heading 1.
        ``end_index`` is exclusive; it is ``len(elements)`` only when the
        chapter genuinely runs to the end of the document.

        Stopping at the next Heading 1 is what keeps unrelated chapters --
        appendices, annexes, revision history, glossaries -- out of the
        comparison entirely: unless a specific section is typed into the
        selection box, only the Requirements chapter is ever compared.

        Returns ``None`` when the document has no such Heading 1.
        """
        keyword = keyword.strip().lower()

        start_i = None
        for i, el in enumerate(elements):
            if (
                el.is_heading
                and el.heading_level == 1
                and not el.is_table_cell
                and keyword in el.text.strip().lower()
            ):
                start_i = i
                break
        if start_i is None:
            return None

        start_level = elements[start_i].heading_level or 1
        end_i = len(elements)
        for j in range(start_i + 1, len(elements)):
            el = elements[j]
            if (
                el.is_heading
                and not el.is_table_cell
                and (el.heading_level or 0) <= start_level
            ):
                end_i = j  # next chapter -> the Requirements chapter ends here
                break

        logger.info(
            "Requirements chapter: %r (elements %d..%d, next chapter: %s)",
            elements[start_i].text.strip(),
            start_i,
            end_i,
            elements[end_i].text.strip() if end_i < len(elements) else "(end of document)",
        )
        return start_i, end_i

    def find_section_bounds(
        self,
        elements: List[DocumentElement],
        target: str,
        numberer: Optional["HeadingNumberer"] = None,
    ) -> Tuple[int, int]:
        """Locate a named section and return ``(start_index, end_index)`` into
        ``elements`` -- the section start through everything up to (but not
        including) the next sibling-or-higher section. ``end_index`` is
        exclusive and equals ``len(elements)`` when the section runs to the end.

        Matching behaves like a very forgiving Ctrl+F. It tolerates differences
        in capitalisation, whitespace (spaces vs tabs vs double spaces),
        punctuation/symbols, dashes and section-numbering style, and it finds a
        section whether it is a real Heading or merely a numbered *paragraph*
        typed into the body text (e.g. "3.2.1 ..."). Auto-numbered headings are
        matched by the number the reader sees, not just their stored text.

        Raises:
            SectionNotFoundError: If nothing matches well enough.
        """
        bounds = self.try_find_section_bounds(elements, target, numberer)
        if bounds is None:
            available = ", ".join(self.available_section_labels(elements)[:40])
            raise SectionNotFoundError(
                f"Could not find a section matching '{target}'.\n\n"
                f"Sections detected in the document: {available or '(none)'}."
            )
        return bounds

    def try_find_section_bounds(
        self,
        elements: List[DocumentElement],
        target: str,
        numberer: Optional["HeadingNumberer"] = None,
    ) -> Optional[Tuple[int, int]]:
        """Like :meth:`find_section_bounds` but returns ``None`` instead of
        raising when the section cannot be located (used so a section that
        exists in only one of the two documents -- freshly added or removed --
        can still be compared rather than aborting)."""
        # 1) Prefer a real Heading match, bounded by heading levels.
        hi = self._find_heading_index(elements, target, numberer)
        if hi is not None:
            start_level = elements[hi].heading_level or 0
            end_i = len(elements)
            for j in range(hi + 1, len(elements)):
                el = elements[j]
                if (
                    el.is_heading
                    and not el.is_table_cell
                    and (el.heading_level or 0) <= start_level
                ):
                    end_i = j
                    break
            return hi, end_i

        # 2) Fallback: a numbered paragraph typed into the body (not a Heading).
        ai = self._find_any_index(elements, target, numberer)
        if ai is None:
            return None
        start_num = self._leading_number(elements[ai].text) or self._leading_number(target)
        # Requirement-style ids ("REQ-SDDD_CLSW_LOADSW_LUP-0087") carry no
        # dotted section number, so the sibling that ends the block is the next
        # id sharing the same prefix rather than the next numbered item.
        id_prefix = self._id_prefix(elements[ai].text) or self._id_prefix(target)
        end_i = len(elements)
        for j in range(ai + 1, len(elements)):
            el = elements[j]
            num = self._leading_number(el.text)
            is_body_head = el.is_heading and not el.is_table_cell
            if is_body_head:
                # A real heading ends the body-section unless it nests under it.
                if start_num and num and self._is_number_descendant(num, start_num):
                    continue
                end_i = j
                break
            if id_prefix and self._id_prefix(el.text) == id_prefix:
                end_i = j  # the next requirement id ends this one
                break
            if start_num and num and not self._is_number_descendant(num, start_num):
                end_i = j  # a sibling-or-higher numbered item ends the section
                break
        return ai, end_i

    def available_section_labels(self, elements: List[DocumentElement]) -> List[str]:
        """Human-readable labels of the headings detected, for error messages."""
        labels = []
        for el in elements:
            if el.is_heading and not el.is_table_cell and el.text.strip():
                labels.append(el.text.strip())
        return labels

    def extract_section(
        self,
        elements: List[DocumentElement],
        target: str,
        numberer: Optional["HeadingNumberer"] = None,
    ) -> List[DocumentElement]:
        """Return only the elements belonging to a named section."""
        start_i, end_i = self.find_section_bounds(elements, target, numberer)
        return elements[start_i:end_i]

    # ------------------------------------------------------------------ #
    # Robust heading matching
    # ------------------------------------------------------------------ #

    def _find_any_index(
        self,
        elements: List[DocumentElement],
        target: str,
        numberer: Optional["HeadingNumberer"],
    ) -> Optional[int]:
        """Ctrl+F-style search over *all* paragraphs (not just headings), used
        when no Heading matched. Prefers an exact leading-number match, then a
        strong title match, so section numbers typed as ordinary text are still
        found without matching unrelated prose."""
        t_num = self._leading_number(target)
        t_title = self._norm_ws(self._drop_leading_number(target))
        t_title_key = self._alnum_key(self._drop_leading_number(target))

        # (a) Exact leading-number match anywhere in the body.
        if t_num:
            for i, el in enumerate(elements):
                if el.is_table_cell:
                    continue
                if self._leading_number(el.text) == t_num:
                    return i

        # (b) Strong title match (high threshold to avoid matching prose).
        if t_title_key:
            best_i, best = None, 90.0
            for i, el in enumerate(elements):
                if el.is_table_cell or not el.text.strip():
                    continue
                body = self._norm_ws(el.text)
                body_title = self._norm_ws(self._drop_leading_number(el.text))
                body_key = self._alnum_key(self._drop_leading_number(el.text))
                score = 0.0
                if t_title_key and t_title_key == body_key:
                    score = 99.0
                elif (
                    t_title and body_title
                    and len(t_title) >= 4 and len(body_title) >= 4
                    and (t_title in body_title or body_title in t_title)
                ):
                    score = 93.0
                else:
                    score = fuzz.ratio(t_title, body_title)
                if score > best:
                    best, best_i = score, i
            return best_i
        return None

    @staticmethod
    def _leading_number(text: str) -> Optional[str]:
        """Return a leading dotted section number ("3", "3.2", "3.2.1"),
        normalised to dot separators, or ``None`` if the text does not start
        with one."""
        m = re.match(r"\s*(\d+(?:[.\-]\d+)*)", text or "")
        if not m:
            return None
        return m.group(1).replace("-", ".")

    @staticmethod
    def _id_prefix(text: str) -> Optional[str]:
        """For a requirement-style identifier at the start of a paragraph
        ("REQ-SDDD_CLSW_LOADSW_LUP-0087") return everything before its trailing
        number ("REQ-SDDD_CLSW_LOADSW_LUP-"); ``None`` for ordinary prose.

        This is what lets one requirement be isolated from the next when the
        ids are typed as plain paragraphs instead of Word headings."""
        parts = (text or "").strip().split()
        if not parts:
            return None
        first = parts[0].rstrip(":.;,")
        m = re.match(r"^([A-Za-z][0-9A-Za-z_.\-]*[-_.])(\d{2,})$", first)
        if not m:
            return None
        return m.group(1)

    @staticmethod
    def _is_number_descendant(num: str, ancestor: str) -> bool:
        """True if ``num`` is ``ancestor`` itself or nested beneath it
        ("3.2.1" is under "3.2"; "3.3" and "4" are not)."""
        return num == ancestor or num.startswith(ancestor + ".")

    def _find_heading_index(
        self,
        elements: List[DocumentElement],
        target: str,
        numberer: Optional["HeadingNumberer"],
        threshold: float = 78.0,
    ) -> Optional[int]:
        """Return the index of the body heading that best matches ``target``
        (above ``threshold``), or ``None`` if nothing matches well enough."""
        best_i: Optional[int] = None
        best_score = threshold
        for i, el in enumerate(elements):
            if not el.is_heading or el.is_table_cell:
                continue
            score = self._heading_match_score(target, el, numberer)
            if score > best_score:
                best_score = score
                best_i = i
        return best_i

    def _heading_match_score(
        self,
        target: str,
        el: DocumentElement,
        numberer: Optional["HeadingNumberer"],
    ) -> float:
        """Score 0..100 for how well ``target`` matches a heading element."""
        title = el.text.strip()
        number = numberer.number_for(el.paragraph._p) if numberer else None
        full = f"{number} {title}" if number else title

        # Candidate forms of the heading.
        full_forms = {self._norm_ws(full), self._norm_ws(title)}
        title_forms = {
            self._norm_ws(self._drop_leading_number(full)),
            self._norm_ws(self._drop_leading_number(title)),
        }
        number_forms = set()
        if number:
            number_forms.add(self._norm_ws(number))
        lit = re.match(r"^\s*(\d+(?:[.\-]\d+)*)", title)
        if lit:
            number_forms.add(self._norm_ws(lit.group(1)))

        # Candidate forms of the user's target.
        t_full = self._norm_ws(target)
        t_title = self._norm_ws(self._drop_leading_number(target))
        t_key = self._alnum_key(target)
        t_key_title = self._alnum_key(self._drop_leading_number(target))

        full_keys = {self._alnum_key(f) for f in full_forms}
        title_keys = {self._alnum_key(f) for f in title_forms}

        best = 0.0
        if t_full and t_full in full_forms:
            best = max(best, 100.0)
        if t_title and t_title in title_forms:
            best = max(best, 97.0)
        if t_full and t_full in number_forms:
            best = max(best, 95.0)
        if t_key and t_key in full_keys:
            best = max(best, 94.0)
        if t_key_title and t_key_title in title_keys:
            best = max(best, 92.0)
        # Containment (the user typed part of the heading, or vice versa).
        for f in full_forms:
            if t_full and f and (t_full in f or f in t_full):
                best = max(best, 88.0)
        # Fuzzy, as a last resort, over both full and title-only forms.
        for f in full_forms | title_forms:
            if f:
                best = max(best, fuzz.ratio(t_full, f))
                if t_title:
                    best = max(best, fuzz.ratio(t_title, f))
        return best

    @staticmethod
    def _norm_ws(text: str) -> str:
        """Lowercase and collapse all runs of whitespace (spaces, tabs,
        newlines) to a single space."""
        return re.sub(r"\s+", " ", (text or "").strip()).lower()

    @staticmethod
    def _alnum_key(text: str) -> str:
        """Reduce to lowercase letters/digits only -- so differences in
        whitespace and symbols (_, -, /, :, tabs, ...) don't defeat a match."""
        return re.sub(r"[^0-9a-z]+", "", (text or "").lower())

    @staticmethod
    def _drop_leading_number(text: str) -> str:
        """Strip a leading section number ("3", "3.1", "3.1.2", "A.2", "3.1)")
        while leaving purely textual titles untouched (the leading token must
        contain a digit to be treated as a number)."""
        return re.sub(
            r"^\s*(?=[0-9A-Za-z.\-]*\d)[0-9A-Za-z]+(?:[.\-][0-9A-Za-z]+)*[.)\-]?\s+",
            "",
            text or "",
        ).strip()

    # ------------------------------------------------------------------ #
    # Internal helpers
    # ------------------------------------------------------------------ #

    @staticmethod
    def _heading_level(style_name: str) -> Optional[int]:
        match = _HEADING_STYLE_RE.match(style_name or "")
        if match:
            return int(match.group(1))
        return None


# Matches ONLY something that really looks like a section number at the start
# of a heading: "3 ", "3.1 ", "3.1.2. ", "A.1 ", "A) ", "B. ".
# Deliberately stricter than "first word": the old pattern also swallowed a
# plain leading word, so "Functional Requirements" and "Performance
# Requirements" both collapsed to "requirements" and a genuine renaming was
# reported as *no* change.
_SECTION_NUMBER_RE = re.compile(
    r"^[\s\u00a0]*(?:"
    r"\d+(?:[.\-]\d+)*[.)]?"          # 3 | 3.1 | 3-1-2 | 3.1.2.
    r"|[A-Za-z][.\-]\d+(?:[.\-]\d+)*[.)]?"  # A.1 | B-2.3
    r"|[A-Za-z][.)]"                  # A. | B)
    r")[\s\u00a0]+"
)


def _strip_section_number(text: str) -> str:
    """Drop a leading section number ("3.1 ", "A.2 ", ...) from a heading."""
    return _SECTION_NUMBER_RE.sub("", text)


# =========================================================================== #
#                                                                             #
#  comparison_engine                                                          #
#  - paragraph-level pairing + word/character level change ranges            #
#                                                                             #
# =========================================================================== #

logger = get_logger("comparison_engine")

# Paragraphs whose fuzzy similarity is at or above this threshold are
# considered "the same paragraph, edited" rather than "deleted + added".
MOVE_MATCH_THRESHOLD = 55.0

# Tokenizer that preserves whitespace runs as their own tokens so that
# whitespace-only changes (e.g. a doubled space) are still detected.
_TOKEN_RE = re.compile(r"\s+|\S+")


class ChangeType(str, Enum):
    ADDED = "added"
    DELETED = "deleted"
    MODIFIED = "modified"
    MOVED = "moved"


@dataclass
class ElementMark:
    """A highlight instruction for one paragraph of one source document:
    the character ranges (in that paragraph's text) that must be shaded."""

    element: DocumentElement
    ranges: List[Tuple[int, int]]
    change_type: ChangeType


@dataclass
class DiffResult:
    before_marks: List[ElementMark]
    after_marks: List[ElementMark]
    counts: Dict[ChangeType, int]

    @property
    def total(self) -> int:
        return sum(self.counts.values())


class DiffEngine:
    """Pairs original/modified paragraphs and computes, for each changed
    pair, the exact character ranges that differ -- so that only the actual
    modified words/phrases get highlighted, never the whole line."""

    def __init__(self, move_threshold: float = MOVE_MATCH_THRESHOLD) -> None:
        self.move_threshold = move_threshold

    @staticmethod
    def _only_renumbered(o_el: DocumentElement, m_el: DocumentElement) -> bool:
        """True when two paired headings differ *only* by their leading
        section number ("5.2 Scope" vs "6.2 Scope"). Inserting a chapter
        renumbers every heading after it; that is not an edit, so it must not
        be counted or highlighted."""
        if not (o_el.is_heading and m_el.is_heading):
            return False
        return (
            _strip_section_number(o_el.text).strip()
            == _strip_section_number(m_el.text).strip()
        )

    # ------------------------------------------------------------------ #
    # Public API
    # ------------------------------------------------------------------ #

    def compare(
        self,
        original: List[DocumentElement],
        modified: List[DocumentElement],
    ) -> DiffResult:
        orig = [e for e in original if e.text.strip()]
        mod = [e for e in modified if e.text.strip()]

        matcher = difflib.SequenceMatcher(
            a=[e.text for e in orig], b=[e.text for e in mod], autojunk=False
        )

        before_marks: List[ElementMark] = []
        after_marks: List[ElementMark] = []
        counts: Dict[ChangeType, int] = {t: 0 for t in ChangeType}
        pending_deletes: List[DocumentElement] = []
        pending_inserts: List[DocumentElement] = []

        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == "equal":
                continue
            if tag == "replace":
                self._handle_replace(
                    orig[i1:i2], mod[j1:j2],
                    before_marks, after_marks, counts,
                    pending_deletes, pending_inserts,
                )
            elif tag == "delete":
                pending_deletes.extend(orig[i1:i2])
            elif tag == "insert":
                pending_inserts.extend(mod[j1:j2])

        # Reconcile leftover deletes/inserts across the whole document as
        # moved (and possibly edited) paragraphs before treating them as
        # pure removals/additions.
        leftover_d, leftover_i = self._match_moved(
            pending_deletes, pending_inserts, before_marks, after_marks, counts
        )

        for el in leftover_d:
            before_marks.append(
                ElementMark(el, [(0, len(el.text))], ChangeType.DELETED)
            )
            counts[ChangeType.DELETED] += 1

        for el in leftover_i:
            after_marks.append(
                ElementMark(el, [(0, len(el.text))], ChangeType.ADDED)
            )
            counts[ChangeType.ADDED] += 1

        return DiffResult(before_marks, after_marks, counts)

    # ------------------------------------------------------------------ #
    # Internal: replace-block handling
    # ------------------------------------------------------------------ #

    def _handle_replace(
        self,
        orig_slice: List[DocumentElement],
        mod_slice: List[DocumentElement],
        before_marks: List[ElementMark],
        after_marks: List[ElementMark],
        counts: Dict[ChangeType, int],
        pending_deletes: List[DocumentElement],
        pending_inserts: List[DocumentElement],
    ) -> None:
        if len(orig_slice) == 1 and len(mod_slice) == 1:
            pairs = [(orig_slice[0], mod_slice[0])]
            unmatched_o: List[DocumentElement] = []
            unmatched_m: List[DocumentElement] = []
        else:
            pairs, unmatched_o, unmatched_m = self._best_pairing(orig_slice, mod_slice)

        for o_el, m_el in pairs:
            if self._only_renumbered(o_el, m_el):
                continue  # same heading, new number -> not a change
            b_ranges, a_ranges = self.diff_ranges(o_el.text, m_el.text)
            if not b_ranges and not a_ranges:
                continue
            if b_ranges:
                before_marks.append(ElementMark(o_el, b_ranges, ChangeType.MODIFIED))
            if a_ranges:
                after_marks.append(ElementMark(m_el, a_ranges, ChangeType.MODIFIED))
            counts[ChangeType.MODIFIED] += 1

        # Leftovers may still find a partner elsewhere in the document
        # (moved paragraphs), so defer the delete/add decision.
        pending_deletes.extend(unmatched_o)
        pending_inserts.extend(unmatched_m)

    def _best_pairing(
        self,
        orig_slice: List[DocumentElement],
        mod_slice: List[DocumentElement],
    ) -> Tuple[
        List[Tuple[DocumentElement, DocumentElement]],
        List[DocumentElement],
        List[DocumentElement],
    ]:
        """Greedy best-similarity pairing between two small lists of
        paragraphs (used inside a single difflib replace block)."""
        candidates = []
        for o in orig_slice:
            for m in mod_slice:
                score = fuzz.ratio(o.text, m.text)
                candidates.append((score, o, m))
        candidates.sort(key=lambda c: c[0], reverse=True)

        pairs: List[Tuple[DocumentElement, DocumentElement]] = []
        used_o, used_m = set(), set()
        for score, o, m in candidates:
            if o.index in used_o or m.index in used_m:
                continue
            if score < self.move_threshold:
                continue
            pairs.append((o, m))
            used_o.add(o.index)
            used_m.add(m.index)

        leftover_o = [o for o in orig_slice if o.index not in used_o]
        leftover_m = [m for m in mod_slice if m.index not in used_m]
        return pairs, leftover_o, leftover_m

    # ------------------------------------------------------------------ #
    # Internal: moved-paragraph detection
    # ------------------------------------------------------------------ #

    def _match_moved(
        self,
        deletes: List[DocumentElement],
        inserts: List[DocumentElement],
        before_marks: List[ElementMark],
        after_marks: List[ElementMark],
        counts: Dict[ChangeType, int],
    ) -> Tuple[List[DocumentElement], List[DocumentElement]]:
        if not deletes or not inserts:
            return deletes, inserts

        candidates = []
        for d_el in deletes:
            for i_el in inserts:
                score = fuzz.ratio(d_el.text, i_el.text)
                if score >= self.move_threshold:
                    candidates.append((score, d_el, i_el))
        candidates.sort(key=lambda c: c[0], reverse=True)

        used_d, used_i = set(), set()
        for score, d_el, i_el in candidates:
            if d_el.index in used_d or i_el.index in used_i:
                continue
            used_d.add(d_el.index)
            used_i.add(i_el.index)

            if self._only_renumbered(d_el, i_el):
                # A heading that merely got a new section number because
                # something was inserted/removed earlier in the document:
                # not a real move, so nothing is highlighted.
                counts[ChangeType.MOVED] += 1
                continue

           
            before_marks.append(
                ElementMark(d_el, [(0, len(d_el.text))], ChangeType.MOVED)
            )
            after_marks.append(
                ElementMark(i_el, [(0, len(i_el.text))], ChangeType.MOVED)
            )
            counts[ChangeType.MOVED] += 1

        leftover_deletes = [d for d in deletes if d.index not in used_d]
        leftover_inserts = [i for i in inserts if i.index not in used_i]
        return leftover_deletes, leftover_inserts

    # ------------------------------------------------------------------ #
    # Internal: word / character level text diff -> character ranges
    # ------------------------------------------------------------------ #

    @staticmethod
    def _tokenize(text: str) -> List[str]:
        return _TOKEN_RE.findall(text)

    def diff_ranges(
        self, before: str, after: str
    ) -> Tuple[List[Tuple[int, int]], List[Tuple[int, int]]]:
        """Return the changed character ranges of a paired paragraph:
        (ranges in *before*, ranges in *after*).

        The comparison works at word granularity: only the whole words (or
        numbers/phrases) that actually changed are marked, never the enclosing
        line or paragraph. A change from "100" to "200" highlights exactly
        "100" on the BEFORE side and "200" on the AFTER side -- not just the
        differing digit and not the surrounding sentence.
        """
        before_tokens = self._tokenize(before)
        after_tokens = self._tokenize(after)

        matcher = difflib.SequenceMatcher(
            a=before_tokens, b=after_tokens, autojunk=False
        )

        b_ranges: List[Tuple[int, int]] = []
        a_ranges: List[Tuple[int, int]] = []
        b_pos = a_pos = 0

        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            b_chunk = "".join(before_tokens[i1:i2])
            a_chunk = "".join(after_tokens[j1:j2])

            if tag == "equal":
                pass
            elif tag == "delete":
                b_ranges.append((b_pos, b_pos + len(b_chunk)))
            elif tag == "insert":
                a_ranges.append((a_pos, a_pos + len(a_chunk)))
            elif tag == "replace":
                # Mark the entire changed word/phrase on each side (word-level,
                # per the spec) rather than drilling down to single characters.
                b_ranges.append((b_pos, b_pos + len(b_chunk)))
                a_ranges.append((a_pos, a_pos + len(a_chunk)))

            b_pos += len(b_chunk)
            a_pos += len(a_chunk)

        return self._merge_ranges(b_ranges), self._merge_ranges(a_ranges)

    @staticmethod
    def _merge_ranges(ranges: List[Tuple[int, int]]) -> List[Tuple[int, int]]:
        """Merge overlapping/adjacent ranges and drop empty ones."""
        cleaned = sorted((s, e) for s, e in ranges if e > s)
        merged: List[Tuple[int, int]] = []
        for s, e in cleaned:
            if merged and s <= merged[-1][1]:
                merged[-1] = (merged[-1][0], max(merged[-1][1], e))
            else:
                merged.append((s, e))
        return merged


# =========================================================================== #
#                                                                             #
#  inline_highlighter                                                         #
#  - applies change highlights directly inside copies of the ORIGINAL        #
#    documents, so every bit of formatting (fonts, styles, tables, lists,    #
#    numbering, images, headers/footers, TOC, bookmarks, margins, ...) is    #
#    preserved exactly as-is; only the changed characters get a shaded       #
#    background                                                              #
#                                                                             #
# =========================================================================== #

logger = get_logger("inline_highlighter")

# GitHub-diff-inspired shades applied to the exact changed tokens only.
BEFORE_FILL = "FFC1C9"   # red    : removed / original wording
AFTER_FILL = "ABF2BC"    # green  : added / updated wording
MOVED_FILL = "C7D2FE"    # indigo : unchanged content that changed POSITION


class InlineHighlighter:
    """Shades exact character ranges inside existing paragraphs, splitting
    runs at range boundaries while cloning each run's own formatting
    (``w:rPr``) so nothing about the visual appearance changes except the
    background of the changed characters."""

    def highlight_pnode(self, p_elem, ranges: List[Tuple[int, int]], fill: str) -> None:
        """Shade the given character ``ranges`` inside a single ``<w:p>``
        element (works equally on a live source paragraph or a deep-copied
        paragraph already inserted into the report)."""
        if not ranges:
            return
        try:
            self._highlight_paragraph(p_elem, ranges, fill)
        except Exception:
            # A single pathological paragraph must never abort the whole
            # report; log and continue.
            logger.exception("Failed to highlight a paragraph while building report")

    # ------------------------------------------------------------------ #
    # Internal
    # ------------------------------------------------------------------ #

    def _highlight_paragraph(
        self, p_elem, ranges: List[Tuple[int, int]], fill: str
    ) -> None:
        if not ranges:
            return
        # Snapshot the run list up front; splitting mutates the tree.
        runs = paragraph_runs(p_elem)
        pos = 0
        for r_elem in runs:
            text = run_text(r_elem)
            if not text:
                continue  # images, field codes, empty runs: never touched
            start, end = pos, pos + len(text)
            pos = end

            local = [
                (max(s, start) - start, min(e, end) - start)
                for s, e in ranges
                if s < end and e > start
            ]
            if not local:
                continue

            if local[0] == (0, len(text)) and len(local) == 1:
                self._shade_whole_run(r_elem, fill)
            elif self._has_complex_content(r_elem):
                # Runs carrying drawings/fields alongside text are shaded
                # whole rather than split, to avoid disturbing the embedded
                # object.
                self._shade_whole_run(r_elem, fill)
            else:
                self._split_and_shade(r_elem, text, local, fill)

    def _has_complex_content(self, r_elem) -> bool:
        tags = _tags()
        allowed = {tags["t"], tags["br"], tags["cr"], tags["tab"], tags["rPr"]}
        return any(child.tag not in allowed for child in r_elem)

    def _split_and_shade(
        self, r_elem, text: str, local_ranges: List[Tuple[int, int]], fill: str
    ) -> None:
        """Replace one run with several runs, each an exact formatting clone
        of the original, shading only the pieces inside ``local_ranges``."""
        tags = _tags()
        rpr = r_elem.find(tags["rPr"])

        points = sorted({0, len(text), *(p for rg in local_ranges for p in rg)})
        pieces: List[Tuple[str, bool]] = []
        for a, b in zip(points, points[1:]):
            if a == b:
                continue
            shaded = any(s <= a and b <= e for s, e in local_ranges)
            pieces.append((text[a:b], shaded))

        parent = r_elem.getparent()
        insert_at = list(parent).index(r_elem)
        for offset, (piece_text, shaded) in enumerate(pieces):
            parent.insert(
                insert_at + offset,
                self._build_run(rpr, piece_text, shaded, fill),
            )
        parent.remove(r_elem)

    def _build_run(self, rpr_source, text: str, shaded: bool, fill: str):
        """Create a new <w:r> carrying a deep copy of the source run's
        formatting, with \\n/\\t converted back into <w:br>/<w:tab>."""
        r = OxmlElement("w:r")
        rpr = deepcopy(rpr_source) if rpr_source is not None else None
        if shaded:
            if rpr is None:
                rpr = OxmlElement("w:rPr")
            self._set_rpr_shading(rpr, fill)
        if rpr is not None:
            r.append(rpr)

        for token in re.split(r"([\n\t])", text):
            if token == "":
                continue
            if token == "\n":
                r.append(OxmlElement("w:br"))
            elif token == "\t":
                r.append(OxmlElement("w:tab"))
            else:
                t = OxmlElement("w:t")
                t.text = token
                t.set(qn("xml:space"), "preserve")
                r.append(t)
        return r

    def _shade_whole_run(self, r_elem, fill: str) -> None:
        tags = _tags()
        rpr = r_elem.find(tags["rPr"])
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            r_elem.insert(0, rpr)
        self._set_rpr_shading(rpr, fill)

    @staticmethod
    def _set_rpr_shading(rpr, fill: str) -> None:
        for existing in rpr.findall(qn("w:shd")):
            rpr.remove(existing)
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)
        rpr.append(shd)


# =========================================================================== #
#                                                                             #
#  report_builder                                                             #
#  - splits each document into heading-delimited sections, pairs them         #
#    across the two versions, and assembles ONE report holding only the       #
#    sections that changed. Each changed section is emitted as a BEFORE       #
#    block (copied verbatim from the original) and an AFTER block (copied     #
#    verbatim from the modified document) with only the changed words         #
#    highlighted. Tables, images, lists, fonts, colours, spacing etc. are     #
#    preserved by deep-copying the section's Office Open XML and carrying      #
#    over the styles / numbering / images it references.                      #
#                                                                             #
# =========================================================================== #

logger = get_logger("report_builder")

# A relationship reference (r:embed / r:id / r:link) lives in this namespace.
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Banner + separator styling for the report scaffolding.
_BEFORE_BANNER_FILL = "F8D7DA"   # soft red
_AFTER_BANNER_FILL = "D4EDDA"    # soft green
_SECTION_TITLE_COLOR = RGBColor(0x1F, 0x24, 0x30)


def _split_marks_by_type(
    marks: "List[ElementMark]",
) -> "Tuple[Dict[int, List[Tuple[int, int]]], Dict[int, List[Tuple[int, int]]]]":
    """Split a list of ElementMark into two paragraph-id -> ranges maps:
    ``(normal_map, moved_map)``. Ranges that come from a MOVED mark are kept
    separate so the report can shade them with :data:`MOVED_FILL` instead of
    the ordinary BEFORE/AFTER red/green -- a paragraph that only changed
    POSITION must never look like it changed wording."""
    normal: "Dict[int, List[Tuple[int, int]]]" = {}
    moved: "Dict[int, List[Tuple[int, int]]]" = {}
    for mk in marks:
        pid = id(mk.element.paragraph._p)
        bucket = moved if mk.change_type == ChangeType.MOVED else normal
        bucket.setdefault(pid, []).extend(mk.ranges)
    return normal, moved


@dataclass
class Section:
    """A heading-delimited slice of one document: the heading paragraph plus
    every top-level block (paragraph or table) up to -- but not including --
    the next heading. Blocks are live references to the source XML so they
    can be deep-copied into the report with all formatting intact."""

    heading_text: str
    heading_level: Optional[int]
    blocks: List[object]                       # top-level <w:p> / <w:tbl>
    has_heading: bool = True                    # False for leading, pre-heading content
    elements: List[DocumentElement] = field(default_factory=list)

    @property
    def heading_block(self):
        """The heading's own ``<w:p>`` (the first block), or ``None`` for a
        leading section that has no heading of its own."""
        if self.has_heading and self.blocks:
            return self.blocks[0]
        return None

    @property
    def content_blocks(self) -> List[object]:
        """Every block of the section except its heading paragraph."""
        if self.has_heading and self.blocks:
            return self.blocks[1:]
        return list(self.blocks)


def _is_heading_pnode(p_elem, elem_by_pnode: Dict[int, DocumentElement]) -> bool:
    el = elem_by_pnode.get(id(p_elem))
    return bool(el and el.is_heading and not el.is_table_cell)


def split_sections(
    document: DocxDocument,
    start_pnode,
    elem_by_pnode: Dict[int, DocumentElement],
    end_pnode=None,
) -> List[Section]:
    """Split the document body into a flat list of heading-delimited
    :class:`Section` objects.

    The walk begins at ``start_pnode`` (``None`` = the very first block) and
    runs until ``end_pnode`` is reached, exclusive (``None`` = to the end of
    the document). ``end_pnode`` is what limits a "specific section" comparison
    to just that section's subtree instead of everything through to the end."""
    tags = _tags()
    p_tag, tbl_tag = tags["p"], tags["tbl"]

    sections: List[Section] = []
    current: Optional[Section] = None
    started = start_pnode is None  # None => start from the very first block

    for child in document.element.body.iterchildren():
        if child is start_pnode:
            started = True
        if not started:
            continue
        if end_pnode is not None and child is end_pnode:
            break  # reached the section's end boundary
        if child.tag not in (p_tag, tbl_tag):
            continue  # sectPr, bookmarks at body level, etc.

        if child.tag == p_tag and _is_heading_pnode(child, elem_by_pnode):
            el = elem_by_pnode.get(id(child))
            current = Section(
                heading_text=(el.text.strip() if el else ""),
                heading_level=(el.heading_level if el else None),
                blocks=[child],
                has_heading=True,
            )
            sections.append(current)
        else:
            if current is None:
                # Content before the first heading in range: keep it in an
                # untitled leading section so nothing is silently dropped.
                current = Section(
                    heading_text="", heading_level=None, blocks=[],
                    has_heading=False,
                )
                sections.append(current)
            current.blocks.append(child)

    # Attach the flattened DocumentElements belonging to each section (the
    # heading plus every paragraph inside its blocks, including table cells),
    # in document order, for the word-level diff.
    for sec in sections:
        for blk in sec.blocks:
            for pnode in blk.iter(p_tag):
                el = elem_by_pnode.get(id(pnode))
                if el is not None:
                    sec.elements.append(el)
    return sections


def _structural_signature(sec: Optional["Section"]) -> Tuple[int, int]:
    """Count embedded images and tables in a section, so a change that has no
    textual footprint (e.g. an image or a table added/removed with the
    surrounding wording untouched) is still recognised as a change."""
    tags = _tags()
    drawing, pict, tbl = qn("w:drawing"), qn("w:pict"), tags["tbl"]
    images = tables = 0
    if sec:
        for blk in sec.blocks:
            for el in blk.iter():
                if el.tag in (drawing, pict):
                    images += 1
                elif el.tag == tbl:
                    tables += 1
    return images, tables


def _eff_level(sec: "Section") -> int:
    """Effective outline level of a section for hierarchy purposes. Real
    Heading 1-9 map to 1-9; a Title/Subtitle-style heading (which python-docx
    reports with no numeric level) is treated as top level."""
    if sec.heading_level:
        return sec.heading_level
    return 1


def compute_ancestors(sections: List["Section"]) -> Dict[int, List["Section"]]:
    """For each section return the ordered chain of ancestor heading sections
    (outermost first, immediate parent last), derived from the heading levels
    in document order. Leading pre-heading content has no ancestors."""
    ancestors: Dict[int, List["Section"]] = {}
    stack: List[Tuple[int, "Section"]] = []
    for sec in sections:
        if not sec.has_heading:
            ancestors[id(sec)] = []
            continue
        level = _eff_level(sec)
        while stack and stack[-1][0] >= level:
            stack.pop()
        ancestors[id(sec)] = [s for _lvl, s in stack]
        stack.append((level, sec))
    return ancestors


def _sec_key(sec: "Section") -> Tuple[int, str]:
    """Identity of a heading for de-duplicating the ancestor chain across
    change entries (so siblings group under one shared parent heading). The
    full heading text -- which includes its section number -- keeps distinct
    headings distinct."""
    return (_eff_level(sec), (sec.heading_text or "").strip().lower())


def _norm_heading(text: str) -> str:
    """Normalise a heading for matching: drop a leading section number so
    "3.1 Capacity" and "4.1 Capacity" still pair up after renumbering."""
    stripped = _strip_section_number(text)
    return re.sub(r"\s+", " ", stripped).strip().lower()


def _section_blob(sec: Section) -> str:
    return sec.heading_text + " \u241f " + " ".join(e.text for e in sec.elements)


def _greedy_pair(
    orig: List[Section], mod: List[Section], threshold: float = 40.0
) -> List[Tuple[Optional[Section], Optional[Section]]]:
    """Greedy best-similarity pairing within a difflib 'replace' block,
    where headings genuinely differ and positional pairing is unreliable."""
    cands = []
    for oi, o in enumerate(orig):
        ob = _section_blob(o)
        for mi, m in enumerate(mod):
            cands.append((fuzz.ratio(ob, _section_blob(m)), oi, mi))
    cands.sort(key=lambda c: c[0], reverse=True)

    used_o: set = set()
    used_m: set = set()
    pairs: List[Tuple[Optional[Section], Optional[Section]]] = []
    for score, oi, mi in cands:
        if oi in used_o or mi in used_m or score < threshold:
            continue
        used_o.add(oi)
        used_m.add(mi)
        pairs.append((orig[oi], mod[mi]))
    for oi, o in enumerate(orig):
        if oi not in used_o:
            pairs.append((o, None))
    for mi, m in enumerate(mod):
        if mi not in used_m:
            pairs.append((None, m))
    return pairs


def pair_sections(
    orig_sections: List[Section], mod_sections: List[Section]
) -> List[Tuple[Optional[Section], Optional[Section]]]:
    """Align the two section lists in document order, returning
    ``(original, modified)`` tuples where either side may be ``None`` for a
    wholly added / removed section."""
    a = [_norm_heading(s.heading_text) for s in orig_sections]
    b = [_norm_heading(s.heading_text) for s in mod_sections]
    matcher = difflib.SequenceMatcher(a=a, b=b, autojunk=False)

    pairs: List[Tuple[Optional[Section], Optional[Section]]] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                pairs.append((orig_sections[i1 + k], mod_sections[j1 + k]))
        elif tag == "replace":
            pairs.extend(_greedy_pair(orig_sections[i1:i2], mod_sections[j1:j2]))
        elif tag == "delete":
            for k in range(i1, i2):
                pairs.append((orig_sections[k], None))
        elif tag == "insert":
            for k in range(j1, j2):
                pairs.append((None, mod_sections[k]))
    return pairs


def group_top_level(sections: List[Section]) -> List[Section]:
    """Merge the fine-grained (one-per-heading) sections into *top-level
    units*: each unit is the shallowest-level heading in scope together with
    its entire subtree (all deeper subsection headings and their content).

    This is what lets the report show one BEFORE / one AFTER per requested
    section, with every subsection nested inside, instead of a separate
    before/after per leaf."""
    levels = [(_eff_level(s)) for s in sections if s.has_heading]
    top = min(levels) if levels else 1

    units: List[Section] = []
    current: Optional[Section] = None
    for s in sections:
        is_top = s.has_heading and _eff_level(s) == top
        if is_top or current is None:
            current = Section(
                heading_text=s.heading_text,
                heading_level=s.heading_level,
                blocks=list(s.blocks),
                has_heading=s.has_heading,
                elements=list(s.elements),
            )
            units.append(current)
        else:
            current.blocks.extend(s.blocks)
            current.elements.extend(s.elements)
    return units


def group_at_level(sections: List[Section], level: int) -> List[Section]:
    """Merge fine-grained (one-per-heading) sections into units anchored at
    ``level``: each unit is a heading at that level together with every
    deeper subheading and its content, folded into one flat block list.

    ``group_top_level`` is the special case of this at the document's
    shallowest level; the same merge is reused one level down to regroup a
    top-level unit's own content back into its immediate child subsections.
    """
    units: List[Section] = []
    current: Optional[Section] = None
    for s in sections:
        is_target = s.has_heading and _eff_level(s) == level
        if is_target or current is None:
            current = Section(
                heading_text=s.heading_text,
                heading_level=s.heading_level,
                blocks=list(s.blocks),
                has_heading=s.has_heading,
                elements=list(s.elements),
            )
            units.append(current)
        else:
            current.blocks.extend(s.blocks)
            current.elements.extend(s.elements)
    return units


def _flat_subsections(sec: "Section") -> List["Section"]:
    """Re-split a grouped section's own content (excluding its own heading)
    back into one Section per heading paragraph found inside it, at
    whatever levels are present -- the same shape ``split_sections``
    produces, but without re-walking the document body.

    Every paragraph inside ``sec`` already has its DocumentElement captured
    in ``sec.elements`` (built once by ``split_sections``), so we key those
    by the raw ``<w:p>`` node to recognise heading blocks."""
    tags = _tags()
    p_tag = tags["p"]

    elem_by_pnode = {
        id(e.paragraph._p): e
        for e in sec.elements
        if getattr(e, "paragraph", None) is not None
    }

    flat: List[Section] = []
    current: Optional[Section] = None
    for blk in sec.content_blocks:
        is_heading = blk.tag == p_tag and _is_heading_pnode(blk, elem_by_pnode)
        if is_heading:
            el = elem_by_pnode[id(blk)]
            current = Section(
                heading_text=el.text.strip(),
                heading_level=el.heading_level,
                blocks=[blk],
                has_heading=True,
            )
            flat.append(current)
        else:
            if current is None:
                current = Section(
                    heading_text="", heading_level=None, blocks=[],
                    has_heading=False,
                )
                flat.append(current)
            current.blocks.append(blk)

    for sub in flat:
        for blk in sub.blocks:
            for pnode in blk.iter(p_tag):
                pe = elem_by_pnode.get(id(pnode))
                if pe is not None:
                    sub.elements.append(pe)
    return flat


def split_into_subsections(sec: Optional["Section"]) -> List["Section"]:
    """Split a grouped top-level Section back into its immediate child
    subsections (one level below whatever heading level appears first
    inside it) -- the inverse of the merge ``group_top_level`` did to
    build the unit in the first place.

    If ``sec`` has no internal heading structure at all, this still
    returns a single Section wrapping all of its content (so callers can
    treat "no subsections" and "one unchanged/changed subsection"
    uniformly). Returns ``[]`` only for ``None`` or an empty section."""
    if sec is None:
        return []
    flat = _flat_subsections(sec)
    if not flat:
        return []
    levels = [_eff_level(s) for s in flat if s.has_heading]
    target_level = min(levels) if levels else 0
    return group_at_level(flat, target_level)


def _section_text(sec: Optional["Section"]) -> str:
    """The comparable text of a section: every non-empty paragraph, in order,
    with whitespace normalised and -- for headings only -- the leading section
    number removed, so that a pure renumbering (chapter 5 becoming chapter 6
    because something was inserted earlier) is NOT mistaken for an edit."""
    if sec is None:
        return ""
    parts: List[str] = []
    for el in sec.elements:
        text = re.sub(r"\s+", " ", el.text).strip()
        if el.is_heading:
            text = _strip_section_number(text).strip()
        if text:
            parts.append(text)
    return "\n".join(parts)


def _section_changed(
    o_sec: Optional["Section"], m_sec: Optional["Section"]
) -> bool:
    """Whether a paired (original, modified) section really differs.

    The verdict is based on the section's *content*, not on diff bookkeeping:
    a section is changed only when it was added/removed, when its text differs
    (ignoring pure renumbering), or when it gained/lost an image or a table.
    Anything else is an untouched section and never reaches the report."""
    if o_sec is None or m_sec is None:
        return True
    if _structural_signature(o_sec) != _structural_signature(m_sec):
        return True
    return _section_text(o_sec) != _section_text(m_sec)


def _subsection_changed(
    o_sub: Optional["Section"], m_sub: Optional["Section"],
    before_map: Dict[int, object], after_map: Dict[int, object],
) -> bool:
    """Whether a paired (original, modified) subsection actually differs:
    added, removed, its text changed (renumbering aside), or a structural
    change (image / table) with no textual footprint.

    The content comparison is authoritative and the diff marks are only a
    secondary signal. Marks are produced by aligning paragraphs across the
    *whole* chapter, so a paragraph in an untouched 8.1 can pick up a mark
    merely because it fuzzy-matched something that moved into 8.3 -- which
    used to drag the untouched subsection into the report."""
    if not _section_changed(o_sub, m_sub):
        return False
    if o_sub is None or m_sub is None:
        return True

    tags = _tags()
    p_tag = tags["p"]
    for blk in o_sub.blocks:
        for pnode in blk.iter(p_tag):
            if id(pnode) in before_map:
                return True
    for blk in m_sub.blocks:
        for pnode in blk.iter(p_tag):
            if id(pnode) in after_map:
                return True

    return True


# --------------------------------------------------------------------------- #
# Cross-document resource import (styles / numbering / images)
# --------------------------------------------------------------------------- #


def import_missing_styles(report_doc: DocxDocument, source_doc: DocxDocument) -> None:
    """Copy every style defined in ``source_doc`` but absent from
    ``report_doc`` (matched by ``w:styleId``). Copying wholesale also brings
    across any ``basedOn`` / ``link`` / ``next`` dependencies for free."""
    try:
        rep_styles = report_doc.styles.element
        src_styles = source_doc.styles.element
    except Exception:
        return
    existing = {
        s.get(qn("w:styleId")) for s in rep_styles.findall(qn("w:style"))
    }
    for style in src_styles.findall(qn("w:style")):
        sid = style.get(qn("w:styleId"))
        if sid and sid not in existing:
            rep_styles.append(deepcopy(style))
            existing.add(sid)


class NumberingImporter:
    """Brings the original document's list-numbering definitions into the
    report and remaps the numId references on BEFORE (original) content to
    the imported ids.

    The whole numbering part is imported at once -- every ``<w:abstractNum>``
    and every ``<w:num>`` -- under fresh ids that cannot collide with the
    modified document's own definitions. Importing wholesale (rather than
    cloning only the single ``abstractNum`` a list points at) is what keeps
    multi-item lists numbering correctly instead of restarting at 1 on every
    line: Word list *styles* are stored as a pair of abstract definitions
    linked by ``styleLink`` / ``numStyleLink``, and the actual level
    definitions (start value, number format, indentation) live in the
    ``styleLink`` half. If only the ``numStyleLink`` half is copied, the copy
    has no levels and Word shows "1." for every item. Copying both halves
    keeps the link intact so the list continues 1, 2, 3, ..."""

    def __init__(self, report_doc: DocxDocument, source_doc: DocxDocument) -> None:
        self._rep_num = self._numbering_element(report_doc)
        self._src_num = self._numbering_element(source_doc)
        self._num_id_map: Dict[str, str] = {}
        self._imported = False

    @staticmethod
    def _numbering_element(document: DocxDocument):
        try:
            return document.part.numbering_part.element
        except Exception:
            return None

    def remap(self, block) -> None:
        """Rewrite the numId of every numbered paragraph in ``block`` to the
        imported definition (importing the whole part on first use)."""
        if self._src_num is None or self._rep_num is None:
            return
        self._ensure_imported()
        for numpr in block.iter(qn("w:numPr")):
            numid_el = numpr.find(qn("w:numId"))
            if numid_el is None:
                continue
            old = numid_el.get(qn("w:val"))
            new = self._num_id_map.get(old)
            if new is not None:
                numid_el.set(qn("w:val"), new)

    # -- internals ------------------------------------------------------- #

    def _ensure_imported(self) -> None:
        if self._imported:
            return
        self._imported = True
        abs_map = self._copy_abstract_nums()
        self._copy_nums(abs_map)

    def _copy_abstract_nums(self) -> Dict[str, str]:
        """Deep-copy every source ``<w:abstractNum>`` under a fresh id,
        returning ``{old_abstractNumId: new_abstractNumId}``. New definitions
        are inserted before the first ``<w:num>`` so the required ordering
        (all abstractNums before all nums) is preserved."""
        next_id = self._next_id("w:abstractNum", "w:abstractNumId", default=0)
        first_num = self._rep_num.find(qn("w:num"))
        abs_map: Dict[str, str] = {}
        for absn in self._src_num.findall(qn("w:abstractNum")):
            old = absn.get(qn("w:abstractNumId"))
            if old is None:
                continue
            new = str(next_id)
            next_id += 1
            abs_map[old] = new
            clone = deepcopy(absn)
            clone.set(qn("w:abstractNumId"), new)
            # A shared <w:nsid> makes Word treat two abstractNums as the same
            # list identity; drop it so the imported copy stays independent of
            # the modified document's definitions.
            for nsid in clone.findall(qn("w:nsid")):
                clone.remove(nsid)
            if first_num is not None:
                first_num.addprevious(clone)
            else:
                self._rep_num.append(clone)
        return abs_map

    def _copy_nums(self, abs_map: Dict[str, str]) -> None:
        """Deep-copy every source ``<w:num>`` under a fresh id, repointing its
        ``<w:abstractNumId>`` at the freshly imported abstractNums, and record
        the numId remapping used to rewrite BEFORE content."""
        next_id = self._next_id("w:num", "w:numId", default=1)
        for num in self._src_num.findall(qn("w:num")):
            old = num.get(qn("w:numId"))
            if old is None:
                continue
            new = str(next_id)
            next_id += 1
            self._num_id_map[old] = new
            clone = deepcopy(num)
            clone.set(qn("w:numId"), new)
            abs_ref = clone.find(qn("w:abstractNumId"))
            if abs_ref is not None:
                mapped = abs_map.get(abs_ref.get(qn("w:val")))
                if mapped is not None:
                    abs_ref.set(qn("w:val"), mapped)
            self._rep_num.append(clone)

    def _next_id(self, tag: str, attr: str, default: int) -> int:
        ids = []
        for el in self._rep_num.findall(qn(tag)):
            try:
                ids.append(int(el.get(qn(attr))))
            except (TypeError, ValueError):
                continue
        return (max(ids) + 1) if ids else default


class MediaImporter:
    """Copies images / external hyperlink targets referenced by BEFORE
    (original) content into the report and rewrites the relationship ids so
    the copied XML keeps pointing at the right resource."""

    def __init__(self, report_doc: DocxDocument, source_doc: DocxDocument) -> None:
        self._report = report_doc
        self._source = source_doc
        self._map: Dict[str, Optional[str]] = {}

    def import_refs(self, block) -> None:
        prefix = "{%s}" % _REL_NS
        for el in block.iter():
            for attr, val in list(el.attrib.items()):
                if attr.startswith(prefix) and val:
                    new_rid = self._map_rid(val)
                    if new_rid and new_rid != val:
                        el.set(attr, new_rid)

    def _map_rid(self, rid: str) -> Optional[str]:
        if rid in self._map:
            return self._map[rid]
        new_rid: Optional[str] = None
        try:
            rel = self._source.part.rels[rid]
            if rel.is_external:
                new_rid = self._report.part.relate_to(
                    rel.target_ref, rel.reltype, is_external=True
                )
            else:
                blob = rel.target_part.blob
                image = Image.from_blob(blob)
                partname = self._report.part.package.next_partname(
                    "/word/media/image%d." + image.ext
                )
                image_part = ImagePart.from_image(image, partname)
                new_rid = self._report.part.relate_to(image_part, RT.IMAGE)
        except Exception:
            logger.exception("Could not import relationship %s from original", rid)
            new_rid = None
        self._map[rid] = new_rid
        return new_rid


class HeadingNumberer:
    """Computes the *displayed* multilevel number of every heading paragraph
    in a document (e.g. "3", "3.1", "3.1.2", "4.2.1").

    Word headings are very often auto-numbered: the number a reader sees is
    generated by a multilevel list attached to the Heading styles and is NOT
    stored in the paragraph text. When such headings are copied into a report
    that deliberately omits their unchanged siblings, Word would recompute the
    numbers (giving wrong values) or show none at all. To keep the report
    faithful to the source, we compute each heading's true number here -- by
    replaying the list counters over the whole source document -- so the
    caller can freeze it as literal text on the copied heading.

    Headings that are numbered manually (the number already sits in the text)
    have no numbering definition to replay, so they yield ``None`` and are
    left exactly as-is (avoiding any double numbering)."""

    _ROMAN = [
        (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"),
        (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"),
        (5, "v"), (4, "iv"), (1, "i"),
    ]

    def __init__(self, document: DocxDocument) -> None:
        self._num_by_pid: Dict[int, str] = {}
        try:
            self._build(document)
        except Exception:
            logger.exception("Heading numbering computation failed; "
                             "numbers left to Word to render")

    def number_for(self, p_elem) -> Optional[str]:
        """Return the frozen number string for a heading ``<w:p>`` element, or
        ``None`` if the heading is not auto-numbered."""
        return self._num_by_pid.get(id(p_elem))

    # -- construction ---------------------------------------------------- #

    def _build(self, document: DocxDocument) -> None:
        numbering = self._numbering_element(document)
        abstract_by_id: Dict[str, object] = {}
        num_to_abs: Dict[str, str] = {}
        abs_to_num: Dict[str, str] = {}
        if numbering is not None:
            for a in numbering.findall(qn("w:abstractNum")):
                abstract_by_id[a.get(qn("w:abstractNumId"))] = a
            for n in numbering.findall(qn("w:num")):
                ar = n.find(qn("w:abstractNumId"))
                if ar is not None:
                    num_to_abs[n.get(qn("w:numId"))] = ar.get(qn("w:val"))
                    abs_to_num.setdefault(ar.get(qn("w:val")), n.get(qn("w:numId")))

        def resolve_abs(absel):
            """Follow numStyleLink -> styleLink so the level definitions are
            found even for style-linked lists."""
            seen = set()
            while absel is not None:
                nsl = absel.find(qn("w:numStyleLink"))
                if nsl is None:
                    return absel
                target = nsl.get(qn("w:val"))
                if target in seen:
                    return absel
                seen.add(target)
                nxt = None
                for a in abstract_by_id.values():
                    sl = a.find(qn("w:styleLink"))
                    if sl is not None and sl.get(qn("w:val")) == target:
                        nxt = a
                        break
                if nxt is None:
                    return absel
                absel = nxt
            return absel

        # Numbering attached to a paragraph *style* (styleId -> (numId, ilvl)).
        style_numpr: Dict[str, Tuple[str, int]] = {}
        styles_el = None
        try:
            styles_el = document.styles.element
        except Exception:
            styles_el = None
        if styles_el is not None:
            for st in styles_el.findall(qn("w:style")):
                sid = st.get(qn("w:styleId"))
                ppr = st.find(qn("w:pPr"))
                if ppr is None:
                    continue
                npr = ppr.find(qn("w:numPr"))
                if npr is None:
                    continue
                nid = npr.find(qn("w:numId"))
                il = npr.find(qn("w:ilvl"))
                if nid is not None and sid:
                    style_numpr[sid] = (
                        nid.get(qn("w:val")),
                        int(il.get(qn("w:val"))) if il is not None else 0,
                    )

        # Numbering that references heading styles from within the list levels
        # (lvl/pStyle) -> styleId -> (numId, ilvl).
        style_from_lvl: Dict[str, Tuple[str, int]] = {}
        for aid, a in abstract_by_id.items():
            ra = resolve_abs(a)
            for lvl in ra.findall(qn("w:lvl")):
                ps = lvl.find(qn("w:pStyle"))
                if ps is None:
                    continue
                nid = abs_to_num.get(aid)
                if nid is not None:
                    style_from_lvl.setdefault(
                        ps.get(qn("w:val")), (nid, int(lvl.get(qn("w:ilvl"))))
                    )

        def lvl_info(num_id: str, ilvl: int):
            aid = num_to_abs.get(num_id)
            if aid is None:
                return None
            a = abstract_by_id.get(aid)
            if a is None:
                return None
            for lvl in resolve_abs(a).findall(qn("w:lvl")):
                if int(lvl.get(qn("w:ilvl"))) == ilvl:
                    return lvl
            return None

        counters: Dict[str, Dict[int, int]] = {}

        for para in document.paragraphs:
            style = para.style
            style_name = style.name if style else "Normal"
            if not _HEADING_STYLE_RE.match(style_name or ""):
                continue
            style_id = getattr(style, "style_id", None)

            num_id: Optional[str] = None
            ilvl = 0
            ppr = para._p.find(qn("w:pPr"))
            if ppr is not None:
                npr = ppr.find(qn("w:numPr"))
                if npr is not None:
                    nid = npr.find(qn("w:numId"))
                    il = npr.find(qn("w:ilvl"))
                    if nid is not None and nid.get(qn("w:val")) is not None:
                        num_id = nid.get(qn("w:val"))
                        ilvl = int(il.get(qn("w:val"))) if il is not None else 0
            if num_id is None and style_id in style_numpr:
                num_id, ilvl = style_numpr[style_id]
            if num_id is None and style_id in style_from_lvl:
                num_id, ilvl = style_from_lvl[style_id]
            if not num_id or num_id == "0":
                continue  # manually numbered / unnumbered: leave text untouched

            info = lvl_info(num_id, ilvl)
            if info is None:
                continue

            key = num_to_abs.get(num_id, num_id)
            cs = counters.setdefault(key, {})

            start = 1
            si = info.find(qn("w:start"))
            if si is not None:
                try:
                    start = int(si.get(qn("w:val")))
                except (TypeError, ValueError):
                    start = 1
            cs[ilvl] = start if ilvl not in cs else cs[ilvl] + 1
            for deeper in [k for k in cs if k > ilvl]:
                del cs[deeper]

            number = self._render(num_id, ilvl, cs, lvl_info)
            if number:
                self._num_by_pid[id(para._p)] = number

    # -- rendering ------------------------------------------------------- #

    def _render(self, num_id, ilvl, counters, lvl_info_fn) -> str:
        info = lvl_info_fn(num_id, ilvl)
        if info is None:
            return ""
        lt = info.find(qn("w:lvlText"))
        text = lt.get(qn("w:val")) if lt is not None else None
        if not text:
            nf = info.find(qn("w:numFmt"))
            return self._fmt(
                counters.get(ilvl, 1),
                nf.get(qn("w:val")) if nf is not None else "decimal",
            )

        def repl(match):
            idx = int(match.group(1)) - 1
            li = lvl_info_fn(num_id, idx)
            nf = li.find(qn("w:numFmt")) if li is not None else None
            return self._fmt(
                counters.get(idx, 1),
                nf.get(qn("w:val")) if nf is not None else "decimal",
            )

        return re.sub(r"%(\d)", repl, text)

    def _fmt(self, n: int, numfmt: str) -> str:
        if numfmt == "upperRoman":
            return self._to_roman(n).upper()
        if numfmt == "lowerRoman":
            return self._to_roman(n)
        if numfmt == "upperLetter":
            return self._to_letter(n).upper()
        if numfmt == "lowerLetter":
            return self._to_letter(n)
        if numfmt == "decimalZero":
            return f"{n:02d}"
        if numfmt == "none":
            return ""
        return str(n)

    def _to_roman(self, n: int) -> str:
        out = []
        for value, sym in self._ROMAN:
            while n >= value:
                out.append(sym)
                n -= value
        return "".join(out)

    def _to_letter(self, n: int) -> str:
        out = ""
        while n > 0:
            n, rem = divmod(n - 1, 26)
            out = chr(97 + rem) + out
        return out

    @staticmethod
    def _numbering_element(document: DocxDocument):
        try:
            return document.part.numbering_part.element
        except Exception:
            return None


# --------------------------------------------------------------------------- #
# Appendix / annex detection  (issue 3: never compare or render these)
# --------------------------------------------------------------------------- #

_APPENDIX_RE = re.compile(r"\b(?:appendix|appendices|annexe?s?)\b", re.IGNORECASE)


def _is_appendix_heading(text: str) -> bool:
    """True for a heading that introduces an appendix / annex. Such a heading
    and everything beneath it is ignored entirely -- never compared, never
    rendered."""
    return bool(_APPENDIX_RE.search(_strip_section_number(text or "")))


def _eff_level_node(node: "Optional[_Node]") -> int:
    """Effective outline level of a tree node (Title/Subtitle -> top level)."""
    if node is not None and node.heading_level:
        return node.heading_level
    return 1


class _Node:
    """One heading in the section tree.

    Holds the heading paragraph, the heading's OWN direct content (the
    paragraphs / tables between this heading and its first sub-heading, i.e.
    NOT its descendants), and the child subsection nodes. Keeping own content
    and descendants separate is what lets the report emit only the changed
    leaf while showing its ancestor headings as breadcrumbs."""

    __slots__ = ("heading_text", "heading_level", "heading_block",
                 "own_blocks", "own_elements", "children")

    def __init__(self, heading_text, heading_level, heading_block,
                 own_blocks, own_elements, children=None):
        self.heading_text = heading_text
        self.heading_level = heading_level
        self.heading_block = heading_block          # source <w:p> (or None)
        self.own_blocks = own_blocks                # own <w:p>/<w:tbl> blocks
        self.own_elements = own_elements            # own DocumentElements
        self.children = children if children is not None else []


# --------------------------------------------------------------------------- #
# The report assembler
# --------------------------------------------------------------------------- #

class ReportBuilder:
    """Builds the single comparison report document."""

    # Fixed outline depths used by the multi-section ("Specific section(s)")
    # layout, so every generated report has the same shape no matter how deep
    # the requested sections sat in the source document.
    _TITLE_LEVEL = 1     # the report title ("Impact SDDD CSC ... CR UCAPSW-370")
    _SECTION_LEVEL = 2   # each requested section / requirement
    _BANNER_LEVEL = 3    # the Before / After labels

    def __init__(self, highlighter: "InlineHighlighter", temp_dir: "Path") -> None:
        self._hl = highlighter
        self._temp_dir = temp_dir

    def build(
        self,
        original_doc: "DocxDocument",
        modified_doc: "DocxDocument",
        orig_elems: "List[DocumentElement]",
        mod_elems: "List[DocumentElement]",
        orig_start_pnode,
        mod_start_pnode,
        engine: "DiffEngine",
        orig_end_pnode=None,
        mod_end_pnode=None,
        title: "Optional[str]" = None,
    ):
        """Return ``(report_document, change_counts, changed_section_count)``.

        For every changed top-level section the report shows, in order:

            <Section heading>          (original style + number, navigable)
                BEFORE                 (real heading, nested under the section)
                    <the whole original section: every subsection, nested>
                AFTER                  (real heading)
                    <the whole modified section: every subsection, nested>

        with only the changed words highlighted. Unchanged top-level sections
        never appear. ``orig_end_pnode`` / ``mod_end_pnode`` bound the walk to
        a single requested chapter (the "Requirements" chapter by default, or
        the section typed into the "specific section" box)."""
        elem_by_pnode_o = {id(e.paragraph._p): e for e in orig_elems}
        elem_by_pnode_m = {id(e.paragraph._p): e for e in mod_elems}

        orig_sections = group_top_level(split_sections(
            original_doc, orig_start_pnode, elem_by_pnode_o, orig_end_pnode
        ))
        mod_sections = group_top_level(split_sections(
            modified_doc, mod_start_pnode, elem_by_pnode_m, mod_end_pnode
        ))
        pairs = pair_sections(orig_sections, mod_sections)

        # Base the report on a clean copy of the MODIFIED document so every
        # style, numbering definition, theme font and media part the AFTER
        # content relies on is already present and correct.
        report = self._new_report_from(modified_doc)
        sectpr = report.element.body.find(qn("w:sectPr"))

        if title:
            self._add_document_title(report, sectpr, title)
        

        # BEFORE content comes from the original: make its styles available
        # up front (numbering / images are imported lazily per block).
        import_missing_styles(report, original_doc)
        numbering = NumberingImporter(report, original_doc)
        media = MediaImporter(report, original_doc)

        # Freeze auto-generated heading numbers ("3", "3.1.2", ...) computed
        # over each *whole* source document, so copied headings keep their real
        # numbers even though other sections are omitted.
        mod_numberer = HeadingNumberer(modified_doc)
        orig_numberer = HeadingNumberer(original_doc)

        counts = {t: 0 for t in ChangeType}
        changed_sections = 0

        for o_sec, m_sec in pairs:
            # Issue 3: appendices / annexes are never compared or rendered.
            _anchor_sec = m_sec if m_sec is not None else o_sec
            if _anchor_sec is not None and _is_appendix_heading(_anchor_sec.heading_text):
                continue
            # Cheap, authoritative check first: an untouched chapter is never
            # even diffed, let alone written to the report. A pure renumbering
            # (chapter 5 becoming chapter 6) is not a change.
            if not _section_changed(o_sec, m_sec):
                continue

            diff = engine.compare(
                o_sec.elements if o_sec else [],
                m_sec.elements if m_sec else [],
            )

            # _emit_unit filters again at subsection level and reports back
            # whether it actually wrote anything, so a chapter that has
            # nothing left to show never leaves a bare heading behind.
            if not self._emit_unit(
                report, sectpr, o_sec, m_sec, diff,
                numbering, media, orig_numberer, mod_numberer,
            ):
                continue

            for t in ChangeType:
                counts[t] += diff.counts[t]
            changed_sections += 1

        if changed_sections == 0:
            self._add_note(
                report, sectpr,
                "No differences were found in the compared section(s).",
            )

        return report, counts, changed_sections

    # ------------------------------------------------------------------ #
    # Multi-section mode: one BEFORE/AFTER pair per requested section     #
    # ------------------------------------------------------------------ #

    def build_sections(
        self,
        original_doc: "DocxDocument",
        modified_doc: "DocxDocument",
        orig_elems: "List[DocumentElement]",
        mod_elems: "List[DocumentElement]",
        scopes: "List[SectionScope]",
        engine: "DiffEngine",
        title: "Optional[str]" = None,
    ):
        """Build the report for an explicit list of requested sections.

        Unlike :meth:`build` -- which hunts for whatever changed inside one
        chapter and prunes everything untouched -- this walks the sections the
        user actually asked for, in the order they were typed, and emits each
        one in full:

            <report title>            Heading 1
                <section heading>     Heading 2
                    Before            Heading 3
                        <the whole original section>
                    After             Heading 3
                        <the whole modified section>

        Only the changed words are highlighted; a section missing from one side
        shows an explicit note there instead of blank space, and a section that
        did not change at all is written with a "no changes" note rather than
        two identical blocks.

        Returns ``(report_document, change_counts, emitted_section_count,
        unchanged_section_labels)``.
        """
        elem_by_pnode_o = {id(e.paragraph._p): e for e in orig_elems}
        elem_by_pnode_m = {id(e.paragraph._p): e for e in mod_elems}

        report = self._new_report_from(modified_doc)
        sectpr = report.element.body.find(qn("w:sectPr"))

        import_missing_styles(report, original_doc)
        numbering = NumberingImporter(report, original_doc)
        media = MediaImporter(report, original_doc)

        orig_numberer = HeadingNumberer(original_doc)
        mod_numberer = HeadingNumberer(modified_doc)

        # The report's own scaffolding levels must exist as real Heading styles
        # so Word AND Google Docs both nest the outline correctly.
        for lvl in range(self._TITLE_LEVEL, self._BANNER_LEVEL + 2):
            self._ensure_heading_style(report, lvl)

        if title:
            self._add_document_title(report, sectpr, title)
        

        counts = {t: 0 for t in ChangeType}
        emitted = 0
        unchanged: List[str] = []

        for scope in scopes:
            o_unit = self._unit_for(
                original_doc, scope.orig_start, scope.orig_end, elem_by_pnode_o
            )
            m_unit = self._unit_for(
                modified_doc, scope.mod_start, scope.mod_end, elem_by_pnode_m
            )

            diff = engine.compare(
                o_unit.elements if o_unit else [],
                m_unit.elements if m_unit else [],
            )
            for t in ChangeType:
                counts[t] += diff.counts[t]

            if not self._emit_section_pair(
                report, sectpr, scope, o_unit, m_unit, diff,
                numbering, media, orig_numberer, mod_numberer,
            ):
                unchanged.append(scope.label)
            emitted += 1

        if emitted == 0:
            self._add_note(
                report, sectpr,
                "None of the requested sections could be located.",
            )
        return report, counts, emitted, unchanged

    @staticmethod
    def _unit_for(document, start_pnode, end_pnode, elem_by_pnode) -> "Optional[Section]":
        """Collect one requested section (its heading plus everything under it,
        up to its end boundary) as a single :class:`Section`.

        ``None`` is returned when the section does not exist in this document
        (the caller passes a sentinel start node in that case, which matches no
        body child, so the walk yields nothing)."""
        if start_pnode is None:
            return None
        sections = split_sections(document, start_pnode, elem_by_pnode, end_pnode)
        if not sections:
            return None
        units = group_top_level(sections)
        if not units:
            return None
        unit = units[0]
        # Defensive: fold any sibling units back in so nothing inside the
        # requested range is silently dropped.
        for extra in units[1:]:
            unit.blocks.extend(extra.blocks)
            unit.elements.extend(extra.elements)
        return unit

    def _emit_section_pair(
        self, report, sectpr, scope, o_unit, m_unit, diff,
        numbering, media, orig_numberer, mod_numberer,
    ) -> bool:
        """Emit one requested section: its heading once, then Before, then
        After.

        Returns True when the section actually changed. An untouched section is
        still written to the report -- the user explicitly asked for it, so its
        absence would be ambiguous -- but instead of two identical Before/After
        blocks it gets a single explicit "no changes" note under its heading."""
        before_map, before_moved_map = _split_marks_by_type(diff.before_marks)
        after_map, after_moved_map = _split_marks_by_type(diff.after_marks)

        # --- the section's own heading, shown once above Before/After ------
        anchor, anchor_from_original = None, False
        if m_unit is not None and m_unit.has_heading and m_unit.heading_block is not None:
            anchor = m_unit
        elif o_unit is not None and o_unit.has_heading and o_unit.heading_block is not None:
            anchor, anchor_from_original = o_unit, True

        if anchor is not None:
            src_level = anchor.heading_level or 1
            self._copy_block(
                report, sectpr, anchor.heading_block, {}, AFTER_FILL,
                from_original=anchor_from_original,
                numbering=numbering, media=media,
                freeze_numberer=(orig_numberer if anchor_from_original else mod_numberer),
                heading_shift=self._SECTION_LEVEL - src_level,
            )
        else:
            # The target is a numbered paragraph / requirement id rather than a
            # real Word heading: synthesise a heading from what the user typed
            # so the report still has a navigable entry per section.
            self._add_plain_heading(report, sectpr, scope.label, self._SECTION_LEVEL)

        # --- untouched section: say so instead of two identical blocks -----
        if not _section_changed(o_unit, m_unit):
            self._add_note(
                report, sectpr,
                "No changes were made in this section.",
            )
            self._add_separator(report, sectpr)
            return False

        # --- re-level the section's inner headings under Before/After ------
        min_level = self._min_heading_level([o_unit, m_unit])
        content_shift = (self._BANNER_LEVEL + 1 - min_level) if min_level else 0

        # --- BEFORE ---------------------------------------------------------
        self._add_banner(report, sectpr, "Before", _BEFORE_BANNER_FILL,
                         heading_level=self._BANNER_LEVEL)
        if o_unit is not None and o_unit.content_blocks:
            for blk in o_unit.content_blocks:
                self._copy_block(
                    report, sectpr, blk, before_map, BEFORE_FILL,
                    from_original=True, numbering=numbering, media=media,
                    freeze_numberer=orig_numberer, heading_shift=content_shift,
                    moved_map=before_moved_map,
                )
        else:
            self._add_note(
                report, sectpr,
                "(this section does not exist in the original document)"
                if o_unit is None
                else "(this section is empty in the original document)",
            )

        # --- AFTER ----------------------------------------------------------
        self._add_banner(report, sectpr, "After", _AFTER_BANNER_FILL,
                         heading_level=self._BANNER_LEVEL)
        if m_unit is not None and m_unit.content_blocks:
            for blk in m_unit.content_blocks:
                self._copy_block(
                    report, sectpr, blk, after_map, AFTER_FILL,
                    from_original=False, numbering=numbering, media=media,
                    freeze_numberer=mod_numberer, heading_shift=content_shift,
                    moved_map=after_moved_map,
                )
        else:
            self._add_note(
                report, sectpr,
                "(this section does not exist in the modified document)"
                if m_unit is None
                else "(this section is empty in the modified document)",
            )

        self._add_separator(report, sectpr)
        return True

    @staticmethod
    def _min_heading_level(units) -> "Optional[int]":
        """Shallowest Heading level found in the *content* of the given units,
        used to re-level a section's subtree onto a fixed depth in the report
        regardless of how deep it sat in the source document."""
        levels = []
        p_tag = qn("w:p")
        for unit in units:
            if unit is None:
                continue
            for blk in unit.content_blocks:
                for pnode in blk.iter(p_tag):
                    ppr = pnode.find(qn("w:pPr"))
                    if ppr is None:
                        continue
                    ps = ppr.find(qn("w:pStyle"))
                    if ps is None:
                        continue
                    m = re.match(r"Heading([1-9])$", ps.get(qn("w:val")) or "")
                    if m:
                        levels.append(int(m.group(1)))
        return min(levels) if levels else None

    def _add_plain_heading(self, report, sectpr, text: str, level: int) -> None:
        """A real Heading-styled paragraph carrying literal text (used when the
        requested section is not a Word heading in the source)."""
        level = max(1, min(int(level), 9))
        self._ensure_heading_style(report, level)
        para = self._insert_paragraph(report, sectpr)
        self._apply_heading_style(para._p, level)
        self._set_outline_level(para._p, level - 1)
        self._suppress_numbering(para._p)
        para.paragraph_format.space_before = Pt(12)
        para.add_run(text)

    def _add_document_title(self, report, sectpr, title: str) -> None:
        """The report's own title line (Heading 1, centred) -- e.g.
        'Impact SDDD CSC WDOG CR UCAPSW-370'."""
        self._ensure_heading_style(report, self._TITLE_LEVEL)
        para = self._insert_paragraph(report, sectpr)
        self._apply_heading_style(para._p, self._TITLE_LEVEL)
        self._set_outline_level(para._p, self._TITLE_LEVEL - 1)
        self._suppress_numbering(para._p)
        ppr = self._get_or_make_ppr(para._p)
        jc = self._ppr_set(ppr, "jc")
        jc.attrib.clear()
        jc.set(qn("w:val"), "center")
        para.paragraph_format.space_after = Pt(18)
        run = para.add_run(title)
        run.bold = True
        try:
            report.core_properties.title = title
        except Exception:
            pass


     

    # ------------------------------------------------------------------ #
    # Section emission  (recursive: only changed leaves + their breadcrumb)
    # ------------------------------------------------------------------ #

    def _emit_unit(
        self, report, sectpr, o_sec, m_sec, diff,
        numbering: "NumberingImporter", media: "MediaImporter",
        orig_numberer: "HeadingNumberer", mod_numberer: "HeadingNumberer",
    ) -> bool:
        """Emit a changed unit (chapter). For every changed subsection the
        report shows that subsection's heading as a navigable header, then a
        BEFORE and an AFTER block. Each BEFORE/AFTER contains ONLY the
        breadcrumb path (ancestor headings) down to the actually changed leaf,
        plus that leaf's content -- unchanged sibling subsections never appear.
        Appendices / annexes are dropped entirely. Returns True iff anything
        was written."""
        # Per-unit memo caches (node identities are stable for this call).
        self._pairs_cache = {}
        self._changed_cache = {}

        before_map, before_moved_map = _split_marks_by_type(diff.before_marks)
        after_map, after_moved_map = _split_marks_by_type(diff.after_marks)

        o_root = self._build_root(o_sec)
        m_root = self._build_root(m_sec)

        wrote = False

        # (a) The unit's own leading content (prose right under the chapter
        #     heading, before its first subsection), only if it changed.
        if self._own_changed(o_root, m_root):
            anchor = m_root if m_sec is not None else o_root
            if anchor.heading_block is not None:
                self._copy_block(
                    report, sectpr, anchor.heading_block, {},
                    AFTER_FILL if m_sec is not None else BEFORE_FILL,
                    from_original=(m_sec is None),
                    numbering=numbering, media=media,
                    freeze_numberer=(mod_numberer if m_sec is not None else orig_numberer),
                )
            banner_level = min(_eff_level_node(anchor) + 1, 9)
            if o_root.own_blocks:
                self._add_banner(report, sectpr, "BEFORE", _BEFORE_BANNER_FILL,
                                 heading_level=banner_level)
                for blk in o_root.own_blocks:
                    self._copy_block(report, sectpr, blk, before_map, BEFORE_FILL,
                                     from_original=True, numbering=numbering, media=media,
                                     freeze_numberer=orig_numberer, demote=True,
                                     moved_map=before_moved_map)
            if m_root.own_blocks:
                self._add_banner(report, sectpr, "AFTER", _AFTER_BANNER_FILL,
                                 heading_level=banner_level)
                for blk in m_root.own_blocks:
                    self._copy_block(report, sectpr, blk, after_map, AFTER_FILL,
                                     from_original=False, numbering=numbering, media=media,
                                     freeze_numberer=mod_numberer, demote=True,
                                     moved_map=after_moved_map)
            self._add_separator(report, sectpr)
            wrote = True

        # (b) Each changed subsection, in document order.
        for o_child, m_child in self._children_pairs(o_root, m_root):
            if not self._node_changed(o_child, m_child):
                continue
            if self._emit_change_entry(
                report, sectpr, o_child, m_child, before_map, after_map,
                numbering, media, orig_numberer, mod_numberer,
                before_moved_map, after_moved_map,
            ):
                wrote = True

        return wrote

    def _emit_change_entry(
        self, report, sectpr, o_child, m_child, before_map, after_map,
        numbering, media, orig_numberer, mod_numberer,
        before_moved_map=None, after_moved_map=None,
    ) -> bool:
        """Emit one changed subsection: its heading as the section header, then
        BEFORE / AFTER each holding the breadcrumb-to-leaf content."""
        anchor = m_child if m_child is not None else o_child
        if anchor is None:
            return False

        before_blocks, after_blocks = self._collect_content(o_child, m_child)
        banner_level = min(_eff_level_node(anchor) + 1, 9)

        # Section header: the subsection's own heading, kept at its natural
        # level and shown once above BEFORE/AFTER (issue 1).
        if anchor.heading_block is not None:
            heading_changed = (
                o_child is not None and m_child is not None
                and _norm_heading(o_child.heading_text)
                != _norm_heading(m_child.heading_text)
            )
            head_side_map = {}
            if heading_changed:
                head_side_map = after_map if m_child is not None else before_map
            self._copy_block(
                report, sectpr, anchor.heading_block, head_side_map,
                AFTER_FILL if m_child is not None else BEFORE_FILL,
                from_original=(m_child is None),
                numbering=numbering, media=media,
                freeze_numberer=(mod_numberer if m_child is not None else orig_numberer),
            )

        if before_blocks:
            self._add_banner(report, sectpr, "BEFORE", _BEFORE_BANNER_FILL,
                             heading_level=banner_level)
            for blk in before_blocks:
                self._copy_block(report, sectpr, blk, before_map, BEFORE_FILL,
                                 from_original=True, numbering=numbering, media=media,
                                 freeze_numberer=orig_numberer, demote=True,
                                 moved_map=before_moved_map)
        if after_blocks:
            self._add_banner(report, sectpr, "AFTER", _AFTER_BANNER_FILL,
                             heading_level=banner_level)
            for blk in after_blocks:
                self._copy_block(report, sectpr, blk, after_map, AFTER_FILL,
                                 from_original=False, numbering=numbering, media=media,
                                 freeze_numberer=mod_numberer, demote=True,
                                 moved_map=after_moved_map)

        self._add_separator(report, sectpr)
        return True

    # ------------------------------------------------------------------ #
    # Section tree construction + change detection
    # ------------------------------------------------------------------ #

    def _build_root(self, sec: "Optional[Section]") -> "_Node":
        """Build the section tree for one grouped unit. The root's own content
        is the unit's leading (pre-subheading) prose; its children are the
        immediate subsections, recursively. Appendix / annex subsections and
        their whole subtree are dropped."""
        if sec is None:
            return _Node("", None, None, [], [], [])
        flat = _flat_subsections(sec)
        own_blocks, own_elems, children = self._assemble_nodes(flat)
        return _Node(sec.heading_text, sec.heading_level,
                     sec.heading_block, own_blocks, own_elems, children)

    @staticmethod
    def _assemble_nodes(flat: "List[Section]"):
        """Nest a flat, one-per-heading section list into a node forest,
        returning ``(root_own_blocks, root_own_elements, top_children)``.
        Appendix / annex headings (and everything beneath them) are skipped."""
        top_children: "List[_Node]" = []
        stack: "List[Tuple[int, _Node]]" = []
        root_own_blocks: "List[object]" = []
        root_own_elems: "List[DocumentElement]" = []
        appendix_level: "Optional[int]" = None

        for fs in flat:
            if not fs.has_heading:
                if appendix_level is None:
                    root_own_blocks.extend(fs.blocks)
                    root_own_elems.extend(fs.elements)
                continue

            lvl = _eff_level(fs)

            # Still inside an appendix subtree? skip until a sibling-or-higher
            # heading closes it.
            if appendix_level is not None:
                if lvl > appendix_level:
                    continue
                appendix_level = None

            if _is_appendix_heading(fs.heading_text):
                appendix_level = lvl
                continue

            node = _Node(fs.heading_text, fs.heading_level,
                         fs.blocks[0] if fs.blocks else None,
                         fs.content_blocks, fs.elements, [])
            while stack and stack[-1][0] >= lvl:
                stack.pop()
            if stack:
                stack[-1][1].children.append(node)
            else:
                top_children.append(node)
            stack.append((lvl, node))

        return root_own_blocks, root_own_elems, top_children

    def _node_section(self, node: "_Node") -> "Section":
        """A lightweight Section wrapping a node's OWN content only (its heading
        plus its direct blocks, excluding descendants) for text/structure
        comparison and heading-based pairing."""
        blocks = ([node.heading_block] if node.heading_block is not None else []) \
            + list(node.own_blocks)
        return Section(
            heading_text=node.heading_text,
            heading_level=node.heading_level,
            blocks=blocks,
            has_heading=node.heading_block is not None,
            elements=list(node.own_elements),
        )

    def _own_changed(self, o: "Optional[_Node]", m: "Optional[_Node]") -> bool:
        """Whether a node's OWN content differs (added/removed always counts).
        Descendants are NOT considered here."""
        if (o is None) != (m is None):
            return True
        if o is None:
            return False
        o_sec, m_sec = self._node_section(o), self._node_section(m)
        if _structural_signature(o_sec) != _structural_signature(m_sec):
            return True
        return _section_text(o_sec) != _section_text(m_sec)

    def _pair_child_nodes(self, o_children, m_children):
        """Align two node lists by heading (renumbering-tolerant), returning
        ``(o_node|None, m_node|None)`` pairs."""
        o_secs = [self._node_section(n) for n in o_children]
        m_secs = [self._node_section(n) for n in m_children]
        o_map = {id(s): n for s, n in zip(o_secs, o_children)}
        m_map = {id(s): n for s, n in zip(m_secs, m_children)}
        result = []
        for os_, ms_ in pair_sections(o_secs, m_secs):
            result.append((
                o_map.get(id(os_)) if os_ is not None else None,
                m_map.get(id(ms_)) if ms_ is not None else None,
            ))
        return result

    def _children_pairs(self, o: "Optional[_Node]", m: "Optional[_Node]"):
        """Cached child pairing for a node pair."""
        key = (id(o), id(m))
        cached = self._pairs_cache.get(key)
        if cached is not None:
            return cached
        pairs = self._pair_child_nodes(
            o.children if o is not None else [],
            m.children if m is not None else [],
        )
        self._pairs_cache[key] = pairs
        return pairs

    def _node_changed(self, o: "Optional[_Node]", m: "Optional[_Node]") -> bool:
        """Whether a node changed anywhere in its subtree (own content or any
        descendant). Cached."""
        key = (id(o), id(m))
        cached = self._changed_cache.get(key)
        if cached is not None:
            return cached
        result = self._own_changed(o, m) or any(
            self._node_changed(co, cm) for co, cm in self._children_pairs(o, m)
        )
        self._changed_cache[key] = result
        return result

    def _collect(self, o: "Optional[_Node]", m: "Optional[_Node]"):
        """Recursively gather ``(before_blocks, after_blocks)`` for a subsection.

        * own content changed        -> emit this heading AND its own content;
        * only a descendant changed   -> emit this heading ONLY (a breadcrumb);
        * nothing changed             -> emit nothing.

        This is what guarantees unchanged siblings never appear while the reader
        still sees the full path (issues 1 & 2)."""
        self_ch = self._own_changed(o, m)
        child_pairs = self._children_pairs(o, m)
        desc_ch = any(self._node_changed(co, cm) for co, cm in child_pairs)
        if not self_ch and not desc_ch:
            return [], []

        before, after = [], []
        if o is not None and o.heading_block is not None:
            before.append(o.heading_block)
        if m is not None and m.heading_block is not None:
            after.append(m.heading_block)
        if self_ch:
            if o is not None:
                before.extend(o.own_blocks)
            if m is not None:
                after.extend(m.own_blocks)
        for co, cm in child_pairs:
            cb, ca = self._collect(co, cm)
            before.extend(cb)
            after.extend(ca)
        return before, after

    def _collect_content(self, o: "Optional[_Node]", m: "Optional[_Node]"):
        """Like :meth:`_collect` but for a subsection whose heading is emitted
        separately as the section header: its own content (if changed) plus its
        descendants, WITHOUT re-emitting its own heading."""
        before, after = [], []
        if self._own_changed(o, m):
            if o is not None:
                before.extend(o.own_blocks)
            if m is not None:
                after.extend(m.own_blocks)
        for co, cm in self._children_pairs(o, m):
            cb, ca = self._collect(co, cm)
            before.extend(cb)
            after.extend(ca)
        return before, after

    # ------------------------------------------------------------------ #
    # Report scaffolding
    # ------------------------------------------------------------------ #

    def _new_report_from(self, modified_doc: "DocxDocument") -> "DocxDocument":
        """Save a clean copy of the modified document to a temp file, reopen
        it as the report, and strip its body (keeping the final sectPr so the
        page layout -- size, margins, orientation -- is preserved)."""
        seed = self._temp_dir / "report_seed.docx"
        modified_doc.save(str(seed))
        report = Document(str(seed))

        body = report.element.body
        for child in list(body.iterchildren()):
            if child.tag == qn("w:sectPr"):
                continue
            body.remove(child)

        report.core_properties.title = "Document Comparison Report"
        return report

    def _insert_paragraph(self, report: "DocxDocument", sectpr):
        """Append a fresh empty <w:p> just before the trailing sectPr and
        return it wrapped as a python-docx Paragraph."""
        p = OxmlElement("w:p")
        if sectpr is not None:
            sectpr.addprevious(p)
        else:
            report.element.body.append(p)
        return Paragraph(p, report._body)

    def _add_note(self, report, sectpr, text: str) -> None:
        """A plain (non-heading) informational line, e.g. the 'no differences'
        message. Deliberately not a Heading style so it never pollutes the
        Navigation Pane."""
        para = self._insert_paragraph(report, sectpr)
        para.paragraph_format.space_before = Pt(12)
        para.paragraph_format.space_after = Pt(6)
        run = para.add_run(text)
        run.bold = True
        run.font.size = Pt(13)
        run.font.color.rgb = _SECTION_TITLE_COLOR

    def _add_banner(self, report, sectpr, label: str, fill: str,
                    heading_level: "Optional[int]" = None) -> None:
        """Emit a BEFORE / AFTER label.

        The label is a REAL Heading-styled paragraph (at ``heading_level``),
        not merely a Normal paragraph with an outline level. Word keys its
        Navigation Pane off the outline level, but Google Docs only lists
        paragraphs whose actual paragraph *style* is a Heading -- so a real
        Heading style is what makes BEFORE/AFTER show up (and wrap their
        subsections) in both editors. The coloured shading and bold label are
        kept as direct formatting on top of the heading style."""
        para = self._insert_paragraph(report, sectpr)
        para.paragraph_format.space_before = Pt(8)
        para.paragraph_format.space_after = Pt(4)

        if heading_level is not None:
            level = max(1, min(int(heading_level), 9))
            self._ensure_heading_style(report, level)
            self._apply_heading_style(para._p, level)
            self._set_outline_level(para._p, level - 1)
            # A heading style may carry automatic list numbering; BEFORE/AFTER
            # are labels, so make sure they never pick up an auto number.
            self._suppress_numbering(para._p)

        self._shade_paragraph(para._p, fill)
        run = para.add_run(label)
        run.bold = True
        run.font.size = Pt(11)
        run.font.color.rgb = _SECTION_TITLE_COLOR  # keep it readable on the fill

    def _add_separator(self, report, sectpr) -> None:
        para = self._insert_paragraph(report, sectpr)
        para.paragraph_format.space_before = Pt(6)
        para.paragraph_format.space_after = Pt(6)

    def _copy_block(
        self, report, sectpr, block, marks_map, fill,
        from_original, numbering: "NumberingImporter", media: "MediaImporter",
        freeze_numberer: "Optional[HeadingNumberer]" = None,
        demote: bool = False,
        heading_shift: "Optional[int]" = None,
        moved_map: "Optional[Dict[int, List[Tuple[int, int]]]]" = None,
    ) -> None:
        """Deep-copy one top-level block (paragraph or table) into the report
        just before the trailing sectPr, highlighting the changed words and
        (for BEFORE content) importing the resources it references.

        ``marks_map``/``fill`` highlight ordinary added/deleted/modified text
        (red on the BEFORE side, green on the AFTER side). ``moved_map`` is a
        SEPARATE map of paragraphs whose content did not change but whose
        POSITION did -- e.g. a paragraph and a table swapped order. Those
        ranges are always shaded with :data:`MOVED_FILL` regardless of which
        side (before/after) they're on, so a pure reorder is never confused
        with an actual wording change and is never left uncoloured either.

        For every heading paragraph in the block: if ``freeze_numberer`` is
        given, its real (possibly auto-generated) number is frozen in as literal
        text so it survives out of context; its heading *style* is then shifted
        by ``heading_shift`` levels (``demote=True`` is the shorthand for a
        shift of +1) so it nests under the BEFORE/AFTER label in both Word and
        Google Docs. A negative shift promotes the heading, which is what
        re-levels a requested section's subtree onto a fixed depth."""
        new = deepcopy(block)

        shift = heading_shift if heading_shift is not None else (1 if demote else 0)

        p_tag = qn("w:p")
        src_ps = list(block.iter(p_tag))
        new_ps = list(new.iter(p_tag))
        for src_p, new_p in zip(src_ps, new_ps):
            ranges = marks_map.get(id(src_p))
            if ranges:
                self._hl.highlight_pnode(new_p, ranges, fill)
            moved_ranges = (moved_map or {}).get(id(src_p))
            if moved_ranges:
                full_text = paragraph_text(new_p)
                if full_text:
                    self._hl.highlight_pnode(new_p, [(0, len(full_text))], fill)
                self._append_moved_note(new_p)
            if self._is_heading_p(new_p):
                if freeze_numberer is not None:
                    number = freeze_numberer.number_for(src_p)
                    if number:
                        self._freeze_heading_number(new_p, number)
                if shift:
                    self._demote_heading_style(new_p, report, offset=shift)

        if from_original:
            try:
                numbering.remap(new)
            except Exception:
                logger.exception("Numbering remap failed for a BEFORE block")
            try:
                media.import_refs(new)
            except Exception:
                logger.exception("Media import failed for a BEFORE block")

        if sectpr is not None:
            sectpr.addprevious(new)
        else:
            report.element.body.append(new)

    @staticmethod
    def _append_moved_note(p_elem) -> None:
        """Append a small, muted note flagging a position-only change.

        Moved content now shares the same red/green fill as an ordinary
        edit (before=red, after=green), so this note is what keeps a reader
        from mistaking "this moved" for "this wording changed"."""
        r = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        rpr.append(OxmlElement("w:i"))
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), "16")  # 8pt
        rpr.append(sz)
        color = OxmlElement("w:color")
        color.set(qn("w:val"), "6B7280")
        rpr.append(color)
        r.append(rpr)
        t = OxmlElement("w:t")
        t.text = " (position changed)"
        t.set(qn("xml:space"), "preserve")
        r.append(t)
        p_elem.append(r)
    @staticmethod
    def _is_heading_p(p_elem) -> bool:
        ppr = p_elem.find(qn("w:pPr"))
        if ppr is None:
            return False
        ps = ppr.find(qn("w:pStyle"))
        if ps is not None and re.match(r"Heading[1-9]$", ps.get(qn("w:val")) or ""):
            return True
        return ppr.find(qn("w:outlineLvl")) is not None

    # ------------------------------------------------------------------ #
    # Heading demotion (by STYLE, so Google Docs nests correctly)
    # ------------------------------------------------------------------ #

    def _demote_heading_style(self, p_elem, report, offset: int = 1) -> None:
        """Shift a heading paragraph by ``offset`` levels *by changing its
        paragraph style* (e.g. Heading2 -> Heading3), not just its outline
        level. This is what makes the subsection nest under the BEFORE/AFTER
        heading in Google Docs, which builds its outline from paragraph styles
        and ignores the ``w:outlineLvl`` override that Word honours. A negative
        offset promotes instead of demoting.

        The heading's real (possibly auto-generated) number has already been
        frozen into the text by the caller, so any list numbering the target
        style might carry is suppressed to avoid a second, wrong number."""
        ppr = p_elem.find(qn("w:pPr"))
        ps = ppr.find(qn("w:pStyle")) if ppr is not None else None
        cur = None
        if ps is not None:
            m = re.match(r"Heading([1-9])$", ps.get(qn("w:val")) or "")
            if m:
                cur = int(m.group(1))
        if cur is None:
            # Not a heading-styled paragraph (only an outline level): fall back
            # to an outline bump so Word at least still nests it.
            self._bump_outline(p_elem)
            return
        new_level = max(1, min(cur + offset, 9))
        if new_level == cur:
            return
        self._ensure_heading_style(report, new_level)
        self._apply_heading_style(p_elem, new_level)
        self._set_outline_level(p_elem, new_level - 1)
        self._suppress_numbering(p_elem)

    def _bump_outline(self, p_elem) -> None:
        """Fallback used only when a paragraph has no recognisable Heading
        style: push its outline level one deeper so Word (which keys off the
        outline level) still nests it under the BEFORE/AFTER label."""
        ppr = p_elem.find(qn("w:pPr"))
        if ppr is None:
            return
        cur = None
        o = ppr.find(qn("w:outlineLvl"))
        if o is not None:
            try:
                cur = int(o.get(qn("w:val")))
            except (TypeError, ValueError):
                cur = None
        if cur is None:
            ps = ppr.find(qn("w:pStyle"))
            if ps is not None:
                m = re.match(r"Heading([1-9])$", ps.get(qn("w:val")) or "")
                if m:
                    cur = int(m.group(1)) - 1
        if cur is None:
            return
        self._set_outline_level(p_elem, min(cur + 1, 8))

    def _ensure_heading_style(self, report, level: int) -> None:
        """Guarantee the report defines a real ``Heading{level}`` paragraph
        style (correct styleId + built-in name), cloning the nearest shallower
        heading for its look. Google Docs maps a paragraph to an outline entry
        by this style, so it must actually exist for the demoted subsections
        and the BEFORE/AFTER labels to appear nested there."""
        level = max(1, min(int(level), 9))
        style_id = f"Heading{level}"
        try:
            styles_el = report.styles.element
        except Exception:
            return

        def find_style(sid):
            for st in styles_el.findall(qn("w:style")):
                if st.get(qn("w:styleId")) == sid:
                    return st
            return None

        if find_style(style_id) is not None:
            return

        template = None
        base_id = None
        for k in range(level - 1, 0, -1):
            template = find_style(f"Heading{k}")
            if template is not None:
                base_id = f"Heading{k}"
                break

        clone = deepcopy(template) if template is not None else OxmlElement("w:style")
        clone.set(qn("w:type"), "paragraph")
        clone.set(qn("w:styleId"), style_id)
        if clone.get(qn("w:default")) is not None:
            del clone.attrib[qn("w:default")]

        # Fresh built-in name ("heading N", matching Word's convention) so both
        # Word and Google Docs recognise the heading level.
        for nm in clone.findall(qn("w:name")):
            clone.remove(nm)
        name = OxmlElement("w:name")
        name.set(qn("w:val"), f"heading {level}")
        clone.insert(0, name)

        # Base it on the shallower heading; drop a shared linked char style so
        # two paragraph styles don't fight over the same link.
        for tag in ("w:basedOn", "w:link"):
            for e in clone.findall(qn(tag)):
                clone.remove(e)
        if base_id is not None:
            based = OxmlElement("w:basedOn")
            based.set(qn("w:val"), base_id)
            name.addnext(based)

        # Make sure it's visible (not latent/hidden) and unnumbered.
        for tag in ("w:semiHidden", "w:unhideWhenUsed"):
            for e in clone.findall(qn(tag)):
                clone.remove(e)
        ppr = clone.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            clone.append(ppr)
        for np in ppr.findall(qn("w:numPr")):
            ppr.remove(np)
        o = self._ppr_set(ppr, "outlineLvl")
        o.attrib.clear()
        o.set(qn("w:val"), str(level - 1))

        styles_el.append(clone)

    # ------------------------------------------------------------------ #
    # Frozen heading numbers
    # ------------------------------------------------------------------ #

    @staticmethod
    def _freeze_heading_number(p_elem, number: str) -> None:
        """Remove a heading paragraph's automatic list numbering and prepend
        the pre-computed number as literal text (number + space)."""
        ppr = p_elem.find(qn("w:pPr"))
        if ppr is not None:
            for numpr in ppr.findall(qn("w:numPr")):
                ppr.remove(numpr)
        run = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = f"{number} "
        t.set(qn("xml:space"), "preserve")
        run.append(t)
        if ppr is not None:
            ppr.addnext(run)
        else:
            p_elem.insert(0, run)

    # ------------------------------------------------------------------ #
    # Low-level pPr helpers (keep the OOXML schema order valid)
    # ------------------------------------------------------------------ #

    # CT_PPr child element order per the OOXML schema (the subset we touch is
    # what matters); new children are inserted so this order is never broken.
    _PPR_ORDER = (
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
        "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
        "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
        "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
        "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
        "suppressOverlap", "jc", "textDirection", "textAlignment",
        "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
        "pPrChange",
    )

    @staticmethod
    def _get_or_make_ppr(p_elem):
        ppr = p_elem.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            p_elem.insert(0, ppr)
        return ppr

    @classmethod
    def _ppr_set(cls, ppr, local_name: str):
        """Return the existing ``w:<local_name>`` child of ``ppr`` or create it
        at the schema-correct position so the pPr stays valid for Word."""
        existing = ppr.find(qn(f"w:{local_name}"))
        if existing is not None:
            return existing
        new = OxmlElement(f"w:{local_name}")
        try:
            my_idx = cls._PPR_ORDER.index(local_name)
        except ValueError:
            ppr.append(new)
            return new
        for child in ppr:
            child_local = child.tag.split("}")[-1]
            if child_local in cls._PPR_ORDER and cls._PPR_ORDER.index(child_local) > my_idx:
                child.addprevious(new)
                return new
        ppr.append(new)
        return new

    @classmethod
    def _apply_heading_style(cls, p_elem, level: int) -> None:
        ppr = cls._get_or_make_ppr(p_elem)
        ps = cls._ppr_set(ppr, "pStyle")
        ps.attrib.clear()
        ps.set(qn("w:val"), f"Heading{level}")

    @classmethod
    def _suppress_numbering(cls, p_elem) -> None:
        """Force-disable automatic list numbering on a paragraph (numId 0),
        overriding any numbering inherited from its (heading) style."""
        ppr = cls._get_or_make_ppr(p_elem)
        numpr = cls._ppr_set(ppr, "numPr")
        for child in list(numpr):
            numpr.remove(child)
        numid = OxmlElement("w:numId")
        numid.set(qn("w:val"), "0")
        numpr.append(numid)

    @classmethod
    def _set_outline_level(cls, p_elem, val: int) -> None:
        ppr = cls._get_or_make_ppr(p_elem)
        o = cls._ppr_set(ppr, "outlineLvl")
        o.attrib.clear()
        o.set(qn("w:val"), str(val))

    @classmethod
    def _shade_paragraph(cls, p_elem, fill: str) -> None:
        ppr = cls._get_or_make_ppr(p_elem)
        shd = cls._ppr_set(ppr, "shd")
        shd.attrib.clear()
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)


# =========================================================================== #
#                                                                             #
#  controller                                                                 #
#  - orchestrates parse -> scope -> compare -> single-report build -> save    #
#                                                                             #
# =========================================================================== #

logger = get_logger("controller")

ProgressCallback = Callable[[str, float], None]  # (message, fraction 0..1)

# Sentinel start node meaning "this section does not exist in this document":
# passed to split_sections it matches no body child, yielding an empty side so
# a section present in only one document still compares (as added / removed).
_ABSENT_SECTION = object()


class CompareScope(str, Enum):
    AUTO_REQUIREMENTS = "auto"   # from first Heading 1 containing "Requirements"
    FULL = "full"
    SECTION = "section"


@dataclass
class SectionScope:
    """One requested section, resolved to its body-node boundaries in each
    document. An absent side carries the ``_ABSENT_SECTION`` sentinel as its
    start node, which makes the section walk yield nothing there (so the
    section is reported as wholly added or removed)."""

    label: str
    orig_start: object = None
    orig_end: object = None
    mod_start: object = None
    mod_end: object = None
    found_in_original: bool = False
    found_in_modified: bool = False


@dataclass
class ComparisonRequest:
    original_path: Path
    modified_path: Path
    output_folder: Path
    scope: CompareScope = CompareScope.AUTO_REQUIREMENTS
    target_section: Optional[str] = None            # legacy single-section field
    target_sections: List[str] = field(default_factory=list)
    report_title: Optional[str] = None
    output_basename: str = "Comparison_Report"

    def __post_init__(self) -> None:
        # Accept either the legacy single section or the new list, and keep
        # both in sync so older callers keep working unchanged.
        if not self.target_sections and self.target_section:
            self.target_sections = [self.target_section]
        self.target_sections = [
            s.strip() for s in self.target_sections if s and s.strip()
        ]
        if self.target_sections and not self.target_section:
            self.target_section = self.target_sections[0]


@dataclass
class ComparisonResult:
    report_path: Path
    counts: Dict[ChangeType, int]
    changed_sections: int = 0
    scope: Optional["CompareScope"] = None
    target_section: Optional[str] = None
    target_sections: List[str] = field(default_factory=list)
    missing_sections: List[str] = field(default_factory=list)
    unchanged_sections: List[str] = field(default_factory=list)

    @property
    def total(self) -> int:
        return sum(self.counts.values())


class ComparisonController:
    """Coordinates the end-to-end document comparison workflow."""

    def __init__(self) -> None:
        self._temp_dir = Path(tempfile.mkdtemp(prefix="word_diff_"))
        self._parser = DocumentParser(temp_dir=self._temp_dir)
        self._engine = DiffEngine()
        self._highlighter = InlineHighlighter()
        self._report_builder = ReportBuilder(self._highlighter, self._temp_dir)

    def run(
        self,
        request: ComparisonRequest,
        progress: Optional[ProgressCallback] = None,
    ) -> ComparisonResult:
        """Execute the full compare-and-report pipeline.

        Produces ONE report document. In "Requirements" (auto) mode it contains
        only the sections that changed; in "specific section(s)" mode it
        contains each requested section, in the order requested, shown as a
        Before block (copied from the original) and an After block (copied from
        the modified document) -- with formatting preserved and only the changed
        words highlighted.
        """

        def report(msg: str, frac: float) -> None:
            logger.info(msg)
            if progress:
                progress(msg, frac)

        report("Validating input files...", 0.05)
        original_path = validate_input_file(request.original_path)
        modified_path = validate_input_file(request.modified_path)
        output_folder = validate_output_folder(request.output_folder)

        report("Opening original document...", 0.12)
        original_doc = self._parser.load(original_path)
        report("Opening modified document...", 0.20)
        modified_doc = self._parser.load(modified_path)

        report("Extracting content structure...", 0.30)
        original_elements = self._parser.extract_elements(original_doc)
        modified_elements = self._parser.extract_elements(modified_doc)

        missing_sections: List[str] = []
        unchanged_sections: List[str] = []

        if request.scope == CompareScope.SECTION and request.target_sections:
            # One Before/After pair per requested section, in the order typed.
            scopes, missing_sections = self._resolve_section_scopes(
                request, original_doc, modified_doc,
                original_elements, modified_elements, report,
            )
            report("Building comparison report (requested sections)...", 0.60)
            (
                report_doc, counts, changed_sections, unchanged_sections,
            ) = self._report_builder.build_sections(
                original_doc=original_doc,
                modified_doc=modified_doc,
                orig_elems=original_elements,
                mod_elems=modified_elements,
                scopes=scopes,
                engine=self._engine,
                title=request.report_title,
            )
        else:
            # Determine the comparison region and its start/end body nodes.
            (
                orig_start_pnode, orig_end_pnode,
                mod_start_pnode, mod_end_pnode,
            ) = self._resolve_scope(
                request, original_doc, modified_doc,
                original_elements, modified_elements, report,
            )

            report("Building comparison report (changed sections only)...", 0.60)
            report_doc, counts, changed_sections = self._report_builder.build(
                original_doc=original_doc,
                modified_doc=modified_doc,
                orig_elems=original_elements,
                mod_elems=modified_elements,
                orig_start_pnode=orig_start_pnode,
                mod_start_pnode=mod_start_pnode,
                engine=self._engine,
                orig_end_pnode=orig_end_pnode,
                mod_end_pnode=mod_end_pnode,
                title=request.report_title,
            )

        report("Saving comparison report...", 0.90)
        base = request.output_basename.strip() or "Comparison_Report"
        report_path = self._ensure_unique_path(output_folder / f"{base}.docx")
        self._save(report_doc, report_path)

        report("Done.", 1.0)
        return ComparisonResult(
            report_path=report_path,
            counts=counts,
            changed_sections=changed_sections,
            scope=request.scope,
            target_section=request.target_section,
            target_sections=list(request.target_sections),
            missing_sections=missing_sections,
            unchanged_sections=unchanged_sections,
        )

    def _resolve_section_scopes(
        self,
        request: ComparisonRequest,
        original_doc: DocxDocument,
        modified_doc: DocxDocument,
        original_elements: List[DocumentElement],
        modified_elements: List[DocumentElement],
        report: Callable[[str, float], None],
    ) -> Tuple[List[SectionScope], List[str]]:
        """Resolve every requested section to its boundaries in both documents.

        Returns ``(scopes, missing)``: the sections that were located (in the
        order the user typed them) and the ones found in neither document. A
        section present in only one document is kept -- it is reported as wholly
        added or removed rather than aborting the whole comparison."""
        # Heading numbers are computed once per document so an auto-numbered
        # heading can be matched by the number the reader sees ("3.1.2").
        orig_numberer = HeadingNumberer(original_doc)
        mod_numberer = HeadingNumberer(modified_doc)

        scopes: List[SectionScope] = []
        missing: List[str] = []
        total = max(len(request.target_sections), 1)

        for i, target in enumerate(request.target_sections, start=1):
            report(f"Locating section '{target}' ({i}/{total})...",
                   0.35 + 0.20 * (i / total))

            o_bounds = self._parser.try_find_section_bounds(
                original_elements, target, orig_numberer
            )
            m_bounds = self._parser.try_find_section_bounds(
                modified_elements, target, mod_numberer
            )
            if o_bounds is None and m_bounds is None:
                logger.warning("Section not found in either document: %s", target)
                missing.append(target)
                continue

            o_start, o_end = self._section_pnodes(original_elements, o_bounds)
            m_start, m_end = self._section_pnodes(modified_elements, m_bounds)
            scopes.append(
                SectionScope(
                    label=target,
                    orig_start=o_start, orig_end=o_end,
                    mod_start=m_start, mod_end=m_end,
                    found_in_original=o_bounds is not None,
                    found_in_modified=m_bounds is not None,
                )
            )

        if not scopes:
            available = ", ".join(
                self._parser.available_section_labels(modified_elements)[:40]
            ) or ", ".join(
                self._parser.available_section_labels(original_elements)[:40]
            )
            raise SectionNotFoundError(
                "None of the requested sections could be found in either "
                f"document: {', '.join(missing)}.\n\n"
                f"Sections detected: {available or '(none)'}."
            )
        return scopes, missing

    def _resolve_scope(
        self,
        request: ComparisonRequest,
        original_doc: DocxDocument,
        modified_doc: DocxDocument,
        original_elements: List[DocumentElement],
        modified_elements: List[DocumentElement],
        report: Callable[[str, float], None],
    ):
        """Return ``(orig_start, orig_end, mod_start, mod_end)`` body nodes
        delimiting the comparison region in each document. A ``start`` of
        ``None`` means "from the very beginning"; an ``end`` of ``None`` means
        "through to the end of the document"."""
        if request.scope == CompareScope.AUTO_REQUIREMENTS:
            report("Locating 'Requirements' section (Heading 1)...", 0.45)
            o_bounds = self._parser.find_requirements_bounds(original_elements)
            m_bounds = self._parser.find_requirements_bounds(modified_elements)

            if o_bounds is None and m_bounds is None:
                raise SectionNotFoundError(
                    "Could not find a Heading 1 containing the word "
                    "'Requirements' in either document. Check that the "
                    "documents use the built-in 'Heading 1' style for their "
                    "chapter titles, or type the exact section you want into "
                    "the 'Specific section' box."
                )

            # The end bound is what confines the comparison to the
            # Requirements chapter: anything after it (appendices, annexes,
            # revision history, ...) is never looked at, let alone reported.
            orig_start, orig_end = self._section_pnodes(original_elements, o_bounds)
            mod_start, mod_end = self._section_pnodes(modified_elements, m_bounds)
            return orig_start, orig_end, mod_start, mod_end

        if request.scope == CompareScope.SECTION and request.target_section:
            report(f"Isolating section '{request.target_section}'...", 0.45)
            # Compute heading numbers so an auto-numbered heading can be found
            # by the number the reader sees (e.g. "3.1.2"), not just its text.
            orig_numberer = HeadingNumberer(original_doc)
            mod_numberer = HeadingNumberer(modified_doc)

            o_bounds = self._parser.try_find_section_bounds(
                original_elements, request.target_section, orig_numberer
            )
            m_bounds = self._parser.try_find_section_bounds(
                modified_elements, request.target_section, mod_numberer
            )

            # The section must exist in at least one document; if it exists in
            # only one, the other side is treated as empty so the section shows
            # up as wholly added or removed instead of aborting the comparison.
            if o_bounds is None and m_bounds is None:
                available = ", ".join(
                    self._parser.available_section_labels(modified_elements)[:40]
                ) or ", ".join(
                    self._parser.available_section_labels(original_elements)[:40]
                )
                raise SectionNotFoundError(
                    f"Could not find a section matching "
                    f"'{request.target_section}' in either document.\n\n"
                    f"Sections detected: {available or '(none)'}."
                )

            orig_start, orig_end = self._section_pnodes(original_elements, o_bounds)
            mod_start, mod_end = self._section_pnodes(modified_elements, m_bounds)
            return orig_start, orig_end, mod_start, mod_end

        # Whole document.
        return None, None, None, None

    @staticmethod
    def _section_pnodes(elements, bounds):
        """Translate ``(start_i, end_i)`` bounds into ``(start_pnode,
        end_pnode)`` body nodes. When ``bounds`` is ``None`` (the section is
        absent from this document) a sentinel start node is returned that makes
        the section walk yield nothing, so the section is compared against an
        empty side (wholly added / removed)."""
        if bounds is None:
            return _ABSENT_SECTION, None
        start_i, end_i = bounds
        start = elements[start_i].paragraph._p
        end = elements[end_i].paragraph._p if end_i < len(elements) else None
        return start, end

    @staticmethod
    def _save(doc: DocxDocument, path: Path) -> None:
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            doc.save(str(path))
        except PermissionError as exc:
            raise PermissionDeniedError(
                f"Cannot write '{path}'. Close the file if it is open "
                "elsewhere, or choose a different output folder."
            ) from exc
        logger.info("Saved %s", path)

    @staticmethod
    def _ensure_unique_path(path: Path) -> Path:
        """Avoid clobbering an existing file by appending a counter."""
        if not path.exists():
            return path
        stem, suffix = path.stem, path.suffix
        counter = 1
        while True:
            candidate = path.with_name(f"{stem} ({counter}){suffix}")
            if not candidate.exists():
                return candidate
            counter += 1

    def cleanup(self) -> None:
        """Remove any temporary files created during conversion."""
        shutil.rmtree(self._temp_dir, ignore_errors=True)


# =========================================================================== #
#                                                                             #
#  gui                                                                        #
#  - Tkinter desktop interface                                                #
#                                                                             #
# =========================================================================== #

logger = get_logger("gui")

_FILETYPES = [
    ("Word Documents", "*.doc *.docx *.docm"),
    ("Word Macro-Enabled (*.docm)", "*.docm"),
    ("Word 2007+ (*.docx)", "*.docx"),
    ("Legacy Word (*.doc)", "*.doc"),
    ("All files", "*.*"),
]

_BG = "#F5F6F8"
_CARD_BG = "#FFFFFF"
_ACCENT = "#2563EB"
_ACCENT_HOVER = "#1D4ED8"
_TEXT = "#1F2430"
_MUTED = "#6B7280"
_BORDER = "#E2E5EA"


class WordDiffApp:
    """Main application window."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Word Document Change Review")
        self.root.geometry("760x780")
        self.root.minsize(620, 420)
        self.root.configure(bg=_BG)

        self.original_path = tk.StringVar()
        self.modified_path = tk.StringVar()
        self.output_folder = tk.StringVar(value=str(Path.home() / "Documents"))
        self.output_basename = tk.StringVar(value="Comparison_Report")
        self.compare_mode = tk.StringVar(value=CompareScope.AUTO_REQUIREMENTS.value)
        self.report_title = tk.StringVar()
        self.status_text = tk.StringVar(value="Ready.")
        self._section_rows: List[Dict[str, object]] = []

        self._controller = ComparisonController()
        self._worker: Optional[threading.Thread] = None

        self._build_style()
        self._build_layout()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------ #
    # Style / layout
    # ------------------------------------------------------------------ #

    def _build_style(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("TFrame", background=_BG)
        style.configure("Card.TFrame", background=_CARD_BG)
        style.configure("TLabel", background=_BG, foreground=_TEXT, font=("Segoe UI", 10))
        style.configure("Card.TLabel", background=_CARD_BG, foreground=_TEXT, font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background=_CARD_BG, foreground=_MUTED, font=("Segoe UI", 9))
        style.configure("Title.TLabel", background=_BG, foreground=_TEXT, font=("Segoe UI Semibold", 18))
        style.configure("Subtitle.TLabel", background=_BG, foreground=_MUTED, font=("Segoe UI", 10))
        style.configure("Section.TLabel", background=_CARD_BG, foreground=_TEXT, font=("Segoe UI Semibold", 11))

        style.configure(
            "Accent.TButton",
            background=_ACCENT,
            foreground="white",
            font=("Segoe UI Semibold", 11),
            padding=(16, 10),
            borderwidth=0,
        )
        style.map("Accent.TButton", background=[("active", _ACCENT_HOVER), ("disabled", "#93A6C9")])

        style.configure(
            "Secondary.TButton",
            background="#EEF1F6",
            foreground=_TEXT,
            font=("Segoe UI", 10),
            padding=(10, 6),
            borderwidth=0,
        )
        style.map("Secondary.TButton", background=[("active", "#E2E6ED")])

        style.configure("TEntry", padding=6)
        style.configure("TRadiobutton", background=_CARD_BG, foreground=_TEXT, font=("Segoe UI", 10))
        style.configure("Horizontal.TProgressbar", background=_ACCENT, troughcolor="#EEF1F6", thickness=10)

    def _build_scroll_container(self) -> "ttk.Frame":
        """Put the whole form on a scrollable canvas.

        The form is taller than a short screen (and taller than the window once
        several section fields are added), so without this the Compare button
        and the Output fields below the fold are simply unreachable. Returns the
        frame every widget should be built inside."""
        container = tk.Frame(self.root, bg=_BG)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, bg=_BG, highlightthickness=0, bd=0)
        vbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        vbar.pack(side="right", fill="y")

        body = ttk.Frame(canvas)
        window = canvas.create_window((0, 0), window=body, anchor="nw")

        def on_body_configure(_event=None) -> None:
            # The form grew or shrank (e.g. a section field was added/removed).
            canvas.configure(scrollregion=canvas.bbox("all"))

        def on_canvas_configure(event) -> None:
            canvas.itemconfigure(window, width=event.width)

        body.bind("<Configure>", on_body_configure)
        canvas.bind("<Configure>", on_canvas_configure)

        # Mouse wheel: Windows/macOS deliver <MouseWheel>, X11 Button-4/5.
        def on_wheel(event) -> None:
            if getattr(event, "num", None) == 4:
                step = -3
            elif getattr(event, "num", None) == 5:
                step = 3
            else:
                step = -3 if event.delta > 0 else 3
            canvas.yview_scroll(step, "units")

        def bind_wheel(_event=None) -> None:
            canvas.bind_all("<MouseWheel>", on_wheel)
            canvas.bind_all("<Button-4>", on_wheel)
            canvas.bind_all("<Button-5>", on_wheel)

        def unbind_wheel(_event=None) -> None:
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        canvas.bind("<Enter>", bind_wheel)
        canvas.bind("<Leave>", unbind_wheel)

        self._canvas = canvas
        self._scroll_body = body
        return body

    def _build_layout(self) -> None:
        outer = ttk.Frame(self._build_scroll_container(), padding=24)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Word Document Change Review", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text="Outputs ONE report with a Before/After block per section, changed words highlighted.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(0, 12))

        card = tk.Frame(outer, bg=_CARD_BG, highlightbackground=_BORDER, highlightthickness=1)
        card.pack(fill="both", expand=True)
        inner = ttk.Frame(card, padding=20, style="Card.TFrame")
        inner.pack(fill="both", expand=True)

        self._build_file_picker(inner, "Original Document", self.original_path, row=0)
        self._build_file_picker(inner, "Modified Document", self.modified_path, row=1)

        # -- comparison scope --------------------------------------------- #
        ttk.Label(inner, text="Comparison scope", style="Section.TLabel").grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(14, 0)
        )
        scope_frame = ttk.Frame(inner, style="Card.TFrame")
        scope_frame.grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Radiobutton(
            scope_frame,
            text='From "Requirements" section (auto)',
            value=CompareScope.AUTO_REQUIREMENTS.value,
            variable=self.compare_mode,
            command=self._on_scope_change,
        ).pack(side="left", padx=(0, 14))
        ttk.Radiobutton(
            scope_frame,
            text="Specific section(s)",
            value=CompareScope.SECTION.value,
            variable=self.compare_mode,
            command=self._on_scope_change,
        ).pack(side="left")

        # -- one input field per requested section ------------------------ #
        self.sections_frame = ttk.Frame(inner, style="Card.TFrame")
        self.sections_frame.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        self.sections_frame.columnconfigure(0, weight=1)

        add_frame = ttk.Frame(inner, style="Card.TFrame")
        add_frame.grid(row=7, column=0, columnspan=3, sticky="w", pady=(2, 0))
        self.add_section_button = ttk.Button(
            add_frame, text="+ Add another section", style="Secondary.TButton",
            command=self._add_section_row,
        )
        self.add_section_button.pack(side="left")

        ttk.Label(
            inner,
            text='Auto mode compares ONLY the chapter whose Heading 1 contains "Requirements". '
                 'In "Specific section(s)" mode, type one section per field -- a heading '
                 '("3.2 Watchdog"), a section number ("3.2.1") or a requirement id '
                 '("REQ-SDDD_CLSW_LOADSW_LUP-0087"). The report shows them in the order '
                 'entered, each with its own Before / After.',
            style="Muted.TLabel",
            wraplength=640,
        ).grid(row=8, column=0, columnspan=3, sticky="w", pady=(6, 0))

        ttk.Separator(inner, orient="horizontal").grid(
            row=9, column=0, columnspan=3, sticky="ew", pady=14
        )

        # -- output -------------------------------------------------------- #
        ttk.Label(inner, text="Output", style="Section.TLabel").grid(
            row=10, column=0, columnspan=3, sticky="w"
        )

        title_frame = ttk.Frame(inner, style="Card.TFrame")
        title_frame.grid(row=11, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        title_frame.columnconfigure(1, weight=1)
        ttk.Label(title_frame, text="Report title:", style="Card.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        self.title_entry = ttk.Entry(title_frame, textvariable=self.report_title)
        self.title_entry.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        ttk.Label(
            inner,
            text='Printed as the title on the report\'s first page, e.g. '
                 '"Impact SDDD CSC WDOG CR UCAPSW-370". This is only the heading '
                 'inside the document -- the saved file is named from the base '
                 'file name below.',
            style="Muted.TLabel",
            wraplength=640,
        ).grid(row=12, column=0, columnspan=3, sticky="w", pady=(2, 0))

        out_frame = ttk.Frame(inner, style="Card.TFrame")
        out_frame.grid(row=13, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        out_frame.columnconfigure(0, weight=1)
        ttk.Entry(out_frame, textvariable=self.output_folder).grid(row=0, column=0, sticky="ew")
        ttk.Button(
            out_frame, text="Browse...", style="Secondary.TButton",
            command=self._pick_output_folder,
        ).grid(row=0, column=1, padx=(8, 0))

        name_frame = ttk.Frame(inner, style="Card.TFrame")
        name_frame.grid(row=14, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        ttk.Label(name_frame, text="Base file name:", style="Card.TLabel").pack(side="left")
        ttk.Entry(name_frame, textvariable=self.output_basename, width=32).pack(
            side="left", padx=(8, 0)
        )
        ttk.Label(
            name_frame, text="-> <name>.docx (single comparison report)", style="Muted.TLabel"
        ).pack(side="left", padx=(8, 0))

        inner.columnconfigure(1, weight=1)

        # -- action row ---------------------------------------------------- #
        action_frame = ttk.Frame(outer)
        action_frame.pack(fill="x", pady=(18, 0))
        self.compare_button = ttk.Button(
            action_frame, text="Compare Documents", style="Accent.TButton",
            command=self._on_compare_clicked,
        )
        self.compare_button.pack(side="left")

        self.progress = ttk.Progressbar(
            outer, mode="determinate", maximum=1.0, style="Horizontal.TProgressbar"
        )
        self.progress.pack(fill="x", pady=(16, 6))

        self.status_label = ttk.Label(outer, textvariable=self.status_text, style="Subtitle.TLabel")
        self.status_label.pack(anchor="w")

        # Start with a single (initially disabled) section field.
        self._add_section_row()
        self._on_scope_change()

    def _build_file_picker(self, parent: ttk.Frame, label: str, var: tk.StringVar, row: int) -> None:
        ttk.Label(parent, text=label, style="Card.TLabel").grid(
            row=row * 2, column=0, sticky="w", pady=(0 if row == 0 else 10, 2)
        )
        entry = ttk.Entry(parent, textvariable=var)
        entry.grid(row=row * 2 + 1, column=0, columnspan=2, sticky="ew", padx=(0, 8))
        ttk.Button(
            parent, text="Select...", style="Secondary.TButton",
            command=lambda: self._pick_file(var),
        ).grid(row=row * 2 + 1, column=2, sticky="e")
        parent.columnconfigure(0, weight=1)

    # ------------------------------------------------------------------ #
    # Dynamic "specific section" fields
    # ------------------------------------------------------------------ #

    def _add_section_row(self, value: str = "") -> None:
        """Append one more section input field."""
        row = ttk.Frame(self.sections_frame, style="Card.TFrame")
        row.pack(fill="x", pady=(0, 4))
        row.columnconfigure(1, weight=1)

        index_label = ttk.Label(row, text="", style="Muted.TLabel", width=3)
        index_label.grid(row=0, column=0, sticky="w")

        var = tk.StringVar(value=value)
        entry = ttk.Entry(row, textvariable=var)
        entry.grid(row=0, column=1, sticky="ew")

        remove = ttk.Button(
            row, text="X", width=3, style="Secondary.TButton",
            command=lambda: self._remove_section_row(row),
        )
        remove.grid(row=0, column=2, padx=(8, 0))

        self._section_rows.append(
            {"frame": row, "var": var, "entry": entry,
             "remove": remove, "label": index_label}
        )
        self._sync_section_state()
        if self.compare_mode.get() == CompareScope.SECTION.value:
            entry.focus_set()

    def _remove_section_row(self, frame) -> None:
        """Delete a section field (the last remaining one is only cleared)."""
        if len(self._section_rows) <= 1:
            self._section_rows[0]["var"].set("")
            return
        for i, r in enumerate(self._section_rows):
            if r["frame"] is frame:
                r["frame"].destroy()
                self._section_rows.pop(i)
                break
        self._sync_section_state()

    def _sync_section_state(self) -> None:
        """Enable/disable the section fields to match the selected scope and
        renumber their labels."""
        enabled = self.compare_mode.get() == CompareScope.SECTION.value
        state = "normal" if enabled else "disabled"
        for i, r in enumerate(self._section_rows, start=1):
            r["label"].configure(text=f"{i}.")
            r["entry"].configure(state=state)
            r["remove"].configure(state=state)
        self.add_section_button.configure(state=state)

    def _collect_sections(self) -> List[str]:
        """The non-empty section targets, de-duplicated, in field order. A
        single field may also hold several targets separated by ';'."""
        seen: set = set()
        out: List[str] = []
        for r in self._section_rows:
            raw = str(r["var"].get())
            for piece in re.split(r"[;\n]+", raw):
                piece = piece.strip()
                if piece and piece.lower() not in seen:
                    seen.add(piece.lower())
                    out.append(piece)
        return out

    # ------------------------------------------------------------------ #
    # Event handlers
    # ------------------------------------------------------------------ #

    def _on_scope_change(self) -> None:
        self._sync_section_state()

    def _pick_file(self, var: tk.StringVar) -> None:
        path = filedialog.askopenfilename(title="Select Word Document", filetypes=_FILETYPES)
        if path:
            var.set(path)

    def _pick_output_folder(self) -> None:
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            self.output_folder.set(path)

    def _on_compare_clicked(self) -> None:
        if self._worker and self._worker.is_alive():
            return

        original = self.original_path.get().strip()
        modified = self.modified_path.get().strip()
        out_folder = self.output_folder.get().strip()
        basename = self.output_basename.get().strip() or "Comparison_Report"
        basename = re.sub(r"\.docx?$", "", basename, flags=re.IGNORECASE)

        if not original or not modified:
            messagebox.showerror("Missing Files", "Please select both the original and modified documents.")
            return
        if not out_folder:
            messagebox.showerror("Missing Output Folder", "Please choose an output folder.")
            return

        scope = CompareScope(self.compare_mode.get())
        target_sections: List[str] = []
        if scope == CompareScope.SECTION:
            target_sections = self._collect_sections()
            if not target_sections:
                messagebox.showerror(
                    "Missing Section",
                    "Please enter at least one section/heading to compare "
                    "(use '+ Add another section' for several), or switch to "
                    "another comparison scope.",
                )
                return

        # The report title is purely the heading printed on the report's first
        # page; the saved file is always named from the "Base file name" field.
        report_title = self.report_title.get().strip() or None

        request = ComparisonRequest(
            original_path=Path(original),
            modified_path=Path(modified),
            output_folder=Path(out_folder),
            scope=scope,
            target_sections=target_sections,
            report_title=report_title,
            output_basename=basename,
        )

        self._set_busy(True)
        self.progress["value"] = 0
        self.status_text.set("Starting comparison...")

        self._worker = threading.Thread(target=self._run_comparison, args=(request,), daemon=True)
        self._worker.start()

    def _run_comparison(self, request: ComparisonRequest) -> None:
        def on_progress(message: str, fraction: float) -> None:
            self.root.after(0, self._update_progress, message, fraction)

        try:
            result = self._controller.run(request, progress=on_progress)
        except DocumentComparisonError as exc:
            logger.error("Comparison failed: %s", exc)
            self.root.after(0, self._on_failure, str(exc))
        except Exception as exc:  # unexpected/unhandled errors
            logger.exception("Unexpected error during comparison")
            self.root.after(0, self._on_failure, f"An unexpected error occurred: {exc}")
        else:
            self.root.after(0, self._on_success, result)

    # ------------------------------------------------------------------ #
    # UI state updates (must run on main thread via `after`)
    # ------------------------------------------------------------------ #

    def _update_progress(self, message: str, fraction: float) -> None:
        self.progress["value"] = fraction
        self.status_text.set(message)

    def _on_success(self, result: ComparisonResult) -> None:
        self._set_busy(False)
        self.progress["value"] = 1.0
        c = result.counts

        warning = ""
        if result.unchanged_sections:
            warning += (
                "\n\nNo changes in:\n  - "
                + "\n  - ".join(result.unchanged_sections)
            )
        if result.missing_sections:
            warning += (
                "\n\nNot found in either document (skipped):\n  - "
                + "\n  - ".join(result.missing_sections)
            )

        # Section mode always writes the requested sections, so a zero here can
        # only happen in auto mode: nothing changed at all.
        if result.changed_sections == 0:
            self.status_text.set("No changes found.")
            messagebox.showinfo(
                "No Changes",
                "No changes were made in the compared content.\n\n"
                "There is nothing to show, so no before/after content was "
                "generated." + warning,
            )
            return

        if result.scope == CompareScope.SECTION:
            written = result.changed_sections
            changed = written - len(result.unchanged_sections)
            headline = (
                f"{written} section(s) written to the report, "
                f"{changed} of them changed."
            )
            self.status_text.set(f"Done - {changed}/{written} section(s) changed.")
        else:
            headline = f"{result.changed_sections} changed section(s) written to the report."
            self.status_text.set(f"Done - {result.changed_sections} changed section(s).")

        messagebox.showinfo(
            "Comparison Complete",
            f"{headline}\n\n"
            f"Word-level changes: {c[ChangeType.ADDED]} added, "
            f"{c[ChangeType.DELETED]} deleted, {c[ChangeType.MODIFIED]} modified, "
            f"{c[ChangeType.MOVED]} moved.\n\n"
            f"Report saved to:\n{result.report_path}" + warning,
        )

    def _on_failure(self, message: str) -> None:
        self._set_busy(False)
        self.progress["value"] = 0
        self.status_text.set("Failed.")
        messagebox.showerror("Comparison Failed", message)

    def _set_busy(self, busy: bool) -> None:
        self.compare_button.configure(state="disabled" if busy else "normal")

    def _on_close(self) -> None:
        try:
            self._controller.cleanup()
        finally:
            self.root.destroy()


def launch() -> None:
    """Entry point used by main() to start the GUI event loop."""
    if not _TK_AVAILABLE:
        raise RuntimeError(
            "Tkinter is not available in this Python installation, so the "
            "graphical interface cannot start. Install the Tk bindings (e.g. "
            "the 'python3-tk' package on Debian/Ubuntu) or drive "
            "ComparisonController.run(...) directly from a script."
        )
    root = tk.Tk()
    WordDiffApp(root)
    root.mainloop()


# =========================================================================== #
#                                                                             #
#  main                                                                       #
#  - application entry point                                                  #
#                                                                             #
# =========================================================================== #


def main() -> int:
    log_dir = Path.home() / ".word_diff_reviewer" / "logs"
    setup_logging(log_file=log_dir / "app.log")
    logger = get_logger("main")
    logger.info("Starting Word Document Change Review application")

    try:
        launch()
    except Exception:
        logger.exception("Fatal error, application terminated")
        raise
    return 0


if __name__ == "__main__":
    sys.exit(main())