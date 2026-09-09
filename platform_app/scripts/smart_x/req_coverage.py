"""Requirement / coverage ID style check, from `REQ_COV_Detector/check_requirements.py`.

The detection rules are the upstream ones unchanged: the same two regexes, the
same "only scan inside a Requirements section" state machine, the same style
names counted as correct (`exi id` for requirement IDs, `exi traice` / `exi
trace` for coverage IDs), and the same red/green fills in a two-sheet workbook.

Three things changed so it can run as a platform tool:

  * it returns the counts instead of only printing them, so the run panel can
    show how many IDs carry the wrong style;
  * a document that fails to load now raises instead of printing and returning
    None, so the run is reported as failed rather than as an empty success;
  * the macro-enabled repack writes to a caller-supplied scratch directory.
"""

from __future__ import annotations

import os
import re
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Dict, List

import docx
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# Upstream patterns, unchanged.
REQ_PATTERN = re.compile(r'(REQ[-_]SDDD_[A-Za-z0-9_-]+)')
COV_PATTERN = re.compile(r'(COV\.REQ[A-Za-z0-9_.-]+)')

REQ_STYLE = "exi id"
COV_STYLES = ("exi traice", "exi trace")

# Word formats whose content type has to be rewritten before python-docx will
# open them.
MACRO_SUFFIXES = ("docm", "dotm", "dot", "doc")


def _open_document(word_file: Path, scratch_dir: Path | None):
    """Open a Word file, repacking macro-enabled formats python-docx refuses."""
    ext = word_file.name.lower().rsplit(".", 1)[-1]
    if ext not in MACRO_SUFFIXES:
        return docx.Document(str(word_file)), None

    handle = tempfile.NamedTemporaryFile(
        suffix=".docx", delete=False, dir=str(scratch_dir) if scratch_dir else None
    )
    tmp_path = Path(handle.name)
    handle.close()

    with zipfile.ZipFile(word_file, "r") as zin:
        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    data = data.replace(
                        b"application/vnd.ms-word.document.macroEnabled.main+xml",
                        b"application/vnd.openxmlformats-officedocument."
                        b"wordprocessingml.document.main+xml",
                    )
                    data = data.replace(
                        b"application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
                        b"application/vnd.openxmlformats-officedocument."
                        b"wordprocessingml.document.main+xml",
                    )
                zout.writestr(item, data)

    return docx.Document(str(tmp_path)), tmp_path


def _iter_block_items(parent):
    """Paragraphs and tables in document order."""
    for child in parent.element.body.iterchildren():
        if child.tag.endswith("p"):
            yield docx.text.paragraph.Paragraph(child, parent)
        elif child.tag.endswith("tbl"):
            yield docx.table.Table(child, parent)


def check_styles_and_export(
    word_file: Path,
    excel_file: Path,
    scratch_dir: Path | None = None,
) -> Dict[str, Any]:
    """Scan `word_file` for requirement and coverage IDs and write the report.

    Returns counts plus the two result lists.
    """
    word_file = Path(word_file)
    doc, tmp_path = _open_document(word_file, scratch_dir)

    try:
        req_results: List[Dict[str, Any]] = []
        cov_results: List[Dict[str, Any]] = []
        state = {"in_requirements": False}

        def process_text(text: str, style_name: str) -> None:
            stripped = text.strip()
            if not stripped:
                return

            # A "Requirements" heading opens the scanned region ...
            if re.search(r'(?i)^\s*(?:\d+(?:\.\d+)*\s+)?Requirements\s*$', stripped):
                state["in_requirements"] = True
                return

            # ... and any other numbered heading closes it again.
            if re.search(r'(?i)^\s*\d+(?:\.\d+)*\s+[A-Za-z]', stripped) or (
                style_name
                and "heading" in style_name.lower()
                and "requirements" not in stripped.lower()
            ):
                state["in_requirements"] = False
                return

            if not state["in_requirements"]:
                return

            for req in REQ_PATTERN.findall(text):
                req_results.append({
                    "id": req,
                    "style": style_name,
                    "is_valid": style_name.strip().lower() == REQ_STYLE,
                })

            for cov in COV_PATTERN.findall(text):
                cov_results.append({
                    "id": cov,
                    "style": style_name,
                    "is_valid": style_name.strip().lower() in COV_STYLES,
                })

        for block in _iter_block_items(doc):
            if isinstance(block, docx.text.paragraph.Paragraph):
                process_text(block.text, block.style.name if block.style else "")
            else:
                for row in block.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            process_text(para.text, para.style.name if para.style else "")

        req_unique = _dedupe(req_results)
        cov_unique = _dedupe(cov_results)

        _write_report(Path(excel_file), req_unique, cov_unique)

        return {
            "requirements": req_unique,
            "coverage": cov_unique,
            "req_total": len(req_unique),
            "req_bad": sum(1 for r in req_unique if not r["is_valid"]),
            "cov_total": len(cov_unique),
            "cov_bad": sum(1 for c in cov_unique if not c["is_valid"]),
        }
    finally:
        if tmp_path and tmp_path.exists():
            os.remove(tmp_path)


def _dedupe(results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    seen: set = set()
    unique = []
    for item in results:
        if item["id"] not in seen:
            seen.add(item["id"])
            unique.append(item)
    return unique


def _write_report(excel_file: Path, req_rows, cov_rows) -> None:
    wb = Workbook()
    red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    green = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

    for title, header, rows in (
        ("LLR ID", "LLR ID", req_rows),
        ("HLR ID", "HLR ID", cov_rows),
    ):
        ws = wb.active if title == "LLR ID" else wb.create_sheet()
        ws.title = title
        ws.append([header, "style"])
        for row_idx, res in enumerate(rows, start=2):
            fill = green if res["is_valid"] else red
            for col, value in ((1, res["id"]), (2, res["style"])):
                cell = ws.cell(row=row_idx, column=col, value=value)
                cell.fill = fill

    wb.save(excel_file)
