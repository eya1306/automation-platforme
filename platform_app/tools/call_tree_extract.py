"""Adapter for the call-tree extraction engine.

Wraps `CCodeExtractor` from the merged `smart_x` package and the workbook
writer lifted out of that project's desktop window. The engine walks a
directory, so loose uploads and zipped trees are both flattened into one
working directory first.
"""

from __future__ import annotations

import re
import shutil
from pathlib import Path
from typing import Any, Dict, List

from ..registry import Artifact, Progress, RunResult
from ..scripts.smart_x import call_tree_writer
from ..scripts.smart_x.call_tree_parser import CCodeExtractor
from . import _sources


def _safe_name(raw: str, fallback: str = "Call_Tree") -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._ -]+", "", str(raw or "")).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    if cleaned.lower().endswith((".xlsx", ".xlsm")):
        cleaned = cleaned.rsplit(".", 1)[0]
    return cleaned or fallback


def run(values: Dict[str, Any], output_dir: Path, say: Progress) -> RunResult:
    uploads: List[Path] = values.get("source_files") or []
    template_upload: Path | None = values.get("template_file")
    sheet_name = (values.get("sheet_name") or call_tree_writer.SHEET_NAME).strip()

    say("Collecting the source files ...", 0.08)
    sources_dir = _sources.work_dir(output_dir)
    files, archives = _sources.gather(uploads, sources_dir)

    if not files:
        raise ValueError(
            "No C source files were found in that upload. Expected .c or .h "
            "files, either loose or inside a .zip."
        )

    if archives:
        say(f"Expanded {archives} archive(s) to {len(files)} source file(s).", 0.16)
    else:
        say(f"Reading {len(files)} source file(s).", 0.16)

    say("Walking the call graph ...", 0.35)
    extractor = CCodeExtractor(str(sources_dir))
    rows = extractor.extract()

    if not rows:
        raise ValueError(
            "No function calls were found in those sources. Check that the "
            "upload holds the .c files themselves and not only headers."
        )

    say(f"{len(rows)} call-tree rows extracted.", 0.65)

    out_path = output_dir / f"{_safe_name(values.get('output_name'))}.xlsx"

    # Start from the supplied structure if there is one, otherwise from the
    # empty workbook the upstream project ships.
    start_from = template_upload or call_tree_writer.EMPTY_TEMPLATE
    if Path(start_from).is_file():
        shutil.copy(start_from, out_path)
        say(f"Filling {Path(start_from).name} ...", 0.78)
    else:
        say("No template available — writing a fresh workbook.", 0.78)

    written = call_tree_writer.write_call_tree(out_path, rows, sheet_name=sheet_name)
    say(f"Wrote {written} rows to the '{sheet_name}' sheet.", 0.94)

    callers = {str(r[0]) for r in rows if len(r) > 0 and r[0]}
    callees = {str(r[5]) for r in rows if len(r) > 5 and r[5]}
    conditional = sum(
        1 for r in rows
        if len(r) > 6 and r[6] and str(r[6]).strip() not in ("", "None")
    )

    stats = [
        {"label": "Call-tree rows", "value": written, "tone": "accent"},
        {"label": "Source files", "value": len(files), "tone": "neutral"},
        {"label": "Calling CSCs", "value": len(callers), "tone": "neutral"},
        {"label": "Called functions", "value": len(callees), "tone": "neutral"},
        {"label": "Conditional calls", "value": conditional, "tone": "neutral"},
    ]

    preview = [
        {
            "cells": [str(c) if c not in (None, "") else "—" for c in list(row[:7])],
            "tone": "neutral",
        }
        for row in rows[:40]
    ]

    table = {
        "title": "Extracted call tree",
        "columns": call_tree_writer.HEADERS,
        "rows": preview,
        "truncated": len(rows) > 40,
    }

    notes = [
        f"The sheet is named '{sheet_name}'. To compare two of these, run "
        "Impact Call Tree with one workbook as before and one as after."
    ]

    return RunResult(
        artifacts=[Artifact(out_path, "Call tree workbook", role="report")],
        stats=stats,
        table=table,
        notes=notes,
    )
