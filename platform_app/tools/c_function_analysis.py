"""Adapter for the C function analysis engine.

Wraps `analyze_function` from the merged `smart_x` package, then hands the
result to whichever exporter the run asked for: the standalone multi-sheet
workbook, the Data Dictionary template, or both.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List

from ..registry import Artifact, Progress, RunResult
from ..scripts.smart_x import c_parser_core, excel_exporter, template_exporter
from . import _sources

DEFAULT_TEMPLATE = Path(template_exporter.__file__).resolve().parent / "assets" / "data_dictionary_template.xlsx"


def _safe_name(raw: str, fallback: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._ -]+", "", str(raw or "")).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    if cleaned.lower().endswith(".xlsx"):
        cleaned = cleaned[:-5]
    return cleaned or fallback


def run(values: Dict[str, Any], output_dir: Path, say: Progress) -> RunResult:
    uploads: List[Path] = values.get("source_files") or []
    func_name = (values.get("function_name") or "").strip()
    export_mode = values.get("export_mode") or "workbook"
    template_upload: Path | None = values.get("template_file")

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

    say(f"Looking for {func_name}() ...", 0.30)
    info = c_parser_core.analyze_function(func_name, [str(p) for p in files])

    if not info.found:
        raise ValueError(
            f"{func_name}() was not defined in any of the {len(files)} files "
            "supplied. Check the spelling, and that the file holding it was "
            "included."
        )

    say(
        f"Found in {Path(info.source_file).name}, lines {info.start_line}–{info.end_line}.",
        0.45,
    )

    artifacts: List[Artifact] = []
    base = _safe_name(values.get("output_name"), f"Analysis_{func_name}")

    if export_mode in ("workbook", "both"):
        say("Writing the analysis workbook ...", 0.62)
        out_path = output_dir / f"{base}.xlsx"
        excel_exporter.export_to_excel(info, str(out_path))
        artifacts.append(Artifact(out_path, "Function analysis", role="report"))

    if export_mode in ("template", "both"):
        template_path = template_upload or DEFAULT_TEMPLATE
        if not Path(template_path).is_file():
            raise ValueError(
                "No Data Dictionary template available. Upload one, or pick "
                "the standalone workbook instead."
            )
        say(f"Filling {Path(template_path).name} ...", 0.80)
        filled = output_dir / f"{base}_DataDictionary.xlsx"
        template_exporter.export_to_template(info, str(template_path), str(filled))
        artifacts.append(Artifact(filled, "Data Dictionary", role="after"))

    say("Done.", 0.95)

    stats = [
        {"label": "Parameters", "value": len(info.parameters), "tone": "accent"},
        {"label": "Local variables", "value": len(info.local_variables), "tone": "neutral"},
        {"label": "Globals used", "value": len(info.global_variables), "tone": "neutral"},
        {"label": "Conditions", "value": len(info.conditions), "tone": "neutral"},
        {"label": "Calls out", "value": len(info.function_calls), "tone": "neutral"},
        {"label": "Lines", "value": max(0, info.end_line - info.start_line + 1), "tone": "neutral"},
    ]

    rows = [
        {"cells": [p.name, p.type_, "parameter", p.line_number or ""], "tone": "neutral"}
        for p in info.parameters
    ] + [
        {"cells": [v.name, v.type_, "local", v.line_number or ""], "tone": "quiet"}
        for v in info.local_variables
    ]

    table = {
        "title": f"{func_name}() variables",
        "columns": ["Name", "Type", "Scope", "Line"],
        "rows": rows[:80],
        "truncated": len(rows) > 80,
    } if rows else None

    notes: List[str] = [
        f"Return type: {info.return_type or 'unknown'} · "
        f"defined in {Path(info.source_file).name}"
    ]
    if not info.function_calls:
        notes.append("This function calls nothing else — it is a leaf.")

    return RunResult(artifacts=artifacts, stats=stats, table=table, notes=notes)
