"""Adapter for the Impact SUTC LLT engine.

Wraps ``diff_workbooks`` from ``scripts/excel_highlight_engine.py``. The engine
is untouched; this module only translates the platform's job vocabulary into
the arguments it expects and turns its ``DiffResult`` back into artifacts and
stats the page can render.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List

from ..registry import Artifact, Progress, RunResult
from ..scripts import excel_highlight_engine as engine

# The engine reports free-text progress with no fraction. These are the phases
# it announces, in order, used to move the bar honestly rather than fake it.
_PHASE_HINTS = [
    ("loading", 0.15),
    ("reading the template", 0.20),
    ("comparing", 0.45),
    ("aligning", 0.55),
    ("writing", 0.80),
    ("saving", 0.90),
]


def _fraction_for(message: str, seen: List[float]) -> float:
    low = message.lower()
    for token, frac in _PHASE_HINTS:
        if token in low:
            seen.append(frac)
            break
    return max(seen) if seen else 0.05


def _short(value: Any, limit: int = 60) -> str:
    if value is None:
        return ""
    text = str(value).replace("\n", " ").strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"


def run(values: Dict[str, Any], output_dir: Path, say: Progress) -> RunResult:
    old_path: Path = values["old_file"]
    new_path: Path = values["new_file"]
    template: Path | None = values.get("template_file")

    seen: List[float] = []

    def relay(message: str) -> None:
        say(message, _fraction_for(message, seen))

    if template:
        say(f"Using {template.name} as the blank template.", 0.04)

    result = engine.diff_workbooks(
        old_path, new_path, output_dir, template=template, progress=relay
    )

    artifacts: List[Artifact] = []
    if result.out_old:
        artifacts.append(
            Artifact(Path(result.out_old), "Original, changes in red", role="before")
        )
    if result.out_new:
        artifacts.append(
            Artifact(Path(result.out_new), "Modified, changes in green", role="after")
        )

    stats = [
        {"label": "Changed cells", "value": result.cell_changes, "tone": "accent"},
        {"label": "Modified", "value": result.count("modified"), "tone": "neutral"},
        {"label": "Added", "value": result.count("added"), "tone": "after"},
        {"label": "Deleted", "value": result.count("deleted"), "tone": "before"},
    ]
    if result.added_rows:
        stats.append({"label": "Rows added", "value": len(result.added_rows), "tone": "after"})
    if result.deleted_rows:
        stats.append({"label": "Rows deleted", "value": len(result.deleted_rows), "tone": "before"})

    rows = []
    for change in result.changes[:400]:
        if change.kind in ("sheet-added", "sheet-deleted"):
            continue
        rows.append(
            {
                "cells": [
                    change.sheet,
                    change.old_cell or change.cell or "",
                    change.new_cell or change.cell or "",
                    _short(change.old_value),
                    _short(change.new_value),
                ],
                "tone": {"deleted": "before", "added": "after"}.get(change.kind, "neutral"),
                "kind": change.kind,
            }
        )

    table = {
        "title": "Changed cells",
        "columns": ["Sheet", "Old cell", "New cell", "Was", "Now"],
        "rows": rows,
        "truncated": len(result.changes) > len(rows),
    } if rows else None

    notes: List[str] = []
    if result.sheets_only_in_old:
        notes.append("Sheets only in the original: " + ", ".join(result.sheets_only_in_old))
    if result.sheets_only_in_new:
        notes.append("Sheets only in the modified file: " + ", ".join(result.sheets_only_in_new))
    if not result.changes:
        notes.append("No differences found. The two workbooks match.")

    for line in engine.summary_lines(result, limit=12):
        say(line, 0.98)

    return RunResult(artifacts=artifacts, stats=stats, table=table, notes=notes)
