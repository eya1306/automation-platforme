"""Adapter for the Impact SDDD engine.

Wraps ``ComparisonController`` from ``scripts/word_diff_engine.py``. The engine
already reports progress as ``(message, fraction)``, which is exactly what the
job store wants, so it is handed straight through.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List

from ..registry import Artifact, Progress, RunResult
from ..scripts import word_diff_engine as engine


def _safe_name(raw: str, fallback: str = "Comparison_Report") -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._ -]+", "", str(raw or "")).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    if cleaned.lower().endswith(".docx"):
        cleaned = cleaned[:-5]
    return cleaned or fallback


_TONE_BY_CHANGE = {
    "added": "after",
    "deleted": "before",
    "modified": "neutral",
    "moved": "accent",
}


def run(values: Dict[str, Any], output_dir: Path, say: Progress) -> RunResult:
    scope_value = str(values.get("scope") or "auto")
    sections: List[str] = [s for s in (values.get("target_sections") or []) if s.strip()]

    if scope_value == "section" and not sections:
        raise ValueError(
            "Specific sections mode needs at least one section heading. "
            "Type one per line, or switch to the Requirements chapter."
        )

    scope = engine.CompareScope(scope_value)
    request = engine.ComparisonRequest(
        original_path=values["original_file"],
        modified_path=values["modified_file"],
        output_folder=output_dir,
        scope=scope,
        target_sections=sections,
        report_title=(values.get("report_title") or "").strip() or None,
        output_basename=_safe_name(values.get("output_name")),
    )

    controller = engine.ComparisonController()
    try:
        result = controller.run(request, progress=say)
    finally:
        controller.cleanup()

    counts = {str(getattr(k, "value", k)): v for k, v in (result.counts or {}).items()}
    stats = [
        {"label": "Changed sections", "value": result.changed_sections, "tone": "accent"},
        {"label": "Added", "value": counts.get("added", 0), "tone": "after"},
        {"label": "Deleted", "value": counts.get("deleted", 0), "tone": "before"},
        {"label": "Modified", "value": counts.get("modified", 0), "tone": "neutral"},
    ]
    if counts.get("moved"):
        stats.append({"label": "Moved", "value": counts["moved"], "tone": "accent"})

    rows = [
        {"cells": [label.title(), value], "tone": _TONE_BY_CHANGE.get(label, "neutral")}
        for label, value in sorted(counts.items())
        if value
    ]
    table = {
        "title": "Changes by kind",
        "columns": ["Kind", "Count"],
        "rows": rows,
        "truncated": False,
    } if rows else None

    notes: List[str] = []
    if result.missing_sections:
        notes.append(
            "Not found in either document: " + ", ".join(result.missing_sections)
        )
    if result.unchanged_sections:
        notes.append(
            "Requested but unchanged: " + ", ".join(result.unchanged_sections)
        )
    if not result.total:
        notes.append("No changes found in the compared scope.")

    artifacts = [Artifact(Path(result.report_path), "Section comparison report", role="report")]
    return RunResult(artifacts=artifacts, stats=stats, table=table, notes=notes)
