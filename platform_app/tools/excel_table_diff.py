"""Adapter for the Impact DD engine.

Wraps ``compare_files`` from ``scripts/excel_table_engine.py``. That engine
runs start to finish in one call with no progress hook, so the phases reported
here bracket the call rather than come from inside it.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List

from ..registry import Artifact, Progress, RunResult
from ..scripts import excel_table_engine as engine


def _safe_name(raw: str, fallback: str = "Change_Report") -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._ -]+", "", str(raw or "")).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    if cleaned.lower().endswith(".xlsx"):
        cleaned = cleaned[:-5]
    return cleaned or fallback


def run(values: Dict[str, Any], output_dir: Path, say: Progress) -> RunResult:
    before_path: Path = values["before_file"]
    after_path: Path = values["after_file"]
    key_index = int(values.get("key_column") or 0)
    out_path = output_dir / f"{_safe_name(values.get('output_name'))}.xlsx"

    say(f"Matching rows on column index {key_index}.", 0.10)
    say(f"Reading {before_path.name} and {after_path.name} ...", 0.25)

    summary = engine.compare_files(
        before_path, after_path, out_path, key_column_index=key_index
    )

    say("Writing the report workbook ...", 0.85)

    removed = sum(row[3] for row in summary)
    added = sum(row[4] for row in summary)
    modified = sum(row[5] for row in summary)

    stats = [
        {"label": "Sheets compared", "value": len(summary), "tone": "accent"},
        {"label": "Rows removed", "value": removed, "tone": "before"},
        {"label": "Rows added", "value": added, "tone": "after"},
        {"label": "Rows modified", "value": modified, "tone": "neutral"},
    ]

    rows = []
    for name, rows_before, rows_after, n_removed, n_added, n_modified in summary:
        changed = n_removed + n_added + n_modified
        rows.append(
            {
                "cells": [name, rows_before, rows_after, n_removed, n_added, n_modified],
                "tone": "neutral" if changed else "quiet",
            }
        )
        say(
            f"{name}: {n_removed} removed, {n_added} added, {n_modified} modified.",
            0.92,
        )

    table = {
        "title": "Per sheet",
        "columns": ["Sheet", "Rows before", "Rows after", "Removed", "Added", "Modified"],
        "rows": rows,
        "truncated": False,
    } if rows else None

    notes: List[str] = []
    if not (removed or added or modified):
        notes.append("No row differences found. Every key matched with identical values.")

    artifacts = [Artifact(out_path, "Row change report", role="report")]
    return RunResult(artifacts=artifacts, stats=stats, table=table, notes=notes)
