"""Adapter for the Impact Call Tree engine.

Wraps `SimpleCallTreeComparator` from the merged `smart_x` package. That class
already returns its own counts, so this mostly translates its summary dict into
the platform's stats.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List

import openpyxl

from ..registry import Artifact, Progress, RunResult
from ..scripts.smart_x.call_tree_comparator import SimpleCallTreeComparator


def _safe_name(raw: str, fallback: str = "Call_Tree_Comparison") -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._ -]+", "", str(raw or "")).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    if cleaned.lower().endswith(".xlsx"):
        cleaned = cleaned[:-5]
    return cleaned or fallback


def _sheet_names(path: Path) -> List[str]:
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _resolve(path: Path, wanted: str, role: str) -> str:
    """Match the requested sheet, case-insensitively, or explain what is there."""
    names = _sheet_names(path)
    for name in names:
        if name.strip().lower() == wanted.strip().lower():
            return name
    available = ", ".join(f"'{n}'" for n in names)
    raise ValueError(
        f"{path.name} has no sheet called '{wanted}' for the {role} side. "
        f"It holds: {available}."
    )


def run(values: Dict[str, Any], output_dir: Path, say: Progress) -> RunResult:
    before_file: Path = values["before_file"]
    after_file: Path | None = values.get("after_file")
    before_sheet = (values.get("before_sheet") or "Call Tree").strip()
    after_sheet = (values.get("after_sheet") or "Call Tree (Apres modif)").strip()
    threshold = values.get("similarity")

    out_path = output_dir / f"{_safe_name(values.get('output_name'))}.xlsx"

    # With no second file the two sheets are expected to live in the first one.
    if after_file is None:
        say(f"Both sides from {before_file.name}.", 0.12)
        before_sheet = _resolve(before_file, before_sheet, "before")
        after_sheet = _resolve(before_file, after_sheet, "after")
    else:
        say(f"Comparing {before_file.name} against {after_file.name}.", 0.12)
        before_sheet = _resolve(before_file, before_sheet, "before")
        after_sheet = _resolve(after_file, after_sheet, "after")

    comparator = SimpleCallTreeComparator()
    if threshold is not None:
        # The form collects a percentage; the engine wants a 0..1 ratio.
        comparator.similarity_threshold = max(0.0, min(1.0, int(threshold) / 100))

    say(
        f"Matching rows, {int(comparator.similarity_threshold * 100)}% similarity "
        "counts as a rename ...",
        0.35,
    )

    summary = comparator.create_comparison_report(
        str(before_file),
        str(after_file) if after_file else None,
        str(out_path),
        before_sheet=before_sheet,
        after_sheet=after_sheet,
    )

    say("Writing the highlighted report ...", 0.88)

    # The engine falls back to a timestamped name if the target is locked.
    produced = Path(summary.get("output_file") or out_path)

    stats = [
        {"label": "Rows removed", "value": summary["deleted"], "tone": "before"},
        {"label": "Rows added", "value": summary["added"], "tone": "after"},
        {"label": "Rows modified", "value": summary["modified"], "tone": "neutral"},
        {"label": "Unchanged", "value": summary["unchanged"], "tone": "quiet"},
    ]

    notes: List[str] = [
        "The report holds a 'before' sheet with removed rows in red and an "
        "'after' sheet with added rows in green.",
    ]
    changed = summary["deleted"] + summary["added"] + summary["modified"]
    if not changed:
        notes.append("The two call trees are identical.")

    return RunResult(
        artifacts=[Artifact(produced, "Call tree comparison", role="report")],
        stats=stats,
        notes=notes,
    )
