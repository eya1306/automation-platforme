"""Adapter for the requirement / coverage ID style check.

Wraps `check_styles_and_export` from the merged `smart_x` package: it scans the
Requirements sections of a Word document for REQ-SDDD and COV.REQ identifiers
and reports which of them carry the wrong paragraph style.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List

from ..registry import Artifact, Progress, RunResult
from ..scripts.smart_x import req_coverage
from . import _sources


def _safe_name(raw: str, fallback: str = "Requirement_Coverage") -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._ -]+", "", str(raw or "")).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    if cleaned.lower().endswith(".xlsx"):
        cleaned = cleaned[:-5]
    return cleaned or fallback


def run(values: Dict[str, Any], output_dir: Path, say: Progress) -> RunResult:
    document: Path = values["document"]
    out_path = output_dir / f"{_safe_name(values.get('output_name'))}.xlsx"

    say(f"Opening {document.name} ...", 0.15)
    scratch = _sources.work_dir(output_dir, "docx")

    say("Scanning the Requirements sections ...", 0.40)
    summary = req_coverage.check_styles_and_export(document, out_path, scratch_dir=scratch)

    say(
        f"Found {summary['req_total']} requirement ID(s) and "
        f"{summary['cov_total']} coverage ID(s).",
        0.80,
    )

    req_ok = summary["req_total"] - summary["req_bad"]
    cov_ok = summary["cov_total"] - summary["cov_bad"]

    stats = [
        {"label": "Requirement IDs", "value": summary["req_total"], "tone": "accent"},
        {"label": "REQ wrong style", "value": summary["req_bad"],
         "tone": "before" if summary["req_bad"] else "after"},
        {"label": "Coverage IDs", "value": summary["cov_total"], "tone": "accent"},
        {"label": "COV wrong style", "value": summary["cov_bad"],
         "tone": "before" if summary["cov_bad"] else "after"},
    ]

    rows = []
    for item in summary["requirements"]:
        rows.append({
            "cells": [item["id"], "Requirement", item["style"] or "—",
                      "ok" if item["is_valid"] else "wrong style"],
            "tone": "after" if item["is_valid"] else "before",
        })
    for item in summary["coverage"]:
        rows.append({
            "cells": [item["id"], "Coverage", item["style"] or "—",
                      "ok" if item["is_valid"] else "wrong style"],
            "tone": "after" if item["is_valid"] else "before",
        })

    table = {
        "title": "Identifiers found",
        "columns": ["ID", "Kind", "Paragraph style", "Verdict"],
        "rows": rows[:100],
        "truncated": len(rows) > 100,
    } if rows else None

    notes: List[str] = []
    if not rows:
        notes.append(
            "No identifiers were found. The scan only reads inside a section "
            "headed 'Requirements', and looks for REQ-SDDD_… and COV.REQ… IDs."
        )
    else:
        notes.append(
            f"Requirement IDs should carry the '{req_coverage.REQ_STYLE}' style "
            f"({req_ok} of {summary['req_total']} do); coverage IDs should carry "
            f"'{req_coverage.COV_STYLES[0]}' ({cov_ok} of {summary['cov_total']} do)."
        )

    return RunResult(
        artifacts=[Artifact(out_path, "Coverage check", role="report")],
        stats=stats,
        table=table,
        notes=notes,
    )
