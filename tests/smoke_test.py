"""End-to-end check: submits one run per tool through the HTTP API and waits.

    python tests/smoke_test.py

Passes if every tool finishes and produces at least one downloadable file.
"""

from __future__ import annotations

import io
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from platform_app.app import create_app  # noqa: E402

FIXTURES = Path(__file__).parent / "fixtures"


def upload(path: Path):
    return (io.BytesIO(path.read_bytes()), path.name)


CASES = [
    (
        "excel-highlight",
        lambda: {
            "old_file": upload(FIXTURES / "before.xlsx"),
            "new_file": upload(FIXTURES / "after.xlsx"),
        },
    ),
    (
        "excel-table-diff",
        lambda: {
            "before_file": upload(FIXTURES / "before.xlsx"),
            "after_file": upload(FIXTURES / "after.xlsx"),
            "key_column": "0",
            "output_name": "Parameter_Changes",
        },
    ),
    (
        "word-sections",
        lambda: {
            "original_file": upload(FIXTURES / "before.docx"),
            "modified_file": upload(FIXTURES / "after.docx"),
            "scope": "auto",
            "report_title": "Impact CR 1234",
            "output_name": "Requirements_Impact",
        },
    ),
    (
        # Loose sources into the multi-file field.
        "c-function-analysis",
        lambda: {
            "source_files": [
                upload(FIXTURES / "csrc" / "engine.c"),
                upload(FIXTURES / "csrc" / "engine.h"),
                upload(FIXTURES / "csrc" / "diag.c"),
            ],
            "function_name": "Engine_Update",
            "export_mode": "both",
            "output_name": "Engine_Update_Analysis",
        },
    ),
    (
        # The same sources as a zipped tree, to cover the unpacking path.
        "call-tree-extract",
        lambda: {
            "source_files": [upload(FIXTURES / "csrc.zip")],
            "sheet_name": "Call Tree",
            "output_name": "Call_Tree_Extracted",
        },
    ),
    (
        # Both sheets in one workbook, so no after file is supplied.
        "call-tree-compare",
        lambda: {
            "before_file": upload(FIXTURES / "call_tree.xlsx"),
            "before_sheet": "Call Tree",
            "after_sheet": "Call Tree (Apres modif)",
            "similarity": "85",
            "output_name": "Call_Tree_Comparison",
        },
    ),
    (
        "req-coverage-check",
        lambda: {
            "document": upload(FIXTURES / "requirements.docx"),
            "output_name": "Requirement_Coverage",
        },
    ),
]


def wait_for(client, job_id: str, timeout: float = 240.0) -> dict:
    deadline = time.time() + timeout
    while time.time() < deadline:
        job = client.get(f"/api/jobs/{job_id}").get_json()
        if job["status"] in ("done", "failed"):
            return job
        time.sleep(0.4)
    raise TimeoutError(f"{job_id} did not finish in {timeout}s")


def main() -> int:
    app = create_app()
    client = app.test_client()

    assert client.get("/").status_code == 200, "the page did not render"
    # Every registered tool should have a case here, so a new tool that nobody
    # wrote a fixture for fails the run rather than going quietly untested.
    registered = {t["id"] for t in client.get("/api/tools").get_json()}
    covered = {tool_id for tool_id, _ in CASES}
    assert registered == covered, f"no smoke case for: {sorted(registered - covered)}"

    failures = []
    for tool_id, payload in CASES:
        data = payload()
        data["tool_id"] = tool_id
        started = client.post("/api/jobs", data=data, content_type="multipart/form-data")
        if started.status_code != 202:
            failures.append(f"{tool_id}: submit returned {started.status_code} {started.get_json()}")
            continue

        job = wait_for(client, started.get_json()["id"])
        mark = "ok  " if job["status"] == "done" else "FAIL"
        stats = "  ".join(f"{s['label']}={s['value']}" for s in job["stats"])
        print(f"[{mark}] {tool_id:<18} {job['duration']:>6.1f}s  {stats}")

        if job["status"] != "done":
            failures.append(f"{tool_id}: {job['error']}")
            for line in job["log"][-8:]:
                print("        " + line["text"])
            continue

        if not job["artifacts"]:
            failures.append(f"{tool_id}: produced no files")
            continue

        for artifact in job["artifacts"]:
            got = client.get(artifact["url"])
            size = len(got.data)
            if got.status_code != 200 or size < 200:
                failures.append(f"{tool_id}: {artifact['name']} would not download")
            print(f"        {artifact['name']:<38} {size:>8,} bytes")
        for note in job["notes"]:
            print(f"        note: {note}")

    # Bad input should come back as a field error, not a crash.
    rejected = client.post(
        "/api/jobs",
        data={"tool_id": "excel-table-diff", "before_file": upload(FIXTURES / "before.docx")},
        content_type="multipart/form-data",
    )
    if rejected.status_code != 400:
        failures.append(f"wrong file type was accepted ({rejected.status_code})")
    else:
        print(f"[ok  ] validation        rejected a .docx: {rejected.get_json()['fields']}")

    print()
    if failures:
        print("FAILURES")
        for line in failures:
            print(" -", line)
        return 1
    print("All tools passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
