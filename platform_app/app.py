"""The web layer.

Thin on purpose: it validates an upload against the tool's declared inputs,
drops the files into the run's folder, hands the run to the job store and
serves status back. All the domain knowledge sits in the registry and the
adapters.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Tuple

from flask import Flask, abort, jsonify, render_template, request, send_from_directory
from werkzeug.utils import secure_filename

from . import config
from .jobs import store
from .registry import CHECKBOX, FILE, FILES, LINES, NUMBER, SELECT, TOOLS, get_tool


def create_app() -> Flask:
    config.ensure_dirs()
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = config.MAX_UPLOAD_MB * 1024 * 1024
    app.config["JSON_SORT_KEYS"] = False

    # ---------------------------------------------------------------- pages #

    @app.get("/")
    def index():
        return render_template(
            "index.html",
            tools=[t.as_json() for t in TOOLS],
            max_upload_mb=config.MAX_UPLOAD_MB,
            app_name=config.APP_NAME,
            app_tagline=config.APP_TAGLINE,
            org_name=config.ORG_NAME,
            logo_file=config.logo_filename(),
        )

    # ------------------------------------------------------------------ api #

    @app.get("/api/tools")
    def list_tools():
        return jsonify([t.as_json() for t in TOOLS])

    @app.get("/api/jobs")
    def list_jobs():
        limit = request.args.get("limit", type=int) or config.HISTORY_LIMIT
        return jsonify([summarise(job) for job in store.recent(limit)])

    @app.post("/api/jobs")
    def create_job():
        tool_id = request.form.get("tool_id", "")
        tool = get_tool(tool_id)
        if not tool:
            return jsonify({"error": f"No tool called {tool_id!r}."}), 404

        # Clearing expired runs here keeps the workspace from growing without
        # needing a scheduler alongside the app.
        store.sweep()

        job = store.create(tool, _label_for(tool_id))

        values, errors = _collect(tool, job.input_dir)
        if errors:
            store.discard(job.id)
            return jsonify({"error": "Check the highlighted fields.", "fields": errors}), 400

        job.inputs = [p.name for p in _uploaded_paths(values)]
        job.options = {
            k: v for k, v in values.items() if not _is_upload(v)
        }
        job.label = _label_for(tool_id, values)
        store.submit(job, tool, values)
        return jsonify({"id": job.id, "status": job.status}), 202

    @app.get("/api/jobs/<job_id>")
    def job_status(job_id: str):
        job = store.get(job_id)
        if not job:
            return jsonify({"error": "That run is no longer available."}), 404
        since = request.args.get("since", type=int) or 0
        return jsonify(job.as_json(since=since))

    @app.delete("/api/jobs/<job_id>")
    def delete_job(job_id: str):
        if not store.discard(job_id):
            return jsonify({"error": "That run is no longer available."}), 404
        return jsonify({"ok": True})

    @app.get("/api/jobs/<job_id>/files/<path:name>")
    def download(job_id: str, name: str):
        job = store.get(job_id)
        if not job:
            abort(404)
        safe = Path(name).name
        target = job.output_dir / safe
        if not target.exists():
            abort(404)
        return send_from_directory(job.output_dir, safe, as_attachment=True)

    # --------------------------------------------------------------- errors #

    @app.errorhandler(413)
    def too_large(_):
        return (
            jsonify({"error": f"That upload is over the {config.MAX_UPLOAD_MB} MB limit."}),
            413,
        )

    return app


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #


def summarise(job) -> dict:
    return {
        "id": job.id,
        "tool_id": job.tool_id,
        "tool_name": job.tool_name,
        "label": job.label,
        "status": job.status,
        "progress": round(job.progress, 3),
        "created_at": job.created_at.isoformat(),
        "duration": round(job.duration, 1) if job.duration else None,
        "artifact_count": len(job.result.artifacts) if job.result else 0,
    }


def _is_upload(value: Any) -> bool:
    """True for a FILE value and for a FILES value (a list of paths)."""
    if isinstance(value, Path):
        return True
    return isinstance(value, list) and bool(value) and all(
        isinstance(item, Path) for item in value
    )


def _uploaded_paths(values: Dict[str, Any]) -> List[Path]:
    paths: List[Path] = []
    for value in values.values():
        if isinstance(value, Path):
            paths.append(value)
        elif _is_upload(value):
            paths.extend(value)
    return paths


def _label_for(tool_id: str, values: Dict[str, Any] | None = None) -> str:
    if not values:
        return "New run"

    # A multi-file field is labelled by its count, not by listing every name.
    for value in values.values():
        if isinstance(value, list) and len(value) > 1 and _is_upload(value):
            return f"{len(value)} files"

    names = [p.name for p in _uploaded_paths(values)]
    if len(names) >= 2:
        return f"{names[0]} → {names[1]}"
    if names:
        return names[0]
    return "New run"


def _collect(tool, input_dir: Path) -> Tuple[Dict[str, Any], Dict[str, str]]:
    """Pull every declared input off the request, validating as it goes."""
    values: Dict[str, Any] = {}
    errors: Dict[str, str] = {}

    for spec in tool.inputs:
        if spec.kind == FILE:
            uploaded = request.files.get(spec.name)
            if not uploaded or not uploaded.filename:
                if spec.required:
                    errors[spec.name] = "Pick a file."
                continue
            problem = _check_extension(uploaded.filename, spec.accept)
            if problem:
                errors[spec.name] = problem
                continue
            values[spec.name] = _save(uploaded, input_dir, spec.name)

        elif spec.kind == FILES:
            uploads = [f for f in request.files.getlist(spec.name) if f and f.filename]
            if not uploads:
                if spec.required:
                    errors[spec.name] = "Pick at least one file."
                else:
                    values[spec.name] = []
                continue
            problems = {
                f.filename: _check_extension(f.filename, spec.accept)
                for f in uploads
            }
            bad = [name for name, problem in problems.items() if problem]
            if bad:
                shown = ", ".join(bad[:3]) + ("…" if len(bad) > 3 else "")
                errors[spec.name] = f"{problems[bad[0]]} Rejected: {shown}"
                continue
            values[spec.name] = [_save(f, input_dir, spec.name) for f in uploads]

        elif spec.kind == LINES:
            raw = request.form.get(spec.name, "")
            values[spec.name] = [line.strip() for line in raw.splitlines() if line.strip()]

        elif spec.kind == NUMBER:
            raw = request.form.get(spec.name, "")
            if raw == "" and spec.default is not None:
                values[spec.name] = spec.default
                continue
            try:
                values[spec.name] = int(raw)
            except ValueError:
                errors[spec.name] = "Needs to be a whole number."
            else:
                if values[spec.name] < 0:
                    errors[spec.name] = "Cannot be negative."

        elif spec.kind == CHECKBOX:
            values[spec.name] = request.form.get(spec.name) in ("on", "true", "1")

        elif spec.kind == SELECT:
            raw = request.form.get(spec.name) or spec.default
            allowed = {o["value"] for o in spec.options}
            if allowed and raw not in allowed:
                errors[spec.name] = "Pick one of the listed options."
            else:
                values[spec.name] = raw

        else:  # TEXT
            raw = (request.form.get(spec.name) or "").strip()
            if not raw and spec.required:
                errors[spec.name] = "This one is needed."
            values[spec.name] = raw or (spec.default or "")

    return values, errors


def _check_extension(filename: str, accept: str) -> str | None:
    if not accept:
        return None
    allowed = {e.strip().lower() for e in accept.split(",") if e.strip()}
    suffix = Path(filename).suffix.lower()
    if suffix not in allowed:
        readable = ", ".join(sorted(allowed))
        return f"{suffix or 'That file'} is not supported here. Use one of: {readable}."
    return None


def _save(uploaded, input_dir: Path, field_name: str) -> Path:
    """Store the upload under a safe name, keeping the original for display."""
    original = Path(uploaded.filename).name
    safe = secure_filename(original) or f"{field_name}{Path(original).suffix}"
    if not Path(safe).suffix:
        safe = f"{safe}{Path(original).suffix}"
    target = input_dir / safe
    counter = 1
    while target.exists():
        stem, suffix = Path(safe).stem, Path(safe).suffix
        target = input_dir / f"{stem}_{counter}{suffix}"
        counter += 1
    uploaded.save(target)
    return target
