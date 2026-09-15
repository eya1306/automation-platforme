# AutoGeneration Platform

One place to run the team's document, spreadsheet and source-code tools. Pick a
tool, drop in the files, watch the run, download the report. The engines are the
original scripts, copied in unchanged — the platform only feeds them and
collects what they produce.

**Change review**

| Tool | Takes | Gives back |
|------|-------|-----------|
| **Impact SUTC LLT** | two workbooks, optional blank template | two copies of the workbook with differing cells filled red and green |
| **Impact DD / Data Dictionary + DD Appendix** | two workbooks with the same table layout | one report workbook: rows removed, added and modified, per sheet |
| **Impact SDDD** | two Word documents | one report with the before and after of each changed section |

**Source code and traceability** — merged in from [capgemini_smart_x](https://github.com/rounaaa/capgemini_smart_x)

| Tool | Takes | Gives back |
|------|-------|-----------|
| **C function extractor** | .c/.h files (or a .zip), one function name | an analysis workbook, the Data Dictionary template filled in, or both |
| **Call tree extractor** | .c/.h files (or a .zip) | the Call Tree workbook, one row per call with its condition |
| **Call tree comparison** | two call-tree sheets | a before/after report, removals red, additions green, renames matched |
| **Requirement coverage check** | one Word document | REQ and COV identifiers on two sheets, green where the style is right |

## Run it

```bash
pip install -r requirements.txt
python run.py
```

The browser opens at `http://127.0.0.1:8000`. `--host 0.0.0.0` makes it
reachable from other machines, `--port` moves it, `--debug` reloads on code
changes.

For a shared install, put a real server in front of Flask's development one:

```bash
pip install waitress
waitress-serve --port=8000 --call platform_app.app:create_app
```

### What you need installed

- **Python 3.10+**
- **openpyxl** and **python-docx** — the two engines that need them
- **LibreOffice** (optional, recommended) — lets the tools accept `.doc`,
  `.xls`, `.ods` and the rest by converting them first. Without it, only the
  modern formats work.
- **rapidfuzz** (optional) — faster section matching in the Word tool; it falls
  back to the standard library when missing.

Typefaces load from Google Fonts. On a machine with no internet the page falls
back to system faces and everything still works.

## How it fits together

```
run.py                     launcher
platform_app/
  registry.py              every tool, its inputs and where its runner lives
  jobs.py                  runs on a thread pool: status, progress, artifacts
  app.py                   routes, upload validation, downloads
  tools/                   one adapter per tool — translates, never re-implements
  scripts/                 the original change-review engines, unmodified
  scripts/smart_x/         the capgemini_smart_x engines, GUIs stripped
  templates/, static/      the page
workspace/jobs/<id>/       input/ and output/ for one run
tests/                     fixture builder and an end-to-end check
```

The registry is the spine. A `ToolSpec` declares the tool's inputs as `Field`
objects, and everything else is generated from that: the form controls, the
upload validation, the file-type checks and the job record. The page never
hard-codes a form.

Runs execute on a small thread pool (two at a time by default, because the
LibreOffice conversion step is memory-hungry) and the page polls for progress.
Closing the tab does not stop a run; it will be waiting in **Recent runs**.

### Settings

All optional, all environment variables:

| Variable | Default | What it does |
|---|---|---|
| `AUTOMATION_WORKSPACE` | `./workspace` | where uploads and reports live |
| `AUTOMATION_MAX_UPLOAD_MB` | `250` | largest accepted upload |
| `AUTOMATION_MAX_WORKERS` | `2` | how many runs go at once |
| `AUTOMATION_RETENTION_HOURS` | `24` | when finished runs are deleted, files and all |
| `AUTOMATION_HOST` / `AUTOMATION_PORT` | `127.0.0.1` / `8000` | where it listens |
| `AUTOMATION_APP_NAME` | `AutoGeneration Platform` | name in the masthead, tab and start-up banner |
| `AUTOMATION_APP_TAGLINE` | `Document and spreadsheet change review` | the line under the name |
| `AUTOMATION_ORG_NAME` | `Capgemini Engineering` | logo alt text, and the masthead fallback if no logo file is present |

### Branding

The masthead carries the organisation logo on the left, a hairline, then the
platform name. The logo is `platform_app/static/logo.png` — the Capgemini
Engineering lockup, trimmed to its ink so that the 30 px height it is drawn at
is 30 px of actual logo rather than mostly transparent margin. To swap it,
replace that file, or drop in `logo.svg` / `logo.webp` / `logo.jpg`; the first
name that matches wins and nothing else needs changing.

With no such file present the masthead falls back to `templates/_org_mark.html`,
a drawn approximation of the lockup: the two words set in webfonts plus two SVG
paths for the symbol. It is close, not exact. The two webfonts it needs are
requested only when that fallback is actually rendered.

The logo is a Capgemini trademark, included here for use on a Capgemini tool.
The file came from [Wikimedia Commons](https://commons.wikimedia.org/wiki/File:CapgeminiEngineering_82mm.png).

Colours live as custom properties at the top of `platform_app/static/styles.css`
in two groups. The brand group (`--brand` Capgemini Blue `#0070AD`,
`--brand-vivid` `#12ABDB`, `--brand-deep` `#005182`) drives the chrome and can
be re-themed freely. The diff group (`--before` red, `--after` green) is
functional: it mirrors the fills the engines write into the output files, so
changing it makes the page lie about the reports.

## Adding another tool

Three steps, no changes to the page.

**1. Write the adapter** in `platform_app/tools/my_tool.py`. It receives the
form values, the folder to write into, and a progress reporter:

```python
from ..registry import Artifact, Progress, RunResult
from ..scripts import my_engine

def run(values, output_dir, say: Progress) -> RunResult:
    say("Reading the input ...", 0.2)
    out = my_engine.do_the_thing(values["source_file"], output_dir)
    say("Writing the report ...", 0.8)
    return RunResult(
        artifacts=[Artifact(out, "What this file is")],
        stats=[{"label": "Rows checked", "value": 412, "tone": "accent"}],
    )
```

**2. Drop the engine** into `platform_app/scripts/` as it is. Adapters
translate; they do not rewrite. If the engine has a Tkinter window, leave it —
as long as the import is inside a function or wrapped in `try/except`, the
platform never touches it. A `import tkinter` at module scope does need
removing, though: it runs on import, and a server has no display.

Engines that report progress by printing need nothing special. While a run is
executing, that worker thread's stdout goes into the run's log, so `print`
lines appear in the run panel alongside the adapter's own.

**3. Register it** in `platform_app/registry.py`:

```python
ToolSpec(
    id="my-tool",
    name="My tool",
    tagline="One line on what it does",
    description="A paragraph for the tool page.",
    produces="One .xlsx report",
    runner_path="platform_app.tools.my_tool:run",
    inputs=[
        Field("source_file", "Source workbook", FILE, required=True, accept=".xlsx"),
        Field("threshold", "Threshold", NUMBER, default=10, help="Rows under this are ignored."),
    ],
)
```

Field kinds are `FILE`, `FILES`, `TEXT`, `NUMBER`, `SELECT`, `CHECKBOX` and
`LINES` (a textarea that arrives as a list of strings). `FILES` is a multi-file
drop zone and arrives as a `List[Path]` — dropping onto it a second time adds
to the selection rather than replacing it. A field can carry
`visible_when={"scope": ["section"]}` to appear only for a certain choice
elsewhere on the form; this works on drop zones as well as on option fields.

`action="Extract call tree"` sets the submit button's label. It defaults to
"Run comparison", which is wrong for anything that is not a comparison.

For a tool that reads a body of source code, `platform_app/tools/_sources.py`
flattens whatever was uploaded — loose files, a zip, or both — into one
directory for the engine to walk.

`stats` tones are `accent`, `before` (red), `after` (green) and `neutral`. Red
and green mean removed and added throughout the interface — worth keeping to.

## Checking it still works

```bash
python tests/make_fixtures.py   # writes small sample files
python tests/smoke_test.py      # runs all three tools, checks the downloads
```

Expected output:

```
[ok  ] excel-highlight       0.2s  Changed cells=11  Modified=3  Added=4  Deleted=4 ...
[ok  ] excel-table-diff      0.0s  Sheets compared=2  Rows removed=1  Rows added=1 ...
[ok  ] word-sections         0.2s  Changed sections=1  Added=0  Deleted=0  Modified=2
[ok  ] validation        rejected a .docx: ...
All tools passed.
```

## Before putting it on a shared server

The platform assumes everyone reaching it is allowed to use it. Three things to
add for a multi-team install:

- **Sign-in.** There is none. Anyone who can reach the port can run a tool and,
  with a run id, download that run's files.
- **A persistent job store.** Run history lives in memory, so a restart loses
  the list. Files on disk survive; only the index goes.
- **A disk quota.** Retention is time-based only. Uploads plus reports on a busy
  day add up.
