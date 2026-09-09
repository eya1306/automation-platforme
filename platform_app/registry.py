"""Tool registry.

Every automation on the platform is described here as a :class:`ToolSpec`.
The spec is the single source of truth: the web form, the upload validation
and the job runner are all generated from it, so adding a fourth script is a
matter of writing one adapter and appending one entry to ``TOOLS``.

Nothing in this module knows about Flask, and nothing here imports the heavy
engines -- adapters are resolved lazily so the app starts instantly.
"""

from __future__ import annotations

import importlib
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

# A progress reporter handed to every runner: report(message, fraction 0..1).
Progress = Callable[[str, float], None]


# --------------------------------------------------------------------------- #
# Input fields
# --------------------------------------------------------------------------- #

FILE = "file"
FILES = "files"          # multi-file upload -> List[Path]
TEXT = "text"
LINES = "lines"          # textarea, one value per line -> List[str]
NUMBER = "number"
SELECT = "select"
CHECKBOX = "checkbox"


@dataclass
class Field:
    """One control on a tool's form."""

    name: str
    label: str
    kind: str = TEXT
    required: bool = False
    help: str = ""
    default: Any = None
    placeholder: str = ""
    accept: str = ""                       # file inputs: allowed extensions
    options: List[Dict[str, str]] = field(default_factory=list)  # select: value/label
    # Only show this field when another field holds one of these values.
    visible_when: Optional[Dict[str, List[str]]] = None

    def as_json(self) -> dict:
        return {
            "name": self.name,
            "label": self.label,
            "kind": self.kind,
            "required": self.required,
            "help": self.help,
            "default": self.default,
            "placeholder": self.placeholder,
            "accept": self.accept,
            "options": self.options,
            "visible_when": self.visible_when,
        }


# --------------------------------------------------------------------------- #
# Results
# --------------------------------------------------------------------------- #


@dataclass
class Artifact:
    """A file the run produced, offered to the user as a download."""

    path: Path
    label: str
    role: str = "report"       # report | before | after — drives the colour chip


@dataclass
class RunResult:
    """What a runner hands back when it finishes."""

    artifacts: List[Artifact] = field(default_factory=list)
    stats: List[Dict[str, Any]] = field(default_factory=list)   # {label, value, tone}
    table: Optional[Dict[str, Any]] = None                      # {columns, rows}
    notes: List[str] = field(default_factory=list)


# --------------------------------------------------------------------------- #
# Tools
# --------------------------------------------------------------------------- #


@dataclass
class ToolSpec:
    id: str
    name: str
    tagline: str
    description: str
    produces: str
    runner_path: str                      # "module:function", imported on first run
    inputs: List[Field] = field(default_factory=list)
    accent: str = "indigo"
    action: str = "Run comparison"        # label on the submit button

    _runner: Optional[Callable[..., RunResult]] = None

    def runner(self) -> Callable[[Dict[str, Any], Path, Progress], RunResult]:
        if self._runner is None:
            module_name, func_name = self.runner_path.split(":")
            module = importlib.import_module(module_name)
            self._runner = getattr(module, func_name)
        return self._runner

    def file_fields(self) -> List[Field]:
        return [f for f in self.inputs if f.kind in (FILE, FILES)]

    def as_json(self) -> dict:
        return {
            "id": self.id,
            "name": self.name,
            "tagline": self.tagline,
            "description": self.description,
            "produces": self.produces,
            "accent": self.accent,
            "action": self.action,
            "inputs": [f.as_json() for f in self.inputs],
        }


SPREADSHEET_ACCEPT = ".xlsx,.xlsm,.xltx,.xltm,.xls,.xlt,.xlsb,.ods,.csv"
WORD_ACCEPT = ".doc,.docx,.docm"
WORD_STYLED_ACCEPT = ".doc,.docx,.docm,.dot,.dotm"

# The two C tools take either loose sources or a zip of a source tree.
C_SOURCE_ACCEPT = ".c,.h,.cpp,.hpp,.cc,.s,.asm,.zip"


TOOLS: List[ToolSpec] = [
    ToolSpec(
        id="excel-highlight",
        name="Impact SUTC LLT",
        tagline="Two highlighted copies of the workbook, cell by cell",
        description=(
            "Compares two versions of the same workbook and writes back two copies of it "
            "with the differing cells filled in — red on the old side, green on the new — "
            "and a note on each one saying what changed. Rows are aligned first, so an "
            "inserted row is reported as one insertion instead of making everything below "
            "it look modified. Layout, styles, formulas and images survive untouched."
        ),
        produces="Two .xlsx workbooks (old and new, highlighted)",
        runner_path="platform_app.tools.excel_highlight:run",
        inputs=[
            Field("old_file", "Original workbook", FILE, required=True,
                  accept=SPREADSHEET_ACCEPT,
                  help="The version you are comparing from."),
            Field("new_file", "Modified workbook", FILE, required=True,
                  accept=SPREADSHEET_ACCEPT,
                  help="The version you are comparing to."),
            Field("template_file", "Blank template", FILE, required=False,
                  accept=SPREADSHEET_ACCEPT,
                  help="Optional. Supply the empty form both files were built from and "
                       "its own headings and labels are left out of the comparison."),
        ],
    ),
    ToolSpec(
        id="excel-table-diff",
        name="Impact DD / Data Dictionary + DD Appendix",
        tagline="Added, removed and modified rows, matched on a key column",
        description=(
            "For workbooks that hold flat tables. Rows are matched between the two files "
            "on a key column, then split into a before sheet (removed and modified rows, "
            "in red) and an after sheet (added and modified rows, in green) per source "
            "sheet, with a summary sheet counting each. Unchanged rows are left out."
        ),
        produces="One .xlsx report workbook",
        runner_path="platform_app.tools.excel_table_diff:run",
        inputs=[
            Field("before_file", "Before workbook", FILE, required=True,
                  accept=".xlsx,.xlsm", help="Both files need the same sheet and column layout."),
            Field("after_file", "After workbook", FILE, required=True, accept=".xlsx,.xlsm"),
            Field("key_column", "Key column", NUMBER, default=0,
                  help="Which column identifies a row, counting from 0. Column A is 0."),
            Field("output_name", "Report file name", TEXT, default="Change_Report",
                  placeholder="Change_Report"),
        ],
    ),
    ToolSpec(
        id="word-sections",
        name="Impact SDDD",
        tagline="Before and after of every changed section, formatting intact",
        description=(
            "Compares two Word documents and builds one report holding only the sections "
            "that changed, each shown as a Before block and an After block with the changed "
            "words highlighted. Sections are deep-copied out of their source documents with "
            "their styles, numbering and images, so a copied section looks exactly as it did."
        ),
        produces="One .docx comparison report",
        runner_path="platform_app.tools.word_sections:run",
        inputs=[
            Field("original_file", "Original document", FILE, required=True, accept=WORD_ACCEPT),
            Field("modified_file", "Modified document", FILE, required=True, accept=WORD_ACCEPT),
            Field("scope", "What to compare", SELECT, default="auto",
                  options=[
                      {"value": "auto", "label": "Requirements chapter only"},
                      {"value": "full", "label": "The whole document"},
                      {"value": "section", "label": "Specific sections"},
                  ],
                  help="Requirements mode runs from the first Heading 1 containing "
                       "\"Requirements\" through to the next Heading 1."),
            Field("target_sections", "Sections to review", LINES,
                  visible_when={"scope": ["section"]},
                  placeholder="3.2 Interfaces\n4.1 Performance",
                  help="One section heading per line. They appear in the report in the "
                       "order you type them."),
            Field("report_title", "Report title", TEXT,
                  placeholder="Impact SDDD CSC CR 1234",
                  help="Optional. Printed as the first heading of the report."),
            Field("output_name", "Report file name", TEXT, default="Comparison_Report",
                  placeholder="Comparison_Report"),
        ],
    ),
    ToolSpec(
        id="c-function-analysis",
        action="Extract function",
        name="C function extractor",
        tagline="One function pulled apart into a Data Dictionary",
        description=(
            "Reads a body of C source, finds one named function in it and writes out "
            "everything it declares and touches: parameters, local variables, the "
            "globals it reads, the conditions it branches on, the constants and types "
            "it uses and the functions it calls. The result is either a multi-sheet "
            "analysis workbook or the Data Dictionary template filled in, or both."
        ),
        produces="One .xlsx analysis workbook, or a filled Data Dictionary",
        runner_path="platform_app.tools.c_function_analysis:run",
        inputs=[
            Field("source_files", "C source files", FILES, required=True,
                  accept=C_SOURCE_ACCEPT,
                  help="The .c and .h files to search. Drop a .zip of the source tree "
                       "instead if there are many of them."),
            Field("function_name", "Function name", TEXT, required=True,
                  placeholder="Cop_MainFunction",
                  help="Exactly as it is spelled in the source, without brackets."),
            Field("export_mode", "What to produce", SELECT, default="workbook",
                  options=[
                      {"value": "workbook", "label": "Analysis workbook"},
                      {"value": "template", "label": "Data Dictionary template"},
                      {"value": "both", "label": "Both"},
                  ]),
            Field("template_file", "Data Dictionary template", FILE, required=False,
                  accept=".xlsx,.xlsm",
                  visible_when={"export_mode": ["template", "both"]},
                  help="Optional. Leave empty to use the template that shipped with "
                       "the tool."),
            Field("output_name", "Report file name", TEXT, default="",
                  placeholder="Analysis_<function>"),
        ],
    ),
    ToolSpec(
        id="call-tree-extract",
        action="Extract call tree",
        name="Call tree extractor",
        tagline="Who calls what, walked out of the C sources",
        description=(
            "Walks a tree of C sources and records every call it finds as one row: the "
            "calling CSC, CSU and function, the called CSC, CSU and function, and the "
            "condition the call sits under. The rows are written into the Call Tree "
            "sheet of the standard workbook, ready to be compared against another "
            "extraction once the code has moved on."
        ),
        produces="One .xlsx workbook with a filled Call Tree sheet",
        runner_path="platform_app.tools.call_tree_extract:run",
        inputs=[
            Field("source_files", "C source files", FILES, required=True,
                  accept=C_SOURCE_ACCEPT,
                  help="Loose .c and .h files, or a .zip of the whole source tree."),
            Field("template_file", "Call Tree workbook", FILE, required=False,
                  accept=".xlsx,.xlsm",
                  help="Optional. Leave empty to start from the empty Call Tree "
                       "structure that shipped with the tool."),
            Field("sheet_name", "Sheet to fill", TEXT, default="Call Tree",
                  placeholder="Call Tree",
                  help="Use \"Call Tree (Apres modif)\" for the after side, so both "
                       "extractions can live in one workbook."),
            Field("output_name", "Report file name", TEXT, default="Call_Tree",
                  placeholder="Call_Tree"),
        ],
    ),
    ToolSpec(
        id="call-tree-compare",
        name="Call tree comparison",
        tagline="Two call trees, matched row by row with renames spotted",
        description=(
            "Compares two call-tree sheets and writes a report holding both sides in "
            "full: a before sheet with the rows that went away in red, and an after "
            "sheet with the rows that arrived in green. Rows that are merely similar "
            "rather than identical are treated as modified rather than as one deletion "
            "plus one addition, so a renamed function does not read as churn."
        ),
        produces="One .xlsx report with a before and an after sheet",
        runner_path="platform_app.tools.call_tree_compare:run",
        inputs=[
            Field("before_file", "Before workbook", FILE, required=True,
                  accept=".xlsx,.xlsm",
                  help="If both sheets live in this one workbook, leave the after "
                       "workbook empty."),
            Field("after_file", "After workbook", FILE, required=False,
                  accept=".xlsx,.xlsm",
                  help="Optional. Only needed when the two call trees are in "
                       "separate files."),
            Field("before_sheet", "Before sheet", TEXT, default="Call Tree",
                  placeholder="Call Tree"),
            Field("after_sheet", "After sheet", TEXT, default="Call Tree (Apres modif)",
                  placeholder="Call Tree (Apres modif)"),
            Field("similarity", "Rename threshold", NUMBER, default=85,
                  help="How alike two rows must be, as a percentage, before one counts "
                       "as a modified version of the other rather than a separate "
                       "deletion and addition."),
            Field("output_name", "Report file name", TEXT, default="Call_Tree_Comparison",
                  placeholder="Call_Tree_Comparison"),
        ],
    ),
    ToolSpec(
        id="req-coverage-check",
        action="Check document",
        name="Requirement coverage check",
        tagline="Every REQ and COV identifier, and whether its style is right",
        description=(
            "Scans the Requirements sections of a Word document for REQ-SDDD "
            "identifiers and COV.REQ coverage identifiers, and checks that each one "
            "carries the paragraph style it is supposed to. The report lists them on "
            "two sheets, green where the style is right and red where it is not, so "
            "the ones that will be missed by downstream traceability tooling stand out."
        ),
        produces="One .xlsx report, LLR and HLR identifiers on two sheets",
        runner_path="platform_app.tools.req_coverage_check:run",
        inputs=[
            Field("document", "Word document", FILE, required=True,
                  accept=WORD_STYLED_ACCEPT,
                  help="Macro-enabled documents (.docm, .dotm) are handled too."),
            Field("output_name", "Report file name", TEXT, default="Requirement_Coverage",
                  placeholder="Requirement_Coverage"),
        ],
    ),
]

TOOLS_BY_ID: Dict[str, ToolSpec] = {t.id: t for t in TOOLS}


def get_tool(tool_id: str) -> Optional[ToolSpec]:
    return TOOLS_BY_ID.get(tool_id)
