/* AutoGeneration Platform — client.
 *
 * The form is not written by hand. Each tool ships a list of declared inputs
 * from the server registry, and this file renders controls from it, so a new
 * tool appears here the moment it is registered on the back end.
 */

const TOOLS = JSON.parse(document.getElementById("tools-data").textContent);

const el = {
  toolList: document.getElementById("tool-list"),
  toolName: document.getElementById("tool-name"),
  toolProduces: document.getElementById("tool-produces"),
  toolDescription: document.getElementById("tool-description"),
  form: document.getElementById("run-form"),
  files: document.getElementById("field-files"),
  options: document.getElementById("field-options"),
  formError: document.getElementById("form-error"),
  runButton: document.getElementById("run-button"),
  runPanel: document.getElementById("run-panel"),
  runRuleFill: document.getElementById("run-rule-fill"),
  runStatusWord: document.getElementById("run-status-word"),
  runId: document.getElementById("run-id"),
  runLabel: document.getElementById("run-label"),
  runClock: document.getElementById("run-clock"),
  runClose: document.getElementById("run-close"),
  tape: document.getElementById("tape"),
  outcome: document.getElementById("run-outcome"),
  stats: document.getElementById("stat-row"),
  artifacts: document.getElementById("artifact-row"),
  notes: document.getElementById("note-row"),
  table: document.getElementById("result-table"),
  history: document.getElementById("history"),
  refreshHistory: document.getElementById("refresh-history"),
};

let activeTool = null;
let watching = null;      // { id, lines, timer }

/* ----------------------------------------------------------------- helpers */

function node(tag, className, text) {
  const n = document.createElement(tag);
  if (className) n.className = className;
  if (text !== undefined && text !== null) n.textContent = String(text);
  return n;
}

function bytes(n) {
  if (!n) return "";
  const units = ["B", "KB", "MB", "GB"];
  let i = 0;
  while (n >= 1024 && i < units.length - 1) { n /= 1024; i += 1; }
  return `${n < 10 && i > 0 ? n.toFixed(1) : Math.round(n)} ${units[i]}`;
}

function ago(iso) {
  const seconds = Math.max(0, (Date.now() - new Date(iso).getTime()) / 1000);
  if (seconds < 60) return "just now";
  if (seconds < 3600) return `${Math.floor(seconds / 60)}m ago`;
  if (seconds < 86400) return `${Math.floor(seconds / 3600)}h ago`;
  return `${Math.floor(seconds / 86400)}d ago`;
}

/* ------------------------------------------------------------- tool picker */

function renderToolList() {
  el.toolList.replaceChildren();
  TOOLS.forEach((tool) => {
    const card = node("button", "tool-card");
    card.type = "button";
    card.setAttribute("role", "tab");
    card.setAttribute("aria-selected", "false");
    card.dataset.toolId = tool.id;

    const strip = node("span", "card-strip");
    strip.setAttribute("aria-hidden", "true");
    for (let i = 0; i < 6; i += 1) strip.appendChild(node("i"));

    const body = node("span", "card-body");
    body.appendChild(node("span", "card-name", tool.name));
    body.appendChild(node("span", "card-tagline", tool.tagline));

    card.append(strip, body);
    card.addEventListener("click", () => selectTool(tool.id));
    el.toolList.appendChild(card);
  });
}

function runLabel() {
  return (activeTool && activeTool.action) || "Run comparison";
}

function selectTool(toolId) {
  activeTool = TOOLS.find((t) => t.id === toolId) || TOOLS[0];
  el.toolList.querySelectorAll(".tool-card").forEach((card) => {
    card.setAttribute("aria-selected", String(card.dataset.toolId === activeTool.id));
  });
  el.toolName.textContent = activeTool.name;
  el.toolProduces.textContent = `Produces ${activeTool.produces}`;
  el.toolDescription.textContent = activeTool.description;
  el.runButton.querySelector(".run-button-label").textContent = runLabel();
  buildForm(activeTool);
  clearFormError();
}

/* -------------------------------------------------------- form from schema */

function buildForm(tool) {
  el.form.reset();
  el.files.replaceChildren();
  el.options.replaceChildren();

  tool.inputs.forEach((field) => {
    if (field.kind === "file" || field.kind === "files") {
      el.files.appendChild(buildDrop(field));
    } else {
      el.options.appendChild(buildOption(field));
    }
  });

  applyVisibility();
}

function buildDrop(field) {
  const many = field.kind === "files";
  const empty = many ? "Choose files or drop them here" : "Choose a file or drop it here";

  const label = node("label", "drop");
  label.dataset.field = field.name;
  if (field.visible_when) label.dataset.visibleWhen = JSON.stringify(field.visible_when);

  const input = document.createElement("input");
  input.type = "file";
  input.name = field.name;
  if (many) input.multiple = true;
  if (field.accept) input.accept = field.accept;
  if (field.required) input.dataset.required = "true";

  const role = node("span", "drop-role", field.required ? field.label : `${field.label} · optional`);
  const name = node("span", "drop-name", empty);
  const hint = node("span", "drop-hint", field.help || "");

  label.append(input, role, name, hint);

  input.addEventListener("change", () => {
    label.classList.remove("has-error");
    const picked = Array.from(input.files);
    if (!picked.length) {
      label.classList.remove("has-file");
      name.textContent = empty;
      hint.textContent = field.help || "";
      return;
    }
    label.classList.add("has-file");
    const total = picked.reduce((sum, f) => sum + f.size, 0);
    if (picked.length === 1) {
      name.textContent = picked[0].name;
      hint.textContent = bytes(total);
    } else {
      // Too many names to list, so lead with the count and sample a few.
      name.textContent = `${picked.length} files`;
      const sample = picked.slice(0, 3).map((f) => f.name).join(", ");
      hint.textContent = `${bytes(total)} · ${sample}${picked.length > 3 ? ", …" : ""}`;
    }
  });

  ["dragenter", "dragover"].forEach((evt) =>
    label.addEventListener(evt, (e) => { e.preventDefault(); label.classList.add("is-over"); }));
  ["dragleave", "drop"].forEach((evt) =>
    label.addEventListener(evt, () => label.classList.remove("is-over")));
  label.addEventListener("drop", (e) => {
    e.preventDefault();
    if (!e.dataTransfer.files.length) return;
    if (many) {
      // Dropping onto a multi-file zone adds to what is already there.
      const merged = new DataTransfer();
      Array.from(input.files).forEach((f) => merged.items.add(f));
      Array.from(e.dataTransfer.files).forEach((f) => merged.items.add(f));
      input.files = merged.files;
    } else {
      input.files = e.dataTransfer.files;
    }
    input.dispatchEvent(new Event("change"));
  });

  return label;
}

function buildOption(field) {
  const wrap = node("div", field.kind === "lines" ? "field-wide" : "");
  wrap.dataset.field = field.name;

  const label = node("label", "field-label");
  label.setAttribute("for", `f-${field.name}`);
  label.textContent = field.label;
  if (!field.required) label.appendChild(node("span", "optional", "  optional"));
  wrap.appendChild(label);

  let control;
  if (field.kind === "select") {
    control = document.createElement("select");
    field.options.forEach((opt) => {
      const o = document.createElement("option");
      o.value = opt.value;
      o.textContent = opt.label;
      if (opt.value === field.default) o.selected = true;
      control.appendChild(o);
    });
    control.addEventListener("change", applyVisibility);
  } else if (field.kind === "lines") {
    control = document.createElement("textarea");
    control.rows = 4;
    control.placeholder = field.placeholder || "";
  } else {
    control = document.createElement("input");
    control.type = field.kind === "number" ? "number" : "text";
    if (field.kind === "number") control.min = "0";
    control.placeholder = field.placeholder || "";
    if (field.default !== null && field.default !== undefined) control.value = field.default;
  }

  control.id = `f-${field.name}`;
  control.name = field.name;
  wrap.appendChild(control);

  if (field.help) wrap.appendChild(node("p", "field-help", field.help));
  if (field.visible_when) wrap.dataset.visibleWhen = JSON.stringify(field.visible_when);

  return wrap;
}

/* Fields can declare that they only matter for a certain choice elsewhere
   (the section list is meaningless unless the scope is "specific sections"). */
function applyVisibility() {
  // Both rails: a rule can sit on an option field or on a file drop zone.
  el.form.querySelectorAll("[data-visible-when]").forEach((wrap) => {
    const rule = JSON.parse(wrap.dataset.visibleWhen);
    const visible = Object.entries(rule).every(([name, allowed]) => {
      const source = el.form.elements[name];
      return source && allowed.includes(source.value);
    });
    wrap.hidden = !visible;
  });
}

/* -------------------------------------------------------------- submission */

function clearFormError() {
  el.formError.hidden = true;
  el.formError.textContent = "";
  el.form.querySelectorAll(".field-error").forEach((n) => n.remove());
  el.form.querySelectorAll(".drop.has-error").forEach((n) => n.classList.remove("has-error"));
}

function showFieldErrors(fields) {
  Object.entries(fields || {}).forEach(([name, message]) => {
    const drop = el.files.querySelector(`.drop[data-field="${name}"]`);
    if (drop) {
      drop.classList.add("has-error");
      drop.appendChild(node("span", "field-error", message));
      return;
    }
    const wrap = el.options.querySelector(`[data-field="${name}"]`);
    if (wrap) wrap.appendChild(node("p", "field-error", message));
  });
}

el.form.addEventListener("submit", async (event) => {
  event.preventDefault();
  clearFormError();

  const missing = [...el.files.querySelectorAll('input[data-required="true"]')]
    .filter((input) => !input.files.length);
  if (missing.length) {
    showFieldErrors(Object.fromEntries(missing.map((i) => [i.name, "Pick a file."])));
    el.formError.textContent = "Two files are needed before a comparison can run.";
    el.formError.hidden = false;
    return;
  }

  const data = new FormData(el.form);
  data.set("tool_id", activeTool.id);
  // Hidden fields carry no meaning for this run.
  el.options.querySelectorAll("[data-visible-when]").forEach((wrap) => {
    if (wrap.hidden) data.delete(wrap.dataset.field);
  });

  el.runButton.disabled = true;
  el.runButton.querySelector(".run-button-label").textContent = "Uploading…";

  try {
    const response = await fetch("/api/jobs", { method: "POST", body: data });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      el.formError.textContent = payload.error || "The run could not be started.";
      el.formError.hidden = false;
      showFieldErrors(payload.fields);
      return;
    }
    watch(payload.id);
    loadHistory();
  } catch (error) {
    el.formError.textContent = "The server did not answer. Check that it is still running.";
    el.formError.hidden = false;
  } finally {
    el.runButton.disabled = false;
    el.runButton.querySelector(".run-button-label").textContent = runLabel();
  }
});

/* ------------------------------------------------------------- run watcher */

function watch(jobId) {
  stopWatching();
  watching = { id: jobId, lines: 0, timer: null };

  el.runPanel.hidden = false;
  el.runPanel.className = "panel run-panel";
  el.tape.replaceChildren();
  el.outcome.hidden = true;
  el.runRuleFill.style.width = "2%";
  el.runId.textContent = jobId;
  el.runLabel.textContent = "Preparing…";
  el.runStatusWord.textContent = "Queued";
  el.runClock.textContent = "0.0s";
  el.runPanel.scrollIntoView({ behavior: "smooth", block: "nearest" });

  poll();
  watching.timer = setInterval(poll, 800);
}

function stopWatching() {
  if (!watching) return;
  clearInterval(watching.timer);
  watching = null;
}

async function poll() {
  if (!watching) return;
  const { id, lines } = watching;
  let job;
  try {
    const response = await fetch(`/api/jobs/${id}?since=${lines}`);
    if (!response.ok) throw new Error("gone");
    job = await response.json();
  } catch (error) {
    appendLines([{ at: "", text: "Lost contact with the server.", level: "error" }]);
    stopWatching();
    return;
  }

  if (!watching || watching.id !== id) return;
  watching.lines = job.log_count;

  appendLines(job.log);
  el.runLabel.textContent = job.label;
  el.runStatusWord.textContent = { queued: "Queued", running: "Running", done: "Finished", failed: "Failed" }[job.status];
  el.runRuleFill.style.width = `${Math.max(2, job.progress * 100)}%`;
  if (job.duration !== null && job.duration !== undefined) {
    el.runClock.textContent = `${job.duration.toFixed(1)}s`;
  }

  if (job.status === "done" || job.status === "failed") {
    el.runPanel.classList.add(job.status === "done" ? "is-done" : "is-failed");
    el.runRuleFill.style.width = "100%";
    renderOutcome(job);
    stopWatching();
    loadHistory();
  }
}

function appendLines(lines) {
  if (!lines || !lines.length) return;
  const atBottom = el.tape.scrollHeight - el.tape.scrollTop - el.tape.clientHeight < 40;
  lines.forEach((line) => {
    const row = node("div", `tape-line level-${line.level}`);
    const time = document.createElement("time");
    time.textContent = line.at;
    row.append(time, node("span", null, line.text));
    el.tape.appendChild(row);
  });
  if (atBottom) el.tape.scrollTop = el.tape.scrollHeight;
}

function renderOutcome(job) {
  el.stats.replaceChildren();
  el.artifacts.replaceChildren();
  el.notes.replaceChildren();
  el.table.replaceChildren();

  if (job.status === "failed") {
    el.outcome.hidden = false;
    const note = node("div", "note", job.error || "The run failed. The log above says where.");
    note.style.borderLeftColor = "var(--before)";
    note.style.background = "var(--before-soft)";
    el.notes.appendChild(note);
    return;
  }

  (job.stats || []).forEach((stat) => {
    const box = node("div", `stat tone-${stat.tone || "neutral"}`);
    box.appendChild(node("b", null, stat.value));
    box.appendChild(node("span", null, stat.label));
    el.stats.appendChild(box);
  });

  (job.artifacts || []).forEach((artifact) => {
    const link = document.createElement("a");
    link.className = `artifact role-${artifact.role}`;
    link.href = artifact.url;
    link.download = artifact.name;

    const body = node("span");
    body.appendChild(node("span", "artifact-name", artifact.name));
    body.appendChild(node("span", "artifact-meta", `${artifact.label} · ${bytes(artifact.size)}`));

    link.append(node("span", "artifact-mark"), body, node("span", "artifact-get", "Download"));
    el.artifacts.appendChild(link);
  });

  (job.notes || []).forEach((text) => el.notes.appendChild(node("div", "note", text)));

  if (job.table && job.table.rows.length) {
    el.table.appendChild(node("h3", null, job.table.title));
    const scroll = node("div", "table-scroll");
    const table = document.createElement("table");

    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    job.table.columns.forEach((col) => headRow.appendChild(node("th", null, col)));
    thead.appendChild(headRow);

    const tbody = document.createElement("tbody");
    job.table.rows.forEach((row) => {
      const tr = node("tr", `tone-${row.tone || "neutral"}`);
      row.cells.forEach((cell) => tr.appendChild(node("td", null, cell)));
      tbody.appendChild(tr);
    });

    table.append(thead, tbody);
    scroll.appendChild(table);
    el.table.appendChild(scroll);
    if (job.table.truncated) {
      el.table.appendChild(node("p", "table-more", "Showing the first 400 changes. The workbooks hold every one."));
    }
  }

  el.outcome.hidden = false;
}

el.runClose.addEventListener("click", () => {
  stopWatching();
  el.runPanel.hidden = true;
});

/* ----------------------------------------------------------------- history */

async function loadHistory() {
  let jobs = [];
  try {
    const response = await fetch("/api/jobs?limit=20");
    jobs = await response.json();
  } catch (error) {
    return;
  }

  el.history.replaceChildren();
  if (!jobs.length) {
    el.history.appendChild(node("li", "history-empty", "Nothing has run yet."));
    return;
  }

  jobs.forEach((job) => {
    const item = node("li");
    const button = node("button", "history-item");
    button.type = "button";

    const label = node("span");
    label.appendChild(node("span", "history-label", job.label));
    label.appendChild(node("span", "history-tool", job.tool_name));

    button.append(
      node("span", `dot ${job.status}`),
      label,
      node("span", "history-time", ago(job.created_at))
    );
    button.addEventListener("click", () => watch(job.id));
    item.appendChild(button);
    el.history.appendChild(item);
  });
}

el.refreshHistory.addEventListener("click", loadHistory);

/* -------------------------------------------------------------------- boot */

renderToolList();
selectTool(TOOLS[0].id);
loadHistory();
setInterval(loadHistory, 20000);
