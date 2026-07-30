/* global Office, PowerPoint, console */

// ---------------------------------------------------------------------------
// Gantt Chart Builder - task pane logic
// Renders a Gantt chart as native PowerPoint shapes (editable, ThinkCell-style).
//
// Features:
//  - Tasks with MULTIPLE bars / milestones (segments), each with own label+color
//  - Group headings that bundle several tasks
//  - Reorderable rows (up / down)
//  - Unique task names (duplicate -> red)
//  - End-before-start validation (red)
//  - Multiple stacked time axes (years / quarters / months / weeks)
//  - Vertical grid delimiters (never diagonal)
//  - Dashed "today" marker
//  - Real routed dependency arrows through the white space (not through bars)
//  - Fit to slide size and group everything
//  - Per segment date display + net workdays
// ---------------------------------------------------------------------------

const PALETTE = ["#4472C4", "#ED7D31", "#70AD47", "#FFC000", "#5B9BD5", "#A5A5A5", "#264478", "#9E480E"];

// Slide is 960 x 540 pt for a 16:9 deck.
const SLIDE = { w: 960, h: 540 };
const MARGIN = { left: 30, right: 30, top: 70, bottom: 30 };
const LABEL_WIDTH = 175;
const BAND_HEIGHT = 18;          // height of one time-axis band
const MIN_ROW_HEIGHT = 22;
const MAX_ROW_HEIGHT = 40;

let rowCounter = 0;
let segCounter = 0;
let linkCounter = 0;
let dateClipboard = null;        // for copy/paste of dates

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("addTask").onclick = () => { addTaskRow(); afterModelChange(); };
    document.getElementById("addGroup").onclick = () => { addGroupRow(); afterModelChange(); };
    document.getElementById("addLink").onclick = () => { addLinkRow(); renderPreview(); };
    document.getElementById("insert").onclick = insertGantt;
    document.getElementById("sample").onclick = () => { loadSample(); afterModelChange(); };
    document.getElementById("syncSlide").onclick = syncFromSlide;
    document.getElementById("startDate").onchange = renderPreview;
    document.getElementById("endDate").onchange = renderPreview;
    document.querySelectorAll(".scaleOpt, #optDates, #optWorkdays, #optToday, #optFit")
      .forEach((el) => { el.onchange = renderPreview; });
    initDefaults();
    afterModelChange();
  }
});

function initDefaults() {
  const today = new Date();
  const inThree = new Date(today);
  inThree.setMonth(inThree.getMonth() + 3);
  document.getElementById("startDate").value = toInputDate(today);
  document.getElementById("endDate").value = toInputDate(inThree);
  loadSample();
}

function afterModelChange() {
  refreshLinkOptions();
  validateNames();
  renderPreview();
}

function toInputDate(d) {
  const x = new Date(d.getTime() - d.getTimezoneOffset() * 60000);
  return x.toISOString().slice(0, 10);
}

function selectedScales() {
  const order = ["year", "quarter", "month", "week"]; // coarse -> fine
  const chosen = Array.from(document.querySelectorAll(".scaleOpt:checked")).map((c) => c.value);
  const sorted = order.filter((o) => chosen.includes(o));
  return sorted.length ? sorted : ["month"];
}

// ===========================================================================
// Row management (task rows + group rows share one ordered list)
// ===========================================================================

function addGroupRow(group = {}) {
  rowCounter += 1;
  const el = document.createElement("div");
  el.className = "group-item";
  el.dataset.id = rowCounter;
  el.dataset.kind = "group";
  el.innerHTML = `
    <div class="reorder"><button class="up" type="button" title="Nach oben">▲</button><button class="down" type="button" title="Nach unten">▼</button></div>
    <input class="group-name" type="text" placeholder="Überschrift" value="${escapeAttr(group.name || "")}" />
    <button class="task-remove" type="button" title="Entfernen">&times;</button>`;
  wireRow(el);
  el.querySelector(".group-name").oninput = afterModelChange;
  document.getElementById("taskList").appendChild(el);
  return el;
}

function addTaskRow(task = {}) {
  rowCounter += 1;
  const el = document.createElement("div");
  el.className = "task-item";
  el.dataset.id = rowCounter;
  el.dataset.kind = "task";
  el.innerHTML = `
    <div class="task-head">
      <div class="reorder"><button class="up" type="button" title="Nach oben">▲</button><button class="down" type="button" title="Nach unten">▼</button></div>
      <input class="task-name" type="text" placeholder="Task-Name (eindeutig)" value="${escapeAttr(task.name || "")}" />
      <button class="task-remove" type="button" title="Entfernen">&times;</button>
    </div>
    <div class="segments"></div>
    <button class="add-segment" type="button">+ Balken / Milestone</button>`;
  wireRow(el);
  el.querySelector(".task-name").oninput = afterModelChange;
  el.querySelector(".add-segment").onclick = () => { addSegment(el); renderPreview(); };
  document.getElementById("taskList").appendChild(el);

  const segs = task.segments && task.segments.length ? task.segments : [defaultSegment()];
  segs.forEach((s) => addSegment(el, s));
  return el;
}

function wireRow(el) {
  el.querySelector(".task-remove").onclick = () => { el.remove(); afterModelChange(); };
  el.querySelector(".up").onclick = () => { moveRow(el, -1); };
  el.querySelector(".down").onclick = () => { moveRow(el, 1); };
}

function moveRow(el, dir) {
  if (dir < 0 && el.previousElementSibling) {
    el.parentNode.insertBefore(el, el.previousElementSibling);
  } else if (dir > 0 && el.nextElementSibling) {
    el.parentNode.insertBefore(el.nextElementSibling, el);
  }
  afterModelChange();
}

function defaultSegment() {
  // Requirement 7: start = today, end = +1 week
  const start = new Date();
  const end = new Date(start);
  end.setDate(end.getDate() + 7);
  return { type: "bar", start: toInputDate(start), end: toInputDate(end), label: "", color: "", showDate: true };
}

function addSegment(taskEl, seg = {}) {
  segCounter += 1;
  const container = taskEl.querySelector(".segments");
  const idx = container.children.length;
  const s = { ...defaultSegment(), ...seg };
  if (!s.color) s.color = PALETTE[(segCounter + idx) % PALETTE.length];

  const row = document.createElement("div");
  row.className = "segment";
  row.dataset.id = segCounter;
  row.innerHTML = `
    <select class="seg-type" title="Typ">
      <option value="bar" ${s.type === "bar" ? "selected" : ""}>Balken</option>
      <option value="milestone" ${s.type === "milestone" ? "selected" : ""}>Milestone</option>
    </select>
    <input class="seg-start" type="date" value="${s.start || ""}" title="Start" />
    <input class="seg-end" type="date" value="${s.end || ""}" title="Ende (nur Balken)" />
    <input class="seg-label" type="text" placeholder="Beschriftung (optional)" value="${escapeAttr(s.label || "")}" />
    <div class="seg-tools">
      <input class="seg-color" type="color" value="${s.color}" title="Farbe" />
      <label class="seg-date-toggle" title="Datum im Chart anzeigen"><input class="seg-showdate" type="checkbox" ${s.showDate ? "checked" : ""} />📅</label>
      <button class="icon-btn seg-copy" type="button" title="Datum kopieren">⧉</button>
      <button class="icon-btn seg-paste" type="button" title="Datum einfügen">⇩</button>
      <button class="icon-btn seg-remove" type="button" title="Segment entfernen">×</button>
    </div>`;

  const typeSel = row.querySelector(".seg-type");
  const startInp = row.querySelector(".seg-start");
  const endInp = row.querySelector(".seg-end");
  const applyType = () => {
    const isMs = typeSel.value === "milestone";
    endInp.style.visibility = isMs ? "hidden" : "visible";
  };
  applyType();
  typeSel.onchange = () => { applyType(); validateSegment(row); renderPreview(); };

  [startInp, endInp].forEach((inp) => {
    inp.oninput = () => { validateSegment(row); renderPreview(); };
  });
  row.querySelector(".seg-label").oninput = renderPreview;
  row.querySelector(".seg-color").oninput = renderPreview;
  row.querySelector(".seg-showdate").onchange = renderPreview;

  row.querySelector(".seg-copy").onclick = () => {
    dateClipboard = { start: startInp.value, end: endInp.value };
    setStatus("Datum kopiert.", "success");
  };
  row.querySelector(".seg-paste").onclick = () => {
    if (!dateClipboard) { setStatus("Zwischenablage leer – erst ⧉ nutzen.", "error"); return; }
    if (dateClipboard.start) startInp.value = dateClipboard.start;
    if (dateClipboard.end) endInp.value = dateClipboard.end;
    validateSegment(row); renderPreview();
  };
  row.querySelector(".seg-remove").onclick = () => {
    if (container.children.length > 1) { row.remove(); renderPreview(); }
    else setStatus("Eine Task braucht mindestens ein Segment.", "error");
  };

  container.appendChild(row);
  validateSegment(row);
}

// Requirement 13: end must not be numerically before start
function validateSegment(row) {
  const type = row.querySelector(".seg-type").value;
  const startInp = row.querySelector(".seg-start");
  const endInp = row.querySelector(".seg-end");
  startInp.classList.remove("invalid");
  endInp.classList.remove("invalid");
  if (type === "bar" && startInp.value && endInp.value) {
    if (new Date(endInp.value) < new Date(startInp.value)) {
      endInp.classList.add("invalid");
      return false;
    }
  }
  return true;
}

// Requirement 9: task names must be unique
function validateNames() {
  const inputs = Array.from(document.querySelectorAll(".task-name"));
  const counts = {};
  inputs.forEach((i) => {
    const v = i.value.trim().toLowerCase();
    if (v) counts[v] = (counts[v] || 0) + 1;
  });
  inputs.forEach((i) => {
    const v = i.value.trim().toLowerCase();
    i.classList.toggle("invalid", !!v && counts[v] > 1);
  });
}

// ===========================================================================
// Model reading
// ===========================================================================

// Returns ordered rows: {kind:'group', name} | {kind:'task', name, segments:[...]}
function readModel() {
  const rows = [];
  Array.from(document.getElementById("taskList").children).forEach((el) => {
    if (el.dataset.kind === "group") {
      const name = el.querySelector(".group-name").value.trim();
      rows.push({ kind: "group", name });
    } else {
      const name = el.querySelector(".task-name").value.trim();
      const segments = [];
      el.querySelectorAll(".segment").forEach((sr) => {
        const type = sr.querySelector(".seg-type").value;
        const startV = sr.querySelector(".seg-start").value;
        const endV = sr.querySelector(".seg-end").value;
        if (!startV) return;
        if (type === "bar" && !endV) return;
        segments.push({
          type,
          start: new Date(startV),
          end: type === "milestone" ? new Date(startV) : new Date(endV),
          label: sr.querySelector(".seg-label").value.trim(),
          color: sr.querySelector(".seg-color").value,
          showDate: sr.querySelector(".seg-showdate").checked
        });
      });
      rows.push({ kind: "task", name, segments });
    }
  });
  return rows;
}

// Tasks that actually have drawable segments.
function readTasks() {
  return readModel().filter((r) => r.kind === "task" && r.segments.length);
}

function taskNames() {
  return readModel().filter((r) => r.kind === "task" && r.name).map((r) => r.name);
}

// ===========================================================================
// Links (dependencies)
// ===========================================================================

function addLinkRow(link = {}) {
  linkCounter += 1;
  const el = document.createElement("div");
  el.className = "link-item";
  el.dataset.id = linkCounter;
  el.innerHTML = `
    <div class="link-row">
      <select class="link-from" title="Von"></select>
      <span class="link-arrow">&rarr;</span>
      <select class="link-to" title="Nach"></select>
    </div>
    <div class="task-meta">
      <select class="link-style" title="Verbindungsart">
        <option value="finish-start" selected>Ende &rarr; Start</option>
        <option value="start-start">Start &rarr; Start</option>
        <option value="finish-finish">Ende &rarr; Ende</option>
      </select>
      <input class="link-color" type="color" value="${link.color || "#c00000"}" />
    </div>
    <button class="task-remove" type="button" title="Entfernen">&times;</button>`;
  document.getElementById("linkList").appendChild(el);
  populateLinkSelects(el, link);
  el.querySelectorAll("select, input").forEach((c) => { c.onchange = renderPreview; });
  el.querySelector(".task-remove").onclick = () => { el.remove(); renderPreview(); };
}

function populateLinkSelects(wrapper, link = {}) {
  const names = taskNames();
  ["from", "to"].forEach((role) => {
    const sel = wrapper.querySelector(`.link-${role}`);
    const current = sel.value || link[role];
    sel.innerHTML = names.map((n) => `<option value="${escapeAttr(n)}">${escapeHtml(n)}</option>`).join("");
    if (current && names.includes(current)) sel.value = current;
    else if (role === "to" && names.length > 1) sel.selectedIndex = 1;
  });
}

function refreshLinkOptions() {
  document.querySelectorAll(".link-item").forEach((w) => populateLinkSelects(w));
}

function readLinks() {
  const links = [];
  document.querySelectorAll(".link-item").forEach((el) => {
    const from = el.querySelector(".link-from").value;
    const to = el.querySelector(".link-to").value;
    if (!from || !to || from === to) return;
    links.push({
      from, to,
      style: el.querySelector(".link-style").value,
      color: el.querySelector(".link-color").value
    });
  });
  return links;
}

// ===========================================================================
// Sample data
// ===========================================================================

function loadSample() {
  document.getElementById("taskList").innerHTML = "";
  document.getElementById("linkList").innerHTML = "";
  rowCounter = 0; segCounter = 0; linkCounter = 0;

  const base = new Date(document.getElementById("startDate").value || new Date());
  const mk = (offset, len) => {
    const s = new Date(base); s.setDate(s.getDate() + offset);
    const e = new Date(s); e.setDate(e.getDate() + len);
    return { start: toInputDate(s), end: toInputDate(e) };
  };

  addGroupRow({ name: "Konzeptphase" });
  addTaskRow({ name: "Anforderungen", segments: [{ type: "bar", ...mk(0, 20), label: "Spec", color: PALETTE[0], showDate: true }] });
  addTaskRow({ name: "Design", segments: [{ type: "bar", ...mk(15, 30), color: PALETTE[1], showDate: true }] });

  addGroupRow({ name: "Umsetzung" });
  addTaskRow({ name: "Entwicklung", segments: [
    { type: "bar", ...mk(40, 25), label: "Backend", color: PALETTE[2], showDate: true },
    { type: "bar", ...mk(60, 25), label: "Frontend", color: PALETTE[4], showDate: false }
  ] });
  addTaskRow({ name: "Test", segments: [{ type: "bar", ...mk(85, 20), color: PALETTE[3], showDate: true }] });
  addTaskRow({ name: "Go-Live", segments: [{ type: "milestone", ...mk(107, 0), label: "Release", color: PALETTE[7], showDate: true }] });

  addLinkRow({ from: "Anforderungen", to: "Design" });
  addLinkRow({ from: "Design", to: "Entwicklung" });
  addLinkRow({ from: "Entwicklung", to: "Test" });
  addLinkRow({ from: "Test", to: "Go-Live" });
  refreshLinkOptions();
}

// ===========================================================================
// Time axis helpers (multi level)
// ===========================================================================

const MONTHS = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"];

function periodStart(date, scale) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  if (scale === "week") { d.setDate(d.getDate() - ((d.getDay() + 6) % 7)); }
  else if (scale === "month") { d.setDate(1); }
  else if (scale === "quarter") { d.setDate(1); d.setMonth(Math.floor(d.getMonth() / 3) * 3); }
  else if (scale === "year") { d.setMonth(0, 1); }
  return d;
}

function nextPeriod(date, scale) {
  const d = new Date(date);
  if (scale === "week") d.setDate(d.getDate() + 7);
  else if (scale === "month") d.setMonth(d.getMonth() + 1);
  else if (scale === "quarter") d.setMonth(d.getMonth() + 3);
  else if (scale === "year") d.setFullYear(d.getFullYear() + 1);
  return d;
}

function periodLabel(date, scale) {
  const d = new Date(date);
  if (scale === "week") { return `KW${isoWeek(d)}`; }
  if (scale === "month") return MONTHS[d.getMonth()];
  if (scale === "quarter") return `Q${Math.floor(d.getMonth() / 3) + 1}`;
  if (scale === "year") return String(d.getFullYear());
  return "";
}

function isoWeek(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dayNum = (date.getUTCDay() + 6) % 7;
  date.setUTCDate(date.getUTCDate() - dayNum + 3);
  const firstThursday = new Date(Date.UTC(date.getUTCFullYear(), 0, 4));
  const week = 1 + Math.round(((date - firstThursday) / 864e5 - 3 + ((firstThursday.getUTCDay() + 6) % 7)) / 7);
  return week;
}

// Build tick periods for one scale within [start, end]; each has start/end date.
function buildPeriods(start, end, scale) {
  const periods = [];
  let cursor = periodStart(start, scale);
  while (cursor < end) {
    const next = nextPeriod(cursor, scale);
    periods.push({ start: new Date(cursor), end: new Date(next), label: periodLabel(cursor, scale), scale });
    cursor = next;
  }
  return periods;
}

// ===========================================================================
// Layout for a set of model rows (fit to slide)
// ===========================================================================

function computeLayout(modelRows, scales) {
  const bandsHeight = scales.length * BAND_HEIGHT;
  const plotTop = MARGIN.top + bandsHeight;
  const available = SLIDE.h - MARGIN.bottom - plotTop;
  const n = Math.max(modelRows.length, 1);
  let rowH = Math.floor(available / n);
  rowH = Math.max(MIN_ROW_HEIGHT, Math.min(MAX_ROW_HEIGHT, rowH));
  const plotLeft = MARGIN.left + LABEL_WIDTH;
  const plotWidth = SLIDE.w - MARGIN.right - plotLeft;
  return { bandsHeight, plotTop, rowH, plotLeft, plotWidth };
}

// ===========================================================================
// PowerPoint insertion
// ===========================================================================

async function insertGantt() {
  validateNames();
  if (document.querySelectorAll(".task-name.invalid").length) {
    setStatus("Task-Namen müssen eindeutig sein (rot markiert).", "error");
    return;
  }
  if (document.querySelectorAll(".seg-end.invalid").length) {
    setStatus("Enddatum darf nicht vor dem Startdatum liegen (rot markiert).", "error");
    return;
  }

  const model = readModel().filter((r) => r.kind === "group" ? r.name : r.segments.length);
  const tasks = model.filter((r) => r.kind === "task");
  if (tasks.length === 0) {
    setStatus("Bitte mindestens eine Task mit gültigem Datum anlegen.", "error");
    return;
  }
  const links = readLinks();

  const allSegs = tasks.flatMap((t) => t.segments);
  const rangeStart = new Date(document.getElementById("startDate").value);
  const rangeEnd = new Date(document.getElementById("endDate").value);
  const start = isValidDate(rangeStart) ? rangeStart : minDate(allSegs.map((s) => s.start));
  const end = isValidDate(rangeEnd) ? rangeEnd : maxDate(allSegs.map((s) => s.end));
  if (end <= start) { setStatus("Enddatum muss nach dem Startdatum liegen.", "error"); return; }

  const title = document.getElementById("chartTitle").value || "Projekt-Timeline";
  const scales = selectedScales();
  const showWorkdays = document.getElementById("optWorkdays").checked;
  const showToday = document.getElementById("optToday").checked;
  const doFit = document.getElementById("optFit").checked;
  const L = computeLayout(model, scales);

  try {
    setStatus("Diagramm wird eingefügt …");
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shapes = slide.shapes;

      // Update mode: remove previously inserted Gantt shapes first.
      shapes.load("items/name");
      await context.sync();
      shapes.items
        .filter((sh) => sh.name && sh.name.indexOf("Gantt") === 0)
        .forEach((sh) => sh.delete());
      await context.sync();

      const created = [];
      const totalMs = end - start;
      const xFor = (date) => L.plotLeft + ((clampDate(date, start, end) - start) / totalMs) * L.plotWidth;
      const plotBottom = L.plotTop + model.length * L.rowH;

      // Title
      created.push(addText(shapes, title, MARGIN.left, 34, SLIDE.w - MARGIN.left - MARGIN.right, 28,
        { bold: true, size: 18, color: "#1b1b1f" }));

      // ---- Multi-level time axis bands (coarse on top) ----
      scales.forEach((scale, level) => {
        const bandTop = MARGIN.top + level * BAND_HEIGHT;
        const periods = buildPeriods(start, end, scale);
        periods.forEach((p) => {
          const x1 = xFor(p.start < start ? start : p.start);
          const x2 = xFor(p.end > end ? end : p.end);
          const w = Math.max(1, x2 - x1);
          const band = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle,
            { left: x1, top: bandTop, width: w, height: BAND_HEIGHT });
          band.fill.setSolidColor(level % 2 === 0 ? "#F1F4F9" : "#E7ECF4");
          band.lineFormat.color = "#D9D9D9";
          band.lineFormat.weight = 0.5;
          band.name = "GanttAxisBand";
          created.push(band);
          created.push(addText(shapes, p.label, x1 + 2, bandTop + 1, Math.max(10, w - 3), BAND_HEIGHT - 2,
            { size: 8, color: "#4a4a52", valign: "middle", bold: level === 0 }));
        });
      });

      // ---- Vertical grid delimiters (finest scale), always vertical ----
      const finest = scales[scales.length - 1];
      buildPeriods(start, end, finest).forEach((p) => {
        if (p.start <= start) return;
        const x = xFor(p.start);
        created.push(addVLine(shapes, x, L.plotTop, plotBottom, "#E3E3E6", { weight: 0.5 }));
      });

      // ---- Rows (groups + tasks) ----
      const geom = {}; // task name -> {left,right,mid}
      model.forEach((row, i) => {
        const rowTop = L.plotTop + i * L.rowH;
        const rowMid = rowTop + L.rowH / 2;

        if (row.kind === "group") {
          const g = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle,
            { left: MARGIN.left, top: rowTop + 1, width: SLIDE.w - MARGIN.left - MARGIN.right, height: L.rowH - 2 });
          g.fill.setSolidColor("#DCE3EF");
          g.lineFormat.visible = false;
          g.name = "GanttGroup";
          created.push(g);
          created.push(addText(shapes, row.name, MARGIN.left + 6, rowTop, LABEL_WIDTH + 200, L.rowH,
            { size: 11, bold: true, color: "#1b1b1f", valign: "middle" }));
          return;
        }

        // Task label
        created.push(addText(shapes, row.name, MARGIN.left, rowTop, LABEL_WIDTH - 8, L.rowH,
          { size: 11, color: "#1b1b1f", valign: "middle" }));

        let taskLeft = Infinity, taskRight = -Infinity;
        row.segments.forEach((seg) => {
          const drawn = drawSegment(shapes, created, seg, rowTop, rowMid, L, xFor, start, end, showWorkdays);
          taskLeft = Math.min(taskLeft, drawn.left);
          taskRight = Math.max(taskRight, drawn.right);
        });
        geom[row.name] = { left: taskLeft, right: taskRight, mid: rowMid };
      });

      // ---- Dashed today marker ----
      const today = new Date();
      if (showToday && today >= start && today <= end) {
        const tx = xFor(today);
        created.push(addVLine(shapes, tx, L.plotTop - 4, plotBottom + 4, "#C00000", { weight: 1.25, dash: true }));
        created.push(addText(shapes, "Heute", tx + 2, L.plotTop - 14, 40, 12, { size: 8, color: "#C00000" }));
      }

      // ---- Dependency arrows (routed through white space) ----
      links.forEach((link) => {
        const a = geom[link.from];
        const b = geom[link.to];
        if (!a || !b) return;
        drawDependency(shapes, created, a, b, link.style, link.color, L.rowH);
      });

      await context.sync();

      // ---- Fit + group everything ----
      if (doFit) {
        try {
          const grp = shapes.addGroup(created);
          grp.name = "GanttGroup-All";
          await context.sync();
        } catch (e) {
          console.warn("Grouping not supported by this host:", e);
        }
      }
    });
    setStatus(`Eingefügt: ${tasks.length} Task(s), ${links.length} Verbindung(en).`, "success");
  } catch (err) {
    console.error(err);
    setStatus("Fehler: " + (err.message || err), "error");
  }
}

// Draw one segment (bar or milestone) on a task row. Returns {left,right}.
function drawSegment(shapes, created, seg, rowTop, rowMid, L, xFor, start, end, showWorkdays) {
  const barHeight = Math.min(18, L.rowH - 8);
  if (seg.type === "milestone") {
    const cx = xFor(seg.start);
    const s = Math.min(16, L.rowH - 6);
    const ms = shapes.addGeometricShape(PowerPoint.GeometricShapeType.diamond,
      { left: cx - s / 2, top: rowMid - s / 2, width: s, height: s });
    ms.fill.setSolidColor(seg.color);
    ms.lineFormat.visible = false;
    ms.name = "GanttSeg";
    created.push(ms);
    const bits = [];
    if (seg.label) bits.push(seg.label);
    if (seg.showDate) bits.push(fmtDate(seg.start));
    if (bits.length) {
      created.push(addText(shapes, bits.join("  "), cx + s / 2 + 4, rowMid - 7, 150, 14,
        { size: 8, color: "#4a4a52", valign: "middle" }));
    }
    return { left: cx - s / 2, right: cx + s / 2 };
  }

  const barTop = rowMid - barHeight / 2;
  const left = xFor(seg.start);
  const right = xFor(seg.end);
  const width = Math.max(2, right - left);
  const bar = shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundRectangle,
    { left, top: barTop, width, height: barHeight });
  bar.fill.setSolidColor(seg.color);
  bar.lineFormat.visible = false;
  bar.name = "GanttSeg";
  created.push(bar);

  // Label inside the bar
  if (seg.label) {
    created.push(addText(shapes, seg.label, left + 3, barTop, Math.max(10, width - 6), barHeight,
      { size: 8, color: "#ffffff", valign: "middle" }));
  }

  // Date + workdays labels
  if (seg.showDate) {
    created.push(addText(shapes, fmtDate(seg.start), left, barTop + barHeight, 60, 12,
      { size: 8, color: "#6b6b73" }));
    created.push(addText(shapes, fmtDate(seg.end), right - 44, barTop + barHeight, 44, 12,
      { size: 8, color: "#6b6b73", align: "right" }));
  }
  if (showWorkdays) {
    created.push(addText(shapes, `${netWorkdays(seg.start, seg.end)} AT`, right + 4, barTop, 60, barHeight,
      { size: 8, color: "#6b6b73", valign: "middle" }));
  }
  return { left, right };
}

// Requirement 1+2: ONE real connector shape (elbow) with a native arrowhead,
// routed at right angles through the white space between the two elements.
function drawDependency(shapes, created, a, b, style, color, rowH) {
  let x1, x2;
  switch (style) {
    case "start-start": x1 = a.left; x2 = b.left; break;
    case "finish-finish": x1 = a.right; x2 = b.right; break;
    default: x1 = a.right; x2 = b.left; break; // finish-start
  }
  const y1 = a.mid, y2 = b.mid;

  // A single elbow connector = one shape that bends at 90° (never diagonal,
  // never cutting through a bar) and carries a proper triangular arrowhead.
  const line = shapes.addLine(PowerPoint.ConnectorType.elbow, {
    left: Math.min(x1, x2),
    top: Math.min(y1, y2),
    width: Math.max(1, Math.abs(x2 - x1)),
    height: Math.max(1, Math.abs(y2 - y1))
  });

  const lf = line.lineFormat;
  lf.color = color;
  lf.weight = 1.5;

  // Native arrowhead so it renders as a real arrow, not a separate triangle.
  // The connector's "end" corner is bottom-right of its bounding box; put the
  // arrowhead on whichever end represents the target element.
  const targetIsEndCorner = (x2 >= x1) === (y2 >= y1);
  try {
    lf.beginArrowheadStyle = PowerPoint.ArrowheadStyle.none;
    lf.endArrowheadStyle = PowerPoint.ArrowheadStyle.none;
    const head = targetIsEndCorner ? "end" : "begin";
    lf[`${head}ArrowheadStyle`] = PowerPoint.ArrowheadStyle.triangle;
    lf[`${head}ArrowheadLength`] = PowerPoint.ArrowheadLength.medium;
    lf[`${head}ArrowheadWidth`] = PowerPoint.ArrowheadWidth.medium;
  } catch (e) {
    // Older hosts without arrowhead support: at least a clean single elbow line.
  }

  line.name = "GanttLink";
  created.push(line);
}

// ===========================================================================
// PowerPoint shape helpers
// ===========================================================================

function addText(shapes, text, left, top, width, height, opts = {}) {
  const box = shapes.addTextBox(text, { left, top, width, height });
  box.fill.clear();
  box.lineFormat.visible = false;
  box.name = opts.name || "GanttText";
  const range = box.textFrame.textRange;
  const font = range.font;
  if (opts.size) font.size = opts.size;
  if (opts.bold) font.bold = true;
  if (opts.color) font.color = opts.color;
  try {
    box.textFrame.topMargin = 0; box.textFrame.bottomMargin = 0;
    box.textFrame.leftMargin = 1; box.textFrame.rightMargin = 1;
  } catch (e) { /* older API */ }
  if (opts.align) {
    try {
      range.paragraphFormat.horizontalAlignment = opts.align === "right"
        ? PowerPoint.ParagraphHorizontalAlignment.right
        : PowerPoint.ParagraphHorizontalAlignment.left;
    } catch (e) { /* older API */ }
  }
  return box;
}

// Guaranteed-vertical line (width 0) rendered via a thin connector.
function addVLine(shapes, x, y1, y2, color, opts = {}) {
  const line = shapes.addLine(PowerPoint.ConnectorType.straight,
    { left: x, top: Math.min(y1, y2), width: 0, height: Math.abs(y2 - y1) });
  line.lineFormat.color = color;
  line.lineFormat.weight = opts.weight || 0.75;
  if (opts.dash) {
    try { line.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.dash; } catch (e) { /* older API */ }
  }
  line.name = "GanttGrid";
  return line;
}

// ===========================================================================
// Date utilities
// ===========================================================================

function isValidDate(d) { return d instanceof Date && !isNaN(d); }
function clampDate(d, min, max) { return new Date(Math.min(Math.max(d, min), max)); }
function minDate(arr) { return new Date(Math.min(...arr)); }
function maxDate(arr) { return new Date(Math.max(...arr)); }

function netWorkdays(start, end) {
  if (!isValidDate(start) || !isValidDate(end)) return 0;
  let a = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const b = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  if (b < a) return 0;
  let count = 0;
  while (a <= b) {
    const dow = a.getDay();
    if (dow !== 0 && dow !== 6) count += 1;
    a.setDate(a.getDate() + 1);
  }
  return count;
}

function fmtDate(d) {
  if (!isValidDate(d)) return "";
  return `${String(d.getDate()).padStart(2, "0")}.${String(d.getMonth() + 1).padStart(2, "0")}.`;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
function escapeAttr(s) { return escapeHtml(s); }

// ===========================================================================
// Interactive SVG preview
// ===========================================================================

const SVGNS = "http://www.w3.org/2000/svg";
const PV = { labelW: 120, rowH: 26, top: 8, barH: 14, msSize: 14, padRight: 50, band: 15 };

function previewRange(segs) {
  const s = new Date(document.getElementById("startDate").value);
  const e = new Date(document.getElementById("endDate").value);
  const start = isValidDate(s) ? s : (segs.length ? minDate(segs.map((x) => x.start)) : new Date());
  let end = isValidDate(e) ? e : (segs.length ? maxDate(segs.map((x) => x.end)) : new Date());
  if (end <= start) end = new Date(start.getTime() + 30 * 864e5);
  return { start, end };
}

function renderPreview() {
  const host = document.getElementById("ganttPreview");
  if (!host) return;
  const model = readModel();
  const scales = selectedScales();
  const showDatesGlobal = document.getElementById("optDates").checked;
  const showWorkdays = document.getElementById("optWorkdays").checked;
  const showToday = document.getElementById("optToday").checked;

  const allSegs = model.filter((r) => r.kind === "task").flatMap((r) => r.segments);
  const { start, end } = previewRange(allSegs);

  const width = Math.max(host.clientWidth || 320, 320);
  const bandsH = scales.length * PV.band;
  const plotTop = PV.top + bandsH;
  const plotLeft = PV.labelW;
  const plotW = width - PV.labelW - PV.padRight;
  const totalMs = end - start;
  const pxPerMs = plotW / totalMs;
  const height = plotTop + model.length * PV.rowH + 24;
  const xFor = (d) => plotLeft + (clampDate(d, start, end) - start) * pxPerMs;

  const svg = document.createElementNS(SVGNS, "svg");
  svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
  svg.setAttribute("width", width);
  svg.setAttribute("height", height);

  const plotBottom = plotTop + model.length * PV.rowH;

  // Multi-level axis bands
  scales.forEach((scale, level) => {
    const bandTop = PV.top + level * PV.band;
    buildPeriods(start, end, scale).forEach((p) => {
      const x1 = xFor(p.start < start ? start : p.start);
      const x2 = xFor(p.end > end ? end : p.end);
      const r = rect(svg, x1, bandTop, Math.max(1, x2 - x1), PV.band, level % 2 ? "#e7ecf4" : "#f4f6fa", "gp-band");
      r.setAttribute("stroke", "#e2e2e6");
      text(svg, x1 + 3, bandTop + PV.band - 4, p.label, level === 0 ? "gp-axis-label bold" : "gp-axis-label");
    });
  });

  // Vertical delimiters (finest scale)
  const finest = scales[scales.length - 1];
  buildPeriods(start, end, finest).forEach((p) => {
    if (p.start <= start) return;
    line(svg, xFor(p.start), plotTop, xFor(p.start), plotBottom, "gp-grid");
  });

  // today
  const today = new Date();
  if (showToday && today >= start && today <= end) {
    const tx = xFor(today);
    line(svg, tx, plotTop - 3, tx, plotBottom + 3, "gp-today");
  }

  const geom = {};
  const rowEls = Array.from(document.getElementById("taskList").children);

  model.forEach((row, i) => {
    const rowTop = plotTop + i * PV.rowH;
    const rowMid = rowTop + PV.rowH / 2;
    const rowEl = rowEls[i];

    if (row.kind === "group") {
      rect(svg, 0, rowTop + 1, width, PV.rowH - 2, "#eef2f9", "");
      text(svg, 4, rowMid + 4, row.name || "Überschrift", "gp-group-label");
      return;
    }

    text(svg, 4, rowMid + 3, row.name || `Task ${i + 1}`, "gp-row-label");

    let taskLeft = Infinity, taskRight = -Infinity;
    row.segments.forEach((seg, si) => {
      const segEl = rowEl ? rowEl.querySelectorAll(".segment")[si] : null;
      if (seg.type === "milestone") {
        const cx = xFor(seg.start);
        const s = PV.msSize;
        const dia = poly(svg, [[cx, rowMid - s / 2], [cx + s / 2, rowMid], [cx, rowMid + s / 2], [cx - s / 2, rowMid]], seg.color, "gp-ms");
        if (segEl) attachDrag(dia, segEl, { start, end, mode: "move-ms" });
        const bits = [];
        if (seg.label) bits.push(seg.label);
        if (showDatesGlobal && seg.showDate) bits.push(fmtDate(seg.start));
        if (bits.length) text(svg, cx + s / 2 + 4, rowMid + 3, bits.join("  "), "gp-datelbl");
        taskLeft = Math.min(taskLeft, cx - s / 2); taskRight = Math.max(taskRight, cx + s / 2);
        return;
      }
      const barTop = rowMid - PV.barH / 2;
      const x1 = xFor(seg.start);
      const x2 = xFor(seg.end);
      const w = Math.max(3, x2 - x1);
      const bar = rect(svg, x1, barTop, w, PV.barH, seg.color, "gp-bar");
      if (segEl) attachDrag(bar, segEl, { start, end, mode: "move" });
      if (seg.label) text(svg, x1 + 4, barTop + PV.barH - 3, seg.label, "gp-seglbl");
      // resize handles
      if (segEl) {
        const hL = rect(svg, x1 - 3, barTop, 6, PV.barH, "transparent", "gp-handle");
        attachDrag(hL, segEl, { start, end, mode: "resize-start" });
        const hR = rect(svg, x2 - 3, barTop, 6, PV.barH, "transparent", "gp-handle");
        attachDrag(hR, segEl, { start, end, mode: "resize-end" });
      }
      if (showDatesGlobal && seg.showDate) {
        text(svg, x1, barTop + PV.barH + 9, fmtDate(seg.start), "gp-datelbl");
        const el = text(svg, x2, barTop + PV.barH + 9, fmtDate(seg.end), "gp-datelbl");
        el.setAttribute("text-anchor", "end");
      }
      if (showWorkdays) text(svg, x2 + 5, rowMid + 3, `${netWorkdays(seg.start, seg.end)} AT`, "gp-datelbl");
      taskLeft = Math.min(taskLeft, x1); taskRight = Math.max(taskRight, x2);
    });
    if (row.name) geom[row.name] = { left: taskLeft, right: taskRight, mid: rowMid };
  });

  // dependency links (routed)
  readLinks().forEach((lk) => {
    const a = geom[lk.from]; const b = geom[lk.to];
    if (!a || !b) return;
    let ax, bx;
    switch (lk.style) {
      case "start-start": ax = a.left; bx = b.left; break;
      case "finish-finish": ax = a.right; bx = b.right; break;
      default: ax = a.right; bx = b.left; break;
    }
    const ay = a.mid, by = b.mid;
    const channelX = ax + 12;
    const enterX = bx - 12;
    const midX = enterX > channelX ? enterX : bx;
    const d = `M ${ax} ${ay} L ${channelX} ${ay} L ${channelX} ${by} L ${midX} ${by} L ${bx} ${by}`;
    const p = document.createElementNS(SVGNS, "path");
    p.setAttribute("d", d); p.setAttribute("class", "gp-link");
    p.setAttribute("stroke", lk.color); p.setAttribute("marker-end", "url(#gp-arrow)");
    svg.appendChild(p);
  });

  const defs = document.createElementNS(SVGNS, "defs");
  defs.innerHTML = `<marker id="gp-arrow" markerWidth="8" markerHeight="8" refX="6" refY="3" orient="auto"><path d="M0,0 L6,3 L0,6 Z" fill="context-stroke"/></marker>`;
  svg.insertBefore(defs, svg.firstChild);

  host.innerHTML = "";
  host.appendChild(svg);
}

// SVG helpers
function rect(svg, x, y, w, h, fill, cls) {
  const r = document.createElementNS(SVGNS, "rect");
  r.setAttribute("x", x); r.setAttribute("y", y);
  r.setAttribute("width", Math.max(0, w)); r.setAttribute("height", h);
  r.setAttribute("rx", 2); r.setAttribute("fill", fill);
  if (cls) r.setAttribute("class", cls);
  svg.appendChild(r); return r;
}
function line(svg, x1, y1, x2, y2, cls) {
  const l = document.createElementNS(SVGNS, "line");
  l.setAttribute("x1", x1); l.setAttribute("y1", y1);
  l.setAttribute("x2", x2); l.setAttribute("y2", y2);
  if (cls) l.setAttribute("class", cls);
  svg.appendChild(l); return l;
}
function text(svg, x, y, str, cls) {
  const t = document.createElementNS(SVGNS, "text");
  t.setAttribute("x", x); t.setAttribute("y", y);
  if (cls) t.setAttribute("class", cls);
  t.textContent = str; svg.appendChild(t); return t;
}
function poly(svg, pts, fill, cls) {
  const p = document.createElementNS(SVGNS, "polygon");
  p.setAttribute("points", pts.map((q) => q.join(",")).join(" "));
  p.setAttribute("fill", fill);
  if (cls) p.setAttribute("class", cls);
  svg.appendChild(p); return p;
}

// Drag: operates on a segment DOM element.
function attachDrag(el, segEl, ctx) {
  el.addEventListener("pointerdown", (ev) => {
    ev.preventDefault();
    const svg = el.ownerSVGElement;
    const mode = ctx.mode;
    const startX = ev.clientX;
    const startInput = segEl.querySelector(".seg-start");
    const endInput = segEl.querySelector(".seg-end");
    const origStart = startInput.value ? new Date(startInput.value) : new Date();
    const origEnd = endInput.value ? new Date(endInput.value) : new Date(origStart);
    const pxPerMs = (svg.viewBox.baseVal.width - PV.labelW - PV.padRight) / (ctx.end - ctx.start);

    const onMove = (e) => {
      const dxMs = (e.clientX - startX) / pxPerMs;
      const dDays = Math.round(dxMs / 864e5);
      if (mode === "move" || mode === "move-ms") {
        startInput.value = toInputDate(addDays(origStart, dDays));
        if (mode === "move") endInput.value = toInputDate(addDays(origEnd, dDays));
      } else if (mode === "resize-start") {
        const ns = addDays(origStart, dDays);
        if (ns < origEnd) startInput.value = toInputDate(ns);
      } else if (mode === "resize-end") {
        const ne = addDays(origEnd, dDays);
        if (ne > origStart) endInput.value = toInputDate(ne);
      }
      validateSegment(segEl);
      renderPreview();
    };
    const onUp = () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  });
}

function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }

// ===========================================================================
// Sync positions back from the slide (first segment of each task)
// ===========================================================================

async function syncFromSlide() {
  setStatus("Positionsabgleich wird in dieser Version pro Segment nicht unterstützt.", "error");
}

function setStatus(msg, type = "") {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = "status" + (type ? " " + type : "");
}
