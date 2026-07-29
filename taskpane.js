/* global Office, PowerPoint, console */

// ---------------------------------------------------------------------------
// Gantt Chart Builder - task pane logic
// Renders a Gantt chart as native PowerPoint shapes (editable, ThinkCell-style).
// Supports: bars, milestones (diamonds) and dependency arrows between elements.
// ---------------------------------------------------------------------------

const PALETTE = ["#4472C4", "#ED7D31", "#70AD47", "#FFC000", "#5B9BD5", "#A5A5A5", "#264478", "#9E480E"];

// Layout constants (points; slide is ~960x540 pt for 16:9)
const LAYOUT = {
  chartLeft: 40,
  chartTop: 90,
  chartWidth: 860,
  labelWidth: 180,
  rowHeight: 30,
  rowGap: 6,
  headerHeight: 24,
  barHeight: 16,
  milestoneSize: 16
};

let taskCounter = 0;
let linkCounter = 0;

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("addTask").onclick = () => { addTaskRow(); refreshLinkOptions(); renderPreview(); };
    document.getElementById("addLink").onclick = () => { addLinkRow(); renderPreview(); };
    document.getElementById("insert").onclick = insertGantt;
    document.getElementById("sample").onclick = () => { loadSample(); renderPreview(); };
    document.getElementById("syncSlide").onclick = syncFromSlide;
    document.getElementById("startDate").onchange = renderPreview;
    document.getElementById("endDate").onchange = renderPreview;
    document.getElementById("scale").onchange = renderPreview;
    document.getElementById("optDates").onchange = renderPreview;
    document.getElementById("optWorkdays").onchange = renderPreview;
    initDefaults();
    renderPreview();
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

function toInputDate(d) {
  return d.toISOString().slice(0, 10);
}

// --- Task row management ----------------------------------------------------

function addTaskRow(task = {}) {
  taskCounter += 1;
  const id = taskCounter;
  const wrapper = document.createElement("div");
  wrapper.className = "task-item";
  wrapper.dataset.id = id;
  const type = task.type || "bar";
  wrapper.innerHTML = `
    <input class="task-name" type="text" placeholder="Task name" value="${task.name || ""}" />
    <div class="task-dates">
      <select class="task-type" title="Type">
        <option value="bar" ${type === "bar" ? "selected" : ""}>Bar</option>
        <option value="milestone" ${type === "milestone" ? "selected" : ""}>Milestone</option>
      </select>
      <input class="task-start" type="date" value="${task.start || ""}" />
      <input class="task-end" type="date" value="${task.end || ""}" title="End (bars only)" />
    </div>
    <div class="task-meta">
      <label style="margin:0">%</label>
      <input class="task-progress" type="number" min="0" max="100" value="${task.progress ?? 0}" />
      <input class="task-color" type="color" value="${task.color || PALETTE[(id - 1) % PALETTE.length]}" />
    </div>
    <button class="task-remove" type="button" title="Remove">&times;</button>
  `;
  const typeSel = wrapper.querySelector(".task-type");
  const endInput = wrapper.querySelector(".task-end");
  const progInput = wrapper.querySelector(".task-progress");
  const applyType = () => {
    const isMs = typeSel.value === "milestone";
    endInput.style.visibility = isMs ? "hidden" : "visible";
    progInput.disabled = isMs;
  };
  typeSel.onchange = () => { applyType(); renderPreview(); };
  applyType();

  wrapper.querySelector(".task-name").oninput = () => { refreshLinkOptions(); renderPreview(); };
  wrapper.querySelectorAll(".task-start, .task-end, .task-progress, .task-color").forEach((inp) => {
    inp.oninput = renderPreview;
  });
  wrapper.querySelector(".task-remove").onclick = () => { wrapper.remove(); refreshLinkOptions(); renderPreview(); };
  document.getElementById("taskList").appendChild(wrapper);
}

function readTasks() {
  const items = document.querySelectorAll(".task-item:not(.link-item)");
  const tasks = [];
  items.forEach((el) => {
    const name = el.querySelector(".task-name").value.trim();
    const type = el.querySelector(".task-type").value;
    const start = el.querySelector(".task-start").value;
    const end = el.querySelector(".task-end").value;
    if (!name || !start) return;
    if (type === "bar" && !end) return;
    tasks.push({
      name,
      type,
      start: new Date(start),
      end: type === "milestone" ? new Date(start) : new Date(end),
      progress: Number(el.querySelector(".task-progress").value) || 0,
      color: el.querySelector(".task-color").value
    });
  });
  return tasks;
}

// --- Link (dependency) management ------------------------------------------

function addLinkRow(link = {}) {
  linkCounter += 1;
  const id = linkCounter;
  const wrapper = document.createElement("div");
  wrapper.className = "task-item link-item";
  wrapper.dataset.id = id;
  wrapper.innerHTML = `
    <div class="link-row">
      <select class="link-from" title="From"></select>
      <span class="link-arrow">&rarr;</span>
      <select class="link-to" title="To"></select>
    </div>
    <div class="task-meta">
      <select class="link-style" title="Line style">
        <option value="finish-start" selected>Finish &rarr; Start</option>
        <option value="start-start">Start &rarr; Start</option>
        <option value="finish-finish">Finish &rarr; Finish</option>
      </select>
      <input class="link-color" type="color" value="${link.color || "#c00000"}" />
    </div>
    <button class="task-remove" type="button" title="Remove">&times;</button>
  `;
  wrapper.querySelector(".task-remove").onclick = () => wrapper.remove();
  document.getElementById("linkList").appendChild(wrapper);
  populateLinkSelects(wrapper, link);
  wrapper.querySelectorAll("select, input").forEach((c) => { c.onchange = renderPreview; });
  wrapper.querySelector(".task-remove").onclick = () => { wrapper.remove(); renderPreview(); };
}

function taskNames() {
  return Array.from(document.querySelectorAll(".task-item:not(.link-item) .task-name"))
    .map((i) => i.value.trim())
    .filter(Boolean);
}

function populateLinkSelects(wrapper, link = {}) {
  const names = taskNames();
  ["from", "to"].forEach((role) => {
    const sel = wrapper.querySelector(`.link-${role}`);
    const current = sel.value || link[role];
    sel.innerHTML = names.map((n) => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join("");
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
      from,
      to,
      style: el.querySelector(".link-style").value,
      color: el.querySelector(".link-color").value
    });
  });
  return links;
}

function loadSample() {
  document.getElementById("taskList").innerHTML = "";
  document.getElementById("linkList").innerHTML = "";
  taskCounter = 0;
  linkCounter = 0;
  const base = new Date(document.getElementById("startDate").value || new Date());
  const mk = (offset, len) => {
    const s = new Date(base); s.setDate(s.getDate() + offset);
    const e = new Date(s); e.setDate(e.getDate() + len);
    return { start: toInputDate(s), end: toInputDate(e) };
  };
  [
    { name: "Concept & requirements", ...mk(0, 20), progress: 100 },
    { name: "Design", ...mk(15, 30), progress: 70 },
    { name: "Development", ...mk(40, 45), progress: 30 },
    { name: "Testing", ...mk(75, 20), progress: 0 },
    { name: "Go-Live", type: "milestone", ...mk(97, 0) }
  ].forEach((t, i) => addTaskRow({ ...t, color: PALETTE[i % PALETTE.length] }));

  addLinkRow({ from: "Concept & requirements", to: "Design" });
  addLinkRow({ from: "Design", to: "Development" });
  addLinkRow({ from: "Development", to: "Testing" });
  addLinkRow({ from: "Testing", to: "Go-Live" });
  refreshLinkOptions();
}

// --- Rendering --------------------------------------------------------------

async function insertGantt() {
  const tasks = readTasks();
  if (tasks.length === 0) {
    setStatus("Please add at least one task with name and start date.", "error");
    return;
  }
  const links = readLinks();

  const rangeStart = new Date(document.getElementById("startDate").value);
  const rangeEnd = new Date(document.getElementById("endDate").value);
  const start = isValidDate(rangeStart) ? rangeStart : minDate(tasks.map((t) => t.start));
  const end = isValidDate(rangeEnd) ? rangeEnd : maxDate(tasks.map((t) => t.end));

  if (end <= start) {
    setStatus("End date must be after start date.", "error");
    return;
  }

  const title = document.getElementById("chartTitle").value || "Project Timeline";
  const scale = document.getElementById("scale").value;
  const showDates = document.getElementById("optDates").checked;
  const showWorkdays = document.getElementById("optWorkdays").checked;

  try {
    setStatus("Inserting chart…");
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

      const totalMs = end - start;
      const plotLeft = LAYOUT.chartLeft + LAYOUT.labelWidth;
      const plotWidth = LAYOUT.chartWidth - LAYOUT.labelWidth;
      const xFor = (date) => plotLeft + ((date - start) / totalMs) * plotWidth;

      // Geometry per task, so links can connect to real coordinates.
      const geom = {};

      // Title
      addText(shapes, title, LAYOUT.chartLeft, 40, LAYOUT.chartWidth, 28, {
        bold: true, size: 18, color: "#1b1b1f"
      });

      // Time axis ticks + labels
      const ticks = buildTicks(start, end, scale);
      const axisTop = LAYOUT.chartTop - LAYOUT.headerHeight;
      const rowsBottom = LAYOUT.chartTop + tasks.length * (LAYOUT.rowHeight + LAYOUT.rowGap);
      ticks.forEach((tick) => {
        const x = xFor(tick.date);
        addLine(shapes, x, axisTop, x, rowsBottom, "#D9D9D9");
        addText(shapes, tick.label, x + 2, axisTop - 2, 80, LAYOUT.headerHeight, {
          size: 9, color: "#6b6b73"
        });
      });

      // Task rows
      tasks.forEach((task, i) => {
        const rowTop = LAYOUT.chartTop + i * (LAYOUT.rowHeight + LAYOUT.rowGap);
        const rowMid = rowTop + LAYOUT.rowHeight / 2;

        // Row label
        addText(shapes, task.name, LAYOUT.chartLeft, rowTop, LAYOUT.labelWidth - 8, LAYOUT.rowHeight, {
          size: 11, color: "#1b1b1f", valign: "middle"
        });

        if (task.type === "milestone") {
          const cx = xFor(clampDate(task.start, start, end));
          const s = LAYOUT.milestoneSize;
          const ms = shapes.addGeometricShape(PowerPoint.GeometricShapeType.diamond, {
            left: cx - s / 2, top: rowMid - s / 2, width: s, height: s
          });
          ms.fill.setSolidColor(task.color);
          ms.lineFormat.visible = false;
          ms.name = `Gantt-MS-${task.name}`;
          const msLabel = showDates ? `${task.name}  ${fmtDate(task.start)}` : task.name;
          addText(shapes, msLabel, cx + s / 2 + 4, rowMid - LAYOUT.barHeight / 2 - 2,
            160, LAYOUT.barHeight + 4, { size: 9, color: "#6b6b73", valign: "middle" });
          geom[task.name] = { left: cx - s / 2, right: cx + s / 2, mid: rowMid, cx };
          return;
        }

        const barTop = rowTop + (LAYOUT.rowHeight - LAYOUT.barHeight) / 2;
        const barLeft = xFor(clampDate(task.start, start, end));
        const barRight = xFor(clampDate(task.end, start, end));
        const barWidth = Math.max(2, barRight - barLeft);
        const bar = shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundRectangle, {
          left: barLeft, top: barTop, width: barWidth, height: LAYOUT.barHeight
        });
        bar.fill.setSolidColor(task.color);
        bar.lineFormat.visible = false;
        bar.name = `Gantt-${task.name}`;
        geom[task.name] = { left: barLeft, right: barRight, mid: rowMid };

        // Progress overlay
        if (task.progress > 0) {
          const progWidth = Math.max(1, (barWidth * Math.min(task.progress, 100)) / 100);
          const prog = shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundRectangle, {
            left: barLeft, top: barTop, width: progWidth, height: LAYOUT.barHeight
          });
          prog.fill.setSolidColor(darken(task.color));
          prog.lineFormat.visible = false;
          prog.name = `Gantt-${task.name}-progress`;
        }

        // Optional start/end date labels below the bar.
        if (showDates) {
          addText(shapes, fmtDate(task.start), barLeft, barTop + LAYOUT.barHeight, 60,
            12, { size: 8, color: "#6b6b73" });
          addText(shapes, fmtDate(task.end), barRight - 40, barTop + LAYOUT.barHeight, 44,
            12, { size: 8, color: "#6b6b73", align: "right" });
        }

        // Right-side label: progress % and/or net workdays.
        const parts = [];
        if (task.progress > 0) parts.push(`${task.progress}%`);
        if (showWorkdays) parts.push(`${netWorkdays(task.start, task.end)} wd`);
        if (parts.length) {
          addText(shapes, parts.join(" · "), barRight + 4, barTop - 2, 90, LAYOUT.barHeight + 4, {
            size: 9, color: "#6b6b73", valign: "middle"
          });
        }
      });

      // Dependency arrows
      links.forEach((link) => {
        const a = geom[link.from];
        const b = geom[link.to];
        if (!a || !b) return;
        const [x1, y1, x2, y2] = anchorPoints(a, b, link.style);
        drawDependency(shapes, x1, y1, x2, y2, link.color);
      });

      await context.sync();
    });
    setStatus(`Inserted ${tasks.length} task(s) and ${links.length} connection(s).`, "success");
  } catch (err) {
    console.error(err);
    setStatus("Error: " + (err.message || err), "error");
  }
}

// Determine start/end anchor points depending on link style.
function anchorPoints(a, b, style) {
  let x1, x2;
  const y1 = a.mid;
  const y2 = b.mid;
  switch (style) {
    case "start-start": x1 = a.left; x2 = b.left; break;
    case "finish-finish": x1 = a.right; x2 = b.right; break;
    case "finish-start":
    default: x1 = a.right; x2 = b.left; break;
  }
  return [x1, y1, x2, y2];
}

// Draw an elbow (right-angle) connector with an arrowhead, think-cell style.
// Uses thin rectangles for the segments (always axis-aligned, unlike lines
// which PowerPoint renders diagonally for zero-size bounding boxes) plus a
// small triangle as the arrowhead.
function drawDependency(shapes, x1, y1, x2, y2, color) {
  const t = 1.4;               // line thickness (pt)
  const arrow = 6;             // arrowhead size (pt)
  const forward = x2 >= x1;
  const midX = forward ? (x1 + x2) / 2 : x1 + 14;
  const stub = forward ? arrow : 0; // leave room so the arrow tip lands on x2

  // horizontal from source, vertical to target row, horizontal into target
  hSeg(shapes, x1, midX, y1, t, color);
  vSeg(shapes, midX, y1, y2, t, color);
  hSeg(shapes, midX, x2 - stub, y2, t, color);

  // arrowhead pointing toward the target (into the successor bar)
  arrowHead(shapes, x2, y2, forward ? "right" : "left", arrow, color);
}

function hSeg(shapes, xa, xb, y, t, color) {
  const left = Math.min(xa, xb);
  const w = Math.abs(xb - xa);
  if (w < 0.5) return;
  const r = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
    left, top: y - t / 2, width: w, height: t
  });
  r.fill.setSolidColor(color);
  r.lineFormat.visible = false;
  r.name = "GanttLink";
}

function vSeg(shapes, x, ya, yb, t, color) {
  const top = Math.min(ya, yb);
  const h = Math.abs(yb - ya);
  if (h < 0.5) return;
  const r = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
    left: x - t / 2, top, width: t, height: h
  });
  r.fill.setSolidColor(color);
  r.lineFormat.visible = false;
  r.name = "GanttLink";
}

function arrowHead(shapes, x, y, dir, size, color) {
  // isosceles triangle points up by default; rotate to point along travel dir.
  const left = dir === "right" ? x - size : x;
  const tri = shapes.addGeometricShape(PowerPoint.GeometricShapeType.triangle, {
    left,
    top: y - size / 2,
    width: size,
    height: size
  });
  tri.fill.setSolidColor(color);
  tri.lineFormat.visible = false;
  tri.name = "GanttLink";
  try {
    tri.rotation = dir === "right" ? 90 : 270;
  } catch (e) { /* rotation not supported */ }
}


// --- PowerPoint shape helpers ----------------------------------------------

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
  if (opts.align) {
    try { range.paragraphFormat.horizontalAlignment = opts.align === "right"
      ? PowerPoint.ParagraphHorizontalAlignment.right
      : PowerPoint.ParagraphHorizontalAlignment.left; } catch (e) { /* older API */ }
  }
  return box;
}

function addLine(shapes, x1, y1, x2, y2, color, name) {
  const line = shapes.addLine(PowerPoint.ConnectorType.straight, {
    left: Math.min(x1, x2),
    top: Math.min(y1, y2),
    width: Math.abs(x2 - x1),
    height: Math.abs(y2 - y1)
  });
  line.lineFormat.color = color;
  line.lineFormat.weight = 0.75;
  line.name = name || "GanttAxis";
  return line;
}

// --- Date + axis utilities --------------------------------------------------

function buildTicks(start, end, scale) {
  const ticks = [];
  const cursor = new Date(start);
  const fmt = { week: shortDate, month: monthLabel, quarter: quarterLabel }[scale] || monthLabel;

  if (scale === "week") {
    cursor.setDate(cursor.getDate() - cursor.getDay());
    while (cursor <= end) {
      ticks.push({ date: new Date(cursor), label: fmt(cursor) });
      cursor.setDate(cursor.getDate() + 7);
    }
  } else if (scale === "quarter") {
    cursor.setDate(1);
    cursor.setMonth(Math.floor(cursor.getMonth() / 3) * 3);
    while (cursor <= end) {
      ticks.push({ date: new Date(cursor), label: fmt(cursor) });
      cursor.setMonth(cursor.getMonth() + 3);
    }
  } else {
    cursor.setDate(1);
    while (cursor <= end) {
      ticks.push({ date: new Date(cursor), label: fmt(cursor) });
      cursor.setMonth(cursor.getMonth() + 1);
    }
  }
  return ticks;
}

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
function monthLabel(d) { return `${MONTHS[d.getMonth()]} ${String(d.getFullYear()).slice(2)}`; }
function quarterLabel(d) { return `Q${Math.floor(d.getMonth() / 3) + 1} ${String(d.getFullYear()).slice(2)}`; }
function shortDate(d) { return `${d.getDate()}.${d.getMonth() + 1}.`; }

function isValidDate(d) { return d instanceof Date && !isNaN(d); }
function clampDate(d, min, max) { return new Date(Math.min(Math.max(d, min), max)); }
function minDate(arr) { return new Date(Math.min(...arr)); }
function maxDate(arr) { return new Date(Math.max(...arr)); }

// Net working days between two dates (inclusive), counting Mon–Fri only.
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

// Short date like "15.08." for labels.
function fmtDate(d) {
  if (!isValidDate(d)) return "";
  return `${String(d.getDate()).padStart(2, "0")}.${String(d.getMonth() + 1).padStart(2, "0")}.`;
}

function darken(hex) {
  const c = hex.replace("#", "");
  const num = parseInt(c, 16);
  const r = Math.max(0, ((num >> 16) & 255) - 40);
  const g = Math.max(0, ((num >> 8) & 255) - 40);
  const b = Math.max(0, (num & 255) - 40);
  return `#${((r << 16) | (g << 8) | b).toString(16).padStart(6, "0")}`;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// --- Interactive SVG preview (drag to move / resize) -----------------------

const SVGNS = "http://www.w3.org/2000/svg";
const PV = { labelW: 110, rowH: 26, rowGap: 4, top: 26, barH: 14, msSize: 14, padRight: 40 };

function previewRange(tasks) {
  const s = new Date(document.getElementById("startDate").value);
  const e = new Date(document.getElementById("endDate").value);
  const start = isValidDate(s) ? s : (tasks.length ? minDate(tasks.map((t) => t.start)) : new Date());
  let end = isValidDate(e) ? e : (tasks.length ? maxDate(tasks.map((t) => t.end)) : new Date());
  if (end <= start) end = new Date(start.getTime() + 30 * 864e5);
  return { start, end };
}

function taskRows() {
  return Array.from(document.querySelectorAll(".task-item:not(.link-item)"));
}

function renderPreview() {
  const host = document.getElementById("ganttPreview");
  if (!host) return;
  const rows = taskRows();
  const tasks = readTasks();
  const showDates = document.getElementById("optDates").checked;
  const showWorkdays = document.getElementById("optWorkdays").checked;
  const { start, end } = previewRange(tasks);
  const width = Math.max(host.clientWidth || 320, 320);
  const plotLeft = PV.labelW;
  const plotW = width - PV.labelW - PV.padRight;
  const totalMs = end - start;
  const pxPerMs = plotW / totalMs;
  const height = PV.top + rows.length * (PV.rowH + PV.rowGap) + 10;
  const xFor = (d) => plotLeft + (clampDate(d, start, end) - start) * pxPerMs;
  const dateFor = (x) => new Date(start.getTime() + (x - plotLeft) / pxPerMs);

  const svg = document.createElementNS(SVGNS, "svg");
  svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
  svg.setAttribute("width", width);
  svg.setAttribute("height", height);

  // grid (month ticks)
  const ticks = buildTicks(start, end, document.getElementById("scale").value);
  ticks.forEach((tk) => {
    const x = xFor(tk.date);
    line(svg, x, PV.top - 6, x, height - 6, "gp-grid");
    text(svg, x + 2, PV.top - 10, tk.label, "gp-axis-label");
  });
  // today marker
  const today = new Date();
  if (today >= start && today <= end) {
    const tx = xFor(today);
    const tl = line(svg, tx, PV.top - 6, tx, height - 6, "gp-today");
    tl.setAttribute("stroke", "#c00000");
  }

  // geometry cache for links
  const geom = {};

  rows.forEach((row, i) => {
    const t = parseRow(row);
    const rowTop = PV.top + i * (PV.rowH + PV.rowGap);
    const rowMid = rowTop + PV.rowH / 2;
    text(svg, 4, rowMid + 3, t.name || `Task ${i + 1}`, "gp-row-label");
    if (!t.valid) return;

    if (t.type === "milestone") {
      const cx = xFor(t.start);
      const s = PV.msSize;
      const dia = poly(svg, [[cx, rowMid - s / 2], [cx + s / 2, rowMid], [cx, rowMid + s / 2], [cx - s / 2, rowMid]], t.color, "gp-ms");
      dia.dataset.row = i; dia.dataset.mode = "move-ms";
      if (showDates) text(svg, cx + s / 2 + 4, rowMid + 3, fmtDate(t.start), "gp-datelbl");
      geom[t.name] = { left: cx - s / 2, right: cx + s / 2, mid: rowMid };
      attachDrag(dia, row, { dateFor, start, end });
      return;
    }

    const barTop = rowTop + (PV.rowH - PV.barH) / 2;
    const x1 = xFor(t.start);
    const x2 = xFor(t.end);
    const w = Math.max(3, x2 - x1);
    const bar = rect(svg, x1, barTop, w, PV.barH, t.color, "gp-bar");
    bar.dataset.row = i; bar.dataset.mode = "move";
    geom[t.name] = { left: x1, right: x2, mid: rowMid };
    if (t.progress > 0) {
      rect(svg, x1, barTop, Math.max(1, w * Math.min(t.progress, 100) / 100), PV.barH, darken(t.color), "gp-progress");
    }
    // resize handles
    const hL = rect(svg, x1 - 3, barTop, 6, PV.barH, "transparent", "gp-handle");
    hL.dataset.row = i; hL.dataset.mode = "resize-start";
    const hR = rect(svg, x2 - 3, barTop, 6, PV.barH, "transparent", "gp-handle");
    hR.dataset.row = i; hR.dataset.mode = "resize-end";

    attachDrag(bar, row, { dateFor, start, end });
    attachDrag(hL, row, { dateFor, start, end });
    attachDrag(hR, row, { dateFor, start, end });

    // optional labels
    if (showDates) {
      text(svg, x1, barTop + PV.barH + 9, fmtDate(t.start), "gp-datelbl");
      const el = text(svg, x2, barTop + PV.barH + 9, fmtDate(t.end), "gp-datelbl");
      el.setAttribute("text-anchor", "end");
    }
    const parts = [];
    if (t.progress > 0) parts.push(`${t.progress}%`);
    if (showWorkdays) parts.push(`${netWorkdays(t.start, t.end)} wd`);
    if (parts.length) text(svg, x2 + 5, rowMid + 3, parts.join(" · "), "gp-datelbl");
  });

  // dependency links
  readLinks().forEach((lk) => {
    const a = geom[lk.from]; const b = geom[lk.to];
    if (!a || !b) return;
    const [ax, ay, bx, by] = anchorPoints(a, b, lk.style);
    const midX = bx >= ax ? (ax + bx) / 2 : ax + 10;
    const d = `M ${ax} ${ay} L ${midX} ${ay} L ${midX} ${by} L ${bx} ${by}`;
    const p = document.createElementNS(SVGNS, "path");
    p.setAttribute("d", d); p.setAttribute("class", "gp-link");
    p.setAttribute("stroke", lk.color); p.setAttribute("marker-end", "url(#gp-arrow)");
    svg.appendChild(p);
  });

  // arrow marker def
  const defs = document.createElementNS(SVGNS, "defs");
  defs.innerHTML = `<marker id="gp-arrow" markerWidth="8" markerHeight="8" refX="6" refY="3" orient="auto"><path d="M0,0 L6,3 L0,6 Z" fill="context-stroke"/></marker>`;
  svg.insertBefore(defs, svg.firstChild);

  host.innerHTML = "";
  host.appendChild(svg);
}

function parseRow(row) {
  const name = row.querySelector(".task-name").value.trim();
  const type = row.querySelector(".task-type").value;
  const startV = row.querySelector(".task-start").value;
  const endV = row.querySelector(".task-end").value;
  const color = row.querySelector(".task-color").value;
  const progress = Number(row.querySelector(".task-progress").value) || 0;
  const valid = name && startV && (type === "milestone" || endV);
  return {
    name, type, color, progress, valid,
    start: startV ? new Date(startV) : null,
    end: (type === "milestone" ? (startV ? new Date(startV) : null) : (endV ? new Date(endV) : null))
  };
}

// SVG element helpers
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

// Drag interaction: updates the row's date inputs live, then re-renders.
function attachDrag(el, row, ctx) {
  el.addEventListener("pointerdown", (ev) => {
    ev.preventDefault();
    const svg = el.ownerSVGElement;
    const mode = el.dataset.mode;
    const startX = ev.clientX;
    const startInput = row.querySelector(".task-start");
    const endInput = row.querySelector(".task-end");
    const origStart = startInput.value ? new Date(startInput.value) : new Date();
    const origEnd = endInput.value ? new Date(endInput.value) : new Date(origStart);
    const pxPerMs = (svg.viewBox.baseVal.width - PV.labelW - PV.padRight) / (ctx.end - ctx.start);

    const onMove = (e) => {
      const dxMs = (e.clientX - startX) / pxPerMs;
      const dDays = Math.round(dxMs / 864e5);
      if (mode === "move" || mode === "move-ms") {
        const ns = addDays(origStart, dDays);
        startInput.value = toInputDate(ns);
        if (mode === "move") endInput.value = toInputDate(addDays(origEnd, dDays));
      } else if (mode === "resize-start") {
        const ns = addDays(origStart, dDays);
        if (ns < origEnd) startInput.value = toInputDate(ns);
      } else if (mode === "resize-end") {
        const ne = addDays(origEnd, dDays);
        if (ne > origStart) endInput.value = toInputDate(ne);
      }
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

// --- Sync positions back from the slide ------------------------------------

async function syncFromSlide() {
  const tasks = readTasks();
  const { start, end } = previewRange(tasks);
  const rangeStart = new Date(document.getElementById("startDate").value);
  const rangeEnd = new Date(document.getElementById("endDate").value);
  const s = isValidDate(rangeStart) ? rangeStart : start;
  const e = isValidDate(rangeEnd) ? rangeEnd : end;
  const plotLeft = LAYOUT.chartLeft + LAYOUT.labelWidth;
  const plotWidth = LAYOUT.chartWidth - LAYOUT.labelWidth;
  const totalMs = e - s;
  const dateFor = (x) => new Date(s.getTime() + ((x - plotLeft) / plotWidth) * totalMs);

  try {
    setStatus("Reading positions from slide…");
    const map = {};
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shapes = slide.shapes;
      shapes.load("items/name,items/left,items/width");
      await context.sync();
      shapes.items.forEach((sh) => {
        if (!sh.name) return;
        if (sh.name.indexOf("Gantt-MS-") === 0) {
          map[sh.name.slice(9)] = { type: "milestone", left: sh.left, width: sh.width };
        } else if (sh.name.indexOf("Gantt-") === 0 && sh.name.indexOf("-progress") < 0) {
          map[sh.name.slice(6)] = { type: "bar", left: sh.left, width: sh.width };
        }
      });
    });

    let changed = 0;
    taskRows().forEach((row) => {
      const name = row.querySelector(".task-name").value.trim();
      const info = map[name];
      if (!info) return;
      if (info.type === "milestone") {
        const cx = info.left + info.width / 2;
        row.querySelector(".task-start").value = toInputDate(dateFor(cx));
      } else {
        row.querySelector(".task-start").value = toInputDate(dateFor(info.left));
        row.querySelector(".task-end").value = toInputDate(dateFor(info.left + info.width));
      }
      changed += 1;
    });
    renderPreview();
    setStatus(changed ? `Synced ${changed} element(s) from the slide.` : "No Gantt shapes found on the slide.", changed ? "success" : "error");
  } catch (err) {
    console.error(err);
    setStatus("Error: " + (err.message || err), "error");
  }
}

function setStatus(msg, type = "") {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = "status" + (type ? " " + type : "");
}