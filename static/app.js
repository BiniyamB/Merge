"use strict";

const state = { mode: null, files: [], result: null, chart: null };

const $ = (id) => document.getElementById(id);

const dz = $("dropzone");
const fileInput = $("file-input");
const fileList = $("file-list");
const mergeBtn = $("merge-btn");
const errorBanner = $("error-banner");

/* ------------------------------------------------------------------ */
/* Particle system on canvas                                            */
/* ------------------------------------------------------------------ */
(function initParticles() {
  const canvas = $("particle-canvas");
  const ctx = canvas.getContext("2d");
  let W, H;
  const particles = [];
  const PARTICLE_COUNT = 60;
  const COLORS = [
    "rgba(255, 60, 172, 0.5)",
    "rgba(43, 134, 197, 0.4)",
    "rgba(0, 232, 143, 0.4)",
    "rgba(168, 85, 247, 0.45)",
    "rgba(0, 212, 255, 0.35)",
    "rgba(255, 140, 50, 0.35)",
  ];

  function resize() {
    W = canvas.width = window.innerWidth;
    H = canvas.height = window.innerHeight;
  }
  resize();
  window.addEventListener("resize", resize);

  class Particle {
    constructor() { this.reset(true); }
    reset(init) {
      this.x = Math.random() * W;
      this.y = init ? Math.random() * H : H + 10;
      this.r = 1.5 + Math.random() * 2.5;
      this.vx = (Math.random() - 0.5) * 0.3;
      this.vy = -(0.15 + Math.random() * 0.4);
      this.color = COLORS[Math.floor(Math.random() * COLORS.length)];
      this.alpha = 0.3 + Math.random() * 0.7;
      this.pulse = Math.random() * Math.PI * 2;
      this.pulseSpeed = 0.01 + Math.random() * 0.02;
    }
    update() {
      this.x += this.vx;
      this.y += this.vy;
      this.pulse += this.pulseSpeed;
      const flicker = 0.5 + 0.5 * Math.sin(this.pulse);
      this.currentAlpha = this.alpha * (0.4 + 0.6 * flicker);
      if (this.y < -10 || this.x < -10 || this.x > W + 10) this.reset(false);
    }
    draw() {
      ctx.beginPath();
      ctx.arc(this.x, this.y, this.r, 0, Math.PI * 2);
      ctx.fillStyle = this.color;
      ctx.globalAlpha = this.currentAlpha;
      ctx.fill();
      ctx.globalAlpha = 1;
    }
  }

  for (let i = 0; i < PARTICLE_COUNT; i++) particles.push(new Particle());

  function drawConnections() {
    for (let i = 0; i < particles.length; i++) {
      for (let j = i + 1; j < particles.length; j++) {
        const dx = particles[i].x - particles[j].x;
        const dy = particles[i].y - particles[j].y;
        const dist = Math.sqrt(dx * dx + dy * dy);
        if (dist < 120) {
          ctx.beginPath();
          ctx.moveTo(particles[i].x, particles[i].y);
          ctx.lineTo(particles[j].x, particles[j].y);
          ctx.strokeStyle = `rgba(168, 85, 247, ${0.08 * (1 - dist / 120)})`;
          ctx.lineWidth = 0.5;
          ctx.stroke();
        }
      }
    }
  }

  function loop() {
    ctx.clearRect(0, 0, W, H);
    for (const p of particles) { p.update(); p.draw(); }
    drawConnections();
    requestAnimationFrame(loop);
  }
  loop();
})();

/* ------------------------------------------------------------------ */
/* Report type selection                                                */
/* ------------------------------------------------------------------ */
const MODE_INFO = {
  pos_decline: {
    label: "POS Decline",
    subtitle: "Drop one or more <code>.xls</code> / <code>.xlsx</code> POS decline reports, or click to browse.",
    columns: ["ACQUIRER","ISSUER","PAN","TRAN_DATE","TIME","TRANS_TYPE","AMOUNT","RESP_CODE","Reversal","FE UTRNNO","REFNUM","MERCHANT"],
  },
  pos_success: {
    label: "POS Success",
    subtitle: "Drop one or more <code>.xls</code> / <code>.xlsx</code> POS success reports, or click to browse.",
    columns: ["ACQUIRER","ISSUER","PAN","TRAN_DATE","TIME","TRANS_TYPE","AMOUNT","RESP_CODE","REFNUM","UTRNNO","MERCHANT"],
  },
  pos: {
    label: "POS",
    subtitle: "Drop one or more <code>.xls</code> / <code>.xlsx</code> POS transaction reports, or click to browse.",
    columns: ["ACQUIRER","ISSUER","CARD_NUMBER","TRANS_DATE","TRANS_TIME","TRANS_TYPE","AMOUNT","CURRENCY","RESP","RRN","TERMINAL_ID","ADDRESS"],
  },
  atm: {
    label: "ATM",
    subtitle: "Drop one or more <code>.xls</code> / <code>.xlsx</code> ATM transaction reports, or click to browse.",
    columns: ["ACQUIRER","ISSUER","CARD_NUMBER","TRANS_DATE","TRANS_TIME","TRANS_TYPE","AMOUNT","CURRENCY","RESP","RRN","UTRNNO","TERMINAL_ID","ADDRESS_NAME"],
  },
  qr: {
    label: "QR",
    subtitle: "Drop one or more <code>.xls</code> / <code>.xlsx</code> QR transfer-export reports, or click to browse.",
    columns: ["DESTINATION_BANK","SOURCE_BANK","TRX_DATE","DBTR_ACCT","CDTR_ACCT","AMOUNT","TX_ID","STATUS"],
  },
};

/* ------------------------------------------------------------------ */
/* Breadcrumbs                                                          */
/* ------------------------------------------------------------------ */
function setBreadcrumbs(steps) {
  const bc = $("breadcrumbs");
  const s2 = $("crumb-step2");
  const s3 = $("crumb-step3");
  const sep3 = $("crumb-sep3");
  if (!steps || !steps.length) {
    bc.classList.add("hidden");
    return;
  }
  bc.classList.remove("hidden");
  s2.textContent = steps[0] || "";
  s2.classList.toggle("crumb-current", steps.length === 1);
  if (steps.length > 1) {
    sep3.classList.remove("hidden");
    s3.textContent = steps[1];
    s3.classList.add("crumb-current");
  } else {
    sep3.classList.add("hidden");
    s3.textContent = "";
    s3.classList.remove("crumb-current");
  }
}

function resetToPicker() {
  state.mode = null;
  state.files = [];
  state.result = null;
  renderFileList();
  setBreadcrumbs([]);
  errorBanner.classList.add("hidden");
  $("warn-banner").classList.add("hidden");
  $("results-card").classList.add("hidden");
  $("upload-card").classList.add("hidden");
  $("mode-chip").classList.add("hidden");
  $("mode-card").classList.remove("hidden");
}

function populateSortColumns(mode) {
  const sortCol = $("sort-col");
  sortCol.innerHTML = "";
  const def = document.createElement("option");
  def.value = "date_time";
  def.textContent = "Date \u00b7 Time (default)";
  sortCol.appendChild(def);
  const cols = (MODE_INFO[mode] && MODE_INFO[mode].columns) || [];
  for (const col of cols) {
    const opt = document.createElement("option");
    opt.value = col;
    opt.textContent = col;
    sortCol.appendChild(opt);
  }
}

function selectMode(mode) {
  state.mode = mode;
  state.files = [];
  state.result = null;
  renderFileList();
  populateSortColumns(mode);
  setBreadcrumbs([MODE_INFO[mode].label]);
  errorBanner.classList.add("hidden");
  $("warn-banner").classList.add("hidden");
  $("results-card").classList.add("hidden");
  $("mode-card").classList.add("hidden");
  $("upload-card").classList.remove("hidden");
  $("mode-chip").classList.remove("hidden");
  $("mode-chip-label").textContent = MODE_INFO[mode].label;
  $("upload-subtitle").innerHTML = MODE_INFO[mode].subtitle;
  window.scrollTo({ top: 0, behavior: "smooth" });
}

document.querySelectorAll(".mode-option").forEach((btn) =>
  btn.addEventListener("click", () => selectMode(btn.dataset.mode))
);
$("change-mode").addEventListener("click", (e) => {
  e.preventDefault();
  resetToPicker();
});
$("crumb-home").addEventListener("click", (e) => {
  e.preventDefault();
  resetToPicker();
});

/* ------------------------------------------------------------------ */
/* File selection                                                       */
/* ------------------------------------------------------------------ */
function formatBytes(bytes) {
  if (bytes === 0) return "0 B";
  const k = 1024;
  const sizes = ["B", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return `${(bytes / Math.pow(k, i)).toFixed(1)} ${sizes[i]}`;
}

function addFiles(fileListLike) {
  for (const file of fileListLike) {
    if (!/\.(xls|xlsx)$/i.test(file.name)) {
      showError(`"${file.name}" is not an Excel file (.xls / .xlsx).`);
      continue;
    }
    const dup = state.files.some(
      (f) => f.name === file.name && f.size === file.size && f.lastModified === file.lastModified
    );
    if (!dup) state.files.push(file);
  }
  renderFileList();
}

function removeFile(index) {
  state.files.splice(index, 1);
  renderFileList();
}

function renderFileList() {
  fileList.innerHTML = "";
  state.files.forEach((f, i) => {
    const li = document.createElement("li");
    const name = document.createElement("span");
    name.className = "fname";
    name.textContent = f.name;
    name.title = f.name;
    const meta = document.createElement("span");
    meta.className = "fmeta";
    meta.textContent = formatBytes(f.size);
    const x = document.createElement("button");
    x.className = "x";
    x.type = "button";
    x.textContent = "\u2715";
    x.title = "Remove";
    x.addEventListener("click", () => removeFile(i));
    li.append(name, meta, x);
    fileList.appendChild(li);
  });
  mergeBtn.disabled = state.files.length === 0;
}

dz.addEventListener("click", () => fileInput.click());
dz.addEventListener("dragover", (e) => {
  e.preventDefault();
  dz.classList.add("dragover");
});
dz.addEventListener("dragleave", () => dz.classList.remove("dragover"));
dz.addEventListener("drop", (e) => {
  e.preventDefault();
  dz.classList.remove("dragover");
  addFiles(e.dataTransfer.files);
});
fileInput.addEventListener("change", () => {
  addFiles(fileInput.files);
  fileInput.value = "";
});

/* ------------------------------------------------------------------ */
/* Error banner                                                         */
/* ------------------------------------------------------------------ */
let errorTimer = null;
function showError(message) {
  const textEl = $("error-text");
  if (textEl) {
    textEl.textContent = message;
  } else {
    errorBanner.textContent = message;
  }
  errorBanner.classList.remove("hidden");
  clearTimeout(errorTimer);
  errorTimer = setTimeout(() => errorBanner.classList.add("hidden"), 10000);
}

/* ------------------------------------------------------------------ */
/* Merge                                                                */
/* ------------------------------------------------------------------ */
mergeBtn.addEventListener("click", merge);

async function merge() {
  if (state.files.length === 0 || !state.mode) return;

  mergeBtn.disabled = true;
  mergeBtn.querySelector(".btn-label").textContent = "Merging";
  const spinner = document.createElement("span");
  spinner.className = "spinner";
  mergeBtn.appendChild(spinner);
  errorBanner.classList.add("hidden");

  const formData = new FormData();
  state.files.forEach((f) => formData.append("files", f, f.name));
  formData.append("mode", state.mode);
  formData.append("sort_by", $("sort-col").value || "date_time");
  formData.append("sort_dir", $("sort-dir").value || "asc");

  try {
    const res = await fetch("/merge", { method: "POST", body: formData });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Merge failed.");
    state.result = data;
    renderResults(data);
    setBreadcrumbs([MODE_INFO[state.mode].label, "Merged report"]);
    $("results-card").classList.remove("hidden");
    $("results-card").scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (err) {
    showError(err.message || "Something went wrong while merging.");
  } finally {
    spinner.remove();
    mergeBtn.disabled = state.files.length === 0;
    mergeBtn.querySelector(".btn-label").textContent = "Merge reports";
  }
}

/* ------------------------------------------------------------------ */
/* Results rendering                                                    */
/* ------------------------------------------------------------------ */
function el(tag, cls, text) {
  const node = document.createElement(tag);
  if (cls) node.className = cls;
  if (text !== undefined) node.textContent = text;
  return node;
}

function renderResults(data) {
  const sortLabel = (data.sort_by && data.sort_by !== "date_time")
    ? ` \u00b7 sorted by ${data.sort_by} (${data.sort_dir === "desc" ? "desc" : "asc"})`
    : (data.sort_dir === "desc" ? " \u00b7 sorted by date & time (desc)" : "");
  $("result-subtitle").textContent =
    `${data.per_file.filter((p) => p.status === "ok").length} file(s) merged \u00b7 ` +
    `${data.total_rows.toLocaleString()} transactions \u00b7 ${data.from_date} \u2192 ${data.to_date}` +
    sortLabel;
  renderWarnings(data.warnings);
  renderStats(data);
  renderFilesTable(data.per_file);
  previewSort.col = "";
  previewSort.dir = "asc";
  renderPreview(data);
  renderChart(data.resp_counts);
  const dl = $("download-btn");
  dl.href = `/download/${data.token}`;
  dl.setAttribute("download", data.filename);
  initFilterPanel(data);
}

function renderStats(data) {
  const stats = $("stats");
  stats.innerHTML = "";
  const cards = [
    ["Transactions", data.total_rows.toLocaleString()],
    ["Date range", data.from_date === data.to_date ? data.from_date : `${data.from_date} \u2192 ${data.to_date}`],
    ["Files merged", data.per_file.filter((p) => p.status === "ok").length],
    ["Response codes", Object.keys(data.resp_counts).length],
  ];
  for (const [label, value] of cards) {
    const s = el("div", "stat");
    s.appendChild(el("div", "n", value));
    s.appendChild(el("div", "l", label));
    stats.appendChild(s);
  }
}

function renderFilesTable(perFile) {
  const table = $("files-table");
  table.innerHTML = "";
  const thead = document.createElement("thead");
  const hr = document.createElement("tr");
  for (const h of ["File", "Status", "Rows", "Blank cols removed", "Notes"]) {
    hr.appendChild(el("th", null, h));
  }
  thead.appendChild(hr);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (const p of perFile) {
    const tr = document.createElement("tr");
    tr.appendChild(el("td", null, p.filename));
    if (p.status === "ok") {
      tr.appendChild(el("td", null, null)).appendChild(el("span", "badge-ok", "ok"));
      tr.appendChild(el("td", null, String(p.data_rows)));
      tr.appendChild(el("td", null, p.blank_columns.length ? p.blank_columns.join(", ") : "\u2014"));
      const notes = [];
      if (p.dropped_columns.length) notes.push(`unknown cols: ${p.dropped_columns.join(", ")}`);
      if (p.extra_columns && p.extra_columns.length) notes.push(`removed (not in sample): ${p.extra_columns.join(", ")}`);
      notes.push(...p.warnings);
      const td = el("td", "muted");
      if (p.order_mismatch) {
        td.appendChild(el("span", "badge-warn", "reshuffled"));
        td.appendChild(document.createTextNode("  " + (p.column_order || []).join(", ")));
        if (notes.length) td.appendChild(document.createElement("br"));
      }
      if (notes.length) td.appendChild(document.createTextNode(notes.join(" \u00b7 ")));
      if (!td.childNodes.length) td.textContent = "\u2014";
      tr.appendChild(td);
    } else {
      tr.appendChild(el("td", null, null)).appendChild(el("span", "badge-err", "error"));
      tr.appendChild(el("td", null, "\u2014"));
      tr.appendChild(el("td", null, "\u2014"));
      tr.appendChild(el("td", null, p.error));
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
}

function renderPreview(data) {
  $("preview-count").textContent = `first ${Math.min(data.preview.length, 50)} of ${data.total_rows.toLocaleString()} rows`;
  const table = $("preview-table");
  table.innerHTML = "";
  const thead = document.createElement("thead");
  const hr = document.createElement("tr");
  for (const h of data.columns) {
    const th = el("th", "preview-sortable", h);
    th.title = "Click to sort preview";
    th.dataset.col = h;
    if (previewSort.col === h) {
      th.classList.add("preview-sorted");
      th.appendChild(el("span", "preview-sort-arrow", previewSort.dir === "desc" ? " \u25bc" : " \u25b2"));
    }
    th.addEventListener("click", () => togglePreviewSort(h, data.columns));
    hr.appendChild(th);
  }
  thead.appendChild(hr);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  let rows = data.preview;
  if (previewSort.col) {
    rows = rows.slice().sort((a, b) => {
      const av = a[previewSort.col];
      const bv = b[previewSort.col];
      const cmp = smartPreviewCompare(av, bv);
      return previewSort.dir === "desc" ? -cmp : cmp;
    });
  }
  for (const row of rows) {
    const tr = document.createElement("tr");
    for (const col of data.columns) {
      const v = row[col];
      tr.appendChild(el("td", null, v === "" || v === null || v === undefined ? "\u2014" : String(v)));
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
}

const previewSort = { col: "", dir: "asc" };

function togglePreviewSort(col, columns) {
  if (previewSort.col === col) {
    previewSort.dir = previewSort.dir === "asc" ? "desc" : "asc";
  } else {
    previewSort.col = col;
    previewSort.dir = "asc";
  }
  if (state.result) renderPreview(state.result);
}

function smartPreviewCompare(a, b) {
  const ae = a === "" || a === null || a === undefined;
  const be = b === "" || b === null || b === undefined;
  if (ae && be) return 0;
  if (ae) return -1;
  if (be) return 1;
  const sa = String(a).trim();
  const sb = String(b).trim();
  const na = Number(sa);
  const nb = Number(sb);
  if (sa !== "" && sb !== "" && !isNaN(na) && !isNaN(nb) && /^[+-]?\d+(\.\d+)?$/.test(sa) && /^[+-]?\d+(\.\d+)?$/.test(sb)) {
    return na - nb;
  }
  return sa.localeCompare(sb, undefined, { numeric: true, sensitivity: "base" });
}

function renderWarnings(warnings) {
  const banner = $("warn-banner");
  const list = $("warn-list");
  list.innerHTML = "";
  if (!warnings || !warnings.length) {
    banner.classList.add("hidden");
    return;
  }
  for (const w of warnings) {
    const li = document.createElement("li");
    li.textContent = w;
    list.appendChild(li);
  }
  banner.classList.remove("hidden");
}

const CHART_COLORS = [
  "#ff3cac", "#2b86c5", "#00e88f", "#a855f7",
  "#ff8c32", "#ff4757", "#00d4ff", "#ffd93d",
  "#7c3aed", "#a3ff12",
];

function renderChart(respCounts) {
  const canvas = $("resp-chart");
  const fallback = $("chart-fallback");
  const labels = Object.keys(respCounts).sort((a, b) => respCounts[b] - respCounts[a]);
  const values = labels.map((l) => respCounts[l]);

  if (typeof window.Chart === "undefined") {
    canvas.classList.add("hidden");
    fallback.classList.remove("hidden");
    fallback.innerHTML = "";
    const max = Math.max(...values, 1);
    labels.forEach((l, i) => {
      const row = el("div", "bar-row");
      row.appendChild(el("div", null, l));
      const track = el("div", "bar-track");
      const fill = el("div", "bar-fill");
      fill.style.width = `${Math.max((values[i] / max) * 100, 2)}%`;
      track.appendChild(fill);
      row.appendChild(track);
      row.appendChild(el("div", "bar-n", String(values[i])));
      fallback.appendChild(row);
    });
    return;
  }

  fallback.classList.add("hidden");
  canvas.classList.remove("hidden");
  if (state.chart) state.chart.destroy();
  state.chart = new Chart(canvas, {
    type: "bar",
    data: {
      labels,
      datasets: [{
        label: "Transactions",
        data: values,
        backgroundColor: labels.map((_, i) => CHART_COLORS[i % CHART_COLORS.length]),
        borderRadius: 8,
        maxBarThickness: 44,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 800, easing: "easeOutQuart" },
      plugins: { legend: { display: false } },
      scales: {
        y: {
          beginAtZero: true,
          ticks: { precision: 0, color: "#5a6a9a" },
          grid: { color: "rgba(255,255,255,0.04)" },
        },
        x: {
          grid: { display: false },
          ticks: { color: "#5a6a9a" },
        },
      },
    },
  });
}

/* ================================================================== */
/* FILTER PANEL                                                        */
/* ================================================================== */
const filterState = {
  sheets: [],          // [{name, filters: {col: val}, rowCount}, ...]
  allColumns: [],      // column names from the merged result
  allRecords: [],      // preview records from the merged result
  totalRows: 0,        // total merged row count
  uniqueValues: {},    // column -> sorted unique values (from server)
};

const filterColEl = $("filter-col");
const filterValEl = $("filter-val");
const filterCountEl = $("filter-count");
const addSheetBtn = $("add-sheet-btn");
const sheetListEl = $("sheet-list");
const filterActionsEl = $("filter-actions");
const filterDownloadBtn = $("filter-download-btn");
const clearSheetsBtn = $("clear-sheets-btn");

function initFilterPanel(data) {
  filterState.sheets = [];
  filterState.allColumns = data.columns || [];
  filterState.allRecords = data.preview || [];
  filterState.totalRows = data.total_rows || 0;
  filterState.uniqueValues = data.unique_values || {};

  // Populate column dropdown
  filterColEl.innerHTML = '<option value="">Select column...</option>';
  for (const col of filterState.allColumns) {
    const opt = document.createElement("option");
    opt.value = col;
    opt.textContent = col;
    filterColEl.appendChild(opt);
  }
  filterValEl.innerHTML = '<option value="">Select a column first...</option>';
  filterValEl.disabled = true;
  filterCountEl.textContent = "-";
  addSheetBtn.disabled = true;
  renderSheetList();
}

filterColEl.addEventListener("change", () => {
  const col = filterColEl.value;
  filterValEl.innerHTML = "";
  if (!col) {
    filterValEl.innerHTML = '<option value="">Select a column first...</option>';
    filterValEl.disabled = true;
    filterCountEl.textContent = "-";
    addSheetBtn.disabled = true;
    return;
  }
  const vals = filterState.uniqueValues[col] || [];
  filterValEl.innerHTML = '<option value="">All (no filter)</option>';
  for (const v of vals) {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v.length > 60 ? v.slice(0, 57) + "..." : v;
    opt.title = v;
    filterValEl.appendChild(opt);
  }
  filterValEl.disabled = false;
  updateFilterCount();
});

filterValEl.addEventListener("change", updateFilterCount);

function updateFilterCount() {
  const col = filterColEl.value;
  const val = filterValEl.value;
  if (!col) {
    filterCountEl.textContent = "-";
    addSheetBtn.disabled = true;
    return;
  }
  // Show total rows for "All", or count matching preview rows as estimate
  let count;
  if (!val) {
    count = filterState.totalRows;
  } else {
    const valUpper = val.toUpperCase();
    count = filterState.allRecords.filter(
      (r) => String(r[col] || "").toUpperCase() === valUpper
    ).length;
    // If we match the preview count, note it may be more in full data
    const uniqueCount = (filterState.uniqueValues[col] || []).length;
    if (uniqueCount > 0) {
      filterCountEl.textContent = `~${count.toLocaleString()} rows`;
      addSheetBtn.disabled = false;
      return;
    }
  }
  filterCountEl.textContent = `${count.toLocaleString()} rows`;
  addSheetBtn.disabled = false;
}

addSheetBtn.addEventListener("click", () => {
  const col = filterColEl.value;
  const val = filterValEl.value;
  if (!col) return;

  const filters = {};
  let name = "All";
  if (val) {
    filters[col] = val;
    // Truncate value for sheet name (max 31 chars total for Excel)
    const valShort = val.length > 15 ? val.slice(0, 12) + "..." : val;
    name = `${col}_${valShort}`;
  } else {
    name = "All";
  }

  // Count matching rows
  let count;
  if (!val) {
    count = filterState.totalRows;
  } else {
    const valUpper = val.toUpperCase();
    count = filterState.allRecords.filter(
      (r) => String(r[col] || "").toUpperCase() === valUpper
    ).length;
  }

  // Check for duplicate
  const key = JSON.stringify(filters);
  const dup = filterState.sheets.find((s) => JSON.stringify(s.filters) === key);
  if (dup) return;

  // Excel sheet name max 31 chars
  if (name.length > 31) name = name.slice(0, 31);

  filterState.sheets.push({ name, filters, rowCount: count });
  renderSheetList();

  // Reset dropdowns
  filterColEl.value = "";
  filterValEl.innerHTML = '<option value="">Select a column first...</option>';
  filterValEl.disabled = true;
  filterCountEl.textContent = "-";
  addSheetBtn.disabled = true;
});

function removeSheet(index) {
  filterState.sheets.splice(index, 1);
  renderSheetList();
}

function renderSheetList() {
  sheetListEl.innerHTML = "";
  filterActionsEl.style.display = filterState.sheets.length ? "flex" : "none";

  filterState.sheets.forEach((s, i) => {
    const li = el("li", "sheet-item");

    const num = el("span", "sheet-num", String(i + 1));
    li.appendChild(num);

    const nameEl = el("span", "sheet-name", s.name);
    li.appendChild(nameEl);

    const filterDesc = el("span", "sheet-filter-desc");
    if (Object.keys(s.filters).length === 0) {
      filterDesc.textContent = "All rows (unfiltered)";
    } else {
      const parts = Object.entries(s.filters).map(
        ([k, v]) => `${k} = ${v}`
      );
      filterDesc.textContent = parts.join(" + ");
    }
    li.appendChild(filterDesc);

    const rowsEl = el("span", "sheet-rows", `~${s.rowCount.toLocaleString()} rows`);
    li.appendChild(rowsEl);

    const removeBtn = el("button", "sheet-remove");
    removeBtn.type = "button";
    removeBtn.title = "Remove sheet";
    removeBtn.textContent = "\u2715";
    removeBtn.addEventListener("click", () => removeSheet(i));
    li.appendChild(removeBtn);

    sheetListEl.appendChild(li);
  });
}

clearSheetsBtn.addEventListener("click", () => {
  filterState.sheets = [];
  renderSheetList();
});

filterDownloadBtn.addEventListener("click", async () => {
  if (!filterState.sheets.length || !state.result) return;

  filterDownloadBtn.disabled = true;
  filterDownloadBtn.querySelector(".btn-label").textContent = "Generating...";
  const spinner = document.createElement("span");
  spinner.className = "spinner";
  filterDownloadBtn.appendChild(spinner);

  try {
    const formData = new FormData();
    formData.append("token", state.result.token);
    formData.append("sheets", JSON.stringify(filterState.sheets));

    const res = await fetch("/filter-download", { method: "POST", body: formData });
    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      throw new Error(data.error || "Filter download failed.");
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const disposition = res.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename=([^;]+)/);
    a.download = match ? match[1].replace(/"/g, "") : "Filtered_Report.xlsx";
    a.href = url;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  } catch (err) {
    showError(err.message || "Failed to download filtered report.");
  } finally {
    spinner.remove();
    filterDownloadBtn.disabled = false;
    filterDownloadBtn.querySelector(".btn-label").textContent = "Download filtered report";
  }
});
