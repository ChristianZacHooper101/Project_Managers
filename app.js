const STORAGE_KEY = "pm_dashboard_state_v3";

const monthLabels = [
  "Jan", "Feb", "Mar", "Apr", "May", "Jun",
  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
];

const defaultState = {
  projectName: "",
  projectManager: "",
  statusDate: "2021-01-20",
  projectStatus: "On Track",
  progress: 45,
  planned: [50000, 45000, 45000, 50000],
  spent: [25000, 23000, 23000, 23000],
  risks: [2, 3, 2],
  issues: [4, 3, 2],
  pendingActions: 2,
  pendingDecisions: 3,
  openChanges: 5,
  taskNames: ["Analysis", "Design", "Development", "Testing", "Implement"],
  taskStarts: [1, 3, 5, 8, 11],
  taskDurations: [2, 2, 3, 3, 2],
  notes: "Project progress is on track\nResource plan is intact\nBusiness requirements approved\nPM on leave for 2 weeks\nNew resource is onboard\nPublic holiday on 12th Sep",
  sectionVisibility: {
    kpiStrip: true,
    gaugeCard: true,
    budgetCard: true,
    riskCard: true,
    actionsCard: true,
    timelinePanel: true,
    notesPanel: true
  }
};

const state = JSON.parse(JSON.stringify(defaultState));

const byId = (id) => document.getElementById(id);

const els = {
  projectName: byId("projectName"),
  projectManager: byId("projectManager"),
  statusDate: byId("statusDate"),
  statusDatePreview: byId("statusDatePreview"),
  statusDatePrint: byId("statusDatePrint"),
  projectStatus: byId("projectStatus"),
  statusChip: byId("statusChip"),
  projectNamePrint: byId("projectNamePrint"),
  projectManagerPrint: byId("projectManagerPrint"),
  kpiProgress: byId("kpiProgress"),
  kpiUtilization: byId("kpiUtilization"),
  kpiPlanned: byId("kpiPlanned"),
  kpiSpent: byId("kpiSpent"),
  kpiRiskIssues: byId("kpiRiskIssues"),
  kpiActions: byId("kpiActions"),
  budgetTotals: byId("budgetTotals"),
  riskTotals: byId("riskTotals"),
  actionTotals: byId("actionTotals"),
  progress: byId("progress"),
  progressValue: byId("progressValue"),
  gauge: byId("gauge"),
  plannedInput: byId("plannedInput"),
  spentInput: byId("spentInput"),
  riskInput: byId("riskInput"),
  issuesInput: byId("issuesInput"),
  pendingInput: byId("pendingInput"),
  decisionsInput: byId("decisionsInput"),
  changesInput: byId("changesInput"),
  taskNames: byId("taskNames"),
  taskStarts: byId("taskStarts"),
  taskDurations: byId("taskDurations"),
  updateTimeline: byId("updateTimeline"),
  timelineBars: byId("timelineBars"),
  timelineWarning: byId("timelineWarning"),
  monthAxis: byId("monthAxis"),
  notes: byId("notes"),
  notesPreview: byId("notesPreview"),
  saveBtn: byId("saveBtn"),
  resetBtn: byId("resetBtn"),
  printBtn: byId("printBtn"),
  downloadPdfBtn: byId("downloadPdfBtn"),
  saveState: byId("saveState"),
  exportBtn: byId("exportBtn"),
  importBtn: byId("importBtn"),
  importFile: byId("importFile")
};

const toggleButtons = Array.from(document.querySelectorAll(".toggle-btn"));

let budgetChart;
let riskChart;
let actionsChart;
let saveTimer;
let viewportTimer;

function parseNumberList(value) {
  return value
    .split(",")
    .map((v) => Number(v.trim()))
    .filter((v) => Number.isFinite(v));
}

function ensureListLength(list, fallback, size) {
  const output = list.slice(0, size);
  while (output.length < size) output.push(fallback[output.length]);
  return output;
}

function formatDate(isoDate) {
  if (!isoDate) return "";
  const date = new Date(`${isoDate}T00:00:00`);
  if (Number.isNaN(date.getTime())) return "";
  const day = String(date.getDate()).padStart(2, "0");
  const month = monthLabels[date.getMonth()];
  const year = date.getFullYear();
  return `${day} ${month} ${year}`;
}

function setSaveState(text) {
  els.saveState.textContent = text;
}

function sum(list) {
  return list.reduce((acc, value) => acc + (Number(value) || 0), 0);
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString();
}

function updateToggleButton(button, isVisible) {
  const label = button.dataset.label || "Section";
  button.classList.toggle("is-on", isVisible);
  button.textContent = `${label}: ${isVisible ? "ON" : "OFF"}`;
}

function applyViewportMode() {
  const compact = window.innerWidth < 1320 || window.innerHeight < 860;
  document.body.classList.toggle("compact", compact);
}

function queueViewportModeUpdate() {
  window.clearTimeout(viewportTimer);
  viewportTimer = window.setTimeout(applyViewportMode, 60);
}

function applySectionVisibility() {
  toggleButtons.forEach((button) => {
    const targetId = button.dataset.target;
    const target = byId(targetId);
    if (!target) return;
    const isVisible = state.sectionVisibility[targetId] !== false;
    target.classList.toggle("is-hidden", !isVisible);
    updateToggleButton(button, isVisible);
  });
}

function updateDisplayValues() {
  const projectName = state.projectName.trim() || "Add text";
  const projectManager = state.projectManager.trim() || "Add text";
  const dateText = formatDate(state.statusDate) || "No date";

  els.projectNamePrint.textContent = projectName;
  els.projectManagerPrint.textContent = projectManager;
  els.statusDatePrint.textContent = dateText;
}

function updateNotesPreview() {
  const lines = state.notes
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  els.notesPreview.innerHTML = "";

  if (!lines.length) {
    const emptyLine = document.createElement("p");
    emptyLine.textContent = "No notes";
    els.notesPreview.appendChild(emptyLine);
    return;
  }

  lines.forEach((line) => {
    const p = document.createElement("p");
    p.textContent = `* ${line}`;
    els.notesPreview.appendChild(p);
  });
}

function updateKpis() {
  const totalPlanned = sum(state.planned);
  const totalSpent = sum(state.spent);
  const totalRemaining = Math.max(totalPlanned - totalSpent, 0);
  const utilization = totalPlanned > 0 ? (totalSpent / totalPlanned) * 100 : 0;
  const totalRisks = sum(state.risks);
  const totalIssues = sum(state.issues);
  const riskIssues = totalRisks + totalIssues;
  const actions = state.pendingActions + state.pendingDecisions + state.openChanges;

  els.kpiProgress.textContent = `${state.progress}%`;
  els.kpiUtilization.textContent = `${Math.round(utilization)}%`;
  els.kpiPlanned.textContent = formatNumber(totalPlanned);
  els.kpiSpent.textContent = formatNumber(totalSpent);
  els.kpiRiskIssues.textContent = formatNumber(riskIssues);
  els.kpiActions.textContent = formatNumber(actions);

  els.budgetTotals.textContent = `Planned: ${formatNumber(totalPlanned)} | Spent: ${formatNumber(totalSpent)} | Remaining: ${formatNumber(totalRemaining)}`;
  els.riskTotals.textContent = `Open Risks: ${formatNumber(totalRisks)} | Open Issues: ${formatNumber(totalIssues)}`;
  els.actionTotals.textContent = `Actions: ${formatNumber(state.pendingActions)} | Decisions: ${formatNumber(state.pendingDecisions)} | Changes: ${formatNumber(state.openChanges)}`;
}

function queueAutoSave() {
  setSaveState("Saving...");
  window.clearTimeout(saveTimer);
  saveTimer = window.setTimeout(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    setSaveState(`Saved ${new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}`);
  }, 400);
}

function syncInputsFromState() {
  els.projectName.value = state.projectName;
  els.projectManager.value = state.projectManager;
  els.statusDate.value = state.statusDate;
  els.projectStatus.value = state.projectStatus;
  els.progress.value = state.progress;
  els.plannedInput.value = state.planned.join(",");
  els.spentInput.value = state.spent.join(",");
  els.riskInput.value = state.risks.join(",");
  els.issuesInput.value = state.issues.join(",");
  els.pendingInput.value = state.pendingActions;
  els.decisionsInput.value = state.pendingDecisions;
  els.changesInput.value = state.openChanges;
  els.taskNames.value = state.taskNames.join(",");
  els.taskStarts.value = state.taskStarts.join(",");
  els.taskDurations.value = state.taskDurations.join(",");
  els.notes.value = state.notes;
  updateDisplayValues();
  updateNotesPreview();
  updateKpis();
}

function paintGauge() {
  const value = Math.min(Math.max(Number(state.progress), 0), 100);
  const degrees = (value / 100) * 180;
  els.progressValue.textContent = `${value}%`;
  els.gauge.style.background =
    `conic-gradient(from 270deg, #1f4574 0deg, #5f8fc2 ${degrees}deg, #d8e3ef ${degrees}deg, #d8e3ef 180deg, transparent 180deg)`;
}

function paintStatus() {
  const status = state.projectStatus;
  const cssMap = {
    "On Track": "status-on-track",
    "At Risk": "status-at-risk",
    "Delayed": "status-delayed",
    "Critical": "status-critical"
  };

  els.statusChip.textContent = status;
  els.statusChip.className = "status-chip";
  els.statusChip.classList.add(cssMap[status] || "status-on-track");
}

function paintDatePreview() {
  els.statusDatePreview.textContent = formatDate(state.statusDate);
  updateDisplayValues();
}

function buildCharts() {
  if (typeof Chart === "undefined") {
    setSaveState("Charts unavailable (offline CDN)");
    return;
  }

  const budgetCtx = byId("budgetChart");
  const riskCtx = byId("riskChart");
  const actionsCtx = byId("actionsChart");

  const remaining = state.planned.map((p, i) => Math.max(p - (state.spent[i] || 0), 0));

  budgetChart = new Chart(budgetCtx, {
    type: "bar",
    data: {
      labels: ["Resources", "Hardware", "Software", "Other"],
      datasets: [
        { label: "Planned", data: state.planned, backgroundColor: "#1e3c67" },
        { label: "Spent", data: state.spent, backgroundColor: "#5f8fc2" },
        { label: "Remaining", data: remaining, backgroundColor: "#a8c4e2" }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { ticks: { font: { size: 12 } } },
        y: { beginAtZero: true, ticks: { font: { size: 12 } } }
      },
      plugins: { legend: { position: "bottom", labels: { font: { size: 12 } } } }
    }
  });

  riskChart = new Chart(riskCtx, {
    type: "bar",
    data: {
      labels: ["High", "Medium", "Low"],
      datasets: [
        { label: "Open Risks", data: state.risks, backgroundColor: ["#1e3c67", "#5f8fc2", "#a8c4e2"] },
        { label: "Open Issues", data: state.issues, backgroundColor: ["#1e3c67", "#5f8fc2", "#a8c4e2"] }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: "bottom", labels: { font: { size: 12 } } } },
      scales: {
        x: { ticks: { font: { size: 12 } } },
        y: { beginAtZero: true, precision: 0, ticks: { font: { size: 12 } } }
      }
    }
  });

  actionsChart = new Chart(actionsCtx, {
    type: "bar",
    data: {
      labels: ["Actions"],
      datasets: [
        { label: "Pending Action Items", data: [state.pendingActions], backgroundColor: "#a8c4e2", stack: "stack1" },
        { label: "Pending Decisions", data: [state.pendingDecisions], backgroundColor: "#5f8fc2", stack: "stack1" },
        { label: "Open Change Requests", data: [state.openChanges], backgroundColor: "#1e3c67", stack: "stack1" }
      ]
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { beginAtZero: true, stacked: true, ticks: { font: { size: 12 } } },
        y: { stacked: true, ticks: { font: { size: 12 } } }
      },
      plugins: { legend: { position: "bottom", labels: { font: { size: 12 } } } }
    }
  });
}

function refreshCharts() {
  if (!budgetChart || !riskChart || !actionsChart) {
    updateKpis();
    return;
  }

  const remaining = state.planned.map((p, i) => Math.max(p - (state.spent[i] || 0), 0));

  budgetChart.data.datasets[0].data = state.planned;
  budgetChart.data.datasets[1].data = state.spent;
  budgetChart.data.datasets[2].data = remaining;
  budgetChart.update();

  riskChart.data.datasets[0].data = state.risks;
  riskChart.data.datasets[1].data = state.issues;
  riskChart.update();

  actionsChart.data.datasets[0].data = [state.pendingActions];
  actionsChart.data.datasets[1].data = [state.pendingDecisions];
  actionsChart.data.datasets[2].data = [state.openChanges];
  actionsChart.update();

  updateKpis();
}

function resizeChartsForPrint() {
  [budgetChart, riskChart, actionsChart].forEach((chart) => {
    if (!chart) return;
    const parent = chart.canvas.parentElement;
    const width = Math.max((parent && parent.clientWidth) || chart.width || 300, 300);
    const height = 104;
    chart.resize(width, height);
    chart.update("none");
  });
}

function resizeChartsAfterPrint() {
  [budgetChart, riskChart, actionsChart].forEach((chart) => {
    if (!chart) return;
    chart.resize();
    chart.update("none");
  });
}

function printDashboard() {
  resizeChartsForPrint();
  window.setTimeout(() => {
    window.print();
  }, 90);
}

async function downloadPdf() {
  if (typeof window.html2canvas === "undefined" || !window.jspdf || !window.jspdf.jsPDF) {
    window.alert("PDF library not loaded. Falling back to print.");
    printDashboard();
    return;
  }

  const btn = els.downloadPdfBtn;
  if (btn) {
    btn.disabled = true;
    btn.textContent = "Building PDF...";
  }

  try {
    document.body.classList.add("pdf-export");
    resizeChartsAfterPrint();
    resizeChartsForPrint();
    await new Promise((resolve) => window.setTimeout(resolve, 180));

    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
    const pageWidthMm = pdf.internal.pageSize.getWidth();
    const pageHeightMm = pdf.internal.pageSize.getHeight();
    const marginMm = 6;
    const renderWidthMm = pageWidthMm - marginMm * 2;
    const maxContentHeightMm = pageHeightMm - marginMm * 2;
    const blockSelectors = [
      ".page > header",
      ".summary",
      "#kpiStrip",
      ".health-section > h2",
      ".health-row-top",
      ".health-row-bottom",
      ".timeline-notes"
    ];
    const blocks = blockSelectors
      .map((selector) => document.querySelector(selector))
      .filter((el) => el && el.offsetParent !== null && el.getBoundingClientRect().height > 4);

    let cursorY = marginMm;
    let pageIndex = 0;

    for (const block of blocks) {
      const canvas = await window.html2canvas(block, {
        backgroundColor: "#ffffff",
        scale: 2,
        useCORS: true,
        logging: false
      });

      let drawWidthMm = renderWidthMm;
      let drawHeightMm = (canvas.height * drawWidthMm) / canvas.width;

      if (drawHeightMm > maxContentHeightMm) {
        const fitRatio = maxContentHeightMm / drawHeightMm;
        drawHeightMm = maxContentHeightMm;
        drawWidthMm = drawWidthMm * fitRatio;
      }

      if (cursorY + drawHeightMm > pageHeightMm - marginMm && cursorY > marginMm) {
        pdf.addPage();
        pageIndex += 1;
        cursorY = marginMm;
      }

      const x = marginMm + (renderWidthMm - drawWidthMm) / 2;
      const y = cursorY;
      pdf.addImage(canvas.toDataURL("image/png"), "PNG", x, y, drawWidthMm, drawHeightMm, `block-${pageIndex}-${cursorY}`, "FAST");

      cursorY += drawHeightMm + 2.2;
    }

    pdf.save("project-management-dashboard.pdf");
  } catch (_err) {
    window.alert("PDF export failed. Trying browser print instead.");
    printDashboard();
  } finally {
    document.body.classList.remove("pdf-export");
    resizeChartsAfterPrint();
    if (btn) {
      btn.disabled = false;
      btn.textContent = "Download PDF";
    }
  }
}

function renderMonthAxis() {
  els.monthAxis.innerHTML = "";
  monthLabels.forEach((m) => {
    const item = document.createElement("div");
    item.textContent = m;
    els.monthAxis.appendChild(item);
  });
}

function renderTimeline() {
  els.timelineBars.innerHTML = "";

  const count = Math.min(state.taskNames.length, state.taskStarts.length, state.taskDurations.length);
  for (let i = 0; i < count; i += 1) {
    const row = document.createElement("div");
    row.className = "bar-row";

    const bar = document.createElement("div");
    bar.className = "bar";
    bar.textContent = state.taskNames[i];

    const start = Math.min(Math.max(Number(state.taskStarts[i]), 1), 12);
    const duration = Math.min(Math.max(Number(state.taskDurations[i]), 1), 12);
    const maxDuration = Math.min(duration, 13 - start);

    bar.style.left = `${((start - 1) / 12) * 100}%`;
    bar.style.width = `${(maxDuration / 12) * 100}%`;

    row.appendChild(bar);
    els.timelineBars.appendChild(row);
  }
}

function syncSeriesFromInputs() {
  state.planned = ensureListLength(parseNumberList(els.plannedInput.value), defaultState.planned, 4);
  state.spent = ensureListLength(parseNumberList(els.spentInput.value), defaultState.spent, 4);
  state.risks = ensureListLength(parseNumberList(els.riskInput.value), defaultState.risks, 3);
  state.issues = ensureListLength(parseNumberList(els.issuesInput.value), defaultState.issues, 3);
  state.pendingActions = Math.max(Number(els.pendingInput.value) || 0, 0);
  state.pendingDecisions = Math.max(Number(els.decisionsInput.value) || 0, 0);
  state.openChanges = Math.max(Number(els.changesInput.value) || 0, 0);
  updateKpis();
}

function syncTimelineFromInputs() {
  const names = els.taskNames.value.split(",").map((v) => v.trim()).filter(Boolean);
  const starts = parseNumberList(els.taskStarts.value).map((v) => Math.min(Math.max(v, 1), 12));
  const durations = parseNumberList(els.taskDurations.value).map((v) => Math.min(Math.max(v, 1), 12));

  const minCount = Math.min(names.length, starts.length, durations.length);
  state.taskNames = minCount ? names.slice(0, minCount) : defaultState.taskNames.slice();
  state.taskStarts = minCount ? starts.slice(0, minCount) : defaultState.taskStarts.slice();
  state.taskDurations = minCount ? durations.slice(0, minCount) : defaultState.taskDurations.slice();

  if (names.length !== starts.length || names.length !== durations.length) {
    els.timelineWarning.textContent = "Timeline lists have different lengths. Showing only matching items.";
  } else {
    els.timelineWarning.textContent = "";
  }
}

function exportState() {
  const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "project-dashboard-data.json";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function importState(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const parsed = JSON.parse(String(reader.result));
      Object.assign(state, defaultState, parsed);
      state.sectionVisibility = Object.assign({}, defaultState.sectionVisibility, parsed.sectionVisibility || {});
      syncInputsFromState();
      paintGauge();
      paintStatus();
      paintDatePreview();
      refreshCharts();
      renderTimeline();
      updateNotesPreview();
      applySectionVisibility();
      queueAutoSave();
    } catch (_err) {
      window.alert("Invalid JSON file.");
    }
  };
  reader.readAsText(file);
}

function loadFromStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      setSaveState("Not saved");
      return;
    }

    const parsed = JSON.parse(raw);
    Object.assign(state, defaultState, parsed);
    state.sectionVisibility = Object.assign({}, defaultState.sectionVisibility, parsed.sectionVisibility || {});
    setSaveState("Loaded saved data");
  } catch (_e) {
    setSaveState("Load failed");
  }
}

function bindEvents() {
  els.projectName.addEventListener("input", (e) => {
    state.projectName = e.target.value;
    updateDisplayValues();
    queueAutoSave();
  });

  els.projectManager.addEventListener("input", (e) => {
    state.projectManager = e.target.value;
    updateDisplayValues();
    queueAutoSave();
  });

  els.statusDate.addEventListener("change", (e) => {
    state.statusDate = e.target.value;
    paintDatePreview();
    queueAutoSave();
  });

  els.projectStatus.addEventListener("change", (e) => {
    state.projectStatus = e.target.value;
    paintStatus();
    queueAutoSave();
  });

  els.progress.addEventListener("input", (e) => {
    state.progress = Number(e.target.value);
    paintGauge();
    updateKpis();
    queueAutoSave();
  });

  [
    els.plannedInput,
    els.spentInput,
    els.riskInput,
    els.issuesInput,
    els.pendingInput,
    els.decisionsInput,
    els.changesInput
  ].forEach((input) => {
    input.addEventListener("input", () => {
      syncSeriesFromInputs();
      refreshCharts();
      queueAutoSave();
    });
  });

  els.updateTimeline.addEventListener("click", () => {
    syncTimelineFromInputs();
    renderTimeline();
    queueAutoSave();
  });

  els.notes.addEventListener("input", (e) => {
    state.notes = e.target.value;
    updateNotesPreview();
    queueAutoSave();
  });

  toggleButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const targetId = button.dataset.target;
      const current = state.sectionVisibility[targetId] !== false;
      state.sectionVisibility[targetId] = !current;
      applySectionVisibility();
      queueAutoSave();
    });
  });

  els.saveBtn.addEventListener("click", () => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    setSaveState("Saved now");
  });

  els.exportBtn.addEventListener("click", exportState);

  els.importBtn.addEventListener("click", () => {
    els.importFile.click();
  });

  els.importFile.addEventListener("change", () => {
    const file = els.importFile.files && els.importFile.files[0];
    if (!file) return;
    importState(file);
    els.importFile.value = "";
  });

  els.resetBtn.addEventListener("click", () => {
    localStorage.removeItem(STORAGE_KEY);
    window.location.reload();
  });

  els.printBtn.addEventListener("click", printDashboard);
  if (els.downloadPdfBtn) {
    els.downloadPdfBtn.addEventListener("click", () => {
      downloadPdf();
    });
  }

  window.addEventListener("beforeprint", () => {
    window.setTimeout(resizeChartsForPrint, 50);
  });

  window.addEventListener("afterprint", () => {
    resizeChartsAfterPrint();
  });

  window.addEventListener("resize", queueViewportModeUpdate);
}

function init() {
  loadFromStorage();
  syncInputsFromState();
  paintGauge();
  paintStatus();
  paintDatePreview();
  renderMonthAxis();
  syncTimelineFromInputs();
  renderTimeline();
  buildCharts();
  applySectionVisibility();
  applyViewportMode();
  bindEvents();
}

init();
