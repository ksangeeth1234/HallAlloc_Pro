/**
 * app.js
 * ======
 * Main controller for the Lecture Hall Allocation & Availability Web Application.
 * Integrates processing engine, UI interactions, Chart.js visualizations, and report generation.
 */

import {
  validateWorkbook,
  loadWorkbook,
  runHallSearchLogic,
  computeOverallStats,
  computeMonthlySummary,
  computeHallUtilization,
  computeStatusBreakdown,
  parseTimeToMinutes,
  generateMonthOptions,
  formatDateSafe
} from './processing.js';

// Application State
const state = {
  activeView: 'dashboard',
  workbookBundle: null,
  classRows: [],
  halls: [],
  searchParams: {
    day: 'Sunday',
    time: '10:15',
    month: ''
  },
  lastSearchResults: null,
  overallStats: null,
  monthlySummary: [],
  hallUtilization: [],
  statusBreakdown: null,

  // Table pagination state
  filteredSchedule: [],
  currentPage: 1,
  pageSize: 15,

  // Chart instances
  charts: {
    trend: null,
    status: null
  }
};

// DOM Element References
const elements = {
  sidebar: document.getElementById('sidebar'),
  mobileToggle: document.getElementById('mobileToggle'),
  navItems: document.querySelectorAll('.nav-item'),
  views: document.querySelectorAll('.view'),
  pageTitle: document.getElementById('pageTitle'),
  pageSubtitle: document.getElementById('pageSubtitle'),

  // Status and file info
  statusText: document.getElementById('statusText'),
  loadedFileName: document.getElementById('loadedFileName'),
  quickFileInfo: document.getElementById('quickFileInfo'),
  toast: document.getElementById('toast'),
  toastMessage: document.getElementById('toastMessage'),

  // Form Search controls
  searchDay: document.getElementById('searchDay'),
  searchTime: document.getElementById('searchTime'),
  searchMonth: document.getElementById('searchMonth'),
  executeSearchBtn: document.getElementById('executeSearchBtn'),
  topRunSearchBtn: document.getElementById('topRunSearchBtn'),

  // KPI elements
  kpiTotalClasses: document.getElementById('kpiTotalClasses'),
  kpiConflictCount: document.getElementById('kpiConflictCount'),
  kpiPhysicalClasses: document.getElementById('kpiPhysicalClasses'),
  kpiPhysicalPercent: document.getElementById('kpiPhysicalPercent'),
  kpiHybridOnline: document.getElementById('kpiHybridOnline'),
  kpiHybridBreakdown: document.getElementById('kpiHybridBreakdown'),
  kpiOngoingClasses: document.getElementById('kpiOngoingClasses'),
  kpiStatusSubtext: document.getElementById('kpiStatusSubtext'),

  // Dashboard tables & results
  lastSearchTimestamp: document.getElementById('lastSearchTimestamp'),
  quickResultsBody: document.getElementById('quickResultsBody'),
  topHallsBody: document.getElementById('topHallsBody'),

  // File Upload Elements
  dropZone: document.getElementById('dropZone'),
  fileInput: document.getElementById('fileInput'),
  uploadFeedback: document.getElementById('uploadFeedback'),
  processingLoader: document.getElementById('processingLoader'),
  previewSection: document.getElementById('previewSection'),
  previewHead: document.getElementById('previewHead'),
  previewBody: document.getElementById('previewBody'),
  previewCountBadge: document.getElementById('previewCountBadge'),

  // Detailed Schedule Table & Pagination
  scheduleSearchInput: document.getElementById('scheduleSearchInput'),
  modeFilterSelect: document.getElementById('modeFilterSelect'),
  statusFilterSelect: document.getElementById('statusFilterSelect'),
  detailedScheduleBody: document.getElementById('detailedScheduleBody'),
  paginationInfo: document.getElementById('paginationInfo'),
  prevPageBtn: document.getElementById('prevPageBtn'),
  nextPageBtn: document.getElementById('nextPageBtn'),

  // Official Report Elements
  repSelectedDay: document.getElementById('repSelectedDay'),
  repSelectedTime: document.getElementById('repSelectedTime'),
  repSelectedMonth: document.getElementById('repSelectedMonth'),
  repTotalHalls: document.getElementById('repTotalHalls'),
  repFreeCount: document.getElementById('repFreeCount'),
  repOccupiedCount: document.getElementById('repOccupiedCount'),
  repUtilizationRate: document.getElementById('repUtilizationRate'),
  reportTableBody: document.getElementById('reportTableBody'),
  reportTimestamp: document.getElementById('reportTimestamp'),
  printReportBtn: document.getElementById('printReportBtn'),
  exportCsvBtn: document.getElementById('exportCsvBtn')
};

/* ── App Initialization ────────────────────────────────────────────────── */

async function init() {
  bindNavigation();
  bindFileUpload();
  bindSearchForm();
  bindDetailedFilters();
  bindReportActions();

  // Load sample default workbook if present in server
  await loadDefaultSampleWorkbook();
}

/**
 * Automatically fetch and load the default Excel file so the app works out-of-the-box.
 */
async function loadDefaultSampleWorkbook() {
  try {
    const response = await fetch('./Lecture_Hall_Allocation_for_ongoing_programs.xlsx');
    if (!response.ok) return;

    const blob = await response.blob();
    const file = new File([blob], 'Lecture_Hall_Allocation_for_ongoing_programs.xlsx', {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });

    await processUploadedFile(file);
    showToast('Loaded default Excel workbook successfully.', 'info');
  } catch (err) {
    console.log('Sample file auto-load skipped. Waiting for user upload.', err);
  }
}

/* ── Navigation Handler ────────────────────────────────────────────────── */

function bindNavigation() {
  elements.navItems.forEach(item => {
    item.addEventListener('click', () => {
      const view = item.dataset.view;
      switchView(view);
    });
  });

  elements.mobileToggle.addEventListener('click', () => {
    elements.sidebar.classList.toggle('open');
  });
}

function switchView(viewName) {
  state.activeView = viewName;

  // Update active sidebar nav
  elements.navItems.forEach(item => {
    item.classList.toggle('active', item.dataset.view === viewName);
  });

  // Toggle view containers
  elements.views.forEach(v => {
    v.classList.toggle('active', v.id === `${viewName}View`);
  });

  // Dynamic titles
  const titles = {
    dashboard: { title: 'Dashboard Overview', subtitle: 'Real-time lecture hall allocation and availability analysis' },
    upload: { title: 'Excel Workbook Upload', subtitle: 'Import .xlsx, .xls, or .xlsm files to update system schedule' },
    results: { title: 'Schedule Records & Search', subtitle: 'Comprehensive database of all class schedules and status conflicts' },
    report: { title: 'Official Availability Report', subtitle: 'Printable and exportable hall search summary report' }
  };

  if (titles[viewName]) {
    elements.pageTitle.textContent = titles[viewName].title;
    elements.pageSubtitle.textContent = titles[viewName].subtitle;
  }

  // Close mobile sidebar if open
  elements.sidebar.classList.remove('open');
}

/* ── File Upload & Validation Handlers ─────────────────────────────────── */

function bindFileUpload() {
  elements.dropZone.addEventListener('click', () => elements.fileInput.click());

  elements.fileInput.addEventListener('change', e => {
    const file = e.target.files[0];
    if (file) handleFileSelection(file);
  });

  ['dragenter', 'dragover'].forEach(evt => {
    elements.dropZone.addEventListener(evt, e => {
      e.preventDefault();
      elements.dropZone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach(evt => {
    elements.dropZone.addEventListener(evt, e => {
      e.preventDefault();
      elements.dropZone.classList.remove('dragover');
    });
  });

  elements.dropZone.addEventListener('drop', e => {
    const file = e.dataTransfer.files[0];
    if (file) handleFileSelection(file);
  });
}

async function handleFileSelection(file) {
  const validator = validateWorkbook || window.HallAllocationEngine?.validateWorkbook;
  const loader = loadWorkbook || window.HallAllocationEngine?.loadWorkbook;

  const validation = validator(file);
  if (!validation.valid) {
    showUploadFeedback(validation.message, 'danger');
    return;
  }

  showUploadFeedback('', 'hidden');
  elements.processingLoader.classList.remove('hidden');

  try {
    await processUploadedFile(file);
    showUploadFeedback(`Successfully processed "${file.name}"!`, 'success');
    showToast('Excel workbook loaded & macro engine updated.', 'success');
    switchView('dashboard');
  } catch (err) {
    console.error(err);
    showUploadFeedback(`Error reading Excel file: ${err.message}`, 'danger');
  } finally {
    elements.processingLoader.classList.add('hidden');
  }
}

async function processUploadedFile(file) {
  const loader = loadWorkbook || window.HallAllocationEngine?.loadWorkbook;
  const bundle = await loader(file);

  state.workbookBundle = bundle;
  state.classRows = bundle.classRows;
  state.halls = bundle.halls;

  // Update status UI
  elements.loadedFileName.textContent = file.name;
  elements.statusText.textContent = `Loaded: ${file.name}`;
  const dot = document.querySelector('.dot');
  if (dot) dot.className = 'dot success';

  // Compute analytics
  const fnStats = computeOverallStats || window.HallAllocationEngine?.computeOverallStats;
  const fnSummary = computeMonthlySummary || window.HallAllocationEngine?.computeMonthlySummary;
  const fnUtil = computeHallUtilization || window.HallAllocationEngine?.computeHallUtilization;
  const fnBreakdown = computeStatusBreakdown || window.HallAllocationEngine?.computeStatusBreakdown;

  state.overallStats = fnStats(state.classRows);
  state.monthlySummary = fnSummary(state.classRows, state.halls);
  state.hallUtilization = fnUtil(state.classRows);
  state.statusBreakdown = fnBreakdown(state.classRows);

  // Populate month dropdown options
  populateMonthOptions();

  // Render Dashboard
  renderKPIs();
  renderCharts();
  renderTopHalls();
  renderUploadPreview(bundle.previewRows);

  // Render Schedule Table
  state.filteredSchedule = [...state.classRows];
  renderDetailedSchedule();

  // Run initial default search macro
  executeSearchMacro();
}

function showUploadFeedback(msg, type) {
  if (type === 'hidden') {
    elements.uploadFeedback.classList.add('hidden');
    return;
  }
  elements.uploadFeedback.textContent = msg;
  elements.uploadFeedback.className = `alert ${type}`;
}

function renderUploadPreview(previewRows) {
  if (!previewRows || !previewRows.length) return;

  const cols = Object.keys(previewRows[0]);
  elements.previewHead.innerHTML = `<tr>${cols.map(c => `<th>${c}</th>`).join('')}</tr>`;

  elements.previewBody.innerHTML = previewRows.map(r => `
    <tr>${cols.map(c => `<td>${r[c] ?? ''}</td>`).join('')}</tr>
  `).join('');

  elements.previewCountBadge.textContent = `${state.classRows.length} total rows loaded`;
  elements.previewSection.classList.remove('hidden');
}

/* ── Search Form & Macro Execution ─────────────────────────────────────── */

function bindSearchForm() {
  elements.executeSearchBtn.addEventListener('click', () => {
    executeSearchMacro();
    switchView('dashboard');
  });

  elements.topRunSearchBtn.addEventListener('click', () => {
    executeSearchMacro();
    switchView('report');
  });
}

function populateMonthOptions() {
  let minDate = null, maxDate = null;
  state.classRows.forEach(r => {
    const sd = r[15] instanceof Date ? r[15] : null;
    const ed = r[16] instanceof Date ? r[16] : null;
    if (sd && (!minDate || sd < minDate)) minDate = sd;
    if (ed && (!maxDate || ed > maxDate)) maxDate = ed;
  });

  const fnGen = generateMonthOptions || window.HallAllocationEngine?.generateMonthOptions;
  const options = fnGen(minDate, maxDate);
  elements.searchMonth.innerHTML = `<option value="">(Any Month)</option>` +
    options.map(o => `<option value="${o.value}">${o.label}</option>`).join('');
}

function executeSearchMacro() {
  if (!state.workbookBundle) {
    showToast('Please upload an Excel workbook first.', 'danger');
    return;
  }

  const fnTime = parseTimeToMinutes || window.HallAllocationEngine?.parseTimeToMinutes;
  const fnRun = runHallSearchLogic || window.HallAllocationEngine?.runHallSearchLogic;

  state.searchParams = {
    day: elements.searchDay.value,
    time: elements.searchTime.value,
    month: elements.searchMonth.value,
    searchMinutes: fnTime(elements.searchTime.value)
  };

  // Run the JS macro algorithm
  const results = fnRun(state.workbookBundle, state.searchParams);
  state.lastSearchResults = results;

  // Render Dashboard Quick Table & Official Report
  renderQuickResults(results);
  renderOfficialReport(results);

  showToast(`Macro executed for ${state.searchParams.day} at ${state.searchParams.time}.`, 'success');
}

/* ── Dashboard Visualizations & Data Rendering ─────────────────────────── */

function renderKPIs() {
  const stats = state.overallStats;
  if (!stats) return;

  elements.kpiTotalClasses.textContent = stats.total;
  elements.kpiConflictCount.textContent = `${stats.conflicts} Conflicts`;

  elements.kpiPhysicalClasses.textContent = stats.physical;
  const physPct = stats.total ? Math.round((stats.physical / stats.total) * 100) : 0;
  elements.kpiPhysicalPercent.textContent = `${physPct}% of total classes`;

  elements.kpiHybridOnline.textContent = stats.hybrid + stats.online;
  elements.kpiHybridBreakdown.textContent = `${stats.hybrid} Hybrid / ${stats.online} Online`;

  elements.kpiOngoingClasses.textContent = stats.ongoing;
  elements.kpiStatusSubtext.textContent = `Ongoing • ${stats.upcoming} Upcoming`;
}

function renderQuickResults(searchOutput) {
  const { rows, summary } = searchOutput;
  elements.lastSearchTimestamp.textContent = `Searched: ${summary.day} ${summary.time} (${summary.monthLabel})`;

  if (!rows || !rows.length) {
    elements.quickResultsBody.innerHTML = `<tr><td colspan="6" class="empty-state">No results available.</td></tr>`;
    return;
  }

  elements.quickResultsBody.innerHTML = rows.map(r => {
    const isFree = r.status === 'Free';
    const badgeClass = isFree ? 'badge success' : 'badge danger';
    return `
      <tr>
        <td><strong>${r.hall}</strong></td>
        <td>${r.roomType}</td>
        <td><span class="${badgeClass}">${r.status}</span></td>
        <td>${r.occBy || '—'}</td>
        <td>${r.occMode || '—'}</td>
        <td>${r.occTime || '—'}</td>
      </tr>
    `;
  }).join('');
}

function renderTopHalls() {
  const list = state.hallUtilization.slice(0, 7);
  if (!list.length) return;

  const maxVal = list[0].count || 1;

  elements.topHallsBody.innerHTML = list.map(item => {
    const pct = Math.round((item.count / maxVal) * 100);
    return `
      <tr>
        <td><strong>${item.hall}</strong></td>
        <td>${item.count} sessions</td>
        <td>
          <div class="progress-bar-wrap">
            <div class="progress-bar-fill" style="width: ${pct}%"></div>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

function renderCharts() {
  if (!window.Chart) return;

  // Chart 1: Monthly Trends (Bar Chart)
  const summaryData = state.monthlySummary;
  const labels = summaryData.map(d => d.monthLabel);
  const physData = summaryData.map(d => d.physical);
  const hybridData = summaryData.map(d => d.hybrid);

  if (state.charts.trend) state.charts.trend.destroy();
  const ctxTrend = document.getElementById('monthlyTrendChart').getContext('2d');
  state.charts.trend = new Chart(ctxTrend, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        { label: 'Physical Sessions', data: physData, backgroundColor: '#2563eb' },
        { label: 'Hybrid Sessions', data: hybridData, backgroundColor: '#8b5cf6' }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'top' } },
      scales: { x: { stacked: true }, y: { stacked: true } }
    }
  });

  // Chart 2: Status Breakdown (Doughnut)
  const stats = state.overallStats;
  if (state.charts.status) state.charts.status.destroy();
  const ctxStatus = document.getElementById('statusBreakdownChart').getContext('2d');
  state.charts.status = new Chart(ctxStatus, {
    type: 'doughnut',
    data: {
      labels: ['Ongoing', 'Upcoming', 'Completed'],
      datasets: [{
        data: [stats.ongoing, stats.upcoming, stats.completed],
        backgroundColor: ['#10b981', '#f59e0b', '#64748b']
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'right' } }
    }
  });
}

/* ── Detailed Schedule Table & Pagination ──────────────────────────────── */

function bindDetailedFilters() {
  const filterHandler = () => {
    const searchVal = elements.scheduleSearchInput.value.toLowerCase().trim();
    const modeVal = elements.modeFilterSelect.value;
    const statusVal = elements.statusFilterSelect.value;

    state.filteredSchedule = state.classRows.filter(r => {
      const hall = (r[9] || '').toString().toLowerCase();
      const prog = (r[11] || '').toString().toLowerCase();
      const mode = (r[12] || '').toString();
      const status = (r[18] || '').toString();

      const matchesSearch = !searchVal || hall.includes(searchVal) || prog.includes(searchVal);
      const matchesMode = modeVal === 'ALL' || mode === modeVal;
      const matchesStatus = statusVal === 'ALL' || status.includes(statusVal);

      return matchesSearch && matchesMode && matchesStatus;
    });

    state.currentPage = 1;
    renderDetailedSchedule();
  };

  elements.scheduleSearchInput.addEventListener('input', filterHandler);
  elements.modeFilterSelect.addEventListener('change', filterHandler);
  elements.statusFilterSelect.addEventListener('change', filterHandler);

  elements.prevPageBtn.addEventListener('click', () => {
    if (state.currentPage > 1) {
      state.currentPage--;
      renderDetailedSchedule();
    }
  });

  elements.nextPageBtn.addEventListener('click', () => {
    const maxPage = Math.ceil(state.filteredSchedule.length / state.pageSize);
    if (state.currentPage < maxPage) {
      state.currentPage++;
      renderDetailedSchedule();
    }
  });
}

function renderDetailedSchedule() {
  const total = state.filteredSchedule.length;
  const startIdx = (state.currentPage - 1) * state.pageSize;
  const pageRows = state.filteredSchedule.slice(startIdx, startIdx + state.pageSize);

  const fnFmt = formatDateSafe || window.HallAllocationEngine?.formatDateSafe;

  if (!pageRows.length) {
    elements.detailedScheduleBody.innerHTML = `<tr><td colspan="9" class="empty-state">No schedule records found.</td></tr>`;
    elements.paginationInfo.textContent = `Showing 0 of 0 records`;
    elements.prevPageBtn.disabled = true;
    elements.nextPageBtn.disabled = true;
    return;
  }

  elements.detailedScheduleBody.innerHTML = pageRows.map(r => {
    const conflict = r[19] === 'CONFLICT' ? `<span class="badge danger">CONFLICT</span>` : `—`;
    return `
      <tr>
        <td><strong>${r[1] || '—'}</strong></td>
        <td>${r[3] || '—'}</td>
        <td><strong>${r[9] || '—'}</strong></td>
        <td>${r[11] || '—'}</td>
        <td>${r[12] || '—'}</td>
        <td>${fnFmt(r[15])}</td>
        <td>${fnFmt(r[16])}</td>
        <td><span class="badge ${r[18] === 'Ongoing' ? 'success' : 'info'}">${r[18] || '—'}</span></td>
        <td>${conflict}</td>
      </tr>
    `;
  }).join('');

  const endIdx = Math.min(startIdx + state.pageSize, total);
  elements.paginationInfo.textContent = `Showing ${startIdx + 1}-${endIdx} of ${total} records`;

  elements.prevPageBtn.disabled = state.currentPage === 1;
  elements.nextPageBtn.disabled = endIdx >= total;
}

/* ── Official Report View & Actions ───────────────────────────────────── */

function renderOfficialReport(searchOutput) {
  if (!searchOutput) return;

  const { rows, summary } = searchOutput;

  elements.repSelectedDay.textContent = summary.day;
  elements.repSelectedTime.textContent = summary.time;
  elements.repSelectedMonth.textContent = summary.monthLabel;
  elements.repTotalHalls.textContent = summary.totalHalls;

  elements.repFreeCount.textContent = summary.freeHalls;
  elements.repOccupiedCount.textContent = summary.occupiedHalls;
  elements.repUtilizationRate.textContent = `${summary.utilization}%`;

  elements.reportTimestamp.textContent = new Date().toLocaleString();

  elements.reportTableBody.innerHTML = rows.map((r, idx) => `
    <tr>
      <td>${idx + 1}</td>
      <td><strong>${r.hall}</strong></td>
      <td>${r.roomType}</td>
      <td><span class="badge ${r.status === 'Free' ? 'success' : 'danger'}">${r.status}</span></td>
      <td>${r.occBy || '—'}</td>
      <td>${r.occMode || '—'}</td>
      <td>${r.occTime || '—'}</td>
    </tr>
  `).join('');
}

function bindReportActions() {
  elements.printReportBtn.addEventListener('click', () => {
    window.print();
  });

  elements.exportCsvBtn.addEventListener('click', () => {
    if (!state.lastSearchResults) return;
    const { rows, summary } = state.lastSearchResults;

    const headers = ['#', 'Lecture Hall', 'Room Type', 'Status', 'Occupied By', 'Mode', 'Occupied Time Slot'];
    const csvContent = [
      [`"Lecture Hall Allocation Search Report"`],
      [`"Search Criteria: Day=${summary.day}, Time=${summary.time}, Month=${summary.monthLabel}"`],
      [],
      headers.join(','),
      ...rows.map((r, i) => [
        i + 1,
        `"${r.hall}"`,
        `"${r.roomType}"`,
        `"${r.status}"`,
        `"${r.occBy || ''}"`,
        `"${r.occMode || ''}"`,
        `"${r.occTime || ''}"`
      ].join(','))
    ].join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `Hall_Availability_${summary.day}_${summary.time}.csv`;
    link.click();
    URL.revokeObjectURL(link.href);
    showToast('CSV Report exported successfully.', 'success');
  });
}

/* ── Toast Utility ─────────────────────────────────────────────────────── */

function showToast(msg, type = 'info') {
  elements.toastMessage.textContent = msg;
  elements.toast.classList.remove('hidden');

  setTimeout(() => {
    elements.toast.classList.add('hidden');
  }, 4000);
}

// Launch application on DOM ready
document.addEventListener('DOMContentLoaded', init);
