/**
 * app.js
 * ======
 * Main UI Controller for Lecture Hall Allocation & Availability System.
 */

import {
  validateWorkbook, loadWorkbook,
  runHallSearchLogic,
  computeOverallStats, computeMonthlySummary,
  computeHallUtilization, computeStatusBreakdown,
  parseTimeToMinutes, generateMonthOptions,
  formatDateSafe
} from './processing.js';

document.addEventListener('DOMContentLoaded', async () => {
  // DOM Elements
  const appLoader = document.getElementById('appLoader');
  const sidebar = document.getElementById('sidebar');
  const mobileToggle = document.getElementById('mobileToggle');
  const navItems = document.querySelectorAll('.nav-item');
  const views = document.querySelectorAll('.view');
  const pageTitle = document.getElementById('pageTitle');
  const pageSubtitle = document.getElementById('pageSubtitle');
  const loadedFileName = document.getElementById('loadedFileName');
  const statusText = document.getElementById('statusText');
  const fileStatusIndicator = document.getElementById('fileStatusIndicator');
  const toast = document.getElementById('toast');
  const toastMessage = document.getElementById('toastMessage');

  // Dashboard KPI Elements
  const kpiTotalClasses = document.getElementById('kpiTotalClasses');
  const kpiConflictCount = document.getElementById('kpiConflictCount');
  const kpiPhysicalClasses = document.getElementById('kpiPhysicalClasses');
  const kpiPhysicalPercent = document.getElementById('kpiPhysicalPercent');
  const kpiHybridOnline = document.getElementById('kpiHybridOnline');
  const kpiHybridBreakdown = document.getElementById('kpiHybridBreakdown');
  const kpiOngoingClasses = document.getElementById('kpiOngoingClasses');
  const kpiStatusSubtext = document.getElementById('kpiStatusSubtext');

  // Search Controls
  const searchDay = document.getElementById('searchDay');
  const searchTime = document.getElementById('searchTime');
  const searchModeInputs = document.querySelectorAll('input[name="searchMode"]');
  const rangeOnlyControls = document.querySelectorAll('.range-only');
  const searchDateFrom = document.getElementById('searchDateFrom');
  const searchDateTo = document.getElementById('searchDateTo');
  const searchTimeFrom = document.getElementById('searchTimeFrom');
  const searchTimeTo = document.getElementById('searchTimeTo');
  const searchMonth = document.getElementById('searchMonth');
  const executeSearchBtn = document.getElementById('executeSearchBtn');
  const topRunSearchBtn = document.getElementById('topRunSearchBtn');
  const quickResultsBody = document.getElementById('quickResultsBody');
  const lastSearchTimestamp = document.getElementById('lastSearchTimestamp');
  const occupancyBoard = document.getElementById('occupancyBoard');

  // Charts
  let monthlyChartInstance = null;
  let statusChartInstance = null;

  // Tables
  const topHallsBody = document.getElementById('topHallsBody');
  const detailedScheduleBody = document.getElementById('detailedScheduleBody');
  const scheduleSearchInput = document.getElementById('scheduleSearchInput');
  const modeFilterSelect = document.getElementById('modeFilterSelect');
  const statusFilterSelect = document.getElementById('statusFilterSelect');
  const prevPageBtn = document.getElementById('prevPageBtn');
  const nextPageBtn = document.getElementById('nextPageBtn');
  const paginationInfo = document.getElementById('paginationInfo');

  // Upload Elements
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');
  const processingLoader = document.getElementById('processingLoader');
  const previewSection = document.getElementById('previewSection');
  const previewHead = document.getElementById('previewHead');
  const previewBody = document.getElementById('previewBody');
  const previewCountBadge = document.getElementById('previewCountBadge');

  // Report Elements
  const repSelectedDay = document.getElementById('repSelectedDay');
  const repSearchMode = document.getElementById('repSearchMode');
  const repSelectedTime = document.getElementById('repSelectedTime');
  const repSelectedDateRange = document.getElementById('repSelectedDateRange');
  const repSelectedTimeRange = document.getElementById('repSelectedTimeRange');
  const repSelectedMonth = document.getElementById('repSelectedMonth');
  const repTotalHalls = document.getElementById('repTotalHalls');
  const repFreeCount = document.getElementById('repFreeCount');
  const repOccupiedCount = document.getElementById('repOccupiedCount');
  const repUtilizationRate = document.getElementById('repUtilizationRate');
  const reportTableBody = document.getElementById('reportTableBody');
  const reportTimestamp = document.getElementById('reportTimestamp');
  const printReportBtn = document.getElementById('printReportBtn');
  const exportCsvBtn = document.getElementById('exportCsvBtn');

  // Application State
  let loadedWorkbookBundle = null;
  let currentSearchResults = null;
  let currentPage = 1;
  const pageSize = 15;

  function getSearchMode() {
    return document.querySelector('input[name="searchMode"]:checked').value;
  }

  function updateSearchMode() {
    const rangeMode = getSearchMode() === 'range';
    rangeOnlyControls.forEach(control => control.classList.toggle('hidden', !rangeMode));
    searchDay.closest('.form-group').classList.toggle('hidden', rangeMode);
    searchTime.closest('.form-group').classList.toggle('hidden', rangeMode);
  }

  searchModeInputs.forEach(input => input.addEventListener('change', updateSearchMode));
  updateSearchMode();

  // Toast notification
  function showToast(msg) {
    toastMessage.textContent = msg;
    toast.classList.remove('hidden');
    setTimeout(() => toast.classList.add('hidden'), 3500);
  }

  // Navigation Switcher
  navItems.forEach(btn => {
    btn.addEventListener('click', () => {
      const viewName = btn.getAttribute('data-view');
      navItems.forEach(i => i.classList.remove('active'));
      btn.classList.add('active');

      views.forEach(v => v.classList.remove('active'));
      const targetView = document.getElementById(`${viewName}View`);
      if (targetView) targetView.classList.add('active');

      // Update titles
      const titles = {
        dashboard: { title: 'Dashboard Overview', sub: 'Real-time lecture hall allocation and availability analysis' },
        upload: { title: 'Excel Workbook Upload', sub: 'Import custom schedule data files (.xlsx, .xls, .xlsm)' },
        results: { title: 'Detailed Class Schedule Analysis', sub: 'Filter, search and review all scheduled sessions' },
        report: { title: 'Official Availability Summary Report', sub: 'Printable matrix generated via macro calculation engine' }
      };

      if (titles[viewName]) {
        pageTitle.textContent = titles[viewName].title;
        pageSubtitle.textContent = titles[viewName].sub;
      }
    });
  });

  // Mobile Toggle
  mobileToggle.addEventListener('click', () => {
    sidebar.classList.toggle('open');
  });

  // Attempt auto-loading default local file if present
  try {
    const response = await fetch('./Lecture_Hall_Allocation_for_ongoing_programs.xlsx');
    if (response.ok) {
      const blob = await response.blob();
      const file = new File([blob], 'Lecture_Hall_Allocation_for_ongoing_programs.xlsx');
      loadedWorkbookBundle = await loadWorkbook(file);
      updateDashboardUI(loadedWorkbookBundle);
      showToast('Loaded default lecture hall workbook data.');
    }
  } catch (e) {
    console.log('No default local file found, waiting for user upload.');
  } finally {
    if (appLoader) appLoader.classList.add('fade-out');
  }

  // File Upload Handlers
  dropZone.addEventListener('click', () => fileInput.click());

  ['dragenter', 'dragover'].forEach(name => {
    dropZone.addEventListener(name, (e) => { e.preventDefault(); dropZone.style.borderColor = 'var(--primary)'; });
  });
  ['dragleave', 'drop'].forEach(name => {
    dropZone.addEventListener(name, (e) => { e.preventDefault(); dropZone.style.borderColor = '#cbd5e1'; });
  });

  dropZone.addEventListener('drop', (e) => {
    if (e.dataTransfer.files && e.dataTransfer.files.length) {
      handleFileUpload(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length) {
      handleFileUpload(e.target.files[0]);
    }
  });

  async function handleFileUpload(file) {
    const val = validateWorkbook(file);
    if (!val.valid) {
      showToast(val.message);
      return;
    }

    processingLoader.classList.remove('hidden');
    try {
      loadedWorkbookBundle = await loadWorkbook(file);
      updateDashboardUI(loadedWorkbookBundle);
      renderUploadPreview(loadedWorkbookBundle);
      statusText.textContent = 'Custom Workbook Loaded';
      fileStatusIndicator.querySelector('.dot').className = 'dot success';
      loadedFileName.textContent = file.name;
      showToast(`Successfully processed "${file.name}"!`);
    } catch (err) {
      showToast(err.message || 'Error processing Excel file.');
    } finally {
      processingLoader.classList.add('hidden');
    }
  }

  function renderUploadPreview(bundle) {
    previewSection.classList.remove('hidden');
    previewCountBadge.textContent = `${bundle.classRows.length} total rows loaded`;

    previewHead.innerHTML = `
      <tr>
        <th>Weekday</th>
        <th>Time Slot</th>
        <th>Hall</th>
        <th>Program / Batch</th>
        <th>Mode</th>
        <th>Start Date</th>
        <th>End Date</th>
        <th>Status</th>
      </tr>
    `;

    previewBody.innerHTML = bundle.previewRows.map(r => `
      <tr>
        <td>${r.Weekday}</td>
        <td>${r['Time (source)']}</td>
        <td><strong>${r.Hall}</strong></td>
        <td>${r['Program / Batch']}</td>
        <td>${r.Mode}</td>
        <td>${r['Start Date']}</td>
        <td>${r['End Date']}</td>
        <td><span class="badge ${r.Status === 'Ongoing' ? 'primary' : 'info'}">${r.Status}</span></td>
      </tr>
    `).join('');
  }

  // Update UI with Workbook Analytics
  function updateDashboardUI(bundle) {
    const stats = computeOverallStats(bundle.classRows);
    kpiTotalClasses.textContent = stats.total;
    kpiConflictCount.textContent = `${stats.conflicts} Conflicts`;

    kpiPhysicalClasses.textContent = stats.physical;
    const physPct = stats.total ? Math.round((stats.physical / stats.total) * 100) : 0;
    kpiPhysicalPercent.textContent = `${physPct}% of total`;

    kpiHybridOnline.textContent = stats.hybrid + stats.online;
    kpiHybridBreakdown.textContent = `${stats.hybrid} Hybrid / ${stats.online} Online`;

    kpiOngoingClasses.textContent = stats.ongoing;
    kpiStatusSubtext.textContent = `Ongoing • ${stats.upcoming} Upcoming`;

    // Populate month filter dropdown
    let minDate = null, maxDate = null;
    bundle.classRows.forEach(r => {
      const sd = r[15] instanceof Date ? r[15] : null;
      const ed = r[16] instanceof Date ? r[16] : null;
      if (sd && (!minDate || sd < minDate)) minDate = sd;
      if (ed && (!maxDate || ed > maxDate)) maxDate = ed;
    });

    const monthOpts = generateMonthOptions(minDate, maxDate);
    searchMonth.innerHTML = `<option value="">(Any Month)</option>` +
      monthOpts.map(m => `<option value="${m.value}">${m.label}</option>`).join('');

    // Top Occupied Halls
    const topHalls = computeHallUtilization(bundle.classRows).slice(0, 5);
    const maxCount = topHalls.length ? topHalls[0].count : 1;
    topHallsBody.innerHTML = topHalls.map(h => `
      <tr>
        <td><strong>${h.hall}</strong></td>
        <td>${h.count} sessions</td>
        <td>
          <div style="background:#e2e8f0; height:8px; border-radius:4px; overflow:hidden; width:100px;">
            <div style="background:var(--primary); height:100%; width:${Math.round((h.count/maxCount)*100)}%;"></div>
          </div>
        </td>
      </tr>
    `).join('');

    // Detailed Schedule Table
    renderDetailedScheduleTable();

    // Render Charts
    renderMonthlyChart(bundle);
    renderStatusChart(bundle);

    // Initial Search Execution
    executeSearch();
  }

  // Execute Search Macro
  function executeSearch() {
    if (!loadedWorkbookBundle) return;

    const day = searchDay.value;
    const timeStr = searchTime.value;
    const mode = getSearchMode();
    const dateFrom = mode === 'range' ? searchDateFrom.value : '';
    const dateTo = mode === 'range' ? searchDateTo.value : '';
    const timeFrom = mode === 'range' ? searchTimeFrom.value : '';
    const timeTo = mode === 'range' ? searchTimeTo.value : '';
    const month = searchMonth.value;
    const searchMinutes = parseTimeToMinutes(timeStr);

    if (mode === 'range' && (!dateFrom || !dateTo || !timeFrom || !timeTo)) {
      showToast('Select both dates and both times for a range search.');
      return;
    }
    if (mode === 'range' && dateFrom > dateTo) {
      showToast('Date From must be before Date To.');
      return;
    }
    if (mode === 'range' && timeFrom >= timeTo) {
      showToast('Time From must be before Time To.');
      return;
    }

    currentSearchResults = runHallSearchLogic(loadedWorkbookBundle, {
      mode,
      day,
      time: timeStr,
      dateFrom,
      dateTo,
      timeFrom,
      timeTo,
      month,
      searchMinutes
    });

    renderSearchQuickResults(currentSearchResults);
    renderOccupancyBoard(currentSearchResults);
    renderReportView(currentSearchResults);
  }

  executeSearchBtn.addEventListener('click', executeSearch);
  topRunSearchBtn.addEventListener('click', () => {
    document.querySelector('[data-view="dashboard"]').click();
    executeSearch();
  });

  function renderSearchQuickResults(res) {
    lastSearchTimestamp.textContent = `Executed at ${new Date().toLocaleTimeString()}`;
    quickResultsBody.innerHTML = res.rows.map(r => `
      <tr>
        <td><strong>${r.hall}</strong></td>
        <td>${r.roomType}</td>
        <td><span class="badge ${r.status === 'Free' ? 'info' : 'danger'}">${r.status}</span></td>
        <td>${r.occBy || '—'}</td>
        <td>${r.occMode || '—'}</td>
        <td>${r.occTime || '—'}</td>
      </tr>
    `).join('');
  }

  function renderOccupancyBoard(res) {
    occupancyBoard.innerHTML = res.rows.map(r => `
      <div class="occupancy-tile ${r.status.toLowerCase()}" title="${r.status === 'Occupied' ? `${r.occBy || 'Scheduled class'} • ${r.occTime || 'Time unavailable'}` : 'Available for booking'}">
        <div class="occupancy-tile-top">
          <span class="occupancy-status-dot"></span>
          <span class="occupancy-status">${r.status}</span>
        </div>
        <strong>${r.hall}</strong>
        <span class="occupancy-room-type">${r.roomType || 'Lecture hall'}</span>
        <span class="occupancy-booking">${r.status === 'Occupied' ? (r.occBy || 'Scheduled class') : 'Available'}</span>
        <span class="occupancy-time">${r.status === 'Occupied' ? (r.occTime || 'Time unavailable') : 'Ready to book'}</span>
      </div>
    `).join('');
  }

  function renderReportView(res) {
    const s = res.summary;
    repSelectedDay.textContent = s.day;
    repSearchMode.textContent = s.mode === 'range' ? 'Date/time range' : 'Selected time';
    repSelectedTime.textContent = s.time;
    repSelectedDateRange.textContent = s.dateRangeLabel;
    repSelectedTimeRange.textContent = s.timeRangeLabel;
    repSelectedMonth.textContent = s.monthLabel;
    repTotalHalls.textContent = s.totalHalls;
    repFreeCount.textContent = s.freeHalls;
    repOccupiedCount.textContent = s.occupiedHalls;
    repUtilizationRate.textContent = `${s.utilization}%`;
    reportTimestamp.textContent = new Date().toLocaleString();

    reportTableBody.innerHTML = res.rows.map((r, idx) => `
      <tr>
        <td>${idx + 1}</td>
        <td><strong>${r.hall}</strong></td>
        <td>${r.roomType}</td>
        <td><span class="badge ${r.status === 'Free' ? 'info' : 'danger'}">${r.status}</span></td>
        <td>${r.occBy || '—'}</td>
        <td>${r.occMode || '—'}</td>
        <td>${r.occTime || '—'}</td>
      </tr>
    `).join('');
  }

  // Print & CSV Export
  printReportBtn.addEventListener('click', () => window.print());

  exportCsvBtn.addEventListener('click', () => {
    if (!currentSearchResults) return;
    let csv = "Index,Lecture Hall,Room Type,Status,Occupied By,Mode,Time Slot\n";
    currentSearchResults.rows.forEach((r, idx) => {
      csv += `"${idx+1}","${r.hall}","${r.roomType}","${r.status}","${r.occBy}","${r.occMode}","${r.occTime}"\n`;
    });

    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Hall_Availability_Report_${searchDay.value}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  });

  // Render Charts using Chart.js
  function renderMonthlyChart(bundle) {
    const summary = computeMonthlySummary(bundle.classRows, bundle.halls);
    const ctx = document.getElementById('monthlyTrendChart').getContext('2d');

    if (monthlyChartInstance) monthlyChartInstance.destroy();

    monthlyChartInstance = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: summary.map(s => s.monthLabel),
        datasets: [
          { label: 'Physical', data: summary.map(s => s.physical), backgroundColor: '#16a34a' },
          { label: 'Hybrid', data: summary.map(s => s.hybrid), backgroundColor: '#2563eb' },
          { label: 'Online', data: summary.map(s => s.online), backgroundColor: '#9333ea' }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: { x: { stacked: true }, y: { stacked: true } }
      }
    });
  }

  function renderStatusChart(bundle) {
    const breakdown = computeStatusBreakdown(bundle.classRows);
    const ctx = document.getElementById('statusBreakdownChart').getContext('2d');

    if (statusChartInstance) statusChartInstance.destroy();

    statusChartInstance = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: ['Ongoing', 'Upcoming', 'Completed'],
        datasets: [{
          data: [breakdown.Ongoing.Total, breakdown.Upcoming.Total, breakdown.Completed.Total],
          backgroundColor: ['#2563eb', '#0284c7', '#64748b']
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false
      }
    });
  }

  // Detailed Schedule Table Filtering & Pagination
  function renderDetailedScheduleTable() {
    if (!loadedWorkbookBundle) return;

    const searchTerm = scheduleSearchInput.value.toLowerCase();
    const modeFilter = modeFilterSelect.value;
    const statusFilter = statusFilterSelect.value;

    const filtered = loadedWorkbookBundle.classRows.filter(r => {
      const day = String(r[1] || '').toLowerCase();
      const hall = String(r[9] || '').toLowerCase();
      const prog = String(r[11] || '').toLowerCase();
      const mode = String(r[12] || '');
      const sd = r[15] instanceof Date ? r[15] : null;
      const ed = r[16] instanceof Date ? r[16] : null;

      let status = 'Unknown';
      if (sd && ed) {
        const today = new Date(); today.setHours(0,0,0,0);
        if (today > ed) status = 'Completed';
        else if (today < sd) status = 'Upcoming';
        else status = 'Ongoing';
      }

      const matchesSearch = !searchTerm || day.includes(searchTerm) || hall.includes(searchTerm) || prog.includes(searchTerm);
      const matchesMode = modeFilter === 'ALL' || mode.toLowerCase() === modeFilter.toLowerCase();
      const matchesStatus = statusFilter === 'ALL' || status.toLowerCase() === statusFilter.toLowerCase();

      return matchesSearch && matchesMode && matchesStatus;
    });

    const totalPages = Math.ceil(filtered.length / pageSize) || 1;
    if (currentPage > totalPages) currentPage = totalPages;

    const startIdx = (currentPage - 1) * pageSize;
    const paged = filtered.slice(startIdx, startIdx + pageSize);

    paginationInfo.textContent = `Showing ${filtered.length ? startIdx + 1 : 0} to ${Math.min(startIdx + pageSize, filtered.length)} of ${filtered.length} records`;
    prevPageBtn.disabled = currentPage === 1;
    nextPageBtn.disabled = currentPage === totalPages;

    detailedScheduleBody.innerHTML = paged.map(r => {
      const sd = r[15] instanceof Date ? r[15] : null;
      const ed = r[16] instanceof Date ? r[16] : null;
      let status = 'Unknown';
      if (sd && ed) {
        const today = new Date(); today.setHours(0,0,0,0);
        if (today > ed) status = 'Completed';
        else if (today < sd) status = 'Upcoming';
        else status = 'Ongoing';
      }
      const conflict = String(r[19] || '');

      return `
        <tr>
          <td>${r[1] || '—'}</td>
          <td>${r[3] || '—'}</td>
          <td><strong>${r[9] || '—'}</strong></td>
          <td>${r[11] || '—'}</td>
          <td>${r[12] || '—'}</td>
          <td>${formatDateSafe(sd)}</td>
          <td>${formatDateSafe(ed)}</td>
          <td><span class="badge ${status === 'Ongoing' ? 'primary' : status === 'Upcoming' ? 'info' : 'secondary'}">${status}</span></td>
          <td>${conflict ? `<span class="badge danger">${conflict}</span>` : '—'}</td>
        </tr>
      `;
    }).join('');
  }

  scheduleSearchInput.addEventListener('input', () => { currentPage = 1; renderDetailedScheduleTable(); });
  modeFilterSelect.addEventListener('change', () => { currentPage = 1; renderDetailedScheduleTable(); });
  statusFilterSelect.addEventListener('change', () => { currentPage = 1; renderDetailedScheduleTable(); });

  prevPageBtn.addEventListener('click', () => { if (currentPage > 1) { currentPage--; renderDetailedScheduleTable(); } });
  nextPageBtn.addEventListener('click', () => { currentPage++; renderDetailedScheduleTable(); });
});
