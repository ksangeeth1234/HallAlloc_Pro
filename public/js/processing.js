/**
 * processing.js
 * =============
 * JavaScript recreation of the VBA RunHallSearch macro and all Excel formulas.
 *
 * Original sheets used:
 *   "Hall Availability Search"  – hall list + macro inputs/outputs
 *   "Class Schedule Analysis"   – 163 class records (source of truth)
 *
 * Column layout (0-based) in Class Schedule Analysis:
 *   A(0)  ID          B(1)  Weekday       C(2)  Wkday#(1=Mon..7=Sun)
 *   D(3)  Time(src)   E(4)  Start Time    F(5)  End Time
 *   G(6)  Start(min)  H(7)  End(min)      I(8)  Time Issue
 *   J(9)  Hall        K(10) Room Type     L(11) Program/Batch
 *   M(12) Mode        N(13) Mode Issue    O(14) Duration(src)
 *   P(15) Start Date  Q(16) End Date      R(17) Date Status
 *   S(18) Class Status (formula)          T(19) Conflict Flag
 */

'use strict';

/* ── Constants ─────────────────────────────────────────────────────────── */

const REQUIRED_SHEETS = ['Hall Availability Search', 'Class Schedule Analysis'];

const COL = {
  ID: 0, WEEKDAY: 1, WKDAY_NUM: 2, TIME_SRC: 3,
  START_TIME: 4, END_TIME: 5, START_MIN: 6, END_MIN: 7,
  TIME_ISSUE: 8, HALL: 9, ROOM_TYPE: 10, PROGRAM: 11,
  MODE: 12, MODE_ISSUE: 13, DURATION_SRC: 14,
  START_DATE: 15, END_DATE: 16, DATE_STATUS: 17,
  CLASS_STATUS: 18, CONFLICT: 19,
};

/* ── File validation ───────────────────────────────────────────────────── */

function validateWorkbook(file) {
  if (!file) return { valid: false, message: 'No file selected.' };
  const ext = (file.name.split('.').pop() || '').toLowerCase();
  if (!['xlsx', 'xls', 'xlsm'].includes(ext))
    return { valid: false, message: `Unsupported file type ".${ext}". Upload .xlsx, .xls or .xlsm.` };
  if (file.size === 0) return { valid: false, message: 'The file is empty.' };
  if (file.size > 50 * 1024 * 1024) return { valid: false, message: 'File exceeds 50 MB limit.' };
  return { valid: true, message: 'File type valid.' };
}

/* ── Workbook loading ──────────────────────────────────────────────────── */

async function loadWorkbook(file) {
  const buf = await file.arrayBuffer();

  if (typeof XLSX === 'undefined') {
    throw new Error('SheetJS (XLSX) library is not loaded.');
  }

  // cellDates:true → date serials become JS Date objects
  const wb = XLSX.read(buf, { type: 'array', cellDates: true, cellNF: false, cellText: false });

  const missing = REQUIRED_SHEETS.filter(n => !wb.SheetNames.includes(n));
  if (missing.length)
    throw new Error(`Missing sheet(s): ${missing.join(', ')}. The workbook must contain "Hall Availability Search" and "Class Schedule Analysis".`);

  const searchSheet = wb.Sheets['Hall Availability Search'];
  const dataSheet   = wb.Sheets['Class Schedule Analysis'];

  /* raw:true → numbers stay numbers, Date objects stay Dates */
  const classRowsRaw = XLSX.utils.sheet_to_json(dataSheet,  { defval: null, raw: true,  header: 1 });
  /* raw:false → format everything to string for the search sheet (we only need hall names) */
  const searchRowsRaw = XLSX.utils.sheet_to_json(searchSheet, { defval: '',  raw: false, header: 1 });

  // Row 0 = sheet title, Row 1 = headers, Rows 2+ = data
  const classRows = classRowsRaw.slice(2).filter(r => Array.isArray(r) && (r[COL.WEEKDAY] || r[COL.HALL]));
  const halls     = extractHallList(searchRowsRaw);

  if (!classRows.length) throw new Error('No class data found in "Class Schedule Analysis".');
  if (!halls.length)     throw new Error('No lecture halls found in "Hall Availability Search" (expected from row 14).');

  /* Build a human-readable preview (first 8 rows) */
  const previewRows = classRows.slice(0, 8).map(r => ({
    Weekday:          safe(r[COL.WEEKDAY]),
    'Time (source)':  safe(r[COL.TIME_SRC]),
    Hall:             safe(r[COL.HALL]),
    'Program / Batch':safe(r[COL.PROGRAM]),
    Mode:             safe(r[COL.MODE]),
    'Start Date':     formatDateSafe(toDate(r[COL.START_DATE])),
    'End Date':       formatDateSafe(toDate(r[COL.END_DATE])),
    Status:           computeClassStatus(toDate(r[COL.START_DATE]), toDate(r[COL.END_DATE])),
  }));

  return { classRows, halls, previewRows, fileName: file.name };
}

/* ── Hall list extraction ──────────────────────────────────────────────── */

/**
 * Reads the hall table from "Hall Availability Search".
 * Header row contains "Lecture Hall" in col A; data follows immediately.
 * Stops at an empty name or a row beginning with "BROWSE".
 */
function extractHallList(searchRows) {
  let headerIdx = searchRows.findIndex(r =>
    Array.isArray(r) && String(r[0]).trim().toLowerCase() === 'lecture hall'
  );
  if (headerIdx < 0) headerIdx = 12; // fallback: Excel row 13

  const halls = [];
  for (let i = headerIdx + 1; i < searchRows.length; i++) {
    const row      = searchRows[i];
    const hallName = String(row[0] ?? '').trim();
    if (!hallName || hallName.toLowerCase().startsWith('browse')) break;
    halls.push({ hall: hallName, roomType: String(row[1] ?? '').trim() });
  }
  return halls;
}

/* ── VBA RunHallSearch recreation ──────────────────────────────────────── */

/**
 * Exact port of the VBA RunHallSearch procedure.
 *
 * For each hall in the hall list:
 *   Scan every class row in Class Schedule Analysis.
 *   Match: Weekday == searchDay  AND  Hall == hallName  AND  Status ∈ {Ongoing, Upcoming}
 *   If searchMinutes ∈ [Start(min), End(min))  AND  month filter passes → Occupied.
 *   First match wins (identical to VBA "Exit For").
 *
 * @param {{ classRows: any[][], halls: {hall:string,roomType:string}[] }} bundle
 * @param {{ mode:string, day:string, time:string, month:string, searchMinutes:number, dateFrom:string, dateTo:string, timeFrom:string, timeTo:string }} params
 */
function runHallSearchLogic(bundle, params) {
  const { classRows, halls } = bundle;
  const { mode, day, time, month, searchMinutes, dateFrom, dateTo, timeFrom, timeTo } = params;
  const rangeMode = mode === 'range';
  const resultDay = rangeMode ? 'All days' : day;

  const selectedDateFrom = rangeMode && dateFrom ? toDate(dateFrom) : null;
  const selectedDateTo = rangeMode && dateTo ? toDate(dateTo) : null;
  if (selectedDateFrom) selectedDateFrom.setHours(0, 0, 0, 0);
  if (selectedDateTo) selectedDateTo.setHours(23, 59, 59, 999);
  const rangeFromMinutes = parseTimeToMinutes(timeFrom);
  const rangeToMinutes = parseTimeToMinutes(timeTo);
  const useTimeRange = rangeMode && Boolean(timeFrom && timeTo) && !isNaN(rangeFromMinutes) && !isNaN(rangeToMinutes);
  const dateRangeLabel = selectedDateFrom || selectedDateTo
    ? `${dateFrom || 'Any date'} - ${dateTo || 'Any date'}`
    : 'Any Date';
  const timeRangeLabel = useTimeRange
    ? `${timeFrom} - ${timeTo}`
    : `Target time: ${time || 'Any time'}`;

  /* Month filter setup */
  const useMonth  = Boolean(month && month !== '');
  let monthStart  = null, monthEnd = null, monthLabel = 'Any Month';
  if (useMonth) {
    monthStart = new Date(month); monthStart.setHours(0, 0, 0, 0);
    monthEnd   = new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 0);
    monthEnd.setHours(23, 59, 59, 999);
    monthLabel = monthStart.toLocaleString('en-US', { month: 'short', year: 'numeric' });
  }

  const results = halls.map(({ hall, roomType }) => {
    let found = false, occBy = '', occMode = '', occTime = '';

    for (const row of classRows) {
      const dDay  = safe(row[COL.WEEKDAY]);
      const dHall = safe(row[COL.HALL]);
      if (!dDay || !dHall) continue;
      if ((!rangeMode && dDay !== day) || dHall !== hall) continue;

      // Recompute status from dates (same as Excel column S formula)
      const sd = toDate(row[COL.START_DATE]);
      const ed = toDate(row[COL.END_DATE]);
      const status = computeClassStatus(sd, ed);
      const hasDateRange = rangeMode;
      if (!hasDateRange && status !== 'Ongoing' && status !== 'Upcoming') continue;

      if (hasDateRange && (!sd || !ed ||
        (selectedDateFrom && ed < selectedDateFrom) ||
        (selectedDateTo && sd > selectedDateTo))) continue;

      const dStart = toNum(row[COL.START_MIN]);
      const dEnd   = toNum(row[COL.END_MIN]);
      if (isNaN(dStart) || isNaN(dEnd)) continue;

      // A range matches when the class interval overlaps it. Without a range,
      // preserve the original point-in-time search behavior.
      const timeMatches = useTimeRange
        ? rangeFromMinutes < dEnd && rangeToMinutes > dStart
        : searchMinutes >= dStart && searchMinutes < dEnd;
      if (timeMatches) {
        // Month overlap: classStart <= monthEnd AND classEnd >= monthStart
        let monthOk = true;
        if (useMonth && sd && ed) monthOk = sd <= monthEnd && ed >= monthStart;

        if (monthOk) {
          found   = true;
          occBy   = safe(row[COL.PROGRAM]);
          occMode = safe(row[COL.MODE]);
          occTime = `${safe(row[COL.START_TIME]) || '—'} - ${safe(row[COL.END_TIME]) || '—'}`;
          break; // VBA: Exit For
        }
      }
    }

    return {
      hall, roomType,
      status:  found ? 'Occupied' : 'Free',
      occBy:   found ? occBy   : '',
      occMode: found ? occMode : '',
      occTime: found ? occTime : '',
      day: resultDay, month: monthLabel, time, mode,
    };
  });

  const free  = results.filter(r => r.status === 'Free').length;
  const occ   = results.length - free;
  const util  = results.length ? Math.round((occ / results.length) * 100) : 0;

  return {
    rows: results,
    summary: { totalHalls: results.length, freeHalls: free, occupiedHalls: occ, utilization: util, monthLabel, day: resultDay, time, mode, dateRangeLabel, timeRangeLabel },
  };
}

/* ── Overall statistics (recreates Dashboard sheet formulas) ───────────── */

function computeOverallStats(classRows) {
  let total = 0, physical = 0, hybrid = 0, online = 0;
  let completed = 0, ongoing = 0, upcoming = 0, conflicts = 0;

  for (const r of classRows) {
    if (!r[COL.WEEKDAY]) continue;
    total++;
    const mode   = safe(r[COL.MODE]).toLowerCase();
    const status = computeClassStatus(toDate(r[COL.START_DATE]), toDate(r[COL.END_DATE]));
    const conf   = safe(r[COL.CONFLICT]);

    if (mode === 'physical') physical++;
    else if (mode === 'hybrid') hybrid++;
    else if (mode === 'online') online++;

    if (status === 'Completed') completed++;
    else if (status === 'Ongoing') ongoing++;
    else if (status === 'Upcoming') upcoming++;

    if (conf === 'CONFLICT') conflicts++;
  }
  return { total, physical, hybrid, online, completed, ongoing, upcoming, conflicts };
}

/* ── Monthly occurrence summary (recreates Monthly Summary sheet) ─────── */

/**
 * For each calendar month spanned by the data, compute:
 *   – physical / hybrid / online SESSION counts (same formula as Monthly Occurrences Support)
 *   – number of physical halls with zero physical sessions that month (gap)
 */
function computeMonthlySummary(classRows, halls) {
  let minDate = null, maxDate = null;
  for (const r of classRows) {
    const sd = toDate(r[COL.START_DATE]);
    const ed = toDate(r[COL.END_DATE]);
    if (sd && (!minDate || sd < minDate)) minDate = sd;
    if (ed && (!maxDate || ed > maxDate)) maxDate = ed;
  }
  if (!minDate || !maxDate) return [];

  const physHallCount = halls.filter(h => !h.roomType.toLowerCase().includes('online')).length;
  const months        = buildMonthRange(minDate, maxDate);

  return months.map(mDate => {
    const mStart = new Date(mDate.getFullYear(), mDate.getMonth(), 1);
    const mEnd   = new Date(mDate.getFullYear(), mDate.getMonth() + 1, 0);
    mEnd.setHours(23, 59, 59, 999);

    let phys = 0, hybr = 0, onl = 0;
    const hallsWithPhys = new Set();

    for (const r of classRows) {
      const sd  = toDate(r[COL.START_DATE]);
      const ed  = toDate(r[COL.END_DATE]);
      if (!sd || !ed) continue;
      const wkn  = toNum(r[COL.WKDAY_NUM]);
      const mode = safe(r[COL.MODE]).toLowerCase();
      const hall = safe(r[COL.HALL]);
      if (!wkn) continue;

      const cnt = countOccurrencesInMonth(sd, ed, wkn, mStart, mEnd);
      if (cnt <= 0) continue;

      if (mode === 'physical') { phys += cnt; hallsWithPhys.add(hall); }
      else if (mode === 'hybrid') hybr += cnt;
      else if (mode === 'online') onl  += cnt;
    }

    return {
      monthDate:  mDate,
      monthLabel: mDate.toLocaleString('en-US', { month: 'short', year: 'numeric' }),
      physical:   phys, hybrid: hybr, online: onl,
      total:      phys + hybr + onl,
      physGap:    Math.max(0, physHallCount - hallsWithPhys.size),
    };
  });
}

/**
 * Recreates the Excel formula:
 *   =IF(MAX(sd,mStart)>MIN(ed,mEnd), 0,
 *     INT((MIN(ed,mEnd)-(MAX(sd,mStart)+daysToFirstOccurrence))/7)+1)
 * wkdayNum: VBA mode 2 (1=Mon..7=Sun)
 */
function countOccurrencesInMonth(sd, ed, wkdayNum, mStart, mEnd) {
  const effS = new Date(Math.max(sd.getTime(), mStart.getTime()));
  const effE = new Date(Math.min(ed.getTime(), mEnd.getTime()));
  if (effS > effE) return 0;

  // VBA 1=Mon..7=Sun → JS 0=Sun..6=Sat
  const jsDay = wkdayNum === 7 ? 0 : wkdayNum; // 1→1,2→2,…,6→6,7→0
  const dToFirst = (jsDay - effS.getDay() + 7) % 7;
  const firstOcc = new Date(effS);
  firstOcc.setDate(firstOcc.getDate() + dToFirst);
  if (firstOcc > effE) return 0;

  return Math.floor((effE.getTime() - firstOcc.getTime()) / (7 * 86400000)) + 1;
}

/* ── Hall utilization (active-class count per hall) ───────────────────── */

function computeHallUtilization(classRows) {
  const counts = {};
  for (const r of classRows) {
    const hall   = safe(r[COL.HALL]);
    const status = computeClassStatus(toDate(r[COL.START_DATE]), toDate(r[COL.END_DATE]));
    if (!hall || (status !== 'Ongoing' && status !== 'Upcoming')) continue;
    counts[hall] = (counts[hall] || 0) + 1;
  }
  return Object.entries(counts).map(([hall, count]) => ({ hall, count })).sort((a, b) => b.count - a.count);
}

/* ── Status × Mode breakdown (recreates "Completed vs Planned" sheet) ─── */

function computeStatusBreakdown(classRows) {
  const out = {
    Completed: { Physical: 0, Hybrid: 0, Online: 0, Other: 0, Total: 0 },
    Ongoing:   { Physical: 0, Hybrid: 0, Online: 0, Other: 0, Total: 0 },
    Upcoming:  { Physical: 0, Hybrid: 0, Online: 0, Other: 0, Total: 0 },
    Unknown:   { Physical: 0, Hybrid: 0, Online: 0, Other: 0, Total: 0 },
  };
  for (const r of classRows) {
    if (!r[COL.WEEKDAY]) continue;
    const mode   = safe(r[COL.MODE]);
    const status = computeClassStatus(toDate(r[COL.START_DATE]), toDate(r[COL.END_DATE]));
    const sk     = out[status] ? status : 'Unknown';
    const mk     = ['Physical', 'Hybrid', 'Online'].includes(mode) ? mode : 'Other';
    out[sk][mk]++; out[sk].Total++;
  }
  return out;
}

/* ── Utility functions ─────────────────────────────────────────────────── */

/**
 * Recreates Excel column S formula:
 *   =IF(OR(P="",Q=""),"Unknown (missing duration data)",IF(TODAY()>Q,"Completed",IF(TODAY()<P,"Upcoming","Ongoing")))
 */
function computeClassStatus(sd, ed) {
  if (!sd || !ed || !(sd instanceof Date) || !(ed instanceof Date)) return 'Unknown (missing duration data)';
  if (isNaN(sd.getTime()) || isNaN(ed.getTime()))                   return 'Unknown (missing duration data)';
  const today = new Date(); today.setHours(0, 0, 0, 0);
  if (today > ed)  return 'Completed';
  if (today < sd)  return 'Upcoming';
  return 'Ongoing';
}

/** Parse time string ("09:00", "9:30 AM") → minutes since midnight */
function parseTimeToMinutes(val) {
  if (!val && val !== 0) return 0;
  const s = String(val).trim();
  if (!s.includes(':')) return 0;
  const low   = s.toLowerCase();
  const isPM  = low.includes('pm'), isAM = low.includes('am');
  const clean = s.replace(/\s*(am|pm)/i, '').trim();
  let [h, m]  = clean.split(':').map(Number);
  if (isNaN(h) || isNaN(m)) return 0;
  if (isPM && h < 12) h += 12;
  if (isAM && h === 12) h = 0;
  return h * 60 + m;
}

/** Generate month options from a date range (for the dropdown) */
function generateMonthOptions(minDate, maxDate) {
  const now = new Date();
  const s = minDate ? new Date(minDate.getFullYear(), minDate.getMonth(), 1)
                    : new Date(now.getFullYear() - 1, now.getMonth(), 1);
  const e = maxDate ? new Date(maxDate.getFullYear(), maxDate.getMonth(), 1)
                    : new Date(now.getFullYear() + 2, 0, 1);
  const opts = []; const cur = new Date(s);
  while (cur <= e) {
    opts.push({
      value: `${cur.getFullYear()}-${String(cur.getMonth() + 1).padStart(2, '0')}-01`,
      label: cur.toLocaleString('en-US', { month: 'long', year: 'numeric' }),
    });
    cur.setMonth(cur.getMonth() + 1);
  }
  return opts;
}

function buildMonthRange(minDate, maxDate) {
  const months = []; const cur = new Date(minDate.getFullYear(), minDate.getMonth(), 1);
  const end = new Date(maxDate.getFullYear(), maxDate.getMonth(), 1);
  while (cur <= end) { months.push(new Date(cur)); cur.setMonth(cur.getMonth() + 1); }
  return months;
}

/** Safe string extraction (trims, handles null/undefined/Date) */
function safe(val) {
  if (val === undefined || val === null) return '';
  if (typeof val === 'string') return val.trim();
  if (val instanceof Date) return formatDateSafe(val);
  return String(val).trim();
}

/** Convert any value to a JS Date (handles Date objects and Excel serials) */
function toDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  if (typeof val === 'number') {
    const d = new Date((val - 25569) * 86400000); // Excel serial fallback
    return isNaN(d.getTime()) ? null : d;
  }
  if (typeof val === 'string') { const d = new Date(val); return isNaN(d.getTime()) ? null : d; }
  return null;
}

/** Convert any value to a number */
function toNum(val) { return val === null || val === undefined || val === '' ? NaN : Number(val); }

/** Format Date → "DD-Mon-YYYY" */
function formatDateSafe(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return '';
  return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
}

/* ── Export for Browser Global Window fallback & ES Modules ──────────── */
const HallAllocationEngine = {
  validateWorkbook, loadWorkbook,
  runHallSearchLogic,
  computeOverallStats, computeMonthlySummary,
  computeHallUtilization, computeStatusBreakdown,
  parseTimeToMinutes, generateMonthOptions,
  computeClassStatus, formatDateSafe,
};

if (typeof window !== 'undefined') {
  window.HallAllocationEngine = HallAllocationEngine;
}

export {
  validateWorkbook, loadWorkbook,
  runHallSearchLogic,
  computeOverallStats, computeMonthlySummary,
  computeHallUtilization, computeStatusBreakdown,
  parseTimeToMinutes, generateMonthOptions,
  computeClassStatus, formatDateSafe,
};
