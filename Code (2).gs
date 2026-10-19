const SPREADSHEET_ID = '1JXjnNIxa7_B8kkb_2AEKFL7wucPUiZuDem5ratAwTpI';
const DATA_SHEET_NAME = 'KPIs Entries';

const SCORES = {
  level1: 0,
  level2: 0,
  level3: 50,
  level4: 100,
  level5: 100
};


// =====================================================
// HELPERS (team + per-execution cache)
// Every entry is saved with the team the agent was
// working in (calls / tickets). Old rows without a team
// are counted on the agent's saved team.
// =====================================================

let __rowsCache = {};
let __teamCache = null;

let __ss = null;
const __sheets = {};
const SKIPPED_ROW_ = [];

function getSpreadsheet_() {
  if (!__ss) __ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return __ss;
}

// Reads the rows of a sheet. Old months are NOT read (only the date column
// is scanned), so saving stays fast even when the sheet has many rows.
// The returned array keeps the same indexes as the sheet (row = index + 2);
// skipped rows are empty, so every caller simply ignores them.
function readRows_(sheet, lastRow, ncols) {
  const key = sheet.getName() + ':' + ncols;
  if (__rowsCache[key]) return __rowsCache[key];

  const n = lastRow - 1;
  const dates = sheet.getRange(2, 1, n, 1).getValues();

  const now = new Date();
  const limit = Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 1) - 2 * 86400000;

  let start = n;
  for (let i = 0; i < n; i++) {
    const v = dates[i][0];
    if (v === '' || v === null) continue;
    const t = new Date(v).getTime();
    if (isNaN(t) || t >= limit) { start = i; break; }
  }

  const out = [];
  for (let i = 0; i < start; i++) out.push(SKIPPED_ROW_);

  if (start < n) {
    const part = sheet.getRange(2 + start, 1, n - start, ncols).getValues();
    for (let j = 0; j < part.length; j++) out.push(part[j]);
  }

  __rowsCache[key] = out;
  return out;
}

function getUserTeamCached_() {
  if (!__teamCache) __teamCache = getUserTeam_();
  return __teamCache;
}

function rowTeam_(value, fallback) {
  const t = String(value || '').toLowerCase().trim();
  return TEAMS.indexOf(t) !== -1 ? t : fallback;
}

function normalizeTeam_(team) {
  const t = String(team || '').toLowerCase().trim();
  return TEAMS.indexOf(t) !== -1 ? t : getUserTeam_();
}

function ensureColumns_(sheet, n) {
  const max = sheet.getMaxColumns();
  if (max < n) sheet.insertColumnsAfter(max, n - max);
}


// =====================================================
// WEB APP
// =====================================================

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Axis Pay - KPIs  Calculator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// =====================================================
// CURRENT USER
// =====================================================

function getCurrentUserEmail() {

  const email =
    Session.getActiveUser().getEmail();

  if (!email) {
    throw new Error(
      'Unable to identify your Axis Pay account. Please open the tool using your company account.'
    );
  }

  return email.toLowerCase().trim();
}


// =====================================================
// SHEET
// =====================================================

function getDataSheet() {

  if (__sheets[DATA_SHEET_NAME]) return __sheets[DATA_SHEET_NAME];

  const ss =
    getSpreadsheet_();

  let sheet =
    ss.getSheetByName(DATA_SHEET_NAME);

  if (!sheet) {

    sheet =
      ss.insertSheet(DATA_SHEET_NAME);

    sheet
      .getRange(1, 1, 1, 12)
      .setValues([[
        'Date',
        'Time',
        'Level 1',
        'Level 2',
        'Level 3',
        'Level 4',
        'Level 5',
        'Total',
        'Weighted Score',
        'CSAT',
        'User Email',
        'Team'
      ]]);

    sheet
      .getRange(1, 1, 1, 12)
      .setFontWeight('bold');

  } else {

    // Make sure User Email column exists
    if (sheet.getLastColumn() < 11) {

      sheet
        .getRange(1, 11)
        .setValue('User Email')
        .setFontWeight('bold');
    }

    // Make sure Team column exists
    if (sheet.getLastColumn() < 12) {

      ensureColumns_(sheet, 12);

      sheet
        .getRange(1, 12)
        .setValue('Team')
        .setFontWeight('bold');
    }
  }

  __sheets[DATA_SHEET_NAME] = sheet;
  return sheet;
}


// =====================================================
// CALCULATE CSAT
// =====================================================

function calculateCSATValue(counts) {

  const level1 =
    Number(counts.level1) || 0;

  const level2 =
    Number(counts.level2) || 0;

  const level3 =
    Number(counts.level3) || 0;

  const level4 =
    Number(counts.level4) || 0;

  const level5 =
    Number(counts.level5) || 0;

  const total =
    level1 +
    level2 +
    level3 +
    level4 +
    level5;

  const weightedScore =
    level1 * SCORES.level1 +
    level2 * SCORES.level2 +
    level3 * SCORES.level3 +
    level4 * SCORES.level4 +
    level5 * SCORES.level5;

  const csat =
    total > 0
      ? weightedScore / total
      : 0;

  return {
    level1,
    level2,
    level3,
    level4,
    level5,
    total,
    weightedScore,
    csat
  };
}


// =====================================================
// ADD ENTRY
// =====================================================

function addCSATEntry(counts, source, team) {

  const sheet =
    getDataSheet();

  const userEmail =
    getCurrentUserEmail();

  const result =
    calculateCSATValue(counts);

  if (result.total <= 0) {

    throw new Error(
      'Please enter at least one response.'
    );
  }

  const now =
    new Date();

  const entryTeam =
    normalizeTeam_(team);

  sheet.appendRow([
    now,
    now,
    result.level1,
    result.level2,
    result.level3,
    result.level4,
    result.level5,
    result.total,
    result.weightedScore,
    result.csat,
    userEmail,
    entryTeam
  ]);

  const lastRow =
    sheet.getLastRow();

  sheet
    .getRange(lastRow, 1, 1, 2)
    .setNumberFormats([['yyyy-MM-dd', 'HH:mm:ss']]);

  sheet
    .getRange(lastRow, 10)
    .setNumberFormat('0.00%');

  const payload = buildCSATPayload_(entryTeam);
  payload.entry = result;
  return payload;
}


// =====================================================
// GET TODAY ENTRIES
// ONLY CURRENT USER
// =====================================================

function getTodayEntries(team) {

  const sheet =
    getDataSheet();

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const userEmail =
    getCurrentUserEmail();

  const values =
    readRows_(sheet, lastRow, 12);

  const timezone =
    Session.getScriptTimeZone() ||
    'Africa/Cairo';

  const todayKey =
    Utilities.formatDate(
      new Date(),
      timezone,
      'yyyy-MM-dd'
    );

  const entries = [];

  values.forEach(function(row, index) {

    if (!row[0]) return;

    const rowEmail =
      String(row[10] || '')
        .toLowerCase()
        .trim();

    // IMPORTANT:
    // Only show current user's records
    if (rowEmail !== userEmail) return;

    // Only the requested team (old rows = user's saved team)
    if (team && rowTeam_(row[11], getUserTeamCached_()) !== team) return;

    const rowDate =
      new Date(row[0]);

    const dateKey =
      Utilities.formatDate(
        rowDate,
        timezone,
        'yyyy-MM-dd'
      );

    if (dateKey !== todayKey) return;

    entries.push({

      rowNumber:
        index + 2,

      date:
        dateKey,

      time:
        Utilities.formatDate(
          new Date(row[1]),
          timezone,
          'HH:mm:ss'
        ),

      level1:
        Number(row[2]) || 0,

      level2:
        Number(row[3]) || 0,

      level3:
        Number(row[4]) || 0,

      level4:
        Number(row[5]) || 0,

      level5:
        Number(row[6]) || 0,

      total:
        Number(row[7]) || 0,

      weightedScore:
        Number(row[8]) || 0,

      csat:
        Number(row[9]) || 0

    });

  });

  return entries.reverse();
}


// =====================================================
// TODAY SUMMARY
// =====================================================

function getTodaySummary(team) {

  const entries =
    getTodayEntries(team);

  const summary = {

    level1: 0,
    level2: 0,
    level3: 0,
    level4: 0,
    level5: 0,

    total: 0,
    weightedScore: 0,
    csat: 0,

    entriesCount:
      entries.length

  };

  entries.forEach(function(entry) {

    summary.level1 +=
      entry.level1;

    summary.level2 +=
      entry.level2;

    summary.level3 +=
      entry.level3;

    summary.level4 +=
      entry.level4;

    summary.level5 +=
      entry.level5;

    summary.total +=
      entry.total;

    summary.weightedScore +=
      entry.weightedScore;

  });

  if (summary.total > 0) {

    summary.csat =
      summary.weightedScore /
      summary.total;
  }

  return summary;
}


// =====================================================
// MONTHLY SUMMARY
// ONLY CURRENT USER
// =====================================================

function getMonthlySummary(team) {

  const sheet =
    getDataSheet();

  const lastRow =
    sheet.getLastRow();

  const emptySummary = {

    level1: 0,
    level2: 0,
    level3: 0,
    level4: 0,
    level5: 0,

    total: 0,
    weightedScore: 0,
    csat: 0,

    entriesCount: 0,

    dailyBreakdown: [],

    entryComparison: {

      previous: null,
      current: null,
      difference: null

    }

  };

  if (lastRow < 2) {
    return emptySummary;
  }

  const userEmail =
    getCurrentUserEmail();

  const timezone =
    Session.getScriptTimeZone() ||
    'Africa/Cairo';

  const now =
    new Date();

  const currentYear =
    Number(
      Utilities.formatDate(
        now,
        timezone,
        'yyyy'
      )
    );

  const currentMonth =
    Number(
      Utilities.formatDate(
        now,
        timezone,
        'MM'
      )
    );

  const data =
    readRows_(sheet, lastRow, 12);

  const summary = {

    level1: 0,
    level2: 0,
    level3: 0,
    level4: 0,
    level5: 0,

    total: 0,
    weightedScore: 0,
    csat: 0,

    entriesCount: 0,

    dailyBreakdown: [],

    entryComparison: {

      previous: null,
      current: null,
      difference: null

    }

  };

  const dailyMap = {};
  const monthlyEntries = [];

  data.forEach(function(row, index) {

    if (!row[0]) return;

    const rowEmail =
      String(row[10] || '')
        .toLowerCase()
        .trim();

    // Only current user
    if (rowEmail !== userEmail) return;

    // Only the requested team (old rows = user's saved team)
    if (team && rowTeam_(row[11], getUserTeamCached_()) !== team) return;

    const date =
      new Date(row[0]);

    const year =
      Number(
        Utilities.formatDate(
          date,
          timezone,
          'yyyy'
        )
      );

    const month =
      Number(
        Utilities.formatDate(
          date,
          timezone,
          'MM'
        )
      );

    if (
      year !== currentYear ||
      month !== currentMonth
    ) {
      return;
    }

    const level1 =
      Number(row[2]) || 0;

    const level2 =
      Number(row[3]) || 0;

    const level3 =
      Number(row[4]) || 0;

    const level4 =
      Number(row[5]) || 0;

    const level5 =
      Number(row[6]) || 0;

    const total =
      Number(row[7]) || 0;

    const weightedScore =
      Number(row[8]) || 0;

    const csat =
      Number(row[9]) || 0;

    summary.level1 += level1;
    summary.level2 += level2;
    summary.level3 += level3;
    summary.level4 += level4;
    summary.level5 += level5;

    summary.total += total;

    summary.weightedScore +=
      weightedScore;

    summary.entriesCount++;

    const dateKey =
      Utilities.formatDate(
        date,
        timezone,
        'yyyy-MM-dd'
      );

    if (!dailyMap[dateKey]) {

      dailyMap[dateKey] = {

        date: dateKey,

        responses: 0,

        level1: 0,
        level2: 0,
        level3: 0,
        level4: 0,
        level5: 0,

        weightedScore: 0

      };

    }

    dailyMap[dateKey].responses +=
      total;

    dailyMap[dateKey].level1 +=
      level1;

    dailyMap[dateKey].level2 +=
      level2;

    dailyMap[dateKey].level3 +=
      level3;

    dailyMap[dateKey].level4 +=
      level4;

    dailyMap[dateKey].level5 +=
      level5;

    dailyMap[dateKey].weightedScore +=
      weightedScore;

    monthlyEntries.push({

      rowNumber:
        index + 2,

      csat:
        csat

    });

  });

  if (summary.total > 0) {

    summary.csat =
      summary.weightedScore /
      summary.total;

  }

  summary.dailyBreakdown =
    Object.keys(dailyMap)
      .sort()
      .map(function(dateKey) {

        const day =
          dailyMap[dateKey];

        const dayCSAT =
          day.responses > 0
            ? day.weightedScore /
              day.responses
            : 0;

        return {

          date:
            day.date,

          responses:
            day.responses,

          csat:
            dayCSAT,

          level1:
            day.level1,

          level2:
            day.level2,

          level3:
            day.level3,

          level4:
            day.level4,

          level5:
            day.level5,

          positive:
            day.level4 +
            day.level5,

          mid:
            day.level3,

          negative:
            day.level1 +
            day.level2

        };

      });

  monthlyEntries.sort(
    function(a, b) {

      return a.rowNumber -
        b.rowNumber;

    }
  );

  if (monthlyEntries.length > 0) {

    summary.entryComparison.current =
      monthlyEntries[
        monthlyEntries.length - 1
      ].csat;

  }

  if (monthlyEntries.length > 1) {

    summary.entryComparison.previous =
      monthlyEntries[
        monthlyEntries.length - 2
      ].csat;

    summary.entryComparison.difference =
      summary.entryComparison.current -
      summary.entryComparison.previous;

  }

  return summary;
}


// =====================================================
// DELETE SPECIFIC ENTRY
// ONLY CURRENT USER
// =====================================================

function deleteSpecificEntry(rowNumber) {

  const sheet =
    getDataSheet();

  const lastRow =
    sheet.getLastRow();

  const userEmail =
    getCurrentUserEmail();

  rowNumber =
    Number(rowNumber);

  if (
    !rowNumber ||
    rowNumber < 2 ||
    rowNumber > lastRow
  ) {

    throw new Error(
      'Invalid entry.'
    );

  }

  const rowValues =
    sheet
      .getRange(rowNumber, 1, 1, 12)
      .getValues()[0];

  const rowEmail =
    String(rowValues[10] || '')
      .toLowerCase()
      .trim();

  const deletedTeam =
    rowTeam_(rowValues[11], getUserTeamCached_());

  // Security check
  if (rowEmail !== userEmail) {

    throw new Error(
      'You can only delete your own entries.'
    );

  }

  sheet.deleteRow(rowNumber);

  const payload = buildCSATPayload_(deletedTeam);
  payload.success = true;
  payload.message = '✓ Entry Deleted Successfully';
  return payload;

}


// =====================================================
// RESET LAST ENTRY
// ONLY CURRENT USER
// =====================================================

function resetLastEntry() {

  const sheet =
    getDataSheet();

  const lastRow =
    sheet.getLastRow();

  const userEmail =
    getCurrentUserEmail();

  if (lastRow < 2) {

    return {

      success: false,

      message:
        'No saved entry to reset.',

      today:
        getTodaySummary(),

      todayEntries:
        getTodayEntries(),

      monthly:
        getMonthlySummary()

    };

  }

  // Find latest entry belonging to current user
  for (
    let row = lastRow;
    row >= 2;
    row--
  ) {

    const rowEmail =
      String(
        sheet
          .getRange(row, 11)
          .getValue() || ''
      )
        .toLowerCase()
        .trim();

    if (rowEmail === userEmail) {

      sheet.deleteRow(row);

      return {

        success: true,

        message:
          '✓ Last Entry Reset Successfully',

        today:
          getTodaySummary(),

        todayEntries:
          getTodayEntries(),

        monthly:
          getMonthlySummary()

      };

    }

  }

  return {

    success: false,

    message:
      'No saved entry found for your account.',

    today:
      getTodaySummary(),

    todayEntries:
      getTodayEntries(),

    monthly:
      getMonthlySummary()

  };

}


// =====================================================
// LOAD ALL DATA
// =====================================================

function buildCSATPayload_(onlyTeam) {

  __rowsCache = {};

  const userEmail =
    getCurrentUserEmail();

  const team =
    getUserTeamCached_();

  const teams = {};

  (onlyTeam ? [onlyTeam] : TEAMS).forEach(function(t) {

    teams[t] = {
      today: getTodaySummary(t),
      todayEntries: getTodayEntries(t),
      monthly: getMonthlySummary(t)
    };

  });

  return {

    userEmail:
      userEmail,

    team:
      team,

    // data of BOTH teams (the page shows the selected one in the tabs
    // and uses both in the dashboard)
    teams:
      teams,

    // saved team (kept for compatibility, only when it was calculated)
    today:
      teams[team] ? teams[team].today : null,

    todayEntries:
      teams[team] ? teams[team].todayEntries : null,

    monthly:
      teams[team] ? teams[team].monthly : null

  };

}

function loadCSATData(month) {
  return buildCSATPayload_();
}


// =====================================================
// QUALITY + AHT
// Sheets: 'Quality Entries' and 'AHT Entries'
// (created automatically on first use)
// =====================================================

const METRICS = {
  quality: {
    sheet: 'Quality Entries',
    headers: ['Date', 'Time', 'Quality %', 'User Email', 'Team'],
    emailCol: 4,
    teamCol: 5
  },
  aht: {
    sheet: 'AHT Entries',
    headers: ['Date', 'Time', 'AHT (seconds)', 'AHT (mm:ss)', 'User Email', 'Team'],
    emailCol: 5,
    teamCol: 6
  },
  productivity: {
    sheet: 'Productivity Entries',
    headers: ['Date', 'Time', 'Productivity', 'User Email', 'Team'],
    emailCol: 4,
    teamCol: 5
  }
};

function getMetricSheet_(metric) {
  const cfg = METRICS[metric];
  if (!cfg) throw new Error('Unknown metric.');

  if (__sheets[cfg.sheet]) return __sheets[cfg.sheet];

  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(cfg.sheet);

  if (!sheet) {
    sheet = ss.insertSheet(cfg.sheet);
    sheet.getRange(1, 1, 1, cfg.headers.length)
      .setValues([cfg.headers])
      .setFontWeight('bold');
  } else if (sheet.getLastColumn() < cfg.headers.length) {
    // older sheet without the Team column
    ensureColumns_(sheet, cfg.headers.length);
    sheet.getRange(1, cfg.teamCol)
      .setValue('Team')
      .setFontWeight('bold');
  }

  __sheets[cfg.sheet] = sheet;
  return sheet;
}

function parseAHT_(text) {
  const m = String(text).trim().match(/^(\d{1,3}):([0-5]?\d)$/);
  if (!m) throw new Error('AHT must be in mm:ss format, for example 06:42.');
  const seconds = Number(m[1]) * 60 + Number(m[2]);
  if (seconds <= 0) throw new Error('AHT must be greater than 00:00.');
  return seconds;
}

function formatAHT_(seconds) {
  const s = Math.round(seconds);
  const mm = Math.floor(s / 60);
  const ss = s % 60;
  return (mm < 10 ? '0' : '') + mm + ':' + (ss < 10 ? '0' : '') + ss;
}

// =====================================================
// ADD ENTRY (Quality: one evaluation score / AHT: day average)
// =====================================================

function addMetricEntry(metric, rawValue, team) {
  const sheet = getMetricSheet_(metric);
  const email = getCurrentUserEmail();
  const now = new Date();
  const entryTeam = normalizeTeam_(team);
  let row, value;

  if (metric === 'quality') {
    value = Number(String(rawValue).replace('%', '').trim());
    if (!isFinite(value) || value < 0 || value > 100) {
      throw new Error('Quality must be a number between 0 and 100.');
    }
    row = [now, now, value, email, entryTeam];
  } else if (metric === 'productivity') {
    value = Number(String(rawValue).trim());
    if (!isFinite(value) || value < 0) {
      throw new Error('Productivity must be a number (0 or more).');
    }
    row = [now, now, value, email, entryTeam];
  } else {
    value = parseAHT_(rawValue);
    row = [now, now, value, '', email, entryTeam];
  }

  sheet.appendRow(row);
  const lastRow = sheet.getLastRow();

  if (metric === 'quality') {
    sheet.getRange(lastRow, 1, 1, 3).setNumberFormats([['yyyy-MM-dd', 'HH:mm:ss', '0.00']]);
  } else {
    sheet.getRange(lastRow, 1, 1, 2).setNumberFormats([['yyyy-MM-dd', 'HH:mm:ss']]);
  }

  if (metric === 'aht') {
    sheet.getRange(lastRow, 4).setNumberFormat('@').setValue(formatAHT_(value));
  }

  return metricUpdate_(metric, entryTeam);
}

// Only the metric + team that changed is calculated and sent back
function metricUpdate_(metric, team) {
  __rowsCache = {};
  return { update: { metric: metric, team: team, data: getMetricData_(metric, team) } };
}

// =====================================================
// DELETE ENTRY (own entries only)
// =====================================================

function deleteMetricEntry(metric, rowNumber) {
  const sheet = getMetricSheet_(metric);
  const cfg = METRICS[metric];
  const email = getCurrentUserEmail();

  rowNumber = Number(rowNumber);
  if (!rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    throw new Error('Invalid entry.');
  }

  const rowValues = sheet.getRange(rowNumber, 1, 1, cfg.headers.length).getValues()[0];

  const rowEmail = String(rowValues[cfg.emailCol - 1] || '')
    .toLowerCase().trim();

  if (rowEmail !== email) {
    throw new Error('You can only delete your own entries.');
  }

  const deletedTeam = rowTeam_(rowValues[cfg.teamCol - 1], getUserTeamCached_());

  sheet.deleteRow(rowNumber);
  return metricUpdate_(metric, deletedTeam);
}

// =====================================================
// LOAD (today entries + monthly summary, current user only)
// =====================================================

function getMetricData_(metric, team) {
  const cfg = METRICS[metric];
  const sheet = getMetricSheet_(metric);
  const lastRow = sheet.getLastRow();

  const out = {
    todayEntries: [],
    monthly: {
      avg: 0, sum: 0, entriesCount: 0, dailyBreakdown: [],
      maxValue: null, minValue: null,
      entryComparison: { previous: null, current: null, difference: null }
    }
  };

  if (lastRow < 2) return out;

  const email = getCurrentUserEmail();
  const tz = Session.getScriptTimeZone() || 'Africa/Cairo';
  const now = new Date();
  const todayKey = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const monthKey = Utilities.formatDate(now, tz, 'yyyy-MM');

  const values = readRows_(sheet, lastRow, cfg.headers.length);
  const dailyMap = {};
  const sequence = [];

  values.forEach(function(row, i) {
    if (!row[0]) return;
    if (String(row[cfg.emailCol - 1] || '').toLowerCase().trim() !== email) return;

    // Only the requested team (old rows = user's saved team)
    if (team && rowTeam_(row[cfg.teamCol - 1], getUserTeamCached_()) !== team) return;

    const dateKey = Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd');
    if (dateKey.slice(0, 7) !== monthKey) return;

    const value = Number(row[2]) || 0;

    out.monthly.sum += value;
    out.monthly.entriesCount++;
    sequence.push(value);

    // best / worst single entry of the month (Quality: best and worst monitor)
    if (out.monthly.maxValue === null || value > out.monthly.maxValue) out.monthly.maxValue = value;
    if (out.monthly.minValue === null || value < out.monthly.minValue) out.monthly.minValue = value;

    if (!dailyMap[dateKey]) dailyMap[dateKey] = { date: dateKey, count: 0, sum: 0 };
    dailyMap[dateKey].count++;
    dailyMap[dateKey].sum += value;

    if (dateKey === todayKey) {
      out.todayEntries.push({
        rowNumber: i + 2,
        time: Utilities.formatDate(new Date(row[1]), tz, 'HH:mm:ss'),
        value: value
      });
    }
  });

  out.todayEntries.reverse();

  if (out.monthly.entriesCount > 0) {
    out.monthly.avg = out.monthly.sum / out.monthly.entriesCount;
  }

  if (sequence.length > 0) {
    out.monthly.entryComparison.current = sequence[sequence.length - 1];
  }
  if (sequence.length > 1) {
    out.monthly.entryComparison.previous = sequence[sequence.length - 2];
    out.monthly.entryComparison.difference =
      out.monthly.entryComparison.current - out.monthly.entryComparison.previous;
  }

  // Productivity is a count: day value = total of the day,
  // monthly value = average per day. Other KPIs = average of entries.
  const isCount = (metric === 'productivity');

  out.monthly.dailyBreakdown = Object.keys(dailyMap).sort().map(function(k) {
    const d = dailyMap[k];
    return { date: d.date, count: d.count, avg: isCount ? d.sum : d.sum / d.count, sum: d.sum };
  });

  if (isCount && out.monthly.dailyBreakdown.length > 0) {
    out.monthly.avg = out.monthly.sum / out.monthly.dailyBreakdown.length;
  }

  return out;
}

function loadMetricsData() {

  __rowsCache = {};

  const teams = {};

  TEAMS.forEach(function(t) {
    teams[t] = {
      quality: getMetricData_('quality', t),
      aht: getMetricData_('aht', t),
      productivity: getMetricData_('productivity', t)
    };
  });

  return {
    team: getUserTeamCached_(),
    teams: teams,
    settings: readSettings_()
  };
}


// =====================================================
// RESET ALL MY DATA (CURRENT MONTH)
// Deletes ONLY the current user's rows of the current month
// from CSAT Entries, Quality Entries and AHT Entries.
// Other users and previous months are not touched.
// =====================================================

function deleteMyMonthRows_(sheet, emailCol, email, tz, monthKey) {

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const values = sheet.getRange(2, 1, lastRow - 1, emailCol).getValues();
  let count = 0;

  // Go from bottom to top so row numbers stay valid while deleting
  for (let i = values.length - 1; i >= 0; i--) {

    const row = values[i];
    if (!row[0]) continue;

    if (String(row[emailCol - 1] || '').toLowerCase().trim() !== email) continue;

    if (Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM') !== monthKey) continue;

    sheet.deleteRow(i + 2);
    count++;
  }

  return count;
}

function resetAllMyData() {

  const email = getCurrentUserEmail();
  const tz = Session.getScriptTimeZone() || 'Africa/Cairo';
  const monthKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM');

  let deleted = 0;

  deleted += deleteMyMonthRows_(getDataSheet(), 11, email, tz, monthKey);
  deleted += deleteMyMonthRows_(getMetricSheet_('quality'), METRICS.quality.emailCol, email, tz, monthKey);
  deleted += deleteMyMonthRows_(getMetricSheet_('aht'), METRICS.aht.emailCol, email, tz, monthKey);
  deleted += deleteMyMonthRows_(getMetricSheet_('productivity'), METRICS.productivity.emailCol, email, tz, monthKey);

  return {
    success: true,
    message: '✓ Reset done (' + deleted + ' entries deleted)',
    csat: loadCSATData(),
    metrics: loadMetricsData()
  };
}


// =====================================================
// ADD SEVERAL QUALITY SCORES AT ONCE
// (all scores are validated first; if one is invalid, nothing is saved)
// =====================================================

function addMetricEntries(metric, rawValues, team) {

  if (metric !== 'quality') {
    throw new Error('Multiple entries are only supported for Quality.');
  }

  if (!rawValues || !rawValues.length) {
    throw new Error('Please enter at least one score.');
  }

  const values = rawValues.map(function(raw) {
    const v = Number(String(raw).replace('%', '').trim());
    if (!isFinite(v) || v < 0 || v > 100) {
      throw new Error('Each Quality score must be a number between 0 and 100. Invalid: "' + raw + '"');
    }
    return v;
  });

  const sheet = getMetricSheet_(metric);
  const email = getCurrentUserEmail();
  const now = new Date();
  const entryTeam = normalizeTeam_(team);

  const startRow = sheet.getLastRow() + 1;
  const rows = values.map(function(value) { return [now, now, value, email, entryTeam]; });

  sheet.getRange(startRow, 1, rows.length, 5).setValues(rows);
  sheet.getRange(startRow, 1, rows.length, 3).setNumberFormats(
    rows.map(function() { return ['yyyy-MM-dd', 'HH:mm:ss', '0.00']; })
  );

  return metricUpdate_(metric, entryTeam);
}



// =====================================================
// TEAM (Calls / Tickets)
// Stored in the sheet 'Team Settings' (User Email | Team).
// A manager can also edit this sheet directly.
// Default team = calls.
// =====================================================

const TEAM_SHEET_NAME = 'Team Settings';
const TEAMS = ['calls', 'tickets'];

function getTeamSheet_() {

  if (__sheets[TEAM_SHEET_NAME]) return __sheets[TEAM_SHEET_NAME];

  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(TEAM_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(TEAM_SHEET_NAME);
    sheet.getRange(1, 1, 1, 2)
      .setValues([['User Email', 'Team']])
      .setFontWeight('bold');
  }

  __sheets[TEAM_SHEET_NAME] = sheet;
  return sheet;
}

function findTeamRow_(sheet, email) {

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').toLowerCase().trim() === email) {
      return i + 2;
    }
  }

  return 0;
}

function getUserTeam_() {

  const email = getCurrentUserEmail();
  const sheet = getTeamSheet_();
  const row = findTeamRow_(sheet, email);

  if (!row) return 'calls';

  const team = String(sheet.getRange(row, 2).getValue() || '').toLowerCase().trim();
  return TEAMS.indexOf(team) !== -1 ? team : 'calls';
}

function saveUserTeam(team) {

  team = String(team || '').toLowerCase().trim();

  if (TEAMS.indexOf(team) === -1) {
    throw new Error('Invalid team.');
  }

  const email = getCurrentUserEmail();
  const sheet = getTeamSheet_();
  const row = findTeamRow_(sheet, email);

  if (row) {
    sheet.getRange(row, 2).setValue(team);
  } else {
    sheet.appendRow([email, team]);
  }

  return team;
}



// =====================================================
// SETTINGS (weights + targets)
// Sheet tab 'Settings':  Team | KPI | Weight % | Target
// The manager edits the numbers there; the tool reads them
// every time the page is opened.
//   - Weight blank  = the KPI is not counted for that team
//   - Target        = CSAT / Quality in %, Productivity as a number,
//                     AHT as mm:ss (example 04:00)
//   - Weights of a team + its Attendance should total 100
// If the sheet is missing it is created with the default numbers.
// =====================================================

const SETTINGS_SHEET_NAME = 'Settings';

const DEFAULT_SETTINGS_ROWS = [
  ['calls',   'CSAT',         25, 94],
  ['calls',   'AHT',          15, '04:00'],
  ['calls',   'Quality',      35, 95],
  ['calls',   'Productivity', '', ''],
  ['calls',   'Attendance',   25, ''],
  ['tickets', 'CSAT',         25, 90],
  ['tickets', 'AHT',          '', ''],
  ['tickets', 'Quality',      35, 95],
  ['tickets', 'Productivity', 25, 90],
  ['tickets', 'Attendance',   15, '']
];

function getSettingsSheet_() {

  if (__sheets[SETTINGS_SHEET_NAME]) return __sheets[SETTINGS_SHEET_NAME];

  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);

    // Target column as plain text, so 04:00 stays 04:00
    sheet.getRange(2, 4, 50, 1).setNumberFormat('@');

    sheet.getRange(1, 1, 1, 5)
      .setValues([['Team', 'KPI', 'Weight %', 'Target', 'Notes']])
      .setFontWeight('bold');

    sheet.getRange(2, 1, DEFAULT_SETTINGS_ROWS.length, 4)
      .setValues(DEFAULT_SETTINGS_ROWS);

    sheet.getRange(2, 5, 4, 1).setValues([
      ['Team: calls or tickets.'],
      ['Weight: leave it empty if the KPI is not counted for that team.'],
      ['Target: CSAT and Quality in %, Productivity as a number, AHT as mm:ss (example 04:00).'],
      ['The weights of a team + its Attendance should total 100.']
    ]);

    sheet.setFrozenRows(1);
  }

  __sheets[SETTINGS_SHEET_NAME] = sheet;
  return sheet;
}

function parseSettingNumber_(value) {
  if (value === '' || value === null || value === undefined) return null;
  const n = Number(String(value).replace('%', '').trim());
  return isFinite(n) ? n : null;
}

function parseSettingTarget_(cap, value) {

  if (value === '' || value === null || value === undefined) return null;

  if (cap === 'Aht') {

    // a time cell: 04:00 is read as 4 minutes 0 seconds
    if (value instanceof Date) {
      const hhmm = Utilities.formatDate(value, getSpreadsheet_().getSpreadsheetTimeZone(), 'HH:mm').split(':');
      return Number(hhmm[0]) * 60 + Number(hhmm[1]);
    }

    // a plain number = seconds
    if (typeof value === 'number') return value > 0 ? value : null;

    const m = String(value).trim().match(/^(\d{1,3}):([0-5]?\d)$/);
    return m ? Number(m[1]) * 60 + Number(m[2]) : null;
  }

  const n = parseSettingNumber_(value);
  return (n !== null && n > 0) ? n : null;
}

function readSettings_() {

  const out = { teams: {}, warnings: [] };

  try {

    const sheet = getSettingsSheet_();
    const lastRow = sheet.getLastRow();
    const rows = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 4).getValues() : [];

    const CAPS = { csat: 'Csat', aht: 'Aht', quality: 'Quality', productivity: 'Productivity' };

    TEAMS.forEach(function(t) {
      out.teams[t] = { attendance: null, weights: {}, targets: {} };
    });

    rows.forEach(function(r) {

      const team = String(r[0] || '').toLowerCase().trim();
      const kpi = String(r[1] || '').toLowerCase().trim();

      if (!out.teams[team] || !kpi) return;

      const weight = parseSettingNumber_(r[2]);

      if (kpi === 'attendance') {
        if (weight !== null) out.teams[team].attendance = weight;
        return;
      }

      const cap = CAPS[kpi];
      if (!cap) return;

      if (weight !== null) out.teams[team].weights[cap] = weight;

      const target = parseSettingTarget_(cap, r[3]);
      if (target !== null) out.teams[team].targets[cap] = target;
    });

    TEAMS.forEach(function(t) {

      const s = out.teams[t];
      const keys = Object.keys(s.weights);

      // a team without any weight keeps the numbers built into the page
      if (!keys.length) {
        delete out.teams[t];
        out.warnings.push('Settings: no weights found for "' + t + '", the default numbers are used.');
        return;
      }

      const total = keys.reduce(function(sum, k) { return sum + s.weights[k]; }, 0) + (s.attendance || 0);

      if (Math.abs(total - 100) > 0.01) {
        out.warnings.push(
          'Settings: the weights of "' + t + '" total ' + (Math.round(total * 100) / 100) + '% (should be 100%).'
        );
      }
    });

  } catch (e) {
    out.teams = {};
    out.warnings.push('Settings could not be read, the default numbers are used. (' + e.message + ')');
  }

  return out;
}
