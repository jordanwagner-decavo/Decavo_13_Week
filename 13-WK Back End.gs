/**
 * ============================================================================
 *  13-WEEK RELEASE — WEB APP (rationale-enforced editor)
 * ============================================================================
 *  submitBatch logs the groups the front end already computed and showed in the
 *  review modal. getScheduleData returns planned qtys plus actual qty/sales and
 *  the current-week index for the value/actuals display.
 *
 *  Relies on CONFIG, LC, LOG_COLS, MS_PER_WEEK, isPlannedQtyCol_,
 *  getLastLogRowForPart_ from the main script file.
 * ============================================================================
 */

const RATIONALE_PRESETS = [
  'Customer request — date change',
  'Customer request — quantity change',
  'Material / supply shortage',
  'Capacity / resource constraint',
  'Production schedule rebalance',
  'Quality hold / rework',
  'Forecast / demand update',
  'Awaiting customer feedback / change',
  'Data correction'
];

// ─── Fiscal Year Forecast tab ────────────────────────────────────────────────
// Months in col A (e.g. "June 26"), FY budget in col C, Current Sales BKLG written
// to col D, Delta (D − C) written to col E. Data starts at row 3.
const FORECAST = {
  SHEET:      'Fiscal Year Forecast',
  FIRST_ROW:  3,
  MONTH_COL:  1,   // A
  BUDGET_COL: 3,   // C
  BKLG_COL:   4,   // D  (app writes)
  DELTA_COL:  5    // E  (app writes)
};
const MONTH_NAMES = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'];

// ─── P&P (SO export from ERP) ────────────────────────────────────────────────
// Doc No (SO#) col A, Part Number col E, Open Qty col G, Delivery/Due Date col I.
// The tab is renamed "P&P MM.DD" (month.day of the update, no leading zeros)
// each time it's refreshed, so the exact name is resolved at runtime rather than
// hard-coded — see resolvePPSheet_().
const PP = {
  SHEET_PREFIX: 'P&P',   // dated tab "P&P MM.DD" resolved via resolvePPSheet_()
  FIRST_ROW: 2,    // row 1 = headers
  SO_COL:    1,    // A
  PART_COL:  5,    // E
  QTY_COL:   7,    // G  (Open Qty = demand)
  DUE_COL:   9     // I  (Delivery / Due Date)
};

// Normalize any month label/date to a "YYYY-MM" key.
function monthKey_(val) {
  if (val instanceof Date) return val.getFullYear() + '-' + ('0' + (val.getMonth() + 1)).slice(-2);
  var s = String(val).trim();
  if (!s) return '';
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
  // Fallback: "June 26" style → parse name + 2-digit year
  var m = s.match(/^([A-Za-z]+)\s*'?(\d{2,4})$/);
  if (m) {
    var idx = MONTH_NAMES.indexOf(m[1].toUpperCase());
    if (idx === -1) { var abbr = m[1].slice(0,3).toUpperCase(); idx = MONTH_NAMES.map(function(n){return n.slice(0,3);}).indexOf(abbr); }
    if (idx >= 0) { var yy = m[2].length === 2 ? (2000 + Number(m[2])) : Number(m[2]); return yy + '-' + ('0' + (idx + 1)).slice(-2); }
  }
  return s;
}

// Finds the current P&P tab. Its name is "P&P MM.DD" (month.day of the last
// update, no leading zeros), so we match that dated pattern and, if more than
// one exists, keep the most recent by date. Returns the Sheet or null.
function resolvePPSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var re = new RegExp('^' + PP.SHEET_PREFIX.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s+(\\d{1,2})\\.(\\d{1,2})$');
  var best = null, bestKey = -1;
  ss.getSheets().forEach(function (sh) {
    var m = sh.getName().match(re);
    if (!m) return;
    var key = Number(m[1]) * 100 + Number(m[2]);   // MM.DD → sortable number
    if (key > bestKey) { bestKey = key; best = sh; }
  });
  return best;
}

// Reads the P&P export → { 'SO|PART': [ {due:'YYYY-MM-DD', qty:Number}, ... ] }.
function readPP_() {
  var sh = resolvePPSheet_();
  var out = {};
  if (!sh) return out;
  var last = sh.getLastRow();
  if (last < PP.FIRST_ROW) return out;
  var n = last - PP.FIRST_ROW + 1;
  var vals = sh.getRange(PP.FIRST_ROW, 1, n, PP.DUE_COL).getValues();
  for (var i = 0; i < vals.length; i++) {
    var so   = String(vals[i][PP.SO_COL - 1] || '').trim();
    var part = String(vals[i][PP.PART_COL - 1] || '').trim();
    if (!so || !part) continue;
    var qty = Number(vals[i][PP.QTY_COL - 1]) || 0;
    if (qty <= 0) continue;
    var dueRaw = vals[i][PP.DUE_COL - 1];
    var due = (dueRaw instanceof Date) ? dueRaw : new Date(dueRaw);
    if (isNaN(due.getTime())) continue;
    var iso = due.getFullYear() + '-' + ('0' + (due.getMonth() + 1)).slice(-2) + '-' + ('0' + due.getDate()).slice(-2);
    var key = so + '|' + part;
    (out[key] = out[key] || []).push({ due: iso, qty: qty });
  }
  return out;
}

// Reads the forecast tab → { 'YYYY-MM': {budget, row} }.
function readForecast_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(FORECAST.SHEET);
  var out = {};
  if (!sh) return out;
  var last = sh.getLastRow();
  if (last < FORECAST.FIRST_ROW) return out;
  var n = last - FORECAST.FIRST_ROW + 1;
  var vals = sh.getRange(FORECAST.FIRST_ROW, 1, n, FORECAST.DELTA_COL).getValues();
  for (var i = 0; i < vals.length; i++) {
    var key = monthKey_(vals[i][FORECAST.MONTH_COL - 1]);
    if (!key) continue;
    out[key] = { budget: Number(vals[i][FORECAST.BUDGET_COL - 1]) || 0, row: FORECAST.FIRST_ROW + i };
  }
  return out;
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('13-WK Front End')
    .setTitle('DECAVO 13-Week Plan')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Reads the schedule grid for the front end.
 *
 * Each week is a GROUP_SIZE (4) column block starting at the planned-qty column:
 *   offset 0 = Planned Qty, 1 = Planned Sales, 2 = Actual Qty, 3 = Actual Sales.
 * If actuals live at different offsets in your sheet, change the two constants
 * below — they are the only place that mapping is defined.
 */
function getScheduleData() {
  var ACT_QTY_OFF   = 2;   // actual qty   column offset within the 4-col group
  var ACT_SALES_OFF = 3;   // actual sales column offset within the 4-col group

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  if (!sheet) throw new Error('Source sheet "' + CONFIG.SOURCE_SHEET + '" was not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const weeks = [];
  for (let col = CONFIG.FIRST_WEEK_COL; col <= lastCol; col += CONFIG.GROUP_SIZE) {
    const d = sheet.getRange(CONFIG.WEEK_DATE_ROW, col).getDisplayValue();
    if (!d) break;
    weeks.push({ col: col, date: d });
  }

  // Current week = first week whose ending date is today or later.
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var currentWeekIndex = -1;
  for (var wi = 0; wi < weeks.length; wi++) {
    var wd = new Date(weeks[wi].date);
    if (!isNaN(wd.getTime()) && wd >= today) { currentWeekIndex = wi; break; }
  }
  if (currentWeekIndex === -1 && weeks.length) currentWeekIndex = weeks.length - 1;

  var ppDemand = readPP_();

  const rows = [];
  if (lastRow >= CONFIG.FIRST_DATA_ROW) {
    const numRows = lastRow - CONFIG.FIRST_DATA_ROW + 1;
    const all = sheet.getRange(CONFIG.FIRST_DATA_ROW, 1, numRows, lastCol).getValues();
    for (let i = 0; i < all.length; i++) {
      const r  = all[i];
      const so = r[0];
      // Only real order lines: the SO# in col A must be a 5-digit number.
      // Anything else (blank, notes, headers, trailing bad-data rows) is excluded.
      if (!/^\d{5}$/.test(String(so).trim())) continue;

      const qtys = [], actualQtys = [], actualVals = [];
      weeks.forEach(function (w) {
        // Number() on non-numeric cell text yields NaN, which is not valid JSON
        // and makes google.script.run silently hand the client `null`. Coerce any
        // non-numeric planned/actual value to blank/0 so the payload stays safe.
        var qv = r[w.col - 1];
        var qn = Number(qv);
        qtys.push((qv === '' || qv === null || isNaN(qn)) ? '' : qn);
        var aqn = Number(r[w.col - 1 + ACT_QTY_OFF]);
        actualQtys.push(isNaN(aqn) ? 0 : aqn);
        var avn = Number(r[w.col - 1 + ACT_SALES_OFF]);
        actualVals.push(isNaN(avn) ? 0 : avn);
      });

      rows.push({
        rowNum:      CONFIG.FIRST_DATA_ROW + i,
        so:          String(so),
        partNo:      String(r[2]),
        description: String(r[3]),
        unitPrice:   Number(r[CONFIG.UNIT_PRICE_COL - 1]) || 0,
        qtys:        qtys,
        actualQtys:  actualQtys,
        actualVals:  actualVals,
        demand:      ppDemand[String(so) + '|' + String(r[2])] || []
      });
    }
  }

  var forecast = readForecast_();
  var budgets = {};
  Object.keys(forecast).forEach(function (k) { budgets[k] = forecast[k].budget; });

  return {
    sheetName:        CONFIG.SOURCE_SHEET,
    weeks:            weeks.map(function (w) { return w.date; }),
    weekCols:         weeks.map(function (w) { return w.col; }),
    weekMonths:       weeks.map(function (w) { return monthKey_(w.date); }),
    currentWeekIndex: currentWeekIndex,
    rows:             rows,
    budgets:          budgets,
    presets:          RATIONALE_PRESETS
  };
}

/**
 * Applies one Planned Qty change from the web app (single-cell path).
 */
function submitChange(payload) {
  const rationale = (payload.rationale || '').toString().trim();
  if (!rationale) throw new Error('A rationale is required to save this change.');

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const sheet    = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (!sheet || !logSheet) throw new Error('Source or log sheet not found.');

  const row = Number(payload.rowNum);
  const col = Number(payload.col);
  if (!isPlannedQtyCol_(col)) throw new Error('That column is not an editable Planned Qty column.');

  const cell   = sheet.getRange(row, col);
  const oldQty = Number(cell.getValue()) || 0;
  const newQty = (payload.newQty === '' || payload.newQty === null) ? 0 : Number(payload.newQty);
  if (isNaN(newQty)) throw new Error('Quantity must be a number.');

  const delta = newQty - oldQty;
  if (delta === 0) return { ok: true, changed: false };

  cell.setValue(newQty === 0 ? '' : newQty);
  logChangeWithRationale_(sheet, logSheet, row, col, oldQty, newQty, rationale);

  return { ok: true, changed: true, newQty: newQty };
}

/**
 * Applies a batch of changes the front end has already grouped.
 *   { kind:'move',      rationale, cells, sources:[edit], dests:[edit] }
 *   { kind:'synthetic', rationale, cells, qty, sources:[edit] }
 *   { kind:'single',    rationale, cells, edit }
 */
function submitBatch(groups) {
  if (!groups || !groups.length) return { ok: true, applied: 0, errors: [] };

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const sheet    = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (!sheet || !logSheet) throw new Error('Source or log sheet not found.');

  const now    = new Date();
  const user   = Session.getActiveUser().getEmail();
  const errors = [];

  groups.forEach(function (g) {
    (g.cells || []).forEach(function (c) {
      try {
        var col = Number(c.col);
        if (!isPlannedQtyCol_(col)) throw new Error('not an editable column');
        var nq = (c.newQty === '' || c.newQty == null) ? 0 : Number(c.newQty);
        if (isNaN(nq)) throw new Error('quantity is not a number');
        sheet.getRange(Number(c.rowNum), col).setValue(nq === 0 ? '' : nq);
      } catch (e) {
        errors.push('row ' + c.rowNum + ' / col ' + c.col + ': ' + (e.message || e));
      }
    });
  });

  function descOf(rowNum) { return sheet.getRange(Number(rowNum), 4).getDisplayValue(); }

  var logRows = [];

  groups.forEach(function (g) {
    var reason = (g.rationale || '').toString().trim();

    if (g.kind === 'move') {
      var src = g.sources[0];
      var noteSuffix = g.note ? (' — ' + g.note) : '';
      g.dests.forEach(function (dst) {
        var movedQty   = Number(dst.newQty) - Number(dst.oldQty);
        var unitPrice  = Number(dst.unitPrice || src.unitPrice || 0);
        var oDate      = new Date(src.weekDate);
        var nDate      = new Date(dst.weekDate);
        var weeksMoved = Math.round((nDate - oDate) / MS_PER_WEEK);
        var changeType = nDate > oDate ? 'PUSH-OUT' : 'PULL-IN';
        logRows.push([
          src.so, src.partNo, descOf(dst.rowNum), unitPrice,
          movedQty, movedQty * unitPrice,
          src.weekDate, dst.weekDate, weeksMoved,
          changeType, reason + noteSuffix, now, user
        ]);
      });

    } else if (g.kind === 'synthetic') {
      var s0 = g.sources[0];
      var up = Number(s0.unitPrice || 0);
      var qty = Number(g.qty);
      logRows.push([
        s0.so, s0.partNo, descOf(s0.rowNum), up,
        qty, qty * up,
        s0.weekDate, '', '',
        'DECREASE', reason, now, user
      ]);

    } else {
      var e   = g.edit;
      var dlt = Number(e.newQty) - Number(e.oldQty);
      var up2 = Number(e.unitPrice || 0);
      var mq  = Math.abs(dlt);
      var ct  = (Number(e.newQty) === 0) ? 'ZERO' : (dlt > 0 ? 'INCREASE' : 'DECREASE');
      logRows.push([
        e.so, e.partNo, descOf(e.rowNum), up2,
        mq, mq * up2,
        (dlt > 0 ? '' : e.weekDate),
        (dlt > 0 ? e.weekDate : ''),
        '',
        ct, reason, now, user
      ]);
    }
  });

  if (logRows.length) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, LOG_COLS)
            .setValues(logRows);
  }

  try { syncForecastBacklog_(); } catch (e) { Logger.log('forecast sync failed: ' + e.message); }

  return { ok: errors.length === 0, applied: logRows.length, errors: errors };
}

/**
 * Recomputes the full (unfiltered) monthly sales backlog across all order lines —
 * actuals for past weeks, planned for current/future — and writes col D (backlog)
 * and col E (delta = backlog − budget) on the forecast tab.
 */
function syncForecastBacklog_() {
  var ACT_QTY_OFF = 2, ACT_SALES_OFF = 3;
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  var fc    = ss.getSheetByName(FORECAST.SHEET);
  if (!sheet || !fc) return;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  // Week columns + month keys.
  var weekCols = [], weekMonths = [];
  for (var col = CONFIG.FIRST_WEEK_COL; col <= lastCol; col += CONFIG.GROUP_SIZE) {
    var d = sheet.getRange(CONFIG.WEEK_DATE_ROW, col).getDisplayValue();
    if (!d) break;
    weekCols.push(col); weekMonths.push(monthKey_(d));
  }

  // Current week index (today or later).
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var cwi = -1;
  for (var wi = 0; wi < weekCols.length; wi++) {
    var wd = new Date(sheet.getRange(CONFIG.WEEK_DATE_ROW, weekCols[wi]).getDisplayValue());
    if (!isNaN(wd.getTime()) && wd >= today) { cwi = wi; break; }
  }
  if (cwi === -1 && weekCols.length) cwi = weekCols.length - 1;

  // Sum sales per month across all rows.
  var byMonth = {};
  if (lastRow >= CONFIG.FIRST_DATA_ROW) {
    var all = sheet.getRange(CONFIG.FIRST_DATA_ROW, 1, lastRow - CONFIG.FIRST_DATA_ROW + 1, lastCol).getValues();
    for (var r = 0; r < all.length; r++) {
      if (all[r][0] === '' || all[r][0] === null) continue;
      var price = Number(all[r][CONFIG.UNIT_PRICE_COL - 1]) || 0;
      for (var w = 0; w < weekCols.length; w++) {
        var base = weekCols[w] - 1;
        var sales;
        if (w < cwi) {
          sales = Number(all[r][base + ACT_SALES_OFF]) || 0;        // actual sales (past)
        } else {
          var pq = all[r][base]; pq = (pq === '' || pq === null) ? 0 : Number(pq);
          sales = pq * price;                                       // planned sales
        }
        var mk = weekMonths[w];
        byMonth[mk] = (byMonth[mk] || 0) + sales;
      }
    }
  }

  // Write D (backlog) and E (delta) for each forecast row we can match.
  var forecast = readForecast_();
  Object.keys(forecast).forEach(function (mk) {
    var row = forecast[mk].row;
    var bklg = byMonth[mk] || 0;
    var delta = bklg - forecast[mk].budget;
    fc.getRange(row, FORECAST.BKLG_COL).setValue(bklg);
    fc.getRange(row, FORECAST.DELTA_COL).setValue(delta);
  });
}

/**
 * Single-cell logging with move-merge (used by submitChange only).
 */
function logChangeWithRationale_(sheet, logSheet, row, col, oldQty, newQty, rationale) {
  const unitPrice   = sheet.getRange(row, CONFIG.UNIT_PRICE_COL).getValue();
  const weekDate    = sheet.getRange(CONFIG.WEEK_DATE_ROW, col).getDisplayValue();
  const delta       = newQty - oldQty;
  const so          = sheet.getRange(row, 1).getDisplayValue();
  const partNo      = sheet.getRange(row, 3).getDisplayValue();
  const description = sheet.getRange(row, 4).getDisplayValue();
  const now         = new Date();
  const user        = Session.getActiveUser().getEmail();

  const lastRow = getLastLogRowForPart_(logSheet, partNo);

  if (lastRow) {
    const lastData       = logSheet.getRange(lastRow, 1, 1, LOG_COLS).getValues()[0];
    const lastChangeType = lastData[LC.CHANGE_TYPE - 1];
    const lastMovedQty   = Number(lastData[LC.MOVED_QTY - 1]);
    const lastTime       = lastData[LC.DATETIME - 1];
    const lastNotes      = lastData[LC.NOTES - 1];
    const timeDiff       = now - new Date(lastTime);

    const prevWeek = (lastChangeType === 'INCREASE')
      ? lastData[LC.NEW_DATE  - 1]
      : lastData[LC.ORIG_DATE - 1];

    const lastDelta      = (lastChangeType === 'INCREASE') ? lastMovedQty : -lastMovedQty;
    const isPending      = ['INCREASE', 'DECREASE', 'ZERO'].indexOf(lastChangeType) !== -1;
    const deltasOpposite = (delta === -lastDelta);
    const withinWindow   = timeDiff <= CONFIG.MERGE_WINDOW_MS;

    if (isPending && deltasOpposite && withinWindow) {
      let origDate, newDate;
      if (delta > 0) { origDate = prevWeek; newDate = weekDate; }
      else           { origDate = weekDate; newDate = prevWeek; }

      const movedQty    = Math.abs(delta);
      const movedValue  = movedQty * unitPrice;
      const oDate       = new Date(origDate);
      const nDate       = new Date(newDate);
      const weeksMoved  = Math.round((nDate - oDate) / MS_PER_WEEK);
      const changeType  = nDate > oDate ? 'PUSH-OUT' : 'PULL-IN';

      const mergedNotes = (lastNotes && lastNotes !== rationale)
        ? (lastNotes + ' | ' + rationale)
        : rationale;

      logSheet.getRange(lastRow, 1, 1, LOG_COLS).setValues([[
        so, partNo, description, unitPrice,
        movedQty, movedValue, origDate, newDate, weeksMoved,
        changeType, mergedNotes, lastTime, user
      ]]);
      return;
    }
  }

  let changeType;
  if (newQty === 0)   changeType = 'ZERO';
  else if (delta > 0) changeType = 'INCREASE';
  else                changeType = 'DECREASE';

  const movedQty   = Math.abs(delta);
  const movedValue = movedQty * unitPrice;
  const origDate   = (delta > 0) ? '' : weekDate;
  const newDate    = (delta > 0) ? weekDate : '';

  logSheet.appendRow([
    so, partNo, description, unitPrice,
    movedQty, movedValue, origDate, newDate, '',
    changeType, rationale, now, user
  ]);
}