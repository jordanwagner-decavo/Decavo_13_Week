/**
 * ============================================================================
 *  13-WEEK RELEASE — AUTOMATIC CHANGE LOG
 * ============================================================================
 *
 *  PURPOSE
 *    Appends a row to "Change Log" whenever a Planned Qty cell changes.
 *    Quantity moved between weeks is merged into one PUSH-OUT / PULL-IN row.
 *    Actual columns are ignored.
 *
 *  CHANGE TYPES
 *    PUSH-OUT  — qty moved to a later week
 *    PULL-IN   — qty moved to an earlier week
 *    INCREASE  — qty increased on a single week (no paired edit)
 *    DECREASE  — qty decreased on a single week, not to zero (no paired edit)
 *    ZERO      — qty zeroed out on a single week (no paired edit)
 *
 *  STATUS:  LIVE — source "13 Week Release", log "Change Log", forecast
 *           "Fiscal Year Forecast", P&P resolved from the dated "P&P MM.DD" tab.
 *
 *  COLUMN STRUCTURE (Change Log)
 *    A: SO | B: Part No. | C: Description | D: Unit Price
 *    E: Moved Qty | F: Moved Value | G: Original Date | H: New Date
 *    I: Weeks Moved | J: Change Type | K: Notes | L: Date Time | M: User
 *
 *  CONTACT / OWNER:  Jordan Wagner
 *  LAST UPDATED:     2026-06-16
 * ============================================================================
 */

const CONFIG = {
  SOURCE_SHEET:    '13 Week Release',
  LOG_SHEET:       'Change Log',
  WEEK_DATE_ROW:   2,
  HEADER_ROW:      3,
  FIRST_DATA_ROW:  5,
  FIRST_WEEK_COL:  10,   // col J — first week's Planned Qty (cols H/I hold Unscheduled Value/Qty)
  GROUP_SIZE:      4,
  UNIT_PRICE_COL:  5,
  MERGE_WINDOW_MS: 5 * 60 * 1000  // 5 minutes
};

// ─── Column indices in the log sheet (1-based) ───────────────────────────────
const LC = {
  SO:           1,
  PART_NO:      2,
  DESCRIPTION:  3,
  UNIT_PRICE:   4,
  MOVED_QTY:    5,
  MOVED_VALUE:  6,
  ORIG_DATE:    7,
  NEW_DATE:     8,
  WEEKS_MOVED:  9,
  CHANGE_TYPE: 10,
  NOTES:       11,
  DATETIME:    12,
  USER:        13
};

const LOG_COLS = 13;
const MS_PER_WEEK = 7 * 24 * 60 * 60 * 1000;

/**
 * Phase 1: reports what the script sees, without changing anything.
 */
function diagnose() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  const log   = ss.getSheetByName(CONFIG.LOG_SHEET);

  Logger.log('Source sheet found: ' + (sheet ? 'YES' : 'NO — check SOURCE_SHEET'));
  Logger.log('Log sheet found: '    + (log   ? 'YES' : 'NO — check LOG_SHEET'));
  if (!sheet) return;

  const row = CONFIG.FIRST_DATA_ROW;
  Logger.log('--- Row ' + row + ' identity ---');
  Logger.log('SO (A): '          + sheet.getRange(row, 1).getDisplayValue());
  Logger.log('Part No. (C): '    + sheet.getRange(row, 3).getDisplayValue());
  Logger.log('Description (D): ' + sheet.getRange(row, 4).getDisplayValue());
  Logger.log('Unit Price (E): '  + sheet.getRange(row, CONFIG.UNIT_PRICE_COL).getValue());

  Logger.log('--- First two weekly groups ---');
  for (let g = 0; g < 2; g++) {
    const qtyCol    = CONFIG.FIRST_WEEK_COL + g * CONFIG.GROUP_SIZE;
    const colLetter = columnToLetter_(qtyCol);
    const weekDate  = sheet.getRange(CONFIG.WEEK_DATE_ROW, qtyCol).getDisplayValue();
    const heading   = sheet.getRange(CONFIG.HEADER_ROW,    qtyCol).getDisplayValue();
    const cellVal   = sheet.getRange(row, qtyCol).getDisplayValue();
    Logger.log('Group ' + (g + 1) + ': Planned Qty is column ' + colLetter +
               ' | heading="' + heading + '" | week date above="' + weekDate +
               '" | row ' + row + ' value="' + cellVal + '"');
  }
}

/**
 * Phase 2 + 4: single-cell change logging with move-merge.
 * Registered as an installable on-edit trigger.
 */
function onEditInstallable(e) {
  const range = e.range;
  const sheet = range.getSheet();

  if (sheet.getName() !== CONFIG.SOURCE_SHEET) return;

  const row = range.getRow();
  const col = range.getColumn();

  if (row < CONFIG.FIRST_DATA_ROW)   return;
  if (!isPlannedQtyCol_(col))         return;
  if (range.getNumRows() > 1 || range.getNumColumns() > 1) return;  // paste — Phase 3

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (!logSheet) return;

  const unitPrice = sheet.getRange(row, CONFIG.UNIT_PRICE_COL).getValue();
  const weekDate  = sheet.getRange(CONFIG.WEEK_DATE_ROW, col).getDisplayValue();
  const oldQty    = (e.oldValue !== undefined && e.oldValue !== '') ? Number(e.oldValue) : 0;
  const newQty    = (e.value    !== undefined && e.value    !== '') ? Number(e.value)    : 0;
  const delta     = newQty - oldQty;

  if (delta === 0) return;  // no real change

  const so          = sheet.getRange(row, 1).getDisplayValue();
  const partNo      = sheet.getRange(row, 3).getDisplayValue();
  const description = sheet.getRange(row, 4).getDisplayValue();
  const now         = new Date();
  const user        = Session.getActiveUser().getEmail();

  // ── Try to merge with the most recent pending entry for this part ──────────
  const lastRow = getLastLogRowForPart_(logSheet, partNo);

  if (lastRow) {
    const lastData       = logSheet.getRange(lastRow, 1, 1, LOG_COLS).getValues()[0];
    const lastChangeType = lastData[LC.CHANGE_TYPE - 1];
    const lastMovedQty   = Number(lastData[LC.MOVED_QTY - 1]);
    const lastTime       = lastData[LC.DATETIME - 1];
    const timeDiff       = now - new Date(lastTime);

    // A pending loss stored its week in ORIG_DATE; a pending gain in NEW_DATE.
    const prevWeek = (lastChangeType === 'INCREASE')
      ? lastData[LC.NEW_DATE  - 1]
      : lastData[LC.ORIG_DATE - 1];

    // Reconstruct the previous edit's SIGNED delta from its change type.
    const lastDelta = (lastChangeType === 'INCREASE') ? lastMovedQty : -lastMovedQty;

    const isPending      = ['INCREASE', 'DECREASE', 'ZERO'].includes(lastChangeType);
    const deltasOpposite = (delta === -lastDelta);   // equal magnitude, opposite sign
    const withinWindow   = timeDiff <= CONFIG.MERGE_WINDOW_MS;

    if (isPending && deltasOpposite && withinWindow) {
      let origDate, newDate;

      if (delta > 0) {
        // current edit is the GAIN; previous entry was the LOSS (the source)
        origDate = prevWeek;
        newDate  = weekDate;
      } else {
        // current edit is the LOSS (the source); previous entry was the GAIN
        origDate = weekDate;
        newDate  = prevWeek;
      }

      const movedQty    = Math.abs(delta);
      const movedValue  = movedQty * unitPrice;
      const oDate       = new Date(origDate);
      const nDate       = new Date(newDate);
      const weeksMoved  = Math.round((nDate - oDate) / MS_PER_WEEK);
      const changeType  = nDate > oDate ? 'PUSH-OUT' : 'PULL-IN';

      logSheet.getRange(lastRow, 1, 1, LOG_COLS).setValues([[
        so, partNo, description, unitPrice,
        movedQty, movedValue,
        origDate, newDate, weeksMoved,
        changeType,
        '',         // Notes
        lastTime,   // keep original timestamp
        user
      ]]);

      return;
    }
  }

  // ── No merge — append a single-side row ────────────────────────────────────
  let changeType;
  if (newQty === 0)   changeType = 'ZERO';
  else if (delta > 0) changeType = 'INCREASE';
  else                changeType = 'DECREASE';

  const movedQty   = Math.abs(delta);
  const movedValue = movedQty * unitPrice;

  // Gain → record week in New Date; loss/zero → record week in Original Date.
  const origDate = (delta > 0) ? '' : weekDate;
  const newDate  = (delta > 0) ? weekDate : '';

  logSheet.appendRow([
    so, partNo, description, unitPrice,
    movedQty, movedValue,
    origDate, newDate, '',   // Weeks Moved blank until paired
    changeType,
    '',   // Notes
    now, user
  ]);
}

/**
 * Finds the last log row for a given part number, returns row number or null.
 */
function getLastLogRowForPart_(logSheet, partNo) {
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return null;

  const values = logSheet.getRange(2, LC.PART_NO, lastRow - 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]) === String(partNo)) {
      return i + 2;
    }
  }
  return null;
}

/** Returns true if col is a Planned Qty column (offset 0 in each 4-col group). */
function isPlannedQtyCol_(col) {
  if (col < CONFIG.FIRST_WEEK_COL) return false;
  return (col - CONFIG.FIRST_WEEK_COL) % CONFIG.GROUP_SIZE === 0;
}

function columnToLetter_(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}