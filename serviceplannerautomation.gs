/***************************************************************************************
 * PM SERVICE PLANNER – RAW UPLOAD → BRANCH SPLITTER + FORMAT CLONER + ARCHIVER
 * Version: 1.0.0  (2025-10-16)
 * License: MIT
 * Author Credit: Created by Austin Monson (with ChatGPT assistance)
 *
 * -------------------------------------------------------------------------------------
 * VERSION HISTORY
 * 1.0.0 (2025-10-16)
 *  - Initial release:
 *    • Upload XLSX (all branches) via dialog → convert to Google Sheet
 *    • Split by Branch → build 3 tabs: Pending Service - Springfield / West Plains / Villa Ridge
 *    • Clone ALL styling from Springfield template tab: header, widths, validation, CF, etc.
 *    • Carry forward 'Due' + 'Notes' from previous branch tabs (keyed by 'Stock #')
 *    • Exclude rows whose carried 'Due' is 'Corrected' | 'Service not Needed' | 'Removed'
 *    • Archive old branch tabs as hidden _Archive_[Branch]_YYYY-MM-DD_hhmm
 *    • Restore UI to roll back an individual branch to any archived version
 *
 * -------------------------------------------------------------------------------------
 * LICENSE (MIT)
 * Copyright (c) 2025 Austin Monson
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 ***************************************************************************************/

/* =============================================================================
   CONFIGURATION (EDIT SAFELY)
   -----------------------------------------------------------------------------
   • TEMPLATE_SHEET: source of styling (we copy its entire look)
   • BRANCH_COLUMN : column name in raw file indicating branch
   • BRANCHES      : list of branches and their destination tab names
   • HEADER_ORDER  : the exact headers/order we want in destination sheets
   • DUE_FILTERS   : rows with these values in 'Due' are excluded
   • HEADER_ROW    : 1-based header row index in the styled template
============================================================================= */
const CONFIG = {
  TEMPLATE_SHEET: 'Pending Service - Springfield',  // uses this sheet’s formatting as the master template
  BRANCH_COLUMN: 'Branch',                           // required in the raw data file
  BRANCHES: [
    { key: 'Springfield',  tabName: 'Pending Service - Springfield' },
    { key: 'West Plains',  tabName: 'Pending Service - West Plains' },
    { key: 'Villa Ridge',  tabName: 'Pending Service - Villa Ridge' },
  ],
  HEADER_ORDER: [
    'Stock #',
    'Description',
    'Type',
    'Overdue',
    'Serial Number',
    'Manufacturer',
    'Model',
    'Due',
    'Notes'
  ],
  DUE_FILTERS: ['Corrected', 'Service not Needed', 'Removed'],
  HEADER_ROW: 1
};

/* =============================================================================
   MENU SETUP
============================================================================= */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PM Planner')
    .addItem('Upload Raw Data (XLSX → Split + Style + Archive)', 'showUploadDialog')
    .addSeparator()
    .addItem('Restore Archived Sheet…', 'showRestoreDialog')
    .addToUi();
}

/* =============================================================================
   DIALOG LAUNCHERS
============================================================================= */
function showUploadDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(520)
    .setHeight(440);
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload Raw PM Data');
}

function showRestoreDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Restore')
    .setWidth(540)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Restore Archived Branch Sheet');
}

/* =============================================================================
   CORE: HANDLE UPLOAD (from Dialog.html)
   - Accepts form object (with .file: Blob)
   - Stores XLSX temporarily in Drive
   - Converts to Google Sheet
   - Parses → splits → archives → rebuilds 3 styled tabs
============================================================================= */
function handleRawUpload(formObj) {
  _logProgress('Uploading file to Drive…');
  if (!formObj || !formObj.file) {
    throw new Error('No file received. Please select an .xlsx file.');
  }

  // 1) Save XLSX to Drive (temporary)
  const xlsxBlob = formObj.file; // HtmlService form sends Blob
  const tempFile = DriveApp.createFile(xlsxBlob);
  tempFile.setName('PM_Raw_Upload_' + _stampForFilename() + '.xlsx');

  try {
    // 2) Convert XLSX → Google Sheet (Advanced Drive API required)
    _logProgress('Converting to Google Sheet…');
    const fileMeta = {
      title: tempFile.getName(),
      mimeType: 'application/vnd.google-apps.spreadsheet'
    };
    const converted = Drive.Files.copy(fileMeta, tempFile.getId()); // Advanced Drive Service
    const convertedId = converted.id;

    // 3) Open converted SS and read raw data
    _logProgress('Reading raw rows…');
    const raw = SpreadsheetApp.openById(convertedId);
    const rawSheet = raw.getSheets()[0]; // assume first sheet in upload
    const rawValues = _getSheetDataAsObjects(rawSheet);

    // 4) Validate Branch column exists
    const hasBranch = rawValues.length === 0 ? false : (CONFIG.BRANCH_COLUMN in rawValues[0]);
    if (!hasBranch) {
      throw new Error('Branch column "' + CONFIG.BRANCH_COLUMN + '" not found in uploaded file.');
    }

    // 5) Build lookup by branch
    _logProgress('Splitting rows by branch…');
    const byBranch = {};
    CONFIG.BRANCHES.forEach(b => byBranch[b.key] = []);
    rawValues.forEach(row => {
      const b = String(row[CONFIG.BRANCH_COLUMN] || '').trim();
      if (byBranch[b]) byBranch[b].push(row);
    });

    // 6) For each branch: archive old tab → rebuild from template → merge carryover → filter → write
    const ss = SpreadsheetApp.getActive();
    const templateSheet = ss.getSheetByName(CONFIG.TEMPLATE_SHEET);
    if (!templateSheet) {
      throw new Error('Template sheet "' + CONFIG.TEMPLATE_SHEET + '" not found.');
    }

    CONFIG.BRANCHES.forEach(branch => {
      _logProgress('Processing branch: ' + branch.key + '…');

      // 6a) Carryover map from previous tab (if exists): {Stock # → {Due, Notes}}
      const carry = _readCarryoverMap(ss, branch.tabName);

      // 6b) Archive existing tab (if exists)
      _archiveSheetIfExists(ss, branch.tabName);

      // 6c) Create fresh branch sheet by duplicating template (inherit full styling)
      const newSheet = templateSheet.copyTo(ss).setName(branch.tabName);
      ss.setActiveSheet(newSheet);
      _ensureVisible(newSheet); // the active working sheet must be visible

      // 6d) Clear all rows under header (retain header row visuals)
      _clearDataBelowHeader(newSheet, CONFIG.HEADER_ROW);

      // 6e) Compose output rows using HEADER_ORDER, merging Due/Notes from carryover
      const inputRows = byBranch[branch.key] || [];
      const staged = inputRows.map(row => {
        const key = String(row['Stock #'] || '').trim();
        const carried = carry.get(key) || {};
        const due = carried.Due ?? row['Due'] ?? '';
        const notes = carried.Notes ?? row['Notes'] ?? '';

        return CONFIG.HEADER_ORDER.map(h => {
          if (h === 'Due')   return due;
          if (h === 'Notes') return notes;
          return row[h] ?? '';
        });
      });

      // 6f) Filter out rows by Due blacklist (Corrected/Service not Needed/Removed)
      const filtered = staged.filter(arr => {
        const dueVal = String(arr[CONFIG.HEADER_ORDER.indexOf('Due')] || '').trim();
        return CONFIG.DUE_FILTERS.indexOf(dueVal) === -1;
      });

      // 6g) Write data (if any)
      if (filtered.length) {
        const startRow = CONFIG.HEADER_ROW + 1;
        const startCol = 1;
        newSheet.getRange(startRow, startCol, filtered.length, CONFIG.HEADER_ORDER.length).setValues(filtered);
      }

      // 6h) Optional: ensure header labels match HEADER_ORDER
      _writeHeaderLabels(newSheet, CONFIG.HEADER_ROW, CONFIG.HEADER_ORDER);

      _logProgress('Finished ' + branch.key + ' (' + filtered.length + ' rows).');
    });

    // 7) Cleanup temp files
    _logProgress('Cleaning up temporary files…');
    Drive.Files.remove(convertedId); // remove converted intermediate
    tempFile.setTrashed(true);

    _logProgress('All branches updated. Archives created. You can close this dialog.');
    return { ok: true, message: 'Success' };
  } catch (err) {
    // Make sure temp file is deleted on error, too
    try { tempFile.setTrashed(true); } catch (e) {}
    throw err;
  }
}

/* =============================================================================
   RESTORE – ARCHIVE DISCOVERY + REPLACE
============================================================================= */
function listArchives() {
  const ss = SpreadsheetApp.getActive();
  const all = ss.getSheets();
  const results = [];

  all.forEach(sh => {
    const name = sh.getName();
    if (name.startsWith('_Archive_')) {
      // Parse pattern: _Archive_[Branch]_YYYY-MM-DD_hhmm
      results.push({
        name,
        hidden: sh.isSheetHidden(),
      });
    }
  });

  // Present archives grouped by the branch-prefix
  return results;
}

function restoreArchive(archiveSheetName) {
  if (!archiveSheetName || typeof archiveSheetName !== 'string') {
    throw new Error('No archive selected.');
  }
  const ss = SpreadsheetApp.getActive();
  const archiveSheet = ss.getSheetByName(archiveSheetName);
  if (!archiveSheet) {
    throw new Error('Archive sheet "' + archiveSheetName + '" not found.');
  }

  // Infer branch tabName target from archive name
  // Expecting: _Archive_[Branch]_YYYY-MM-DD_hhmm
  const branchKey = archiveSheetName.replace(/^_Archive_/, '').replace(/_\d{4}-\d{2}-\d{2}_\d{4}$/, '');
  const branch = CONFIG.BRANCHES.find(b => b.key === branchKey);
  if (!branch) {
    throw new Error('Could not infer branch from archive name: ' + archiveSheetName);
  }

  // Archive current live sheet first (safety), then replace with a copy of archive
  _archiveSheetIfExists(ss, branch.tabName);

  const copy = archiveSheet.copyTo(ss).setName(branch.tabName);
  _ensureVisible(copy);
  ss.setActiveSheet(copy);

  return { ok: true };
}

/* =============================================================================
   HELPERS – DATA IO + CARRYOVER + ARCHIVE + STYLING
============================================================================= */

/** Read a sheet as array of objects using the first row as headers. */
function _getSheetDataAsObjects(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length === 0) return [];
  const header = values[0].map(h => String(h || '').trim());
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = {};
    for (let c = 0; c < header.length; c++) {
      row[header[c]] = values[r][c];
    }
    // skip fully empty rows
    if (Object.values(row).some(v => v !== '' && v !== null)) out.push(row);
  }
  return out;
}

/** Build a carryover map of { 'Stock #' : {Due, Notes} } from an existing branch tab. */
function _readCarryoverMap(ss, tabName) {
  const map = new Map();
  const sh = ss.getSheetByName(tabName);
  if (!sh) return map;

  const vals = _getSheetDataAsObjects(sh);
  vals.forEach(o => {
    const key = String(o['Stock #'] || '').trim();
    if (!key) return;
    map.set(key, {
      Due:   o['Due'] ?? '',
      Notes: o['Notes'] ?? ''
    });
  });
  return map;
}

/** Archive a sheet if it exists; hide the archive. */
function _archiveSheetIfExists(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  const archName = `_Archive_${_branchKeyFromTab(sheetName)}_${_stampForName()}`;
  const arch = sh.copyTo(ss).setName(archName);
  arch.hideSheet();
  ss.deleteSheet(sh);
}

/** Derive branch key from tab name using CONFIG.BRANCHES mapping. */
function _branchKeyFromTab(tabName) {
  const found = CONFIG.BRANCHES.find(b => b.tabName === tabName);
  return found ? found.key : tabName;
}

/** Clear all data rows below the header row (retain formatting). */
function _clearDataBelowHeader(sheet, headerRow) {
  const last = sheet.getMaxRows();
  const lastCol = sheet.getMaxColumns();
  if (last > headerRow) {
    sheet.getRange(headerRow + 1, 1, last - headerRow, lastCol).clearContent();
  }
}

/** Write the header labels explicitly (does not change formatting). */
function _writeHeaderLabels(sheet, headerRow, headers) {
  sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]);
}

/** Make sure a sheet is visible. */
function _ensureVisible(sheet) {
  try { sheet.showSheet(); } catch (e) {}
}

/** Timestamp helpers. */
function _stampForName() {
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  const hh = String(now.getHours()).padStart(2, '0');
  const mi = String(now.getMinutes()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}_${hh}${mi}`;
}
function _stampForFilename() {
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  const hh = String(now.getHours()).padStart(2, '0');
  const mi = String(now.getMinutes()).padStart(2, '0');
  const ss = String(now.getSeconds()).padStart(2, '0');
  return `${yyyy}${mm}${dd}-${hh}${mi}${ss}`;
}

/** Progress logger (surface text back to dialog). */
function _logProgress(msg) {
  // This is a hook; the HTML dialog polls for latest message via getLastProgress().
  CacheService.getScriptCache().put('pm_progress', msg, 120);
}
function getLastProgress() {
  return CacheService.getScriptCache().get('pm_progress') || '';
}
