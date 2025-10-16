/***************************************************************************************
 * PM SERVICE PLANNER – RAW UPLOAD → BRANCH SPLITTER + FORMAT CLONER + ARCHIVER
 * Version: 1.5.0  (2025-10-16)
 * Author Credit:  Created by Austin Monson (with ChatGPT assistance)
 * License: MIT
 *
 * Repository: https://github.com/austinm-hayden/googlesheets
 *
 * Summary:
 * - Accept an uploaded .xlsx file via in-sheet Dialog (form → doPost).
 * - Convert to Google Sheet and parse raw INCUS5 data.
 * - Split rows into branch tabs (Springfield, West Plains, Villa Ridge).
 * - Clone ALL formatting from a template (Springfield) to each branch.
 * - Carry forward "Due" + "Notes" by Stock # across updates.
 * - Exclude rows marked in Due: Corrected / Service not Needed / Removed.
 * - Archive prior branch tabs as hidden, timestamped sheets; restore via UI.
 * - Build an Information tab from GitHub buildinfo.json + show latest commit SHA.
 ***************************************************************************************/


/**
 * =============================================================================
 * SECTION: GLOBAL CONFIGURATION (EDIT SAFELY)
 * -----------------------------------------------------------------------------
 * - TEMPLATE_SHEET: the master styling source (formatting, validation, CF)
 * - BRANCH_COLUMN : column name in raw data that indicates branch key
 * - BRANCHES      : mapping of branch keys to destination tab names
 * - HEADER_ORDER  : destination column order written under the header row
 * - DUE_FILTERS   : any row with these "Due" statuses is not included
 * - REPO_URL      : GitHub repository hosting buildinfo.json (main branch)
 * =============================================================================
 */
const CONFIG = {
  TEMPLATE_SHEET: 'Pending Service - Springfield',
  BRANCH_COLUMN: 'Branch',
  BRANCHES: [
    { key: 'Springfield', tabName: 'Pending Service - Springfield' },
    { key: 'West Plains', tabName: 'Pending Service - West Plains' },
    { key: 'Villa Ridge', tabName: 'Pending Service - Villa Ridge' },
  ],
  HEADER_ORDER: [
    'Stock #', 'Description', 'Type', 'Overdue',
    'Serial Number', 'Manufacturer', 'Model', 'Due', 'Notes'
  ],
  DUE_FILTERS: ['Corrected', 'Service not Needed', 'Removed'],
  HEADER_ROW: 1,
  REPO_URL: 'https://github.com/austinm-hayden/googlesheets.git'
};


/**
 * =============================================================================
 * SECTION: MENU INITIALIZATION
 * -----------------------------------------------------------------------------
 * Adds the "PM Planner" menu on sheet open. Wrapped in try/catch to avoid
 * errors if `onOpen` is executed outside an interactive UI context.
 * =============================================================================
 */
function onOpen(e) {
  try {
    SpreadsheetApp.getUi()
      .createMenu('PM Planner')
      .addItem('Upload Raw Data (XLSX → Split + Style + Archive)', 'showUploadDialog')
      .addItem('Rebuild Information Tab', 'buildInformationTab')
      .addSeparator()
      .addItem('Restore Archived Sheet…', 'showRestoreDialog')
      .addToUi();
  } catch (err) {
    Logger.log('onOpen skipped (no UI context): ' + err.message);
  }
}


/**
 * =============================================================================
 * SECTION: DIALOG LAUNCHERS (CASE-INSENSITIVE HTML LOOKUPS)
 * -----------------------------------------------------------------------------
 * - showUploadDialog(): injects ScriptApp URL into Dialog template (form action)
 * - showRestoreDialog(): loads Restore dialog (case-insensitive filename)
 * =============================================================================
 */

/** Internal: case-insensitive Html file load (non-templated). */
function _getHtmlFileCaseInsensitive(base) {
  const variants = [
    base, base.toLowerCase(), base.toUpperCase(),
    base.charAt(0).toUpperCase() + base.slice(1).toLowerCase(),
    base + '.html', base.toLowerCase() + '.html'
  ];
  for (const n of variants) {
    try { return HtmlService.createHtmlOutputFromFile(n); } catch (_) {}
  }
  throw new Error('HTML file not found for ' + base + ' (checked: ' + variants.join(', ') + ')');
}

/** Opens the Upload Dialog; uses template so we can inject a valid doPost URL. */
function showUploadDialog() {
  const candidates = ['Dialog', 'dialog', 'DIALOG', 'Dialog.html', 'dialog.html'];
  let templateFile = null;
  for (const name of candidates) {
    try { HtmlService.createTemplateFromFile(name); templateFile = name; break; } catch (_) {}
  }
  if (!templateFile) throw new Error('Could not locate Dialog HTML (tried: ' + candidates.join(', ') + ')');

  const t = HtmlService.createTemplateFromFile(templateFile);
  // This resolves to a valid execution URL in the container-bound project:
  t.webAppUrl = ScriptApp.getService().getUrl();
  const html = t.evaluate().setWidth(520).setHeight(440);
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload Raw PM Data');
}

/** Opens the Restore dialog (case-insensitive filename). */
function showRestoreDialog() {
  const html = _getHtmlFileCaseInsensitive('Restore').setWidth(540).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Restore Archived Branch Sheet');
}


/**
 * =============================================================================
 * SECTION: FILE UPLOAD HANDLER (doPost)
 * -----------------------------------------------------------------------------
 * Receives a multipart/form POST from Dialog.html.
 * - Saves XLSX to Drive temporarily
 * - Triggers processing (convert → split → archive → write)
 * - Rebuilds the Information tab after success
 * NOTE: Requires Advanced Google Service "Drive API" enabled.
 * =============================================================================
 */
function doPost(e) {
  try {
    const blob = e?.files?.file;
    if (!blob) throw new Error('No file uploaded.');

    const temp = DriveApp.createFile(blob);
    temp.setName('PM_Raw_Upload_' + _stampForFilename() + '.xlsx');
    _logProgress('Uploaded file → ' + temp.getName());

    const result = _processUploadedRawFile(temp);

    // Refresh the Information tab using GitHub manifest + latest commit info
    buildInformationTab(true);

    return HtmlService.createHtmlOutput('Upload completed: ' + JSON.stringify(result));
  } catch (err) {
    return HtmlService.createHtmlOutput('Error: ' + err.message);
  }
}


/**
 * =============================================================================
 * SECTION: CORE PROCESSOR – CONVERT, SPLIT, ARCHIVE, CARRYOVER, FORMAT
 * -----------------------------------------------------------------------------
 * Converts the temporary XLSX into a Google Sheet (Drive.Files.copy),
 * parses the first sheet’s data, groups rows by branch, archives old tabs,
 * clones formatting from the template, carries forward "Due" + "Notes",
 * filters excluded "Due" statuses, and writes the refreshed data.
 * =============================================================================
 */
function _processUploadedRawFile(tempFile) {
  const ss = SpreadsheetApp.getActive();

  // Validate template styling source
  const template = ss.getSheetByName(CONFIG.TEMPLATE_SHEET);
  if (!template) throw new Error('Template sheet "' + CONFIG.TEMPLATE_SHEET + '" not found.');

  // Convert XLSX → Google Sheet
  _logProgress('Converting to Google Sheet…');
  const meta = { title: tempFile.getName(), mimeType: 'application/vnd.google-apps.spreadsheet' };
  const converted = Drive.Files.copy(meta, tempFile.getId()); // Advanced Drive Service
  const rawSS = SpreadsheetApp.openById(converted.id);
  const rawSheet = rawSS.getSheets()[0];

  // Parse rows as objects
  const rows = _getSheetDataAsObjects(rawSheet);

  // Cleanup temp intermediates
  Drive.Files.remove(converted.id);
  tempFile.setTrashed(true);

  if (!rows.length) throw new Error('Raw file is empty.');
  if (!(CONFIG.BRANCH_COLUMN in rows[0])) {
    throw new Error('Missing branch column "' + CONFIG.BRANCH_COLUMN + '" in uploaded file.');
  }

  // Group rows by branch
  _logProgress('Splitting rows by branch…');
  const grouped = {};
  CONFIG.BRANCHES.forEach(b => grouped[b.key] = []);
  rows.forEach(r => {
    const key = (r[CONFIG.BRANCH_COLUMN] || '').toString().trim();
    if (grouped[key]) grouped[key].push(r);
  });

  // For each branch: archive → duplicate template → clear body → write
  CONFIG.BRANCHES.forEach(branch => {
    _logProgress('Processing ' + branch.key + '…');

    // 1) Build carryover map from prior working tab (Stock # → {Due, Notes})
    const carry = _readCarryoverMap(ss, branch.tabName);

    // 2) Archive existing working tab (hidden, timestamped)
    _archiveSheetIfExists(ss, branch.tabName);

    // 3) Duplicate template to preserve all formatting
    const sh = template.copyTo(ss).setName(branch.tabName);
    sh.showSheet();
    _clearDataBelowHeader(sh, CONFIG.HEADER_ROW);

    // 4) Stage output rows in destination header order
    const staged = (grouped[branch.key] || []).map(row => {
      const stock = (row['Stock #'] || '').toString().trim();
      const prev = carry.get(stock) || {};
      const due = prev.Due ?? row['Due'] ?? '';
      const notes = prev.Notes ?? row['Notes'] ?? '';

      return CONFIG.HEADER_ORDER.map(h => {
        if (h === 'Due')   return due;
        if (h === 'Notes') return notes;
        return row[h] ?? '';
      });
    })
    // 5) Filter by Due blacklist
    .filter(arr => {
      const d = (arr[CONFIG.HEADER_ORDER.indexOf('Due')] || '').toString().trim();
      return !CONFIG.DUE_FILTERS.includes(d);
    });

    // 6) Write data under header if present
    if (staged.length) {
      sh.getRange(CONFIG.HEADER_ROW + 1, 1, staged.length, CONFIG.HEADER_ORDER.length).setValues(staged);
    }

    // 7) Ensure header labels are exactly as expected (format preserved)
    _writeHeaderLabels(sh, CONFIG.HEADER_ROW, CONFIG.HEADER_ORDER);

    _logProgress(branch.key + ' complete (' + staged.length + ' rows).');
  });

  _logProgress('All branches complete.');
  return { ok: true };
}


/**
 * =============================================================================
 * SECTION: INFORMATION TAB BUILDER (VERBOSE MANIFEST RENDERER)
 * -----------------------------------------------------------------------------
 * Reads `buildinfo.json` from GitHub `main` branch and formats sections:
 * Overview, Architecture (Core Components + Interaction Flow), Features,
 * Planned, Limitations, Version History (rich), Notes, and Metadata.
 * Also fetches the latest commit SHA/date (public GitHub API).
 * - autoTriggered=true → silent mode (no UI alert; safe in background).
 * =============================================================================
 */
function buildInformationTab(autoTriggered = false) {
  const ss = SpreadsheetApp.getActive();
  const sheetName = 'Information';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName); else sheet.clear();

  const info = _getProjectInfoRemote(CONFIG.REPO_URL);
  const commit = _getLatestCommitInfo(CONFIG.REPO_URL);
  const buildStamp = new Date().toLocaleString();

  let row = 1;
  const put = (text, style = {}) => {
    const r = sheet.getRange(row, 1);
    r.setValue(text);
    if (style.bold) r.setFontWeight('bold');
    if (style.size) r.setFontSize(style.size);
    if (style.bg)   r.setBackground(style.bg);
    if (style.color) r.setFontColor(style.color);
    row++;
  };
  const spacer = (n = 1) => { row += n; };

  // — Header / metadata
  put(info.title || 'PM Service Planner – Branch Automation Suite', { bold: true, size: 18, bg: '#dbeafe' });
  put(`Author: ${info.author || 'Austin Monson'}`);
  put(`License: ${info.license || 'MIT'}`);
  put(`Version: ${info.version || 'n/a'} (${info.releaseDate || ''})`);
  spacer();

  const repo = info.repo || CONFIG.REPO_URL;
  sheet.getRange(row, 1)
    .setValue('GitHub Repository')
    .setFontColor('blue').setFontLine('underline')
    .setFormula(`=HYPERLINK("${repo}", "GitHub Repository")`);
  spacer();

  if (commit && commit.sha) {
    const label = `Source Commit: ${commit.sha.substring(0, 7)} (${commit.date || ''})`;
    sheet.getRange(row, 1)
      .setValue(label)
      .setFontColor('blue').setFontLine('underline')
      .setFormula(`=HYPERLINK("${commit.url}", "${label}")`);
    spacer();
  }

  put(`Last Build: ${buildStamp}`);
  spacer(2);

  // — Overview
  if (Array.isArray(info.overview) && info.overview.length) {
    put('Overview:', { bold: true, size: 14, bg: '#e8f0fe' });
    info.overview.forEach(line => put('• ' + line));
    spacer();
  }

  // — Architecture
  if (info.architecture) {
    put('Architecture:', { bold: true, size: 14, bg: '#e8f0fe' });
    if (info.architecture.coreComponents) {
      put('Core Components:', { bold: true });
      Object.entries(info.architecture.coreComponents).forEach(([k, v]) => put(`• ${k}: ${v}`));
      spacer();
    }
    if (Array.isArray(info.architecture.interactionFlow)) {
      put('Interaction Flow:', { bold: true });
      info.architecture.interactionFlow.forEach(step => put('→ ' + step));
      spacer();
    }
  }

  // — Features (Working)
  if (Array.isArray(info.features) && info.features.length) {
    put('Working Features:', { bold: true, size: 14, bg: '#e8f0fe' });
    info.features.forEach(line => put(line));
    spacer();
  }

  // — Planned
  if (Array.isArray(info.planned) && info.planned.length) {
    put('Planned / Pending:', { bold: true, size: 14, bg: '#e8f0fe' });
    info.planned.forEach(line => put(line));
    spacer();
  }

  // — Limitations
  if (Array.isArray(info.limitations) && info.limitations.length) {
    put('Known Limitations:', { bold: true, size: 14, bg: '#e8f0fe' });
    info.limitations.forEach(line => put(line));
    spacer();
  }

  // — Version History (rich)
  if (Array.isArray(info.versionHistory) && info.versionHistory.length) {
    put('Version History:', { bold: true, size: 14, bg: '#e8f0fe' });
    info.versionHistory.forEach(v => {
      put(`${v.version || 'v?'} – ${v.date || ''}: ${v.summary || ''}`, { bold: true });
      if (Array.isArray(v.changes)) v.changes.forEach(ch => put('   • ' + ch));
      spacer();
    });
  }

  // — Notes
  if (Array.isArray(info.notes) && info.notes.length) {
    put('Developer Notes:', { bold: true, size: 14, bg: '#e8f0fe' });
    info.notes.forEach(line => put('• ' + line));
    spacer();
  }

  // — Metadata (from manifest)
  if (info.metadata && typeof info.metadata === 'object') {
    put('Metadata:', { bold: true, size: 14, bg: '#e8f0fe' });
    Object.entries(info.metadata).forEach(([k, v]) => put(`${k}: ${v}`));
    spacer();
  }

  // Final formatting
  const last = row;
  sheet.getRange(1, 1, last, 1).setWrap(true);
  sheet.setColumnWidth(1, 920);
  sheet.getRange('A1:A10').setBackground('#dbeafe');
  sheet.setFrozenRows(1);

  _logProgress('Information tab rebuilt successfully.');
  if (!autoTriggered) {
    try { SpreadsheetApp.getUi().alert('Information tab rebuilt successfully.'); }
    catch (err) { Logger.log('No UI for alert: ' + err.message); }
  }
}


/**
 * =============================================================================
 * SECTION: GITHUB INTEGRATION (BUILD INFO + LATEST COMMIT)
 * -----------------------------------------------------------------------------
 * - _getProjectInfoRemote(repoUrl): fetches /main/buildinfo.json (raw)
 * - _getLatestCommitInfo(repoUrl):  fetches latest commit for "main" branch
 * - _parseGitHub(repoUrl):          extracts {owner, repo} from URL
 * =============================================================================
 */

/** Returns verbose manifest from buildinfo.json; falls back to defaults on error. */
function _getProjectInfoRemote(repoUrl) {
  const fallback = {
    title: "PM Service Planner – Branch Automation Suite",
    author: "Austin Monson",
    license: "MIT",
    version: "local-dev",
    releaseDate: new Date().toISOString().split('T')[0],
    repo: repoUrl,
    overview: ["(Offline mode) Could not load buildinfo.json from GitHub."],
    features: [], planned: [], limitations: [],
    versionHistory: [], notes: [], metadata: {}
  };
  try {
    const { owner, repo } = _parseGitHub(repoUrl);
    const rawUrl = `https://raw.githubusercontent.com/${owner}/${repo}/main/buildinfo.json`;
    const resp = UrlFetchApp.fetch(rawUrl, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return fallback;
    const info = JSON.parse(resp.getContentText());
    // Normalize and ensure repo link exists
    info.repo = info.repo || repoUrl;
    return info;
  } catch (e) {
    Logger.log('Build info fetch error: ' + e);
    return fallback;
  }
}

/** Returns { sha, url, date } for latest commit on main, or {} if unavailable. */
function _getLatestCommitInfo(repoUrl) {
  try {
    const { owner, repo } = _parseGitHub(repoUrl);
    const api = `https://api.github.com/repos/${owner}/${repo}/commits/main`;
    const resp = UrlFetchApp.fetch(api, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return {};
    const json = JSON.parse(resp.getContentText());
    return {
      sha: json.sha,
      url: json.html_url,
      date: (json.commit && json.commit.author && json.commit.author.date) ? json.commit.author.date : ''
    };
  } catch (e) {
    Logger.log('Commit fetch error: ' + e);
    return {};
  }
}

/** Extracts {owner, repo} from a GitHub URL (with or without .git suffix). */
function _parseGitHub(repoUrl) {
  const m = repoUrl.replace(/\.git$/,'').match(/github\.com\/([^\/]+)\/([^\/]+)$/i);
  if (!m) throw new Error('Cannot parse GitHub owner/repo from: ' + repoUrl);
  return { owner: m[1], repo: m[2] };
}


/**
 * =============================================================================
 * SECTION: ARCHIVE & RESTORE
 * -----------------------------------------------------------------------------
 * - listArchives(): returns all sheets whose names start with "_Archive_"
 * - restoreArchive(name): archives current live tab and restores selection
 * =============================================================================
 */
function listArchives() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheets()
    .filter(s => s.getName().startsWith('_Archive_'))
    .map(s => ({ name: s.getName(), hidden: s.isSheetHidden() }));
}

function restoreArchive(name) {
  if (!name) throw new Error('No archive selected.');
  const ss = SpreadsheetApp.getActive();
  const arch = ss.getSheetByName(name);
  if (!arch) throw new Error('Archive not found: ' + name);

  // Name pattern: _Archive_[BranchKey]_YYYY-MM-DD_hhmm
  const branchKey = name.replace(/^_Archive_/, '').replace(/_\d{4}-\d{2}-\d{2}_\d{4}$/, '');
  const branch = CONFIG.BRANCHES.find(b => b.key === branchKey);
  if (!branch) throw new Error('Branch not recognized in archive name: ' + name);

  // Safety: archive current live sheet before restoring
  _archiveSheetIfExists(ss, branch.tabName);

  const copy = arch.copyTo(ss).setName(branch.tabName);
  copy.showSheet();
  ss.setActiveSheet(copy);
  return { ok: true };
}


/**
 * =============================================================================
 * SECTION: UTILITY HELPERS (DATA, SHEETS, TIMESTAMPS, LOGGING)
 * -----------------------------------------------------------------------------
 * - _getSheetDataAsObjects(sheet)
 * - _readCarryoverMap(ss, tabName)
 * - _archiveSheetIfExists(ss, name)
 * - _clearDataBelowHeader(sheet, headerRow)
 * - _writeHeaderLabels(sheet, headerRow, headers)
 * - _stampForName(), _stampForFilename()
 * - _logProgress(msg), getLastProgress()
 * =============================================================================
 */

/** Reads a sheet into an array of row objects with row[header] = value. */
function _getSheetDataAsObjects(sheet) {
  const v = sheet.getDataRange().getValues();
  if (!v.length) return [];
  const h = v[0].map(x => String(x || '').trim());
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const o = {};
    for (let j = 0; j < h.length; j++) o[h[j]] = v[i][j];
    if (Object.values(o).some(x => x !== '' && x != null)) out.push(o); // skip blank rows
  }
  return out;
}

/** Builds a map: Stock # → { Due, Notes } from an existing working tab (if any). */
function _readCarryoverMap(ss, tabName) {
  const map = new Map();
  const sh = ss.getSheetByName(tabName);
  if (!sh) return map;
  _getSheetDataAsObjects(sh).forEach(o => {
    const k = (o['Stock #'] || '').toString().trim();
    if (!k) return;
    map.set(k, { Due: o['Due'] ?? '', Notes: o['Notes'] ?? '' });
  });
  return map;
}

/** Archives the given sheet (copy + timestamped name + hidden) and deletes original. */
function _archiveSheetIfExists(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) return;
  const key = CONFIG.BRANCHES.find(b => b.tabName === name)?.key || name;
  const archName = `_Archive_${key}_${_stampForName()}`;
  const arch = sh.copyTo(ss).setName(archName);
  arch.hideSheet();
  ss.deleteSheet(sh);
}

/** Clears all content below the header row; preserves formatting/validations. */
function _clearDataBelowHeader(sheet, headerRow) {
  const maxR = sheet.getMaxRows();
  const maxC = sheet.getMaxColumns();
  if (maxR > headerRow) {
    sheet.getRange(headerRow + 1, 1, maxR - headerRow, maxC).clearContent();
  }
}

/** Writes the header labels (one row) without altering existing formatting. */
function _writeHeaderLabels(sheet, headerRow, headers) {
  sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]);
}

/** Timestamp for archive sheet names: YYYY-MM-DD_hhmm */
function _stampForName() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const da = String(d.getDate()).padStart(2, '0');
  const h = String(d.getHours()).padStart(2, '0');
  const mi = String(d.getMinutes()).padStart(2, '0');
  return `${y}-${m}-${da}_${h}${mi}`;
}

/** Timestamp for temp filenames: YYYYMMDD_hhmmss */
function _stampForFilename() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}_${String(d.getHours()).padStart(2, '0')}${String(d.getMinutes()).padStart(2, '0')}${String(d.getSeconds()).padStart(2, '0')}`;
}

/** Pushes a short status string to CacheService for Dialog progress polling. */
function _logProgress(msg) {
  CacheService.getScriptCache().put('pm_progress', msg, 120);
}

/** Returns the last progress message (polled by Dialog.html). */
function getLastProgress() {
  return CacheService.getScriptCache().get('pm_progress') || '';
}
