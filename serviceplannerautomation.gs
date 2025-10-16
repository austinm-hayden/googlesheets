/***************************************************************************************
 * PM SERVICE PLANNER – RAW UPLOAD → BRANCH SPLITTER + FORMAT CLONER + ARCHIVER
 * Version: 1.4.0  (2025-10-16)
 * Author Credit:  Created by Austin Monson (with ChatGPT assistance)
 * License: MIT
 *
 * -------------------------------------------------------------------------------------
 * CHANGELOG
 * 1.0.0 – Base release (upload → split → archive)
 * 1.1.0 – Reliable doPost() upload handler
 * 1.2.0 – In-sheet upload dialog; case-insensitive HTML (Restore)
 * 1.2.1 – Case-insensitive HTML (Upload) templating fallback
 * 1.3.0 – Information tab automation + GitHub sync (buildinfo.json)
 * 1.4.0 – PROJECT_INFO fully remote (buildinfo.json) + latest commit SHA via GitHub API
 *
 * -------------------------------------------------------------------------------------
 * LICENSE (MIT)
 * Copyright (c) 2025 Austin Monson
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to do so, subject to the following conditions:
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
 * ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 ***************************************************************************************/

/* =============================================================================
   CONFIGURATION (EDIT SAFELY)
============================================================================= */
const CONFIG = {
  TEMPLATE_SHEET: 'Pending Service - Springfield',       // source of formatting/validations/CF
  BRANCH_COLUMN:  'Branch',                              // column in raw file that names the branch
  BRANCHES: [
    { key: 'Springfield', tabName: 'Pending Service - Springfield' },
    { key: 'West Plains', tabName: 'Pending Service - West Plains' },
    { key: 'Villa Ridge', tabName: 'Pending Service - Villa Ridge' },
  ],
  HEADER_ORDER: [
    'Stock #','Description','Type','Overdue',
    'Serial Number','Manufacturer','Model','Due','Notes'
  ],
  DUE_FILTERS: ['Corrected','Service not Needed','Removed'], // rows with these 'Due' are excluded
  HEADER_ROW: 1,

  // GitHub repo URL (used for buildinfo.json + latest commit)
  REPO_URL: 'https://github.com/austinm-hayden/googlesheets.git'
};

/* =============================================================================
   MENU
============================================================================= */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PM Planner')
    .addItem('Upload Raw Data (XLSX → Split + Style + Archive)','showUploadDialog')
    .addItem('Rebuild Information Tab','buildInformationTab')
    .addSeparator()
    .addItem('Restore Archived Sheet…','showRestoreDialog')
    .addToUi();
}

/* =============================================================================
   DIALOG LAUNCHERS (case-insensitive loaders)
============================================================================= */
function _getHtmlFileCaseInsensitive(base) {
  const variants = [base, base.toLowerCase(), base.toUpperCase(),
    base.charAt(0).toUpperCase() + base.slice(1).toLowerCase(),
    base + '.html', base.toLowerCase() + '.html'];
  for (const n of variants) {
    try { return HtmlService.createHtmlOutputFromFile(n); } catch (e) {}
  }
  throw new Error('HTML file not found for ' + base + ' (checked: ' + variants.join(', ') + ')');
}

// Upload dialog must be templated (to inject ScriptApp URL for doPost action)
function showUploadDialog() {
  const candidates = ['Dialog', 'dialog', 'DIALOG', 'Dialog.html', 'dialog.html'];
  let templateFile = null;
  for (const name of candidates) {
    try { HtmlService.createTemplateFromFile(name); templateFile = name; break; } catch(e){}
  }
  if (!templateFile) throw new Error('Could not locate Dialog HTML file (checked: ' + candidates.join(', ') + ')');

  const t = HtmlService.createTemplateFromFile(templateFile);
  t.webAppUrl = ScriptApp.getService().getUrl();  // evaluated server-side so form gets a valid doPost URL
  const html = t.evaluate().setWidth(520).setHeight(440);
  SpreadsheetApp.getUi().showModalDialog(html,'Upload Raw PM Data');
}

function showRestoreDialog() {
  const html = _getHtmlFileCaseInsensitive('Restore').setWidth(540).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html,'Restore Archived Branch Sheet');
}

/* =============================================================================
   doPost() – Receives XLSX blob from form (Dialog.html)
   REQUIREMENT: Enable Advanced Google Services → Drive API (Drive.Files.*)
============================================================================= */
function doPost(e) {
  try {
    const blob = e?.files?.file;
    if (!blob) throw new Error('No file uploaded.');

    const temp = DriveApp.createFile(blob);
    temp.setName('PM_Raw_Upload_' + _stampForFilename() + '.xlsx');
    _logProgress('Uploaded file → ' + temp.getName());

    const res = _processUploadedRawFile(temp);

    // Refresh the Information tab after successful processing
    buildInformationTab(true);

    return HtmlService.createHtmlOutput('Upload completed: ' + JSON.stringify(res));
  } catch (err) {
    return HtmlService.createHtmlOutput('Error: ' + err.message);
  }
}

/* =============================================================================
   CORE PROCESSOR – convert, split by branch, archive, format, carry Due/Notes
============================================================================= */
function _processUploadedRawFile(tempFile) {
  const ss = SpreadsheetApp.getActive();
  const template = ss.getSheetByName(CONFIG.TEMPLATE_SHEET);
  if (!template) throw new Error('Template sheet "' + CONFIG.TEMPLATE_SHEET + '" not found.');

  _logProgress('Converting to Google Sheet…');
  const meta = { title: tempFile.getName(), mimeType: 'application/vnd.google-apps.spreadsheet' };
  const converted = Drive.Files.copy(meta, tempFile.getId());  // Advanced Drive API
  const raw = SpreadsheetApp.openById(converted.id);
  const rawSheet = raw.getSheets()[0];
  const rows = _getSheetDataAsObjects(rawSheet);

  // Cleanup the converted intermediate & temp
  Drive.Files.remove(converted.id);
  tempFile.setTrashed(true);

  if (!rows.length) throw new Error('Raw file is empty.');
  if (!(CONFIG.BRANCH_COLUMN in rows[0])) throw new Error('Missing branch column "' + CONFIG.BRANCH_COLUMN + '"');

  _logProgress('Splitting rows by branch…');
  const groups = {}; CONFIG.BRANCHES.forEach(b => groups[b.key] = []);
  rows.forEach(r => {
    const key = (r[CONFIG.BRANCH_COLUMN] || '').toString().trim();
    if (groups[key]) groups[key].push(r);
  });

  // For each branch: archive old tab, copy template, write filtered rows, preserve Due/Notes
  CONFIG.BRANCHES.forEach(branch => {
    _logProgress('Processing ' + branch.key + '…');

    // Carryover: Stock # → {Due, Notes}
    const carry = _readCarryoverMap(ss, branch.tabName);

    _archiveSheetIfExists(ss, branch.tabName);

    // Duplicate template to preserve all styling, then clear body rows
    const sheet = template.copyTo(ss).setName(branch.tabName);
    sheet.showSheet();
    _clearDataBelowHeader(sheet, CONFIG.HEADER_ROW);

    const staged = (groups[branch.key] || []).map(row => {
      const stock = (row['Stock #'] || '').toString().trim();
      const prev  = carry.get(stock) || {};
      const due   = prev.Due   ?? row['Due']   ?? '';
      const notes = prev.Notes ?? row['Notes'] ?? '';

      return CONFIG.HEADER_ORDER.map(h => {
        if (h === 'Due')   return due;
        if (h === 'Notes') return notes;
        return row[h] ?? '';
      });
    }).filter(arr => {
      const d = (arr[CONFIG.HEADER_ORDER.indexOf('Due')] || '').toString().trim();
      return !CONFIG.DUE_FILTERS.includes(d);
    });

    if (staged.length) {
      sheet.getRange(CONFIG.HEADER_ROW + 1, 1, staged.length, CONFIG.HEADER_ORDER.length).setValues(staged);
    }
    _writeHeaderLabels(sheet, CONFIG.HEADER_ROW, CONFIG.HEADER_ORDER);

    _logProgress(branch.key + ' done (' + staged.length + ' rows).');
  });

  _logProgress('All branches complete.');
  return { ok: true };
}

/* =============================================================================
   INFORMATION TAB – loads buildinfo.json + latest commit SHA via GitHub API
============================================================================= */

/**
 * Public entry point (menu & auto after upload).
 * If autoTriggered=true, skips UI alerts.
 */
function buildInformationTab(autoTriggered = false) {
  const ss = SpreadsheetApp.getActive();
  const name = 'Information';
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name); else sheet.clear();

  // Pull remote build info; fall back to local default if needed
  const info = _getProjectInfoRemote(CONFIG.REPO_URL);
  info.lastBuild = new Date().toLocaleString();

  // Also fetch latest commit SHA & timestamp from GitHub API (no token needed for public repo)
  const commit = _getLatestCommitInfo(CONFIG.REPO_URL);  // { sha, url, date }
  // ----------------------------------------------------------------------------------
  // LAYOUT
  let r = 1;
  const put = (text, style={}) => {
    const rng = sheet.getRange(r, 1);
    rng.setValue(text);
    if (style.bold) rng.setFontWeight('bold');
    if (style.size) rng.setFontSize(style.size);
    if (style.bg)   rng.setBackground(style.bg);
    r++;
  };

  put(info.title || 'PM Service Planner – Branch Automation Suite', { bold:true, size:16, bg:'#dbeafe' });
  put(`Author: ${info.author || 'Austin Monson'}`);
  put(`License: ${info.license || 'MIT'}`);
  put(`Version: ${info.version || 'n/a'} (${info.releaseDate || ''})`);
  put(`Repository: ${info.repo || CONFIG.REPO_URL}`);
  if (commit.sha) put(`Source Commit: ${commit.sha.substring(0,7)} (${commit.date || ''})`);
  put(`Last Build: ${info.lastBuild}`);
  r++;

  put('Overview:', { bold:true });
  (info.overview || []).forEach(line => put(line));
  r++;

  put('Working Features:', { bold:true });
  (info.working || []).forEach(line => put(line));
  r++;

  put('Planned / Pending Features:', { bold:true });
  (info.pending || []).forEach(line => put(line));
  r++;

  put('Known Limitations:', { bold:true });
  (info.limitations || []).forEach(line => put(line));
  r++;

  put('Version History:', { bold:true });
  (info.changelog || []).forEach(line => put(line));
  r++;

  // Formatting
  const last = r;
  sheet.getRange(1,1,last,1).setWrap(true);
  sheet.setColumnWidth(1, 760);
  sheet.getRange('A1:A7').setBackground('#dbeafe');
  sheet.getRange('A1').setFontSize(18);
  sheet.setFrozenRows(1);

  // Hyperlinks
  // Repo link on its own line
  sheet.getRange(5,1).setValue('GitHub Repository')
    .setFontColor('blue').setFontLine('underline')
    .setFormula(`=HYPERLINK("${info.repo || CONFIG.REPO_URL}", "GitHub Repository")`);

  // Commit link (if available)
  if (commit.url) {
    // place it at the “Source Commit” line (row 6 by current layout)
    sheet.getRange(6,1).setValue(`Source Commit: ${commit.sha.substring(0,7)}`)
      .setFontColor('blue').setFontLine('underline')
      .setFormula(`=HYPERLINK("${commit.url}", "Source Commit: ${commit.sha.substring(0,7)}")`);
  }

  _logProgress('Information tab updated.');
  if (!autoTriggered) SpreadsheetApp.getUi().alert('Information tab rebuilt.');
}

/**
 * Fetch buildinfo.json from GitHub repo root:
 *   https://raw.githubusercontent.com/{owner}/{repo}/main/buildinfo.json
 * Falls back to a minimal default if unavailable.
 */
function _getProjectInfoRemote(repoUrl) {
  const fallback = {
    title: "PM Service Planner – Branch Automation Suite",
    author: "Austin Monson",
    license: "MIT",
    version: "local-dev",
    releaseDate: new Date().toISOString().split('T')[0],
    repo: repoUrl,
    overview: ["(Offline mode) Could not load buildinfo.json from GitHub."],
    changelog: [], working: [], pending: [], limitations: []
  };
  try {
    const { owner, repo } = _parseGitHub(repoUrl);
    const rawUrl = `https://raw.githubusercontent.com/${owner}/${repo}/main/buildinfo.json`;
    const resp = UrlFetchApp.fetch(rawUrl, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return fallback;
    const info = JSON.parse(resp.getContentText());
    // Normalize expected fields
    info.repo = info.repo || repoUrl;
    return info;
  } catch (e) {
    Logger.log('Build info fetch error: ' + e);
    return fallback;
  }
}

/**
 * Fetch latest commit for main branch:
 *   https://api.github.com/repos/{owner}/{repo}/commits/main
 * Returns { sha, url, date } or empty object if not available.
 */
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

/** Parse owner/repo from a standard GitHub URL (with or without .git). */
function _parseGitHub(repoUrl) {
  // e.g., https://github.com/austinm-hayden/googlesheets.git
  const m = repoUrl.replace(/\.git$/,'').match(/github\.com\/([^\/]+)\/([^\/]+)$/i);
  if (!m) throw new Error('Cannot parse GitHub owner/repo from: ' + repoUrl);
  return { owner: m[1], repo: m[2] };
}

/* =============================================================================
   ARCHIVE / RESTORE
============================================================================= */
function listArchives() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheets().filter(s => s.getName().startsWith('_Archive_'))
    .map(s => ({ name: s.getName(), hidden: s.isSheetHidden() }));
}

function restoreArchive(name) {
  if (!name) throw new Error('No archive selected.');
  const ss = SpreadsheetApp.getActive();
  const arch = ss.getSheetByName(name);
  if (!arch) throw new Error('Archive not found: ' + name);

  // Infer branch key from archive name: _Archive_[Branch]_YYYY-MM-DD_hhmm
  const branchKey = name.replace(/^_Archive_/, '').replace(/_\d{4}-\d{2}-\d{2}_\d{4}$/, '');
  const branch = CONFIG.BRANCHES.find(b => b.key === branchKey);
  if (!branch) throw new Error('Branch not recognized in archive name: ' + name);

  _archiveSheetIfExists(ss, branch.tabName);
  const copy = arch.copyTo(ss).setName(branch.tabName);
  copy.showSheet(); ss.setActiveSheet(copy);
  return { ok: true };
}

/* =============================================================================
   HELPERS – data IO, archive ops, stamps, logging
============================================================================= */
function _getSheetDataAsObjects(sheet) {
  const v = sheet.getDataRange().getValues(); if (!v.length) return [];
  const h = v[0].map(x => String(x || '').trim());
  const out = [];
  for (let i=1;i<v.length;i++){
    const o={}; for (let j=0;j<h.length;j++) o[h[j]]=v[i][j];
    if (Object.values(o).some(x => x !== '' && x != null)) out.push(o);
  }
  return out;
}

function _readCarryoverMap(ss, tabName) {
  const map = new Map(); const sh = ss.getSheetByName(tabName); if (!sh) return map;
  _getSheetDataAsObjects(sh).forEach(o => {
    const k = (o['Stock #'] || '').toString().trim(); if (!k) return;
    map.set(k, { Due: o['Due'] ?? '', Notes: o['Notes'] ?? '' });
  });
  return map;
}

function _archiveSheetIfExists(ss, name) {
  const sh = ss.getSheetByName(name); if (!sh) return;
  const key = CONFIG.BRANCHES.find(b => b.tabName === name)?.key || name;
  const archName = `_Archive_${key}_${_stampForName()}`;
  const arch = sh.copyTo(ss).setName(archName);
  arch.hideSheet();
  ss.deleteSheet(sh);
}

function _clearDataBelowHeader(sh, headerRow) {
  const maxR = sh.getMaxRows(), maxC = sh.getMaxColumns();
  if (maxR > headerRow) sh.getRange(headerRow + 1, 1, maxR - headerRow, maxC).clearContent();
}

function _writeHeaderLabels(sh, row, headers) {
  sh.getRange(row, 1, 1, headers.length).setValues([headers]);
}

function _stampForName() {
  const d=new Date(),y=d.getFullYear(),
        m=String(d.getMonth()+1).padStart(2,'0'),
        da=String(d.getDate()).padStart(2,'0'),
        h=String(d.getHours()).padStart(2,'0'),
        mi=String(d.getMinutes()).padStart(2,'0');
  return `${y}-${m}-${da}_${h}${mi}`;
}

function _stampForFilename() {
  const d=new Date();
  return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}_${String(d.getHours()).padStart(2,'0')}${String(d.getMinutes()).padStart(2,'0')}${String(d.getSeconds()).padStart(2,'0')}`;
}

function _logProgress(msg) { CacheService.getScriptCache().put('pm_progress', msg, 120); }
function getLastProgress() { return CacheService.getScriptCache().get('pm_progress') || ''; }
