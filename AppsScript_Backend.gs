/**
 * Advat Report Database - Apps Script backend.
 *
 * Replaces the previous implementation with a tolerant, header-driven parser
 * and built-in diagnostics so empty/misparsed tabs are visible from the client.
 *
 * Query params (all optional):
 *   sheet=<name>     return only that tab
 *   debug=1          include skip reasons and sample skipped rows; disables cache
 *   fresh=1          bypass cache for one call
 *   since=YYYY-MM-DD filter records whose report_release_date is >= that day
 */

const CACHE_TTL_SECONDS = 300;

// Tokens we recognize as "this row is a header". Normalized form (lowercase,
// spaces -> _, slashes -> _, other punctuation stripped).
const HEADER_TOKENS = new Set([
  'sample_name', 'samplename', 'sample',
  'gen_id', 'genid',
  'anderson_id', 'andersonid',
  'test',
  'sno', 's_no', 'no', 'no_',
  'category',
  'single', 'couple', 'trio',
  'single_couple_trio', 'singlecoupletrio', 'single_couple',
  'sample_type', 'sampletype',
  'remarks',
  'report_release_date', 'reportreleasedate', 'release_date', 'releasedate',
  'raw_data_location', 'rawdatalocation', 'raw_data', 'rawdata'
]);

// Matches "Run-46A", "SURFseq-Run-46A", "SURFseq Run 46A", "RUN 1", etc.
const RUN_REGEX = /(?:SURF(?:seq)?[\s-]*)?Run[\s\-#:]*([A-Za-z0-9_.-]+)/i;

function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const requestedSheet = params.sheet || '';
    const debug = params.debug === '1' || params.debug === 'true';
    const noCache = params.fresh === '1' || debug;
    const since = params.since ? parseDateInput(params.since) : null;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cacheKey = 'rdb:' + ss.getId() + ':' + (requestedSheet || 'all') + ':' + (params.since || '');
    const cache = CacheService.getScriptCache();

    if (!noCache) {
      const cached = cache.get(cacheKey);
      if (cached) return jsonOut(cached);
    }

    const result = {
      success: true,
      timestamp: new Date().toISOString(),
      sheets: {},
      diagnostics: {},
      data: []
    };

    ss.getSheets().forEach(function (sheet) {
      const sheetName = sheet.getName();
      if (requestedSheet && sheetName !== requestedSheet) return;

      const parsed = parseSheetData(sheet, sheetName, { debug: debug });
      let records = parsed.records;

      if (since) {
        records = records.filter(function (r) {
          const d = parseDateString(r.report_release_date);
          return d && d >= since;
        });
      }

      result.sheets[sheetName] = records.length;
      result.diagnostics[sheetName] = parsed.diagnostics;
      result.data = result.data.concat(records);
    });

    result.totalRecords = result.data.length;
    const body = JSON.stringify(result);

    if (!noCache && body.length < 95000) {
      try { cache.put(cacheKey, body, CACHE_TTL_SECONDS); } catch (err) { /* size cap, ignore */ }
    }

    return jsonOut(body);
  } catch (error) {
    return jsonOut(JSON.stringify({
      success: false,
      error: error.toString(),
      stack: error.stack || ''
    }));
  }
}

function jsonOut(body) {
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function parseSheetData(sheet, sheetName, opts) {
  opts = opts || {};
  const records = [];
  const diagnostics = {
    headerRowsFound: 0,
    headerRowIndices: [],
    runsDetected: [],
    rowsScanned: 0,
    rowsSkipped: 0,
    skipReasons: {},
    sampleSkippedRows: [],
    datedRows: 0,        // records that have a report_release_date
    linksFound: 0,       // records where a report link was detected
    sampleMissingLink: [] // dated rows with NO link (to inspect misses)
  };

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { records: records, diagnostics: diagnostics };

  const dataRange = sheet.getRange(1, 1, lastRow, lastCol);
  const data = dataRange.getValues();
  let richTextValues = null;
  try { richTextValues = dataRange.getRichTextValues(); } catch (e) { richTextValues = null; }
  let formulas = null;
  try { formulas = dataRange.getFormulas(); } catch (e) { formulas = null; }

  // Primary link source: the Advanced Sheets API exposes hyperlinks on
  // Date/number-typed cells, which getRichTextValues() cannot read. Falls back
  // to rich text / formulas if the advanced service isn't enabled.
  const linkGrid = getHyperlinkGrid(sheet.getParent().getId(), sheetName, lastRow, lastCol);
  diagnostics.advancedLinks = !!(linkGrid && linkGrid.length && !linkGrid._error);
  if (linkGrid && linkGrid._error) diagnostics.advancedLinksError = linkGrid._error;

  let currentRun = '';
  let headers = [];
  let headerSeen = false;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const nonEmpty = row.filter(function (c) { return String(c == null ? '' : c).trim() !== ''; });
    if (nonEmpty.length === 0) continue;
    diagnostics.rowsScanned++;

    // Run-marker detection: a row whose populated cells are basically just the run label.
    const joined = nonEmpty.map(String).join(' ').trim();
    const runMatch = joined.match(RUN_REGEX);
    if (runMatch && nonEmpty.length <= 3 && joined.length < 80) {
      currentRun = runMatch[0].replace(/\s+/g, ' ').trim();
      diagnostics.runsDetected.push(currentRun);
      continue;
    }

    // Header detection: count cells whose normalized form is a recognized token.
    const normalized = row.map(normalizeHeader);
    let hits = 0;
    for (let j = 0; j < normalized.length; j++) {
      if (normalized[j] && HEADER_TOKENS.has(normalized[j])) hits++;
    }
    if (hits >= 2) {
      headers = normalized.map(canonicalHeader);
      headerSeen = true;
      diagnostics.headerRowsFound++;
      diagnostics.headerRowIndices.push(i + 1);

      // Merged group headers (e.g. "SNV" / "CNV" spanning several columns)
      // leave their sub-columns blank in this row and put the real column
      // names ("Pathogenic", "Likely Pathogenic", "VUS", ...) in the row
      // directly below. Detect that continuation row, fold its labels into
      // `headers` prefixed by the group name from this row (so they don't
      // collide with an identically-named sub-column under a different
      // group), and skip it so it isn't parsed as a data row.
      const nextRow = data[i + 1];
      if (nextRow) {
        const nextNormalized = nextRow.map(normalizeHeader);
        let nextHits = 0;
        let subLabelCount = 0;
        for (let j = 0; j < nextNormalized.length; j++) {
          if (nextNormalized[j] && HEADER_TOKENS.has(nextNormalized[j])) nextHits++;
          if (nextNormalized[j]) subLabelCount++;
        }
        if (nextHits === 0 && subLabelCount >= 2) {
          let currentGroup = '';
          for (let j = 0; j < row.length; j++) {
            const rawGroupCell = String(row[j] == null ? '' : row[j]).trim();
            if (rawGroupCell) currentGroup = normalizeHeader(rawGroupCell);
            if (nextNormalized[j]) {
              headers[j] = currentGroup ? (currentGroup + '_' + nextNormalized[j]) : nextNormalized[j];
            }
          }
          diagnostics.headerRowsFound++;
          diagnostics.headerRowIndices.push(i + 2);
          i++; // consume the sub-header row, don't parse it as data
        }
      }
      continue;
    }

    if (!headerSeen) {
      diagnostics.rowsSkipped++;
      bumpReason(diagnostics, 'no_header_yet');
      sampleSkip(diagnostics, i, row);
      continue;
    }

    // Build record from current headers.
    const record = { _run: currentRun, _sheet: sheetName, _report_link: '' };

    for (let idx = 0; idx < headers.length && idx < row.length; idx++) {
      const header = headers[idx];
      if (!header) continue;

      let value = row[idx];
      if (value instanceof Date) {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd-MM-yyyy');
      }
      const text = String(value == null ? '' : value).trim();
      record[header] = text;

      // Report links live as a hyperlink on the report_release_date cell
      // (either a rich-text link or a =HYPERLINK() formula). Capture any
      // http(s) link there; fall back to Drive/Docs links found in any cell.
      const cellLink = getCellLink(linkGrid, richTextValues, formulas, i, idx, text);
      if (cellLink) {
        if (header === 'report_release_date') {
          record._report_link = cellLink;        // highest priority
        } else if (!record._report_link && isReportUrl(cellLink)) {
          record._report_link = cellLink;        // fallback
        }
      }
    }

    // Canonicalize alias header names so the client doesn't have to.
    aliasField(record, 'samplename', 'sample_name');
    aliasField(record, 'genid', 'gen_id');
    aliasField(record, 'andersonid', 'anderson_id');
    aliasField(record, 'sampletype', 'sample_type');
    aliasField(record, 'reportreleasedate', 'report_release_date');
    aliasField(record, 'releasedate', 'report_release_date');
    aliasField(record, 'rawdatalocation', 'raw_data_location');
    aliasField(record, 'rawdata', 'raw_data_location');
    aliasField(record, 'singlecoupletrio', 'single_couple_trio');
    aliasField(record, 'single_couple', 'single_couple_trio');
    aliasField(record, 'sno', 'no');
    aliasField(record, 's_no', 'no');
    aliasField(record, 'no_', 'no');

    const sampleName = record.sample_name || '';
    if (!sampleName) {
      diagnostics.rowsSkipped++;
      bumpReason(diagnostics, 'empty_sample_name');
      sampleSkip(diagnostics, i, row);
      continue;
    }

    // Derive a single human-readable classification from whichever snv_<label>
    // sub-column (from the merged "SNV" group header) is actually populated in
    // this sheet - only labels that exist as real columns can ever appear here.
    record.snv_status = deriveGroupStatus(record, 'snv_');

    // Link-coverage tracking: how many dated rows actually got a link, and a
    // few examples of dated rows that did NOT (so misses can be inspected).
    if (record.report_release_date) {
      diagnostics.datedRows++;
      if (record._report_link) {
        diagnostics.linksFound++;
      } else if (diagnostics.sampleMissingLink.length < 8) {
        const dateIdx = headers.indexOf('report_release_date');
        diagnostics.sampleMissingLink.push({
          rowIndex: i + 1,
          sample: record.sample_name,
          date: record.report_release_date,
          dateCellFormula: (formulas && formulas[i] && dateIdx > -1) ? String(formulas[i][dateIdx] || '') : '',
          dateCellValue: (dateIdx > -1 && row[dateIdx] != null) ? String(row[dateIdx]) : ''
        });
      }
    }

    records.push(record);
  }

  // Trim sample lists if not in debug mode (keep payload small).
  if (!opts.debug) {
    diagnostics.sampleSkippedRows = [];
    diagnostics.sampleMissingLink = [];
  }

  return { records: records, diagnostics: diagnostics };
}

function aliasField(record, from, to) {
  if (record[from] && !record[to]) record[to] = record[from];
}

function normalizeHeader(value) {
  return String(value == null ? '' : value).toLowerCase().trim()
    .replace(/\s+/g, '_')
    .replace(/\//g, '_')
    .replace(/[^a-z0-9_]/g, '');
}

// Map any spelling/punctuation variant of a header to a single canonical key,
// so 2023's "SL.NO" / "SINGLE / COUPLE/TRIO" line up with 2025's "S.NO" /
// "SINGLE/COUPLE/TRIO". Unknown headers keep their normalized name.
function canonicalHeader(norm) {
  if (!norm) return '';
  var h = norm;
  if (h === 'no' || h === 'no_' || h === 'sno' || h === 's_no' || h === 'slno' || h === 'sl_no' || h === 'sl' || h === 'serial' || h === 'serial_no') return 'no';
  if (h.indexOf('sample') > -1 && h.indexOf('name') > -1) return 'sample_name';
  if (h.indexOf('gen') > -1 && h.indexOf('id') > -1) return 'gen_id';
  if (h.indexOf('anderson') > -1) return 'anderson_id';
  if (h.indexOf('sample') > -1 && h.indexOf('type') > -1) return 'sample_type';
  if (h.indexOf('single') > -1 || h.indexOf('couple') > -1 || h.indexOf('trio') > -1) return 'single_couple_trio';
  if (h.indexOf('categ') > -1) return 'category';
  if (h.indexOf('remark') > -1) return 'remarks';
  if ((h.indexOf('report') > -1 && h.indexOf('date') > -1) || h.indexOf('release') > -1) return 'report_release_date';
  if (h.indexOf('raw') > -1 || h.indexOf('location') > -1) return 'raw_data_location';
  if (h === 'test' || h.indexOf('test') > -1) return 'test';
  return h;
}

// Columns under a merged group that are pure free-text annotation, not a
// classification, are excluded from the derived status label. "Additional"
// is kept - it's one of the real classification columns in this sheet.
const GROUP_STATUS_EXCLUDE_SUFFIXES = new Set(['remarks', 'remark', 'notes', 'note']);
const GROUP_STATUS_NEGATIVE_VALUES = new Set(['no', 'n', 'na', 'n_a', 'nil', 'none', '-']);

function deriveGroupStatus(record, prefix) {
  const labels = [];
  Object.keys(record).forEach(function (key) {
    if (key.indexOf(prefix) !== 0) return;
    const suffix = key.slice(prefix.length);
    if (GROUP_STATUS_EXCLUDE_SUFFIXES.has(suffix)) return;
    const val = String(record[key] || '').trim();
    if (!val || GROUP_STATUS_NEGATIVE_VALUES.has(val.toLowerCase())) return;
    labels.push(prettifyGroupStatusLabel(suffix));
  });
  return labels.join(', ');
}

function prettifyGroupStatusLabel(suffix) {
  return suffix.split('_').map(function (w) {
    if (w === 'vus') return 'VUS';
    if (w === 'het') return 'Het';
    return w.charAt(0).toUpperCase() + w.slice(1);
  }).join(' ');
}

function bumpReason(diag, reason) {
  diag.skipReasons[reason] = (diag.skipReasons[reason] || 0) + 1;
}

function sampleSkip(diag, idx, row) {
  if (diag.sampleSkippedRows.length < 5) {
    diag.sampleSkippedRows.push({
      rowIndex: idx + 1,
      preview: row.slice(0, 4).map(function (c) { return String(c == null ? '' : c); }).join(' | ')
    });
  }
}

// Builds a [row][col] grid of hyperlink URLs using the Advanced Sheets API.
// CellData.hyperlink exposes links on Date/number cells that the rich-text API
// silently drops. Returns [] (with ._error set) if the service is unavailable,
// so callers gracefully fall back to rich text.
function getHyperlinkGrid(ssId, sheetName, lastRow, lastCol) {
  const grid = [];
  if (typeof Sheets === 'undefined' || !Sheets.Spreadsheets) {
    grid._error = 'Advanced Sheets service not enabled';
    return grid;
  }
  try {
    const a1 = "'" + String(sheetName).replace(/'/g, "''") + "'!A1:" + colLetter(lastCol) + lastRow;
    const resp = Sheets.Spreadsheets.get(ssId, {
      ranges: [a1],
      includeGridData: true,
      fields: 'sheets(data(rowData(values(hyperlink,textFormatRuns(format(link(uri)))))))'
    });
    const sheetsArr = (resp && resp.sheets) || [];
    if (!sheetsArr.length) return grid;
    const dataArr = sheetsArr[0].data || [];
    if (!dataArr.length) return grid;
    const rowData = dataArr[0].rowData || [];
    for (let r = 0; r < rowData.length; r++) {
      const values = (rowData[r] && rowData[r].values) || [];
      const rowArr = [];
      for (let c = 0; c < values.length; c++) rowArr.push(extractCellDataLink(values[c]));
      grid.push(rowArr);
    }
  } catch (e) {
    grid.length = 0;
    grid._error = String(e);
  }
  return grid;
}

function extractCellDataLink(cell) {
  if (!cell) return '';
  if (cell.hyperlink) return cell.hyperlink;
  const runs = cell.textFormatRuns;
  if (runs) {
    for (let i = 0; i < runs.length; i++) {
      if (runs[i].format && runs[i].format.link && runs[i].format.link.uri) return runs[i].format.link.uri;
    }
  }
  return '';
}

function colLetter(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// Returns the URL linked from a cell, checking every place a link can hide:
//   0. Advanced Sheets API hyperlink grid (works on Date/number cells too)
//   1. whole-cell rich-text link  (Insert > Link on the entire cell)
//   2. any rich-text run link      (link on part of the cell text)
//   3. a =HYPERLINK("url", ...) formula
//   4. a plain URL inside the cell text
// Returns '' if none.
function getCellLink(linkGrid, richTextValues, formulas, i, idx, text) {
  if (linkGrid && linkGrid[i] && linkGrid[i][idx]) return linkGrid[i][idx];
  if (richTextValues && richTextValues[i] && richTextValues[i][idx]) {
    const rt = richTextValues[i][idx];
    try {
      const whole = rt.getLinkUrl();      // whole-cell link
      if (whole) return whole;
    } catch (e) { }
    const runs = rt.getRuns();
    for (let r = 0; r < runs.length; r++) {
      const url = runs[r].getLinkUrl();
      if (url) return url;
    }
  }
  if (formulas && formulas[i] && formulas[i][idx]) {
    const f = String(formulas[i][idx]);
    const m = f.match(/HYPERLINK\(\s*["']([^"']+)["']/i);
    if (m) return m[1];
  }
  const inline = extractAnyUrl(text);
  if (inline) return inline;
  return '';
}

function isReportUrl(url) {
  return /drive\.google\.com|docs\.google\.com|\.pdf(\?|$)|^https?:\/\//i.test(url);
}

function extractDriveLink(text) {
  return extractAnyUrl(text);
}

// Pull the first http(s) URL out of free text (Drive or any other host).
function extractAnyUrl(text) {
  if (!text) return '';
  const m = String(text).match(/https?:\/\/[^\s,)"']+/);
  return m ? m[0] : '';
}

function parseDateString(s) {
  if (!s) return null;
  const m = String(s).match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/);
  if (!m) return null;
  let y = m[3];
  if (y.length === 2) y = '20' + y;
  return new Date(parseInt(y, 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
}

function parseDateInput(s) {
  if (!s) return null;
  const iso = String(s).match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) return new Date(parseInt(iso[1], 10), parseInt(iso[2], 10) - 1, parseInt(iso[3], 10));
  return parseDateString(s);
}

function clearCache() {
  CacheService.getScriptCache().removeAll([]);
}

/**
 * ONE-TIME bulk permission change: sets every linked report PDF (the
 * hyperlink on each row's Report Release Date cell) to "Anyone with the
 * link can view". Run manually from the Apps Script editor - it is NOT
 * exposed through doGet, since it mutates sharing settings and should
 * only ever be triggered deliberately.
 *
 * PRIVACY NOTE: this makes every linked report accessible to anyone who
 * obtains the link (not indexed/searchable, but not private either).
 * These are real patient genetic test reports - only run this if that
 * trade-off is acceptable. The first run will prompt for an additional
 * Drive authorization scope, since the script has so far only touched
 * the Sheet.
 */
function bulkShareAllReportLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fileIds = new Set();

  ss.getSheets().forEach(function (sheet) {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return;
    const range = sheet.getRange(1, 1, lastRow, lastCol);
    const data = range.getValues();
    let richTextValues = null;
    try { richTextValues = range.getRichTextValues(); } catch (e) { }
    let formulas = null;
    try { formulas = range.getFormulas(); } catch (e) { }
    const linkGrid = getHyperlinkGrid(ss.getId(), sheet.getName(), lastRow, lastCol);

    for (let i = 0; i < data.length; i++) {
      for (let j = 0; j < data[i].length; j++) {
        const text = String(data[i][j] == null ? '' : data[i][j]).trim();
        const link = getCellLink(linkGrid, richTextValues, formulas, i, j, text);
        if (link) {
          const m = link.match(/\/d\/([a-zA-Z0-9_-]+)/);
          if (m) fileIds.add(m[1]);
        }
      }
    }
  });

  Logger.log('Found ' + fileIds.size + ' unique linked files.');

  let updated = 0, failed = 0;
  fileIds.forEach(function (id) {
    try {
      DriveApp.getFileById(id).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      updated++;
    } catch (e) {
      failed++;
      Logger.log('Failed for ' + id + ': ' + e);
    }
  });

  Logger.log('Updated: ' + updated + ', Failed: ' + failed);
}

function testDoGet() {
  const result = doGet({ parameter: { debug: '1' } });
  Logger.log(result.getContent().substring(0, 5000));
}

function testDoGet2023() {
  const result = doGet({ parameter: { sheet: '2023', debug: '1' } });
  Logger.log(result.getContent());
}

/**
 * Deep probe for dated rows that the batch reader found NO link on. For each
 * such row it re-reads every cell INDIVIDUALLY (per-cell getRichTextValue,
 * which sometimes exposes links the batch getRichTextValues misses on
 * date-typed cells) and reports any column that carries a link. Run from the
 * editor; pass a sheet name, or leave blank to probe all sheets.
 */
function inspectMissingDeep(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = sheetName ? [ss.getSheetByName(sheetName)] : ss.getSheets();
  const out = {};

  sheets.forEach(function (sheet) {
    if (!sheet) return;
    const name = sheet.getName();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) { out[name] = { empty: true }; return; }

    const range = sheet.getRange(1, 1, lastRow, lastCol);
    const data = range.getValues();
    let rtv = null;
    try { rtv = range.getRichTextValues(); } catch (e) { }

    // header + date column
    let dateCol = -1, headerRow = -1, headerCells = [];
    for (let i = 0; i < data.length && dateCol === -1; i++) {
      for (let j = 0; j < data[i].length; j++) {
        if (canonicalHeader(normalizeHeader(data[i][j])) === 'report_release_date') {
          dateCol = j; headerRow = i; headerCells = data[i].map(function (c) { return String(c); }); break;
        }
      }
    }
    const summary = { dateCol: dateCol, headerCells: headerCells, probed: [] };
    if (dateCol === -1) { out[name] = summary; return; }

    let probes = 0;
    for (let i = headerRow + 1; i < data.length && probes < 8; i++) {
      const v = data[i][dateCol];
      const txt = (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd-MM-yyyy') : String(v == null ? '' : v).trim();
      if (!txt) continue;
      // skip repeated header rows
      if (canonicalHeader(normalizeHeader(txt)) === 'report_release_date') continue;

      // did the BATCH reader see a link on the date cell?
      let batchLink = '';
      if (rtv && rtv[i] && rtv[i][dateCol]) {
        try { batchLink = rtv[i][dateCol].getLinkUrl() || ''; } catch (e) { }
        if (!batchLink) { const runs = rtv[i][dateCol].getRuns(); for (let r = 0; r < runs.length; r++) { if (runs[r].getLinkUrl()) { batchLink = runs[r].getLinkUrl(); break; } } }
      }
      if (batchLink) continue; // not a miss

      // PER-CELL probe of the whole row
      const rowLinks = {};
      for (let c = 0; c < lastCol; c++) {
        let url = '';
        try {
          const rt = sheet.getRange(i + 1, c + 1).getRichTextValue();
          if (rt) {
            url = rt.getLinkUrl() || '';
            if (!url) { const rs = rt.getRuns(); for (let r = 0; r < rs.length; r++) { if (rs[r].getLinkUrl()) { url = rs[r].getLinkUrl(); break; } } }
          }
        } catch (e) { }
        if (url) rowLinks[(headerCells[c] || ('col' + c))] = url;
      }
      summary.probed.push({ row: i + 1, dateText: txt, perCellLinksInRow: rowLinks });
      probes++;
    }
    out[name] = summary;
  });
  Logger.log(JSON.stringify(out, null, 2));
}

/**
 * Confirms WHY links are missing: for every dated data row, classifies it by
 * (has a link?) x (cell is a true Date object vs text). If links cluster on
 * text cells and misses cluster on Date-typed cells, the hyperlinks were lost
 * when Sheets auto-converted DD-MM-YYYY text (day <= 12) into dates.
 */
function inspectLinkVsType() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = {};
  ss.getSheets().forEach(function (sheet) {
    const name = sheet.getName();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) { out[name] = { empty: true }; return; }
    const range = sheet.getRange(1, 1, lastRow, lastCol);
    const data = range.getValues();
    let rtv = null; try { rtv = range.getRichTextValues(); } catch (e) { }

    let dateCol = -1, headerRow = -1;
    for (let i = 0; i < data.length && dateCol === -1; i++) {
      for (let j = 0; j < data[i].length; j++) {
        if (canonicalHeader(normalizeHeader(data[i][j])) === 'report_release_date') { dateCol = j; headerRow = i; break; }
      }
    }
    const s = { linkedDate: 0, linkedText: 0, missDate: 0, missText: 0 };
    if (dateCol === -1) { out[name] = s; return; }

    for (let i = headerRow + 1; i < data.length; i++) {
      const v = data[i][dateCol];
      const txt = (v instanceof Date) ? 'd' : String(v == null ? '' : v).trim();
      if (!txt) continue;
      if (canonicalHeader(normalizeHeader(String(v))) === 'report_release_date') continue;
      let link = '';
      if (rtv && rtv[i] && rtv[i][dateCol]) {
        try { link = rtv[i][dateCol].getLinkUrl() || ''; } catch (e) { }
        if (!link) { const rs = rtv[i][dateCol].getRuns(); for (let r = 0; r < rs.length; r++) { if (rs[r].getLinkUrl()) { link = rs[r].getLinkUrl(); break; } } }
      }
      const isDate = (v instanceof Date);
      if (link && isDate) s.linkedDate++;
      else if (link && !isDate) s.linkedText++;
      else if (!link && isDate) s.missDate++;
      else s.missText++;
    }
    out[name] = s;
  });
  Logger.log(JSON.stringify(out, null, 2));
}

/**
 * Run this from the editor to see, for every sheet, how many dated rows got a
 * link and which detection method found it. Pinpoints exactly which reports
 * are being missed and why. Logs a compact JSON summary.
 */
function inspectLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = {};
  ss.getSheets().forEach(function (sheet) {
    const name = sheet.getName();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) { out[name] = { empty: true }; return; }

    const range = sheet.getRange(1, 1, lastRow, lastCol);
    const data = range.getValues();
    let rtv = null, fx = null;
    try { rtv = range.getRichTextValues(); } catch (e) { }
    try { fx = range.getFormulas(); } catch (e) { }
    const grid = getHyperlinkGrid(sheet.getParent().getId(), name, lastRow, lastCol);

    // Locate the report_release_date column from the first header row.
    let dateCol = -1, headerRow = -1;
    for (let i = 0; i < data.length && dateCol === -1; i++) {
      for (let j = 0; j < data[i].length; j++) {
        if (canonicalHeader(normalizeHeader(data[i][j])) === 'report_release_date') { dateCol = j; headerRow = i; break; }
      }
    }

    const summary = {
      dateCol: dateCol, headerRow: headerRow + 1,
      advancedApi: !!(grid && grid.length && !grid._error),
      advancedApiError: grid && grid._error ? grid._error : undefined,
      dated: 0, linked: 0,
      methods: { advancedApi: 0, wholeCell: 0, run: 0, formula: 0, inlineText: 0 },
      samplesMissing: []
    };
    if (dateCol === -1) { out[name] = summary; return; }

    for (let i = headerRow + 1; i < data.length; i++) {
      const v = data[i][dateCol];
      const txt = (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd-MM-yyyy') : String(v == null ? '' : v).trim();
      if (!txt) continue;
      // skip repeated section-header rows ("Report Release Date") so counts are real
      if (canonicalHeader(normalizeHeader(txt)) === 'report_release_date') continue;
      summary.dated++;

      let method = '';
      if (grid && grid[i] && grid[i][dateCol]) method = 'advancedApi';
      if (!method && rtv && rtv[i] && rtv[i][dateCol]) {
        const rt = rtv[i][dateCol];
        try { if (rt.getLinkUrl()) method = 'wholeCell'; } catch (e) { }
        if (!method) { const runs = rt.getRuns(); for (let r = 0; r < runs.length; r++) { if (runs[r].getLinkUrl()) { method = 'run'; break; } } }
      }
      if (!method && fx && fx[i] && /HYPERLINK\(/i.test(String(fx[i][dateCol] || ''))) method = 'formula';
      if (!method && extractAnyUrl(txt)) method = 'inlineText';

      if (method) { summary.linked++; summary.methods[method]++; }
      else if (summary.samplesMissing.length < 6) {
        summary.samplesMissing.push({
          row: i + 1,
          dateText: txt,
          formula: fx && fx[i] ? String(fx[i][dateCol] || '') : '',
          rawValue: String(v)
        });
      }
    }
    out[name] = summary;
  });
  Logger.log(JSON.stringify(out, null, 2));
}
