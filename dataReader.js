/**
 * ═══════════════════════════════════════════════════════════════════════════
 * engine/dataReader.js  —  IVP IRR Analytics Engine
 * ═══════════════════════════════════════════════════════════════════════════
 *
 * Replaces all Google Apps Script SpreadsheetApp / readTab() calls with
 * their Office.js Excel.run() equivalents.
 *
 * Key contract: every function that was called via google.script.run in
 * the original HTML modules is available here as an async export.
 * The output shapes are IDENTICAL to the GAS versions so no HTML UI
 * logic needs to change.
 *
 * Sheet tab names are the same as in Google Sheets:
 *   Master_Data_2_CF          ← cashflow source
 *   Master_Data_2_valuation   ← valuation / NAV source
 *   IRR_Config                ← active runtime config (key-value)
 *   Config_Library            ← saved named configurations
 *   Sim_Config                ← Anthropic API key + model settings
 *   Simulation_Log            ← simulation run history (append-only)
 *   Mapping_Table             ← transaction type → bucket mapping
 *
 * ═══════════════════════════════════════════════════════════════════════════
 */

'use strict';

import { computeEntityIRR, calcMetrics } from './xirr.js';

// ── Constants (mirrors Code.gs SYS object) ────────────────────────────────
export const PATH_SEP = ' > ';

// Excel date serial epoch offset (days from 1900-01-01 to 1970-01-01)
// Excel incorrectly treats 1900 as a leap year, hence 25569 not 25567
const EXCEL_DATE_OFFSET_DAYS = 25569;
const MS_PER_DAY              = 86400 * 1000;

// ─────────────────────────────────────────────────────────────────────────
//  SECTION 1 — LOW-LEVEL EXCEL UTILITIES
// ─────────────────────────────────────────────────────────────────────────

/**
 * excelReadTab — Reads a worksheet and returns an array of header-keyed objects.
 *
 * This is the direct replacement for Code.gs readTab(sheetName, headerRow).
 * Output contract is identical: array of { colName: value } plain objects.
 *
 * @param {string} sheetName   Worksheet name (case-sensitive, must exist)
 * @param {number} headerRow   1-indexed row containing column headers (default 1)
 * @returns {Promise<Array>}   Array of row objects, empty array if sheet is empty
 */
export async function excelReadTab(sheetName, headerRow = 1) {
  return Excel.run(async (ctx) => {
    // Use getItemOrNullObject — getItem throws at ctx.sync() (not synchronously)
    // so a try-catch around getItem alone would NOT catch missing-sheet errors.
    const ws = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
    ws.load('isNullObject');
    await ctx.sync();
    if (ws.isNullObject) {
      console.warn(`[dataReader] Sheet "${sheetName}" not found — returning [].`);
      return [];
    }

    const range = ws.getUsedRange();
    range.load('values, rowCount, columnCount');
    await ctx.sync();

    if (!range.values || range.rowCount === 0) return [];

    const rows    = range.values;
    const headers = rows[headerRow - 1]; // 1-indexed

    const data = [];
    for (let r = headerRow; r < rows.length; r++) {
      const row = rows[r];

      // Skip entirely empty rows
      if (row.every(cell => cell === '' || cell === null || cell === undefined)) continue;

      const obj = {};
      headers.forEach((h, i) => {
        if (h !== null && h !== '' && h !== undefined) {
          obj[String(h).trim()] = normaliseCell(row[i]);
        }
      });
      data.push(obj);
    }

    return data;
  });
}

/**
 * normaliseCell — Converts Excel cell values to JS-native types.
 *
 * Critical: Excel returns dates as serial numbers (float), not Date objects.
 * This function detects them by range (realistic PE/VC dates = 2000–2050)
 * and converts to JS Date with ACT/365 fidelity.
 *
 * @param {*} val  Raw value from range.values[][]
 * @returns {*}    Normalised value (Date, number, string, boolean, or null)
 */
export function normaliseCell(val) {
  if (val === null || val === undefined || val === '') return null;

  // Booleans
  if (typeof val === 'boolean') return val;

  // Excel date serial: floats in the range 36526–73050 (= year 2000–2099)
  // This range safely avoids false positives from large dollar amounts
  if (typeof val === 'number' && val >= 36526 && val <= 73050) {
    const d = new Date((val - EXCEL_DATE_OFFSET_DAYS) * MS_PER_DAY);
    // Validate it parsed as a plausible date
    if (!isNaN(d.getTime()) && d.getFullYear() >= 1990 && d.getFullYear() <= 2100) {
      // Normalise to UTC midnight to avoid timezone drift in year-frac calcs
      return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
    }
  }

  // Numbers (non-date)
  if (typeof val === 'number') return val;

  // String date formats: "3/31/2025", "2025-03-31", "31-Mar-2025" etc.
  if (typeof val === 'string') {
    const trimmed = val.trim();
    if (trimmed === '') return null;

    // Try to detect date strings
    const datePatterns = [
      /^\d{4}-\d{2}-\d{2}$/,          // ISO: 2025-03-31
      /^\d{1,2}\/\d{1,2}\/\d{4}$/,    // US: 3/31/2025
      /^\d{1,2}-[A-Za-z]{3}-\d{4}$/,  // 31-Mar-2025
    ];
    if (datePatterns.some(p => p.test(trimmed))) {
      const d = new Date(trimmed);
      if (!isNaN(d.getTime()) && d.getFullYear() >= 1990 && d.getFullYear() <= 2100) {
        return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
      }
    }
    return trimmed;
  }

  return val;
}

/**
 * getSheetNames — Returns all worksheet names in the active workbook.
 * Replaces SpreadsheetApp.getSheets().map(s => s.getName())
 *
 * @returns {Promise<string[]>}
 */
export async function getSheetNames() {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load('items/name');
    await ctx.sync();
    return sheets.items.map(s => s.name);
  });
}

/**
 * getTabHeaders — Returns the header row of a sheet as a string array.
 * Replaces Code.gs getTabHeaders(tab, headerRow).
 *
 * @param {string} sheetName
 * @param {number} headerRow  1-indexed
 * @returns {Promise<string[]>}
 */
export async function getTabHeaders(sheetName, headerRow = 1) {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
    ws.load('isNullObject');
    await ctx.sync();
    if (ws.isNullObject) return [];
    const range = ws.getUsedRange();
    range.load('values');
    await ctx.sync();
    if (!range.values || range.values.length < headerRow) return [];
    return range.values[headerRow - 1]
      .filter(h => h !== null && h !== '' && h !== undefined)
      .map(h => String(h).trim());
  });
}

/**
 * appendRow — Appends a single row to the bottom of a sheet.
 * Replaces sheet.appendRow(rowData) in Apps Script.
 *
 * @param {string}  sheetName
 * @param {Array}   rowData     Flat array of cell values
 * @returns {Promise<void>}
 */
export async function appendRow(sheetName, rowData) {
  return Excel.run(async (ctx) => {
    const ws   = ctx.workbook.worksheets.getItem(sheetName);
    const used = ws.getUsedRange();
    used.load('rowCount');
    await ctx.sync();

    // getCell is 0-indexed; rowCount is the next available row
    const startCell = ws.getCell(used.rowCount, 0);
    const range     = startCell.getResizedRange(0, rowData.length - 1);
    range.values    = [rowData.map(v => v instanceof Date ? _dateToExcelSerial(v) : v)];
    await ctx.sync();
  });
}

/**
 * updateRowByKey — Find a row where column[keyCol] === keyValue and update it.
 * Used for upsert patterns (config save, etc.).
 *
 * @param {string} sheetName
 * @param {string} keyCol     Column header name to match on
 * @param {string} keyValue   Value to match
 * @param {Array}  newRowData Full row data to replace with
 * @returns {Promise<boolean>}  true if found+updated, false if not found
 */
export async function updateRowByKey(sheetName, keyCol, keyValue, newRowData) {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
    ws.load('isNullObject');
    await ctx.sync();
    if (ws.isNullObject) return false; // sheet doesn't exist yet
    const range = ws.getUsedRange();
    range.load('values');
    await ctx.sync();

    const rows    = range.values;
    const headers = rows[0];
    const keyIdx  = headers.findIndex(h => h === keyCol);
    if (keyIdx === -1) return false;

    for (let r = 1; r < rows.length; r++) {
      if (String(rows[r][keyIdx]).trim() === String(keyValue).trim()) {
        // 0-indexed in Excel.run, +1 for header offset relative to range start
        const rowRange = ws.getCell(r, 0).getResizedRange(0, newRowData.length - 1);
        rowRange.values = [newRowData.map(v => v instanceof Date ? _dateToExcelSerial(v) : v)];
        await ctx.sync();
        return true;
      }
    }
    return false; // row not found
  });
}

/**
 * deleteRowByKey — Delete the first row where column[keyCol] === keyValue.
 *
 * @returns {Promise<boolean>}  true if deleted, false if not found
 */
export async function deleteRowByKey(sheetName, keyCol, keyValue) {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
    ws.load('isNullObject');
    await ctx.sync();
    if (ws.isNullObject) return false;
    const range = ws.getUsedRange();
    range.load('values, address');
    await ctx.sync();

    const rows    = range.values;
    const headers = rows[0];
    const keyIdx  = headers.findIndex(h => h === keyCol);
    if (keyIdx === -1) return false;

    for (let r = 1; r < rows.length; r++) {
      if (String(rows[r][keyIdx]).trim() === String(keyValue).trim()) {
        // getRow is not in Excel.run — delete by selecting the row range
        const rowRange = ws.getCell(r, 0).getEntireRow();
        rowRange.delete(Excel.DeleteShiftDirection.up);
        await ctx.sync();
        return true;
      }
    }
    return false;
  });
}

// Internal: JS Date → Excel serial number
function _dateToExcelSerial(date) {
  if (!(date instanceof Date)) return date;
  return Math.round(date.getTime() / MS_PER_DAY + EXCEL_DATE_OFFSET_DAYS);
}


// ─────────────────────────────────────────────────────────────────────────
//  SECTION 2 — CONFIG MANAGEMENT
//  Replaces: parseConfig(), getConfigs(), saveConfig(), deleteConfig(),
//            setDefaultConfig(), getTabHeaders()
// ─────────────────────────────────────────────────────────────────────────

/**
 * Config schema (mirrors getUIConfig() from configLibrary.html):
 * {
 *   source:    { cfTab, valTab, cfHeaderRow:1, valHeaderRow:1 },
 *   cfMap:     { date, amount, type, currency, notes },
 *   valMap:    { date, amount },
 *   hierarchy: [{ level, label, cfCol, valCol }],
 *   txnMap:    { 'TypeName': { bucket, bucketName, isCF, sign } }  ← from Mapping_Table
 * }
 *
 * Config_Library sheet columns: Name | Config_JSON | Is_Default | Saved_At
 */

const CONFIG_LIB_SHEET  = 'Config_Library';
const CONFIG_LIB_COLS   = { name: 'Name', json: 'Config_JSON', isDefault: 'Is_Default', savedAt: 'Saved_At' };
const SETTINGS_KEY_CFG  = 'ivp_activeConfig';   // Office.settings key

/**
 * getConfigs — Returns all saved configurations from Config_Library sheet.
 * Replaces google.script.run.getConfigs()
 *
 * @returns {Promise<Array>}  [{name, config, isDefault, savedAt}]
 */
export async function getConfigs() {
  try {
    const rows = await excelReadTab(CONFIG_LIB_SHEET, 1);
    return rows
      .filter(r => r[CONFIG_LIB_COLS.name])
      .map(r => {
        let cfg = null;
        try { cfg = JSON.parse(r[CONFIG_LIB_COLS.json] || 'null'); } catch (e) {}
        return {
          name:      String(r[CONFIG_LIB_COLS.name]).trim(),
          config:    cfg,
          isDefault: String(r[CONFIG_LIB_COLS.isDefault] || '').toLowerCase() === 'true',
          savedAt:   r[CONFIG_LIB_COLS.savedAt] ? String(r[CONFIG_LIB_COLS.savedAt]) : ''
        };
      });
  } catch (e) {
    console.warn('[dataReader] getConfigs error:', e.message);
    return [];
  }
}

/**
 * saveConfig — Upsert a named configuration to Config_Library.
 * Replaces google.script.run.saveConfig(name, cfg, isDef)
 *
 * @param {string}  name      Config name
 * @param {object}  cfg       Config object (same shape as getUIConfig())
 * @param {boolean} isDefault Whether to mark as default
 * @returns {Promise<void>}
 */
export async function saveConfig(name, cfg, isDefault = false) {
  const savedAt   = new Date().toISOString().split('T')[0];
  const cfgJson   = JSON.stringify(cfg);
  const newRow    = [name, cfgJson, isDefault ? 'true' : 'false', savedAt];

  // Try update first; append if not found
  const updated = await updateRowByKey(CONFIG_LIB_SHEET, CONFIG_LIB_COLS.name, name, newRow);
  if (!updated) {
    // Check if header row exists — if sheet is empty, create it first
    await _ensureConfigLibHeaders();
    await appendRow(CONFIG_LIB_SHEET, newRow);
  }

  // If setting as default, unset all others
  if (isDefault) {
    await setDefaultConfig(name);
  }
}

/**
 * deleteConfig — Remove a named configuration from Config_Library.
 * Replaces google.script.run.deleteConfig(name)
 */
export async function deleteConfig(name) {
  return deleteRowByKey(CONFIG_LIB_SHEET, CONFIG_LIB_COLS.name, name);
}

/**
 * setDefaultConfig — Mark one config as default, unmark all others.
 * Replaces google.script.run.setDefaultConfig(name)
 * Also stores in Office.settings for fast retrieval.
 */
export async function setDefaultConfig(name) {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItemOrNullObject(CONFIG_LIB_SHEET);
    ws.load('isNullObject');
    await ctx.sync();
    if (ws.isNullObject) return; // nothing to do — sheet doesn't exist yet
    const range = ws.getUsedRange();
    range.load('values');
    await ctx.sync();

    const rows    = range.values;
    if (rows.length < 2) return;
    const headers = rows[0];
    const defIdx  = headers.findIndex(h => h === CONFIG_LIB_COLS.isDefault);
    const nameIdx = headers.findIndex(h => h === CONFIG_LIB_COLS.name);
    if (defIdx === -1 || nameIdx === -1) return;

    // Update all rows: set Is_Default = true only for the target name
    for (let r = 1; r < rows.length; r++) {
      const cell  = ws.getCell(r, defIdx);
      const rowName = String(rows[r][nameIdx] || '').trim();
      cell.values   = [[rowName === name ? 'true' : 'false']];
    }
    await ctx.sync();

    // Also persist in Office.settings (faster read on next open)
    Office.context.document.settings.set(SETTINGS_KEY_CFG, name);
    await new Promise((res, rej) =>
      Office.context.document.settings.saveAsync(r => r.status === Office.AsyncResultStatus.Succeeded ? res() : rej(r.error))
    );
  });
}

/**
 * getActiveConfig — Load the active (default) config.
 * Called on task-pane open to initialise BUNDLE.
 *
 * Resolution order:
 *   1. Office.settings → name → find in Config_Library
 *   2. Config_Library → first row with Is_Default = true
 *   3. Config_Library → first row (any config)
 *   4. null (no configs saved)
 *
 * @returns {Promise<{name, config}|null>}
 */
export async function getActiveConfig() {
  const configs = await getConfigs();
  if (configs.length === 0) return null;

  // Check Office.settings for stored default name
  const storedName = Office.context.document.settings.get(SETTINGS_KEY_CFG);
  if (storedName) {
    const match = configs.find(c => c.name === storedName);
    if (match) return match;
  }

  // Fall back to first config marked as default in sheet
  const sheetDefault = configs.find(c => c.isDefault);
  if (sheetDefault) return sheetDefault;

  // Fall back to first config
  return configs[0];
}

/**
 * loadMappingTable — Reads Mapping_Table sheet to build txnMap.
 * Maps: Transaction_Type → { bucket, bucketName, isCF, sign }
 *
 * Expected columns: Transaction_Type | Bucket_Number | Bucket_Name | Is_Cashflow | Sign
 */
export async function loadMappingTable() {
  try {
    const rows = await excelReadTab('Mapping_Table', 1);
    const txnMap = {};
    for (const row of rows) {
      const type = row['Transaction_Type'] || row['Type'] || row['transaction_type'];
      if (!type) continue;
      txnMap[String(type).trim()] = {
        bucket:     parseInt(row['Bucket_Number'] || row['Bucket'] || 1, 10),
        bucketName: String(row['Bucket_Name'] || row['Bucket_Label'] || 'Cashflow Impact'),
        isCF:       String(row['Is_Cashflow'] || 'true').toLowerCase() !== 'false',
        sign:       String(row['Sign'] || 'NEGATIVE').toUpperCase()
      };
    }
    return txnMap;
  } catch (e) {
    console.warn('[dataReader] Mapping_Table not found, using defaults:', e.message);
    return {};
  }
}

// Internal: ensure Config_Library has a header row
async function _ensureConfigLibHeaders() {
  return Excel.run(async (ctx) => {
    try {
      const ws    = ctx.workbook.worksheets.getItem(CONFIG_LIB_SHEET);
      const range = ws.getUsedRange();
      range.load('rowCount');
      await ctx.sync();
      if (range.rowCount > 0) return; // headers exist
    } catch (e) {
      // Sheet doesn't exist — create it
      ctx.workbook.worksheets.add(CONFIG_LIB_SHEET);
    }
    const ws = ctx.workbook.worksheets.getItem(CONFIG_LIB_SHEET);
    const header = ws.getCell(0, 0).getResizedRange(0, 3);
    header.values = [[CONFIG_LIB_COLS.name, CONFIG_LIB_COLS.json, CONFIG_LIB_COLS.isDefault, CONFIG_LIB_COLS.savedAt]];
    header.format.font.bold = true;
    header.format.fill.color = '#1a237e';
    header.format.font.color = '#ffffff';
    await ctx.sync();
  });
}


// ─────────────────────────────────────────────────────────────────────────
//  SECTION 3 — DATA NORMALISATION
//  Mirrors normaliseCFRow / normaliseValRow from Code.gs
// ─────────────────────────────────────────────────────────────────────────

/**
 * normaliseCFRow — Convert a raw CF tab row into a normalised transaction object.
 *
 * @param {object} row       Raw row from excelReadTab()
 * @param {object} cfg       Parsed config (cfMap + hierarchy + txnMap)
 * @param {number} idx       Row index (for generating _id)
 * @returns {object|null}    Normalised txn or null if row is invalid
 */
export function normaliseCFRow(row, cfg, idx) {
  const { cfMap, hierarchy, txnMap = {} } = cfg;

  const dateVal   = row[cfMap.date];
  const amountVal = row[cfMap.amount];
  const typeStr   = String(row[cfMap.type] || '').trim();

  if (!dateVal || amountVal === null || amountVal === undefined || amountVal === '') return null;

  const date = dateVal instanceof Date ? dateVal : _parseAnyDate(dateVal);
  if (!date) return null;

  const amount = parseFloat(amountVal);
  if (isNaN(amount)) return null;

  // Resolve transaction type mapping
  const mapEntry = txnMap[typeStr] || {
    bucket:     1,
    bucketName: 'Cashflow Impact',
    isCF:       true,
    sign:       amount < 0 ? 'NEGATIVE' : 'POSITIVE'
  };

  // Build hierarchy path
  const levels = {};
  const path   = [];
  for (const h of hierarchy) {
    const val = String(row[h.cfCol] || '').trim();
    levels[h.level] = val;
    if (val) path.push(val);
  }
  if (path.length === 0) return null;

  const quarter = _assignQuarter(date);

  return {
    _id:        'CF-' + String(idx).padStart(6, '0'),
    date,
    amount,
    source:     'cashflow',
    txnType:    typeStr,
    bucket:     mapEntry.bucket,
    bucketName: mapEntry.bucketName,
    isCashflow: mapEntry.isCF,
    sign:       mapEntry.sign,
    currency:   String(row[cfMap.currency] || 'USD'),
    notes:      String(row[cfMap.notes]    || ''),
    quarter,
    levels,
    path,
    pathKey:    path.join(PATH_SEP),   // e.g. "Fund 1 > Deal 1"
    pathStr:    path.join(PATH_SEP)
  };
}

/**
 * normaliseValRow — Convert a raw Val tab row into a normalised valuation object.
 */
export function normaliseValRow(row, cfg, idx) {
  const { valMap, hierarchy } = cfg;

  const dateVal   = row[valMap.date];
  const amountVal = row[valMap.amount];

  if (!dateVal || amountVal === null || amountVal === undefined || amountVal === '') return null;

  const date = dateVal instanceof Date ? dateVal : _parseAnyDate(dateVal);
  if (!date) return null;

  const amount = parseFloat(amountVal);
  if (isNaN(amount)) return null;

  const levels = {};
  const path   = [];
  for (const h of hierarchy) {
    const val = String(row[h.valCol] || '').trim();
    levels[h.level] = val;
    if (val) path.push(val);
  }
  if (path.length === 0) return null;

  const quarter = _assignQuarter(date);

  return {
    _id:        'VAL-' + String(idx).padStart(6, '0'),
    date,
    amount:     Math.abs(amount), // NAV is always stored as positive
    source:     'valuation',
    txnType:    'Valuation',
    bucket:     2,
    bucketName: 'Valuation Impact',
    isCashflow: false,
    sign:       'POSITIVE',
    currency:   'USD',
    notes:      '',
    quarter,
    levels,
    path,
    pathKey:    path.join(PATH_SEP),
    pathStr:    path.join(PATH_SEP)
  };
}

// Internal: assign a quarter label to a date ("Q1_2024", "Q4_2023" etc.)
function _assignQuarter(date) {
  if (!(date instanceof Date)) return '';
  const y = date.getUTCFullYear();
  const m = date.getUTCMonth() + 1; // 1-12
  const q = Math.ceil(m / 3);
  return `Q${q}_${y}`;
}

// Internal: safe any-format date parse
function _parseAnyDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}


// ─────────────────────────────────────────────────────────────────────────
//  SECTION 4 — BUNDLE BUILDER
//  The main entry point — replaces the server-side bundle injection
//  that was done via <?!= bundle ?> in HtmlService templates.
// ─────────────────────────────────────────────────────────────────────────

/**
 * buildBundle — Read CF + Val tabs, normalise, compute hierarchy + metrics.
 *
 * This replaces the server-side bundle injection (<?!= bundle ?>) in all
 * HTML module files. Call this once on task-pane open; the result is the
 * same BUNDLE object shape that analyticsDashboard.html, irrSimulator.html,
 * reverseIRR.html, managementCommentary.html all consume.
 *
 * BUNDLE shape:
 * {
 *   txnCount:        number,
 *   cfCount:         number,
 *   valCount:        number,
 *   quarters:        [{label, endDate, endDateStr}],
 *   hierarchy:       [{level, label}],
 *   hierarchyValues: { 1: ['Fund1','Fund2'], 2: ['Deal1'] },
 *   cascading:       { 2: { 'Fund1': ['Deal1','Deal2'] } },
 *   deals:           [{name, pathKey, pathStr, irr, moic, dpi, rvpi, nav, invested, distributed, holdMonths}],
 *   fund:            {irr, moic, dpi, rvpi, nav, invested, distributed},
 *   lastQuarter:     'Q4_2024',
 *   lastQuarterEnd:  Date,
 *   pathSep:         ' > ',
 *   maxLevel:        number,
 *   error:           null | string
 * }
 *
 * @param {object} cfg  Parsed config from getActiveConfig().config
 * @returns {Promise<object>}  The BUNDLE
 */
export async function buildBundle(cfg) {
  try {
    if (!cfg || !cfg.source) {
      return _errorBundle('No config provided. Open Config Library and run Save & Compute.');
    }

    const { source, cfMap, valMap, hierarchy } = cfg;

    if (!source.cfTab || !source.valTab) {
      return _errorBundle('Config is missing CF Tab or Val Tab. Open Config Library to configure.');
    }

    // ── Load txnMap from Mapping_Table (optional) ───────────────────────
    const txnMap = await loadMappingTable();
    const fullCfg = { ...cfg, txnMap };

    // ── Read both sheets in parallel ────────────────────────────────────
    const [cfRows, valRows] = await Promise.all([
      excelReadTab(source.cfTab,  source.cfHeaderRow  || 1),
      excelReadTab(source.valTab, source.valHeaderRow || 1)
    ]);

    if (cfRows.length === 0) {
      return _errorBundle(`Cashflow tab "${source.cfTab}" is empty or not found.`);
    }

    // ── Normalise rows ───────────────────────────────────────────────────
    const cfTxns  = cfRows.map((r, i)  => normaliseCFRow(r,  fullCfg, i)).filter(Boolean);
    const valTxns = valRows.map((r, i) => normaliseValRow(r, fullCfg, i)).filter(Boolean);
    const allTxns = [...cfTxns, ...valTxns].sort((a, b) => a.date - b.date);

    if (cfTxns.length === 0) {
      return _errorBundle(`No valid cashflow rows found in "${source.cfTab}". Check column mappings in Config Library.`);
    }

    // ── Build quarter list (auto-detect from data dates) ─────────────────
    const quarters = _buildQuarterList(allTxns);
    const lastQtr  = quarters.length > 0 ? quarters[quarters.length - 1] : null;

    // ── Build hierarchy structure ────────────────────────────────────────
    const { hierarchyValues, cascading } = _buildHierarchy(allTxns, hierarchy);

    // ── Build deal-level metrics (leaf-level paths) ──────────────────────
    const maxLevel = hierarchy.length;
    const leafPaths = _getAllLeafPaths(allTxns, maxLevel);

    const deals = [];
    for (const pathArr of leafPaths) {
      const pathKey = pathArr.join(PATH_SEP);
      const txns    = allTxns.filter(t => t.pathKey === pathKey || t.pathStr === pathKey);

      // Compute IRR up to last quarter
      const cutoff   = lastQtr ? lastQtr.endDate : new Date();
      const irrResult = computeEntityIRR(txns, cutoff);

      // Hold period in months from first CF to cutoff
      const firstCF = txns.find(t => t.isCashflow);
      const holdMonths = firstCF
        ? Math.round((cutoff - firstCF.date) / (30.44 * MS_PER_DAY))
        : 0;

      deals.push({
        name:        pathArr[pathArr.length - 1],  // leaf name
        pathKey,
        pathStr:     pathKey,
        irr:         irrResult.valid ? irrResult.irr : null,
        moic:        irrResult.valid ? (irrResult.metrics?.moic ?? null) : null,
        dpi:         irrResult.valid ? (irrResult.metrics?.dpi  ?? null) : null,
        rvpi:        irrResult.valid ? (irrResult.metrics?.rvpi ?? null) : null,
        nav:         irrResult.valid ? (irrResult.metrics?.nav  ?? 0)   : 0,
        invested:    irrResult.valid ? (irrResult.metrics?.invested    ?? 0) : 0,
        distributed: irrResult.valid ? (irrResult.metrics?.distributed ?? 0) : 0,
        holdMonths,
        bps:         irrResult.valid && irrResult.irr !== null ? Math.round(irrResult.irr * 10000) : null
      });
    }

    // ── Fund-level (all-transactions) IRR ────────────────────────────────
    const cutoff    = lastQtr ? lastQtr.endDate : new Date();
    const fundIRR   = computeEntityIRR(allTxns, cutoff);
    const fund = {
      irr:         fundIRR.valid ? fundIRR.irr : null,
      moic:        fundIRR.valid ? (fundIRR.metrics?.moic ?? null) : null,
      dpi:         fundIRR.valid ? (fundIRR.metrics?.dpi  ?? null) : null,
      rvpi:        fundIRR.valid ? (fundIRR.metrics?.rvpi ?? null) : null,
      nav:         fundIRR.valid ? (fundIRR.metrics?.nav  ?? 0)   : 0,
      invested:    fundIRR.valid ? (fundIRR.metrics?.invested    ?? 0) : 0,
      distributed: fundIRR.valid ? (fundIRR.metrics?.distributed ?? 0) : 0,
      totalNAV:    fundIRR.valid ? (fundIRR.metrics?.nav ?? 0)   : 0
    };

    // ── Assemble and return BUNDLE ────────────────────────────────────────
    return {
      txnCount:        allTxns.length,
      cfCount:         cfTxns.length,
      valCount:        valTxns.length,
      quarters,
      hierarchy:       hierarchy.map(h => ({ level: h.level, label: h.label })),
      hierarchyValues,
      cascading,
      deals,
      fund,
      lastQuarter:     lastQtr ? lastQtr.label : null,
      lastQuarterEnd:  lastQtr ? lastQtr.endDate : null,
      pathSep:         PATH_SEP,
      maxLevel,
      error:           null,
      // Raw transactions stored for module use (sub-period, attribution, etc.)
      _txns:           allTxns
    };

  } catch (e) {
    console.error('[dataReader] buildBundle error:', e);
    return _errorBundle('Bundle build failed: ' + e.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────
//  SECTION 5 — PER-MODULE DATA FUNCTIONS
//  Replaces the server-side functions called by each HTML module
// ─────────────────────────────────────────────────────────────────────────

/**
 * getAnalyticsAllTransactions — Per-entity cashflow detail table.
 * Replaces google.script.run.getAnalyticsAllTransactions(pathStr, quarter)
 *
 * @param {string} pathStr    Entity path e.g. "Fund 1 > Deal 1"
 * @param {string} quarter    Quarter label e.g. "Q4_2024"
 * @param {object} bundle     Current BUNDLE (pass window.IVP_BUNDLE)
 * @returns {object}  { irr, nav, invested, distributed, moic, terminalNAV, rows }
 */
export function getAnalyticsAllTransactions(pathStr, quarter, bundle) {
  const allTxns = bundle._txns || [];
  const qtr     = bundle.quarters.find(q => q.label === quarter);
  const cutoff  = qtr ? qtr.endDate : new Date();

  // Filter by path
  const txns = allTxns.filter(t =>
    (t.pathStr === pathStr || t.pathKey === pathStr) && t.date <= cutoff
  );

  const irrResult = computeEntityIRR(txns, cutoff);

  const rows = txns.map((t, i) => ({
    rowNum:      i + 1,
    date:        t.date instanceof Date ? t.date.toISOString().split('T')[0] : String(t.date),
    txnType:     t.txnType,
    amount:      t.amount,
    isCashflow:  t.isCashflow,
    isTerminal:  false, // terminal NAV is injected separately in the UI
    bucket:      t.bucket,
    bucketName:  t.bucketName,
    quarter:     t.quarter,
    currency:    t.currency || 'USD',
    notes:       t.notes || '',
    sign:        t.sign || (t.amount < 0 ? 'NEGATIVE' : 'POSITIVE'),
    source:      t.source
  }));

  return {
    irr:         irrResult.irr,
    nav:         irrResult.metrics?.nav ?? 0,
    invested:    irrResult.metrics?.invested ?? 0,
    distributed: irrResult.metrics?.distributed ?? 0,
    moic:        irrResult.metrics?.moic ?? null,
    terminalNAV: irrResult.metrics?.nav ?? 0,
    rows
  };
}

/**
 * getAnalyticsSubPeriodTrend — IRR per quarter (sub-period trending).
 * Replaces google.script.run.getAnalyticsSubPeriodTrend(pathStr, quarter)
 *
 * Returns IRR computed inception-to-date for each quarter in sequence.
 *
 * @param {string} pathStr
 * @param {string} endQuarter   Latest quarter to compute through
 * @param {object} bundle
 * @returns {Array}  [{ quarter, irrPct, moic, nav, invested, distributed }]
 */
export function getAnalyticsSubPeriodTrend(pathStr, endQuarter, bundle) {
  const allTxns  = bundle._txns || [];
  const quarters = bundle.quarters;
  const endIdx   = quarters.findIndex(q => q.label === endQuarter);
  const relevant = endIdx >= 0 ? quarters.slice(0, endIdx + 1) : quarters;

  const entityTxns = allTxns.filter(t =>
    t.pathStr === pathStr || t.pathKey === pathStr
  );

  return relevant
    .map(q => {
      const txnsUpTo = entityTxns.filter(t => t.date <= q.endDate);
      if (txnsUpTo.length === 0) return null;

      const r = computeEntityIRR(txnsUpTo, q.endDate);
      if (!r.valid) return null;

      return {
        quarter:     q.label,
        quarterEnd:  q.endDate instanceof Date ? q.endDate.toISOString().split('T')[0] : '',
        irrPct:      r.irr !== null ? parseFloat((r.irr * 100).toFixed(4)) : null,
        moic:        r.metrics?.moic ?? null,
        nav:         r.metrics?.nav  ?? 0,
        invested:    r.metrics?.invested    ?? 0,
        distributed: r.metrics?.distributed ?? 0
      };
    })
    .filter(Boolean);
}

/**
 * getSimConfig — Reads Sim_Config sheet to get Claude API key + model.
 * Replaces SIM_CONFIG bundle injection in irrSimulator.html.
 *
 * Expected columns: Key | Value
 * Rows: apiKey | <key>, apiModel | claude-sonnet-4-20250514, etc.
 *
 * @returns {Promise<{apiKey, apiModel, maxTokens}>}
 */
export async function getSimConfig() {
  try {
    const rows = await excelReadTab('Sim_Config', 1);
    const map  = {};
    for (const row of rows) {
      const k = String(row['Key'] || row['Setting'] || '').trim();
      const v = String(row['Value'] || row['Val']    || '').trim();
      if (k) map[k] = v;
    }
    return {
      apiKey:    map['apiKey']    || map['api_key']     || '',
      apiModel:  map['apiModel']  || map['model']       || 'claude-sonnet-4-20250514',
      maxTokens: parseInt(map['maxTokens'] || map['max_tokens'] || '1000', 10)
    };
  } catch (e) {
    console.warn('[dataReader] Sim_Config not found:', e.message);
    return { apiKey: '', apiModel: 'claude-sonnet-4-20250514', maxTokens: 1000 };
  }
}

/**
 * saveSimulationToLog — Append a simulation run to Simulation_Log sheet.
 * Replaces google.script.run.saveSimulationToLog(payload)
 *
 * @param {object} payload  { simName, mode, scenarioConfig, results, timestamp }
 */
export async function saveSimulationToLog(payload) {
  const row = [
    payload.timestamp || new Date().toISOString(),
    payload.simName   || '(unnamed)',
    payload.mode      || 'deal',
    JSON.stringify(payload.scenarioConfig || {}),
    payload.results?.fund?.irr  != null ? (payload.results.fund.irr  * 100).toFixed(2) + '%' : '',
    payload.results?.fund?.moic != null ?  payload.results.fund.moic.toFixed(2) + 'x'       : '',
    JSON.stringify(payload.results || {})
  ];

  try {
    await appendRow('Simulation_Log', row);
  } catch (e) {
    // Sheet may not exist — create header first
    await _ensureSimLogHeaders();
    await appendRow('Simulation_Log', row);
  }
}

async function _ensureSimLogHeaders() {
  return Excel.run(async (ctx) => {
    try {
      ctx.workbook.worksheets.getItem('Simulation_Log');
      await ctx.sync();
    } catch {
      ctx.workbook.worksheets.add('Simulation_Log');
      await ctx.sync();
      const ws     = ctx.workbook.worksheets.getItem('Simulation_Log');
      const header = ws.getCell(0, 0).getResizedRange(0, 6);
      header.values = [['Timestamp','Sim_Name','Mode','Scenario_Config','Stressed_IRR','Stressed_MOIC','Full_Results']];
      header.format.font.bold  = true;
      header.format.fill.color = '#0a1520';
      header.format.font.color = '#ffffff';
      await ctx.sync();
    }
  });
}


// ─────────────────────────────────────────────────────────────────────────
//  SECTION 6 — INTERNAL HELPERS
// ─────────────────────────────────────────────────────────────────────────

// Build quarter objects from transaction dates
function _buildQuarterList(txns) {
  const quarterMap = {};
  for (const t of txns) {
    if (!(t.date instanceof Date)) continue;
    const label = _assignQuarter(t.date);
    if (!quarterMap[label]) {
      quarterMap[label] = _quarterEndDate(t.date);
    }
  }

  return Object.entries(quarterMap)
    .sort((a, b) => a[1] - b[1])
    .map(([label, endDate]) => ({
      label,
      endDate,
      endDateStr: endDate instanceof Date ? endDate.toISOString().split('T')[0] : ''
    }));
}

// Quarter end date for a given date
function _quarterEndDate(date) {
  const y = date.getUTCFullYear();
  const m = date.getUTCMonth() + 1;
  const q = Math.ceil(m / 3);
  const endMonth = q * 3; // 3, 6, 9, 12
  const lastDay  = new Date(Date.UTC(y, endMonth, 0)).getUTCDate(); // day 0 of next month = last day of this month
  return new Date(Date.UTC(y, endMonth - 1, lastDay));
}

// Build hierarchyValues + cascading lookup from transaction data
function _buildHierarchy(txns, hierarchy) {
  const hierarchyValues = {};
  const cascading       = {};

  for (const h of hierarchy) {
    hierarchyValues[h.level] = [];
  }

  for (const t of txns) {
    for (const h of hierarchy) {
      const val = t.levels[h.level];
      if (val && !hierarchyValues[h.level].includes(val)) {
        hierarchyValues[h.level].push(val);
      }
    }
  }

  // Build cascading: for each level > 1, map parent-value → child values
  for (let lvl = 2; lvl <= hierarchy.length; lvl++) {
    cascading[lvl] = {};
    for (const t of txns) {
      const parentVal = t.levels[lvl - 1];
      const childVal  = t.levels[lvl];
      if (!parentVal || !childVal) continue;
      if (!cascading[lvl][parentVal]) cascading[lvl][parentVal] = [];
      if (!cascading[lvl][parentVal].includes(childVal)) {
        cascading[lvl][parentVal].push(childVal);
      }
    }
  }

  return { hierarchyValues, cascading };
}

// Get all unique leaf-level paths (deepest hierarchy level)
function _getAllLeafPaths(txns, maxLevel) {
  const seen = new Set();
  const paths = [];

  for (const t of txns) {
    if (Object.values(t.levels).filter(Boolean).length !== maxLevel) continue;
    const key = t.pathKey || t.pathStr;
    if (key && !seen.has(key)) {
      seen.add(key);
      paths.push(t.path || key.split(PATH_SEP));
    }
  }

  return paths;
}

// Return a bundle with an error state
function _errorBundle(msg) {
  return {
    txnCount: 0, cfCount: 0, valCount: 0,
    quarters: [], hierarchy: [], hierarchyValues: {},
    cascading: {}, deals: [], fund: {},
    lastQuarter: null, pathSep: PATH_SEP, maxLevel: 0,
    error: msg, _txns: []
  };
}