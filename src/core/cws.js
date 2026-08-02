import {
  STANDALONE_APP_SCRIPT,
  STANDALONE_APP_SHELL,
  STANDALONE_APP_STYLE,
} from "./standalone-template.js";
import {
  cloneLightTable,
  columnKeyForIndex,
  createEmptyLightTable,
  DEFAULT_SHEET_NAME,
  getColumnCount,
  getCellStoredValue,
  getGridLimits,
  getHeaderRowIndex,
  getSheetRowCount,
} from "./table-model.js";

export const CWS_HTML_FORMAT = "CWS_HTML";
export const CWS_SCHEMA_VERSION = 1;
export const CWS_MODEL_SCRIPT_ID = "websheet-model";

const DEFAULT_APP_NAME = "CWS Light Table";
const DEFAULT_APP_SHORT_NAME = "CWS Light Table";
const DEFAULT_GENERATOR_VERSION = "1.1.1";
const DEFAULT_GUIDE_URL = "https://szsumc.github.io/CWS-Light-Table-Docs/cws-html-guide-v1.json";
const DEFAULT_SCHEMA_URL = "https://szsumc.github.io/CWS-Light-Table-Docs/cws-html-workbook-model-v1.schema.json";

export function looksLikeCwsHtml(source = "") {
  const text = String(source || "");
  const hasCwsFormat =
    /\bdata-cws-format=["']CWS_HTML["']/i.test(text) ||
    /<meta\b(?=[^>]*\bname=["']cws:format["'])(?=[^>]*\bcontent=["']CWS_HTML["'])/i.test(text);
  return (
    hasCwsFormat &&
    /<script\b(?=[^>]*\bid=["']websheet-model["'])(?=[^>]*\btype=["']application\/json["'])/i.test(text)
  );
}

export function parseCwsHtmlWorkbook(source, options = {}) {
  const html = String(source || "");
  if (!looksLikeCwsHtml(html)) {
    throw new Error("The selected file is not recognized as CWS HTML.");
  }
  const payload = extractCwsModelScriptText(html);
  let workbook;
  try {
    workbook = JSON.parse(payload);
  } catch (error) {
    throw new Error(`The CWS workbook JSON is invalid: ${error.message}`);
  }
  if (!hasWorkbookSheets(workbook)) {
    throw new Error("The CWS workbook JSON does not contain any sheets.");
  }
  const normalized = cloneJsonLikeValue(workbook);
  normalized.sourceName = sanitizeHtmlFileName(options.fileName || normalized.sourceName || "workbook.cws.html");
  normalized.saveType = "html";
  return normalized;
}

export function listWorkbookSheets(workbook) {
  return (workbook?.sheets || []).map((sheet, index) => ({
    index,
    name: String(sheet?.name || `Sheet${index + 1}`),
    rowCount: normalizePositiveInteger(sheet?.rowCount, 1),
    colCount: normalizePositiveInteger(sheet?.colCount, 1),
  }));
}

export function workbookToLightTable(workbook, sheetRef = undefined) {
  const { sheet, sheetIndex } = resolveWorkbookSheet(workbook, sheetRef);
  const usedRange = inferUsedRange(sheet);
  const savedGridLimits = normalizeGridLimitsDefinition(sheet?.cwsLight?.gridLimits, {
    maxRows: sheet?.rowCount,
    maxColumns: sheet?.colCount,
  });
  const colCount = Math.max(1, usedRange.colCount);
  const rowCount = Math.max(1, usedRange.rowCount);
  const savedHeaderRowIndex = normalizeHeaderRowIndex(
    sheet?.cwsLight?.headerRowIndex,
    Math.max(rowCount, savedGridLimits.maxRows),
  );
  const columns = [];
  for (let columnIndex = 1; columnIndex <= colCount; columnIndex += 1) {
    const headerValue = flattenCellValue(sheet.cells?.[cellKey(savedHeaderRowIndex, columnIndex)]);
    columns.push({
      key: columnKeyForIndex(columnIndex),
      label: normalizeColumnLabel(headerValue),
    });
  }
  const leadingRows = [];
  for (let rowIndex = 1; rowIndex < savedHeaderRowIndex; rowIndex += 1) {
    leadingRows.push(buildRecordFromSheetRow(sheet, columns, rowIndex));
  }
  const rows = [];
  for (let rowIndex = savedHeaderRowIndex + 1; rowIndex <= rowCount; rowIndex += 1) {
    rows.push(buildRecordFromSheetRow(sheet, columns, rowIndex));
  }
  return {
    sourceName: sanitizeHtmlFileName(workbook.sourceName || "workbook.cws.html"),
    sheetName: String(sheet.name || `Sheet${sheetIndex + 1}`),
    columns,
    leadingRows,
    rows,
    gridLimits: {
      maxRows: Math.max(savedGridLimits.maxRows, rowCount, savedHeaderRowIndex),
      maxColumns: Math.max(savedGridLimits.maxColumns, colCount),
    },
    pageSetup: normalizePageSetup(sheet.pageLayout),
  };
}

export function serializeLightTableToCwsHtml(table, options = {}) {
  const workbook = lightTableToWorkbook(table, options);
  const sourceName = sanitizeHtmlFileName(options.fileName || table?.sourceName || workbook.sourceName || "workbook.cws.html");
  const title = escapeHtml(sourceName);
  const payload = safeJson(workbook, 2);
  return `<!doctype html>
<html ${cwsHtmlRootAttributes()}>
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
${cwsHtmlMetaTags()}
<title>${title}</title>
<style>${STANDALONE_APP_STYLE}</style>
</head>
<body data-cws-light-shell="standalone">
${STANDALONE_APP_SHELL}
${cwsAiNoticeComment()}
${cwsModelScriptTag(payload)}
${cwsAiInstructionsScript()}
<script id="cws-light-runtime" data-cws-light-runtime="standalone">${escapeInlineScript(STANDALONE_APP_SCRIPT)}</script>
</body>
</html>`;
}

export function lightTableToWorkbook(table, options = {}) {
  const normalizedTable = normalizeLightTable(table);
  const gridLimits = getGridLimits(normalizedTable);
  const colCount = Math.max(1, getColumnCount(normalizedTable), gridLimits.maxColumns);
  const rowCount = Math.max(1, getSheetRowCount(normalizedTable), gridLimits.maxRows);
  const headerRowIndex = getHeaderRowIndex(normalizedTable);
  const cells = {};
  for (let rowIndex = 1; rowIndex <= rowCount; rowIndex += 1) {
    for (let columnIndex = 1; columnIndex <= colCount; columnIndex += 1) {
      const value = normalizePrimitive(getCellStoredValue(normalizedTable, rowIndex, columnIndex));
      if (value === "") continue;
      cells[cellKey(rowIndex, columnIndex)] = {
        row: rowIndex,
        col: columnIndex,
        raw: value,
        display: value == null ? "" : String(value),
      };
    }
  }
  const sourceName = sanitizeHtmlFileName(options.fileName || normalizedTable.sourceName || "workbook.cws.html");
  return {
    version: 8,
    sourceName,
    saveType: "html",
    createdAt: new Date().toISOString(),
    activeSheetIndex: 0,
    showFormulaBar: true,
    definedNames: [],
    relationships: [],
    workbookProperties: {},
    calcProperties: {},
    sheets: [
      {
        name: normalizedSheetName(normalizedTable.sheetName || DEFAULT_SHEET_NAME),
        state: "visible",
        rowCount,
        colCount,
        defaultColumnWidth: 64,
        defaultRowHeight: 20,
        zoomScale: 100,
        displayScale: 1,
        defaultCellCss: {},
        columns: Array.from({ length: colCount }, () => ({ width: 64, hidden: false })),
        rows: Array.from({ length: rowCount }, () => ({ height: 20, hidden: false })),
        merges: [],
        images: [],
        inkStrokes: [],
        cells,
        autoFilter: null,
        cwsLight: {
          headerRowIndex,
          gridLimits,
        },
        pageLayout: normalizePageSetup(normalizedTable.pageSetup),
      },
    ],
  };
}

export function cellKey(row, col) {
  return `${row}:${col}`;
}

function cwsHtmlRootAttributes() {
  return [
    `lang="en"`,
    `data-cws-format="${escapeAttr(CWS_HTML_FORMAT)}"`,
    `data-cws-schema-version="${escapeAttr(CWS_SCHEMA_VERSION)}"`,
    `data-cws-app="${escapeAttr(DEFAULT_APP_SHORT_NAME)}"`,
  ].join(" ");
}

function cwsHtmlMetaTags() {
  return [
    `<meta name="generator" content="${escapeAttr(`${DEFAULT_APP_NAME} ${DEFAULT_GENERATOR_VERSION}`)}" />`,
    `<meta name="cws:format" content="${escapeAttr(CWS_HTML_FORMAT)}" />`,
    `<meta name="cws:schema-version" content="${escapeAttr(CWS_SCHEMA_VERSION)}" />`,
    `<meta name="cws:guide" content="${escapeAttr(DEFAULT_GUIDE_URL)}" />`,
    `<meta name="cws:schema" content="${escapeAttr(DEFAULT_SCHEMA_URL)}" />`,
  ].join("\n");
}

function cwsModelScriptTag(payload) {
  return `<script id="${CWS_MODEL_SCRIPT_ID}" type="application/json" data-cws-model="workbook" data-cws-format="${escapeAttr(CWS_HTML_FORMAT)}" data-cws-schema-version="${escapeAttr(CWS_SCHEMA_VERSION)}" data-cws-editable="true">${payload}</script>`;
}

function cwsAiNoticeComment() {
  return `<!--
CWS Light Table AI Notice:
This file is a CWS HTML workbook.
The actual workbook data is the JSON inside <script id="${CWS_MODEL_SCRIPT_ID}" type="application/json">.
Only edit that JSON when changing workbook content.
-->`;
}

function cwsAiInstructionsScript() {
  return `<script id="cws-ai-instructions" type="application/json">${safeJson(cwsAiInstructionsPayload(), 2)}</script>`;
}

function cwsAiInstructionsPayload() {
  return {
    notice: "This file is a CWS HTML workbook. The actual workbook data is the JSON inside script#websheet-model. Only edit that JSON when changing workbook content.",
    format: CWS_HTML_FORMAT,
    appName: DEFAULT_APP_NAME,
    shortName: DEFAULT_APP_SHORT_NAME,
    schemaVersion: CWS_SCHEMA_VERSION,
    modelScriptId: CWS_MODEL_SCRIPT_ID,
    editableRegion: `script#${CWS_MODEL_SCRIPT_ID}`,
    guideUrl: DEFAULT_GUIDE_URL,
    schemaUrl: DEFAULT_SCHEMA_URL,
    rules: [
      "Only edit the workbook JSON inside script#websheet-model.",
      "Keep JSON valid when changing workbook content.",
    ],
  };
}

function extractCwsModelScriptText(source) {
  const sanitized = String(source || "").replace(/<!--[\s\S]*?-->/g, "");
  const match = sanitized.match(/<script\b(?=[^>]*\bid=["']websheet-model["'])(?=[^>]*\btype=["']application\/json["'])[^>]*>([\s\S]*?)<\/script>/i);
  if (!match) throw new Error("The CWS workbook model script was not found.");
  return String(match[1] || "").trim();
}

function hasWorkbookSheets(workbook) {
  return Boolean(workbook && typeof workbook === "object" && Array.isArray(workbook.sheets) && workbook.sheets.length > 0);
}

function resolveWorkbookSheet(workbook, sheetRef = undefined) {
  if (!hasWorkbookSheets(workbook)) {
    throw new Error("The workbook does not contain any sheets.");
  }
  let sheetIndex;
  if (typeof sheetRef === "number" && Number.isInteger(sheetRef)) {
    sheetIndex = sheetRef;
  } else if (typeof sheetRef === "string" && sheetRef.trim() !== "") {
    sheetIndex = workbook.sheets.findIndex((sheet) => String(sheet?.name || "") === sheetRef);
  } else {
    sheetIndex = Number.isInteger(workbook.activeSheetIndex) ? workbook.activeSheetIndex : 0;
  }
  if (sheetIndex < 0 || sheetIndex >= workbook.sheets.length) {
    throw new Error("The requested sheet could not be found.");
  }
  const sheet = workbook.sheets[sheetIndex];
  return { sheet, sheetIndex };
}

function inferUsedRange(sheet = {}) {
  let maxRow = 1;
  let maxCol = 1;
  Object.entries(sheet.cells || {}).forEach(([key, cell]) => {
    const parsed = parseCellKey(key);
    if (!parsed) return;
    const value = flattenCellValue(cell);
    if (value === "" && value !== 0 && value !== false) return;
    maxRow = Math.max(maxRow, parsed.row);
    maxCol = Math.max(maxCol, parsed.col);
  });
  return {
    rowCount: maxRow,
    colCount: maxCol,
  };
}

function buildRecordFromSheetRow(sheet, columns, rowIndex) {
  const record = {};
  columns.forEach((column, offset) => {
    const cell = sheet.cells?.[cellKey(rowIndex, offset + 1)];
    record[column.key] = flattenCellValue(cell);
  });
  return record;
}

function parseCellKey(key) {
  const match = String(key || "").match(/^([1-9]\d*):([1-9]\d*)$/);
  if (!match) return null;
  return {
    row: Number(match[1]),
    col: Number(match[2]),
  };
}

function flattenCellValue(cell) {
  if (!cell || typeof cell !== "object") return "";
  if (typeof cell.formula === "string" && cell.formula.startsWith("=")) {
    if (cell.display != null) return normalizePrimitive(cell.display);
    if (cell.cached != null) return normalizePrimitive(cell.cached);
    if (cell.raw != null) return normalizePrimitive(cell.raw);
    return "";
  }
  if (cell.raw != null) return normalizePrimitive(cell.raw);
  if (cell.display != null) return normalizePrimitive(cell.display);
  return "";
}

function normalizePrimitive(value) {
  if (value == null) return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
  return String(value);
}

function normalizeColumnLabel(value) {
  if (value == null) return "";
  return String(value);
}

function normalizeLightTable(table = {}) {
  const base = cloneLightTable(isPlainObject(table) ? table : createEmptyLightTable());
  const columns = base.columns.map((column, index) => ({
    key: String(column?.key || columnKeyForIndex(index + 1)),
    label: normalizeColumnLabel(column?.label),
  }));
  const normalizeRecordList = (records) => records.map((record) => {
    const output = {};
    columns.forEach((column) => {
      output[column.key] = normalizePrimitive(record?.[column.key]);
    });
    return output;
  });
  return {
    sourceName: sanitizeHtmlFileName(base.sourceName || "workbook.cws.html"),
    sheetName: normalizedSheetName(base.sheetName || DEFAULT_SHEET_NAME),
    columns,
    leadingRows: normalizeRecordList(base.leadingRows || []),
    rows: normalizeRecordList(base.rows || []),
    gridLimits: normalizeGridLimitsDefinition(base.gridLimits),
    pageSetup: normalizePageSetup(base.pageSetup),
  };
}

function normalizedSheetName(value) {
  const text = String(value || DEFAULT_SHEET_NAME).trim() || DEFAULT_SHEET_NAME;
  const stripped = text.replace(/[:\\/?*\[\]]/g, " ").trim();
  return (stripped || DEFAULT_SHEET_NAME).slice(0, 31);
}

function sanitizeHtmlFileName(value) {
  const text = String(value || "workbook.cws.html").trim() || "workbook.cws.html";
  const safe = text.replace(/[\\/:*?"<>|]+/g, "-");
  return /\.html?$/i.test(safe) ? safe : `${safe}.html`;
}

function normalizePositiveInteger(value, fallback) {
  const number = Math.max(1, Math.round(Number(value) || fallback));
  return number;
}

function normalizePageSetup(pageLayout = {}) {
  const source = pageLayout && typeof pageLayout === "object" ? pageLayout : {};
  return {
    margins: cloneJsonLikeValue(source.margins || {}),
    paperSize: source.paperSize || "",
    orientation: source.orientation || "",
    printArea: cloneJsonLikeValue(source.printArea || null),
    headerFooter: cloneJsonLikeValue(source.headerFooter || {}),
    background: cloneJsonLikeValue(source.background || null),
  };
}

function normalizeHeaderRowIndex(value, maxRowCount) {
  const upperBound = Math.max(1, normalizePositiveInteger(maxRowCount, 1));
  return Math.max(1, Math.min(upperBound, normalizePositiveInteger(value, 1)));
}

function normalizeGridLimitsDefinition(value, fallback = {}) {
  return {
    maxRows: Math.max(1, normalizePositiveInteger(value?.maxRows, fallback.maxRows || 50)),
    maxColumns: Math.max(1, normalizePositiveInteger(value?.maxColumns, fallback.maxColumns || 30)),
  };
}

function isPlainObject(value) {
  return Boolean(value && typeof value === "object");
}

function cloneJsonLikeValue(value) {
  return JSON.parse(JSON.stringify(value ?? null));
}

function safeJson(value, space = 0) {
  return JSON.stringify(value, null, space)
    .replace(/</g, "\\u003c")
    .replace(/\u2028/g, "\\u2028")
    .replace(/\u2029/g, "\\u2029");
}

function escapeInlineScript(source) {
  return String(source ?? "")
    .replace(/<\/script/gi, "<\\/script")
    .replace(/<!--/g, "<\\!--");
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "\"": "&quot;",
    "'": "&#39;",
  })[char]);
}

function escapeAttr(value) {
  return escapeHtml(value).replace(/`/g, "&#96;");
}
