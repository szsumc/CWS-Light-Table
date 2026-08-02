import { readFile, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const scriptDir = dirname(fileURLToPath(import.meta.url));
const projectDir = resolve(scriptDir, "..");

const styleSource = await readFile(resolve(projectDir, "src/styles.css"), "utf8");
const indexSource = await readFile(resolve(projectDir, "index.html"), "utf8");
const standaloneVendorScriptPayloads = {
  "./vendor/xlsx/dist/xlsx.full.min.js": await readVendorScriptAsBase64("vendor/xlsx/dist/xlsx.full.min.js"),
  "./vendor/codepage/dist/cpexcel.full.js": await readVendorScriptAsBase64("vendor/codepage/dist/cpexcel.full.js"),
  "./vendor/codepage/bits/51932.js": await readVendorScriptAsBase64("vendor/codepage/bits/51932.js"),
  "./vendor/codepage/cputils.js": await readVendorScriptAsBase64("vendor/codepage/cputils.js"),
};
const standaloneVendorScriptPayloadsSource = `const STANDALONE_VENDOR_SCRIPT_PAYLOADS = ${JSON.stringify(standaloneVendorScriptPayloads)};`;

const shellMatch = indexSource.match(/(<div class="app-shell">[\s\S]*?)\s*<script type="module" src="\.\/src\/main\.js(?:\?[^"]*)?"><\/script>/);
if (!shellMatch) {
  throw new Error("Failed to extract the standalone shell from index.html.");
}
const shellSource = shellMatch[1];

const moduleOrder = [
  "src/core/table-model.js",
  "src/core/data-cleaning.js",
  "src/core/viewport-state.js",
  "src/core/grid-metrics.js",
  "src/core/visible-range.js",
  "src/core/editor-state.js",
  "src/core/history.js",
  "src/core/clipboard.js",
  "src/core/grid.js",
  "src/core/importers.js",
  "src/core/exporters.js",
  "src/core/table-operations.js",
  "src/core/cws.js",
];

const standaloneHelpers = [];

for (const relativePath of moduleOrder) {
  const absolutePath = resolve(projectDir, relativePath);
  let source = await readFile(absolutePath, "utf8");
  source = stripImports(source);
  source = source.replace(/\bexport\s+/g, "");

  if (relativePath.endsWith("table-model.js")) {
    const renamed = renameTableModelSource(source);
    standaloneHelpers.push([
      renamed,
      "function createEmptyLightTable(...args) { return tableModelCreateEmptyLightTable(...args); }",
      "function cloneLightTable(...args) { return tableModelCloneLightTable(...args); }",
      "function columnKeyForIndex(...args) { return tableModelColumnKeyForIndex(...args); }",
      "function getColumnCount(...args) { return tableModelGetColumnCount(...args); }",
      "function getLeadingRowCount(...args) { return tableModelGetLeadingRowCount(...args); }",
      "function getGridLimits(...args) { return tableModelGetGridLimits(...args); }",
      "function getHeaderRowIndex(...args) { return tableModelGetHeaderRowIndex(...args); }",
      "function getSheetRowCount(...args) { return tableModelGetSheetRowCount(...args); }",
      "function isLeadingRow(...args) { return tableModelIsLeadingRow(...args); }",
      "function isHeaderRow(...args) { return tableModelIsHeaderRow(...args); }",
      "function isDataRow(...args) { return tableModelIsDataRow(...args); }",
      "function getCellValue(...args) { return tableModelGetCellValue(...args); }",
      "function getCellStoredValue(...args) { return tableModelGetCellStoredValue(...args); }",
      "function getRowValues(...args) { return tableModelGetRowValues(...args); }",
      "function setCellValue(...args) { return tableModelSetCellValue(...args); }",
      "function setHeaderRow(...args) { return tableModelSetHeaderRow(...args); }",
      "function setGridLimits(...args) { return tableModelSetGridLimits(...args); }",
      "function expandGridLimitsToInclude(...args) { return tableModelExpandGridLimitsToInclude(...args); }",
      "function getLastUsedRowIndex(...args) { return tableModelGetLastUsedRowIndex(...args); }",
      "function getLastUsedColumnIndex(...args) { return tableModelGetLastUsedColumnIndex(...args); }",
      "function insertRows(...args) { return tableModelInsertRows(...args); }",
      "function insertColumns(...args) { return tableModelInsertColumns(...args); }",
      "function insertCells(...args) { return tableModelInsertCells(...args); }",
      "function setPageSetup(...args) { return tableModelSetPageSetup(...args); }",
    ].join("\n"));
    continue;
  }

  if (relativePath.endsWith("editor-state.js")) {
    standaloneHelpers.push([
      "const getCellValueFromTableModel = tableModelGetCellValue;",
      "const setCellValueFromTableModel = tableModelSetCellValue;",
      "const columnKeyForIndexFromTableModel = tableModelColumnKeyForIndex;",
      source
        .replace(/\bgetTableCellValue\b/g, "getCellValueFromTableModel")
        .replace(/\bsetTableCellValue\b/g, "setCellValueFromTableModel")
        .replace(/\bcolumnKeyForIndexFromTableModel\b/g, "columnKeyForIndexFromTableModel"),
    ].join("\n"));
    continue;
  }

  if (relativePath.endsWith("grid.js")) {
    standaloneHelpers.push([
      "const createEmptyLightTableModel = tableModelCreateEmptyLightTable;",
      source,
    ].join("\n"));
    continue;
  }

  standaloneHelpers.push(source);
}

let mainSource = await readFile(resolve(projectDir, "src/main.js"), "utf8");
mainSource = stripImports(mainSource);
mainSource = mainSource.replace(/sourceFileName:\s*""\s*,/, "sourceFileName: inferSourceFileName(),");
mainSource = mainSource.replace(
  /(window\.addEventListener\("resize", handleGridViewportResize\);\r?\n(?:window\.addEventListener\("afterprint", handleAfterPrint\);\r?\n)?\r?\n)renderWorkspace\(\);/,
  "$1bootstrapFromEmbeddedWorkbook();\nrenderWorkspace();",
);
if (!mainSource.includes("bootstrapFromEmbeddedWorkbook();")) {
  throw new Error("Failed to inject bootstrapFromEmbeddedWorkbook() into the standalone runtime.");
}

const bootstrapHelpers = `
function bootstrapFromEmbeddedWorkbook() {
  const workbook = parseEmbeddedWorkbook();
  if (!workbook) return;
  const currentFileName = inferCurrentHtmlFileName();
  state.workbook = workbook;
  state.sourceFileName = sanitizeHtmlFileName(currentFileName || workbook.sourceName || state.sourceFileName || inferSourceFileName());
  const sheets = listWorkbookSheets(workbook);
  if (sheets.length > 1) {
    refs.sheetSelect.innerHTML = sheets
      .map((sheet) => \`<option value="\${sheet.index}">\${escapeHtml(sheet.name)} (\${sheet.rowCount}x\${sheet.colCount})</option>\`)
      .join("");
    refs.sheetPickerPanel.classList.remove("hidden");
    state.table = createEmptyLightTable();
    resetGridInteractionState();
    renderWorkspace();
    setStatus(\`Loaded \${state.sourceFileName}. Select one sheet to continue.\`);
    return;
  }
  refs.sheetPickerPanel.classList.add("hidden");
  loadSelectedSheet(sheets[0] ? sheets[0].index : 0);
}

function parseEmbeddedWorkbook() {
  const modelScript = document.getElementById(CWS_MODEL_SCRIPT_ID);
  if (!modelScript) return null;
  const payload = String(modelScript.textContent || "").trim();
  if (!payload) return null;
  try {
    const workbook = JSON.parse(payload);
    if (!hasWorkbookSheets(workbook)) return null;
    const normalized = cloneJsonLikeValue(workbook);
    normalized.sourceName = sanitizeHtmlFileName(normalized.sourceName || state.sourceFileName || inferSourceFileName());
    normalized.saveType = "html";
    return normalized;
  } catch (error) {
    setStatus(\`Failed to read the embedded workbook: \${error.message}\`);
    return null;
  }
}

function inferSourceFileName() {
  const path = String(window.location.pathname || "").trim();
  const parts = path.split(/[\\\\/]/).filter(Boolean);
  return sanitizeHtmlFileName(parts.at(-1) || "workbook.cws.html");
}

function inferCurrentHtmlFileName() {
  const path = String(window.location.pathname || "").trim();
  const parts = path.split(/[\\\\/]/).filter(Boolean);
  const candidate = parts.at(-1) || "";
  return /\\.html?$/i.test(candidate) ? sanitizeHtmlFileName(candidate) : "";
}
`.trim();

const standaloneScript = `
function standaloneBootstrap() {
${indentBlock(standaloneVendorScriptPayloadsSource, 2)}

${indentBlock(standaloneHelpers.join("\n\n"), 2)}

${indentBlock(bootstrapHelpers, 2)}

${indentBlock(mainSource, 2)}
}

standaloneBootstrap();
`.trim();

const output = [
  `export const STANDALONE_APP_STYLE = ${toTemplateLiteral(styleSource)};`,
  "",
  `export const STANDALONE_APP_SHELL = ${toTemplateLiteral(shellSource)};`,
  "",
  `export const STANDALONE_APP_SCRIPT = ${toTemplateLiteral(standaloneScript)};`,
  "",
].join("\n");

await writeFile(resolve(projectDir, "src/core/standalone-template.js"), output, "utf8");

function stripImports(source) {
  return source.replace(/^\s*import[\s\S]*?from\s+["'][^"']+["'];\s*\r?\n/gm, "");
}

function toTemplateLiteral(source) {
  return `\`${String(source)
    .replace(/\\/g, "\\\\")
    .replace(/`/g, "\\`")
    .replace(/\$\{/g, "\\${")}\``;
}

function indentBlock(source, size) {
  const prefix = " ".repeat(size);
  return String(source)
    .split(/\r?\n/)
    .map((line) => (line ? `${prefix}${line}` : line))
    .join("\n");
}

async function readVendorScriptAsBase64(relativePath) {
  const source = await readFile(resolve(projectDir, relativePath));
  return source.toString("base64");
}

function renameTableModelSource(source) {
  return String(source)
    .replace(/\bfunction createEmptyLightTable\(/g, "function tableModelCreateEmptyLightTable(")
    .replace(/\bfunction cloneLightTable\(/g, "function tableModelCloneLightTable(")
    .replace(/\bfunction columnKeyForIndex\(/g, "function tableModelColumnKeyForIndex(")
    .replace(/\bfunction getColumnCount\(/g, "function tableModelGetColumnCount(")
    .replace(/\bfunction getLeadingRowCount\(/g, "function tableModelGetLeadingRowCount(")
    .replace(/\bfunction getGridLimits\(/g, "function tableModelGetGridLimits(")
    .replace(/\bfunction getHeaderRowIndex\(/g, "function tableModelGetHeaderRowIndex(")
    .replace(/\bfunction getSheetRowCount\(/g, "function tableModelGetSheetRowCount(")
    .replace(/\bfunction isLeadingRow\(/g, "function tableModelIsLeadingRow(")
    .replace(/\bfunction isHeaderRow\(/g, "function tableModelIsHeaderRow(")
    .replace(/\bfunction isDataRow\(/g, "function tableModelIsDataRow(")
    .replace(/\bfunction getCellValue\(/g, "function tableModelGetCellValue(")
    .replace(/\bfunction getCellStoredValue\(/g, "function tableModelGetCellStoredValue(")
    .replace(/\bfunction getRowValues\(/g, "function tableModelGetRowValues(")
    .replace(/\bfunction setCellValue\(/g, "function tableModelSetCellValue(")
    .replace(/\bfunction setHeaderRow\(/g, "function tableModelSetHeaderRow(")
    .replace(/\bfunction setGridLimits\(/g, "function tableModelSetGridLimits(")
    .replace(/\bfunction expandGridLimitsToInclude\(/g, "function tableModelExpandGridLimitsToInclude(")
    .replace(/\bfunction getLastUsedRowIndex\(/g, "function tableModelGetLastUsedRowIndex(")
    .replace(/\bfunction getLastUsedColumnIndex\(/g, "function tableModelGetLastUsedColumnIndex(")
    .replace(/\bfunction insertRows\(/g, "function tableModelInsertRows(")
    .replace(/\bfunction insertColumns\(/g, "function tableModelInsertColumns(")
    .replace(/\bfunction insertCells\(/g, "function tableModelInsertCells(")
    .replace(/\bfunction setPageSetup\(/g, "function tableModelSetPageSetup(")
    .replace(/\bcreateEmptyLightTable\(/g, "tableModelCreateEmptyLightTable(")
    .replace(/\bcloneLightTable\(/g, "tableModelCloneLightTable(")
    .replace(/\bcolumnKeyForIndex\(/g, "tableModelColumnKeyForIndex(")
    .replace(/\bgetColumnCount\(/g, "tableModelGetColumnCount(")
    .replace(/\bgetLeadingRowCount\(/g, "tableModelGetLeadingRowCount(")
    .replace(/\bgetGridLimits\(/g, "tableModelGetGridLimits(")
    .replace(/\bgetHeaderRowIndex\(/g, "tableModelGetHeaderRowIndex(")
    .replace(/\bgetSheetRowCount\(/g, "tableModelGetSheetRowCount(")
    .replace(/\bisLeadingRow\(/g, "tableModelIsLeadingRow(")
    .replace(/\bisHeaderRow\(/g, "tableModelIsHeaderRow(")
    .replace(/\bisDataRow\(/g, "tableModelIsDataRow(")
    .replace(/\bgetCellValue\(/g, "tableModelGetCellValue(")
    .replace(/\bgetCellStoredValue\(/g, "tableModelGetCellStoredValue(")
    .replace(/\bgetRowValues\(/g, "tableModelGetRowValues(")
    .replace(/\bsetCellValue\(/g, "tableModelSetCellValue(")
    .replace(/\bsetHeaderRow\(/g, "tableModelSetHeaderRow(")
    .replace(/\bsetGridLimits\(/g, "tableModelSetGridLimits(")
    .replace(/\bexpandGridLimitsToInclude\(/g, "tableModelExpandGridLimitsToInclude(")
    .replace(/\bgetLastUsedRowIndex\(/g, "tableModelGetLastUsedRowIndex(")
    .replace(/\bgetLastUsedColumnIndex\(/g, "tableModelGetLastUsedColumnIndex(")
    .replace(/\binsertRows\(/g, "tableModelInsertRows(")
    .replace(/\binsertColumns\(/g, "tableModelInsertColumns(")
    .replace(/\binsertCells\(/g, "tableModelInsertCells(")
    .replace(/\bsetPageSetup\(/g, "tableModelSetPageSetup(");
}
