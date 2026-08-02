import {
  createEmptyLightTable as createEmptyLightTableModel,
  DEFAULT_GRID_LIMIT_COLUMN_COUNT,
  DEFAULT_GRID_LIMIT_ROW_COUNT,
  getCellValue,
  getColumnCount,
  getGridLimits,
  getHeaderRowIndex,
  getSheetRowCount,
} from "./table-model.js";

export const DEFAULT_VIEWPORT_COLUMN_COUNT = DEFAULT_GRID_LIMIT_COLUMN_COUNT;
export const DEFAULT_VIEWPORT_ROW_COUNT = DEFAULT_GRID_LIMIT_ROW_COUNT;

export function columnLabelFromIndex(columnIndex) {
  let index = Math.trunc(Number(columnIndex));
  if (index < 1) {
    throw new Error("Column index must be 1 or greater.");
  }
  let label = "";
  while (index > 0) {
    const remainder = (index - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    index = Math.floor((index - 1) / 26);
  }
  return label;
}

export function gridCellName(rowIndex, columnIndex) {
  const row = Math.max(1, Math.trunc(Number(rowIndex) || 1));
  return `${columnLabelFromIndex(columnIndex)}${row}`;
}

export function createEmptyLightTable() {
  return createEmptyLightTableModel();
}

export function isViewportPaddingRow(table, rowIndex) {
  const safeTable = isObject(table) ? table : createEmptyLightTable();
  const safeRowIndex = Math.max(1, Math.trunc(Number(rowIndex) || 1));
  return safeRowIndex > getSheetRowCount(safeTable);
}

export function buildSheetGridView(table, options = {}) {
  const safeTable = isObject(table) ? table : createEmptyLightTable();
  const gridLimits = getGridLimits(safeTable);
  const minColumnCount = normalizeViewportSize(options.minColumnCount, gridLimits.maxColumns);
  const minRowCount = normalizeViewportSize(options.minRowCount, gridLimits.maxRows);
  const dataColumnCount = getColumnCount(safeTable);
  const dataRowCount = getSheetRowCount(safeTable);
  const columnCount = Math.max(1, minColumnCount, dataColumnCount);
  const rowCount = Math.max(1, minRowCount, dataRowCount);
  const renderedColumnIndices = resolveRenderedIndices(options.columnIndices, columnCount);
  const columns = renderedColumnIndices.map((columnIndex) => ({
    index: columnIndex,
    key: `c${columnIndex}`,
    letter: columnLabelFromIndex(columnIndex),
    label: normalizeDisplayText(getCellValue(safeTable, getHeaderRowIndex(safeTable), columnIndex)),
  }));
  const renderedRowIndices = resolveRenderedRowIndices(options.rowIndices, rowCount);
  const rows = renderedRowIndices.map((rowIndex) => {
    const safeRowIndex = Math.max(1, Math.min(rowCount, rowIndex));
    return {
      index: safeRowIndex,
      cells: columns.map((column) => ({
        row: safeRowIndex,
        col: column.index,
        value: normalizeDisplayText(getCellValue(safeTable, safeRowIndex, column.index)),
      })),
    };
  });
  return {
    columnCount,
    rowCount,
    headerRowIndex: getHeaderRowIndex(safeTable),
    columns,
    rows,
  };
}

function normalizeViewportSize(value, fallback) {
  return Math.max(1, Math.trunc(Number(value) || fallback));
}

function resolveRenderedRowIndices(rowIndices, rowCount) {
  return resolveRenderedIndices(rowIndices, rowCount);
}

function resolveRenderedIndices(indices, maxCount) {
  if (!Array.isArray(indices) || !indices.length) {
    return Array.from({ length: maxCount }, (_, offset) => offset + 1);
  }
  const rendered = [];
  const seen = new Set();
  indices.forEach((index) => {
    const normalized = Math.max(1, Math.min(maxCount, Math.trunc(Number(index) || 1)));
    if (seen.has(normalized)) return;
    seen.add(normalized);
    rendered.push(normalized);
  });
  return rendered.length ? rendered : [1];
}

function normalizeDisplayText(value) {
  if (value == null) return "";
  return String(value);
}

function isObject(value) {
  return Boolean(value && typeof value === "object");
}
