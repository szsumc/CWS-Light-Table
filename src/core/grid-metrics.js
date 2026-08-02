import {
  getColumnCount,
  getGridLimits,
  getSheetRowCount,
} from "./table-model.js";

export const DEFAULT_GRID_COLUMN_WIDTH = 84;
export const DEFAULT_GRID_ROW_HEIGHT = 22;
export const DEFAULT_ROW_HEADER_WIDTH = 36;
export const DEFAULT_COLUMN_HEADER_HEIGHT = 22;

export function buildGridMetrics(table, options = {}) {
  const gridLimits = getGridLimits(table);
  const explicitColumnCount = normalizeOptionalPositiveInteger(options.columnCount);
  const explicitRowCount = normalizeOptionalPositiveInteger(options.rowCount);
  const minColumnCount = normalizePositiveInteger(options.minColumnCount, 1);
  const minRowCount = normalizePositiveInteger(options.minRowCount, 1);
  const columnCount = Math.max(1, explicitColumnCount || 0, minColumnCount, getColumnCount(table), gridLimits.maxColumns);
  const rowCount = Math.max(1, explicitRowCount || 0, minRowCount, getSheetRowCount(table), gridLimits.maxRows);
  const columnWidth = normalizePositiveInteger(options.columnWidth, DEFAULT_GRID_COLUMN_WIDTH);
  const rowHeight = normalizePositiveInteger(options.rowHeight, DEFAULT_GRID_ROW_HEIGHT);
  const rowHeaderWidth = normalizePositiveInteger(options.rowHeaderWidth, DEFAULT_ROW_HEADER_WIDTH);
  const columnHeaderHeight = normalizePositiveInteger(options.columnHeaderHeight, DEFAULT_COLUMN_HEADER_HEIGHT);
  return {
    columnCount,
    rowCount,
    columnWidth,
    rowHeight,
    rowHeaderWidth,
    columnHeaderHeight,
    contentWidth: rowHeaderWidth + columnCount * columnWidth,
    contentHeight: columnHeaderHeight + rowCount * rowHeight,
  };
}

export function estimateVisibleColumnCount(viewportWidth, metrics) {
  const safeMetrics = buildGridMetrics(null, metrics || {});
  const bodyWidth = Math.max(0, normalizeNonNegativeInteger(viewportWidth) - safeMetrics.rowHeaderWidth);
  return Math.max(1, Math.ceil(bodyWidth / safeMetrics.columnWidth));
}

export function estimateVisibleRowCount(viewportHeight, metrics) {
  const safeMetrics = buildGridMetrics(null, metrics || {});
  const bodyHeight = Math.max(0, normalizeNonNegativeInteger(viewportHeight) - safeMetrics.columnHeaderHeight);
  return Math.max(1, Math.ceil(bodyHeight / safeMetrics.rowHeight));
}

export function columnStartPixel(columnIndex, metrics) {
  const safeMetrics = buildGridMetrics(null, metrics || {});
  const column = normalizePositiveInteger(columnIndex, 1);
  return safeMetrics.rowHeaderWidth + (column - 1) * safeMetrics.columnWidth;
}

export function rowStartPixel(rowIndex, metrics) {
  const safeMetrics = buildGridMetrics(null, metrics || {});
  const row = normalizePositiveInteger(rowIndex, 1);
  return safeMetrics.columnHeaderHeight + (row - 1) * safeMetrics.rowHeight;
}

function normalizePositiveInteger(value, fallback) {
  return Math.max(1, Math.round(Number(value) || fallback));
}

function normalizeOptionalPositiveInteger(value) {
  const numeric = Math.round(Number(value));
  return Number.isFinite(numeric) && numeric >= 1 ? numeric : 0;
}

function normalizeNonNegativeInteger(value) {
  return Math.max(0, Math.round(Number(value) || 0));
}
