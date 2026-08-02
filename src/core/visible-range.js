import {
  buildGridMetrics,
  estimateVisibleColumnCount,
  estimateVisibleRowCount,
} from "./grid-metrics.js";

export const DEFAULT_VISIBLE_RANGE_OVERSCAN = {
  rows: 6,
  columns: 2,
};

export function calculateVisibleRange(metrics, viewportState, overscan = {}) {
  const safeMetrics = buildGridMetrics(null, metrics || {});
  const safeViewport = normalizeViewportState(viewportState);
  const safeOverscan = normalizeOverscan(overscan);
  const visibleRowCount = Math.max(1, estimateVisibleRowCount(safeViewport.height, safeMetrics));
  const visibleColumnCount = Math.max(1, estimateVisibleColumnCount(safeViewport.width, safeMetrics));
  const rowWindowSpan = Math.min(safeMetrics.rowCount, visibleRowCount + safeOverscan.rows * 2);
  let startRow = clamp(
    Math.floor(safeViewport.scrollTop / safeMetrics.rowHeight) + 1 - safeOverscan.rows,
    1,
    safeMetrics.rowCount,
  );
  let endRow = clamp(startRow + rowWindowSpan - 1, startRow, safeMetrics.rowCount);
  startRow = clamp(endRow - rowWindowSpan + 1, 1, endRow);
  const columnWindowSpan = Math.min(safeMetrics.columnCount, visibleColumnCount + safeOverscan.columns * 2);
  let startColumn = clamp(
    Math.floor(safeViewport.scrollLeft / safeMetrics.columnWidth) + 1 - safeOverscan.columns,
    1,
    safeMetrics.columnCount,
  );
  let endColumn = clamp(startColumn + columnWindowSpan - 1, startColumn, safeMetrics.columnCount);
  startColumn = clamp(endColumn - columnWindowSpan + 1, 1, endColumn);
  return {
    startRow,
    endRow,
    startColumn,
    endColumn,
    visibleRowCount,
    visibleColumnCount,
    overscanRows: safeOverscan.rows,
    overscanColumns: safeOverscan.columns,
  };
}

export function visibleRangeSignature(range) {
  const safeRange = range || {};
  return [
    Math.max(1, Math.round(Number(safeRange.startRow) || 1)),
    Math.max(1, Math.round(Number(safeRange.endRow) || 1)),
    Math.max(1, Math.round(Number(safeRange.startColumn) || 1)),
    Math.max(1, Math.round(Number(safeRange.endColumn) || 1)),
  ].join(":");
}

function normalizeViewportState(viewportState) {
  return {
    scrollTop: Math.max(0, Math.round(Number(viewportState?.scrollTop) || 0)),
    scrollLeft: Math.max(0, Math.round(Number(viewportState?.scrollLeft) || 0)),
    width: Math.max(0, Math.round(Number(viewportState?.width) || 0)),
    height: Math.max(0, Math.round(Number(viewportState?.height) || 0)),
  };
}

function normalizeOverscan(overscan) {
  return {
    rows: Math.max(0, Math.round(Number(overscan?.rows) || DEFAULT_VISIBLE_RANGE_OVERSCAN.rows)),
    columns: Math.max(0, Math.round(Number(overscan?.columns) || DEFAULT_VISIBLE_RANGE_OVERSCAN.columns)),
  };
}

function clamp(value, minimum, maximum) {
  return Math.min(Math.max(value, minimum), maximum);
}
