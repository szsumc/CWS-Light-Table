import {
  columnKeyForIndex as columnKeyForIndexFromTableModel,
  getCellValue as getTableCellValue,
  setCellValue as setTableCellValue,
} from "./table-model.js";

export function getCellValue(table, rowIndex, columnIndex) {
  return getTableCellValue(table, rowIndex, columnIndex);
}

export function setCellValue(table, rowIndex, columnIndex, value) {
  return setTableCellValue(table, rowIndex, columnIndex, value);
}

export function createCellSelection(cell) {
  const normalized = normalizeCell(cell);
  return {
    mode: "cell",
    anchor: normalized,
    focus: normalized,
  };
}

export function createRangeSelection(anchor, focus) {
  return {
    mode: "range",
    anchor: normalizeCell(anchor),
    focus: normalizeCell(focus),
  };
}

export function createRowSelection(rowIndex, columnCount, focusRowIndex = rowIndex) {
  const row = normalizePositiveIndex(rowIndex);
  const focusRow = normalizePositiveIndex(focusRowIndex);
  const endCol = normalizePositiveIndex(columnCount);
  return {
    mode: "row",
    anchor: { row, col: 1 },
    focus: { row: focusRow, col: endCol },
  };
}

export function createColumnSelection(columnIndex, rowCount, focusColumnIndex = columnIndex) {
  const col = normalizePositiveIndex(columnIndex);
  const focusCol = normalizePositiveIndex(focusColumnIndex);
  const endRow = normalizePositiveIndex(rowCount);
  return {
    mode: "column",
    anchor: { row: 1, col },
    focus: { row: endRow, col: focusCol },
  };
}

export function getSelectionBounds(selection) {
  const anchor = normalizeCell(selection?.anchor);
  const focus = normalizeCell(selection?.focus);
  return {
    startRow: Math.min(anchor.row, focus.row),
    endRow: Math.max(anchor.row, focus.row),
    startCol: Math.min(anchor.col, focus.col),
    endCol: Math.max(anchor.col, focus.col),
  };
}

export function isCellSelected(selection, rowIndex, columnIndex) {
  const bounds = getSelectionBounds(selection);
  return (
    rowIndex >= bounds.startRow &&
    rowIndex <= bounds.endRow &&
    columnIndex >= bounds.startCol &&
    columnIndex <= bounds.endCol
  );
}

export function isRowHeaderSelected(selection, rowIndex) {
  const bounds = getSelectionBounds(selection);
  return rowIndex >= bounds.startRow && rowIndex <= bounds.endRow;
}

export function isColumnHeaderSelected(selection, columnIndex) {
  const bounds = getSelectionBounds(selection);
  return columnIndex >= bounds.startCol && columnIndex <= bounds.endCol;
}

export function moveActiveCell(activeCell, direction, limits) {
  const cell = normalizeCell(activeCell);
  const rowCount = normalizePositiveIndex(limits?.rowCount);
  const columnCount = normalizePositiveIndex(limits?.columnCount);
  const deltas = {
    left: { row: 0, col: -1 },
    right: { row: 0, col: 1 },
    up: { row: -1, col: 0 },
    down: { row: 1, col: 0 },
    tab: { row: 0, col: 1 },
    "shift+tab": { row: 0, col: -1 },
    enter: { row: 1, col: 0 },
    "shift+enter": { row: -1, col: 0 },
  };
  const delta = deltas[direction] || { row: 0, col: 0 };
  return {
    row: clampNumber(cell.row + delta.row, 1, rowCount),
    col: clampNumber(cell.col + delta.col, 1, columnCount),
  };
}

export function normalizeCell(cell) {
  return {
    row: normalizePositiveIndex(cell?.row),
    col: normalizePositiveIndex(cell?.col),
  };
}

export function columnKeyForIndex(columnIndex) {
  return columnKeyForIndexFromTableModel(columnIndex);
}

export function getKeyboardEditIntent(eventLike) {
  const key = String(eventLike?.key ?? "");
  if (!key) return null;
  if (eventLike?.ctrlKey || eventLike?.metaKey) return null;
  if (key === "F2") {
    return { type: "edit-current", text: "" };
  }
  if ((key === "Backspace" || key === "Delete") && !eventLike?.altKey) {
    return { type: "replace-all", text: "" };
  }
  if (key === "Enter" && eventLike?.altKey) {
    return { type: "append", text: "\n" };
  }
  if (key.length === 1 && !eventLike?.altKey) {
    return { type: "replace-all", text: key };
  }
  return null;
}

export function applyKeyboardEditIntent(currentValue, intent) {
  const value = String(currentValue ?? "");
  if (!intent) return value;
  if (intent.type === "replace-all") {
    return String(intent.text ?? "");
  }
  if (intent.type === "append") {
    return `${value}${String(intent.text ?? "")}`;
  }
  return value;
}

function normalizePositiveIndex(value) {
  return Math.max(1, Math.trunc(Number(value) || 1));
}

function clampNumber(value, min, max) {
  return Math.max(min, Math.min(max, Math.trunc(Number(value) || min)));
}
