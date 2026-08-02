import {
  createRangeSelection,
  getCellValue,
  getSelectionBounds,
  normalizeCell,
} from "./editor-state.js";
import { applyMatrixToLightTable } from "./table-model.js";

export function selectionToMatrix(table, selection) {
  const bounds = getSelectionBounds(selection);
  const matrix = [];
  for (let rowIndex = bounds.startRow; rowIndex <= bounds.endRow; rowIndex += 1) {
    const row = [];
    for (let columnIndex = bounds.startCol; columnIndex <= bounds.endCol; columnIndex += 1) {
      row.push(getCellValue(table, rowIndex, columnIndex));
    }
    matrix.push(row);
  }
  return matrix;
}

export function selectionToDelimitedText(table, selection, options = {}) {
  const delimiter = options.delimiter ?? "\t";
  const lineEnding = options.lineEnding ?? "\r\n";
  return matrixToDelimitedText(selectionToMatrix(table, selection), {
    delimiter,
    lineEnding,
  });
}

export function matrixToDelimitedText(matrix, options = {}) {
  const delimiter = options.delimiter ?? "\t";
  const lineEnding = options.lineEnding ?? "\r\n";
  return matrix
    .map((row) => row.map((value) => encodeDelimitedCell(value, delimiter)).join(delimiter))
    .join(lineEnding);
}

export function parseDelimitedText(text, options = {}) {
  const source = String(text ?? "");
  if (source === "") {
    return options.preserveEmptyCell ? [[""]] : [];
  }
  const delimiter = options.delimiter ?? "\t";
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    const nextChar = source[index + 1];
    if (inQuotes) {
      if (char === "\"" && nextChar === "\"") {
        cell += "\"";
        index += 1;
      } else if (char === "\"") {
        inQuotes = false;
      } else {
        cell += char;
      }
      continue;
    }

    if (char === "\"") {
      inQuotes = true;
      continue;
    }
    if (char === delimiter) {
      row.push(cell);
      cell = "";
      continue;
    }
    if (char === "\r" || char === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
      if (char === "\r" && nextChar === "\n") {
        index += 1;
      }
      continue;
    }
    cell += char;
  }

  row.push(cell);
  rows.push(row);
  return trimTrailingBlankRow(rows, source);
}

export function applyMatrixToTable(table, startCell, matrix) {
  return applyMatrixToLightTable(table, normalizeCell(startCell), matrix);
}

export function selectionFromMatrix(startCell, matrix) {
  const origin = normalizeCell(startCell);
  const rowCount = Math.max(1, matrix.length);
  const columnCount = Math.max(1, Math.max(...matrix.map((row) => row.length), 1));
  return createRangeSelection(origin, {
    row: origin.row + rowCount - 1,
    col: origin.col + columnCount - 1,
  });
}

export function fillSelectionByRepeat(table, selection, targetCell) {
  const sourceBounds = getSelectionBounds(selection);
  const focus = normalizeCell(targetCell);
  const targetBounds = {
    startRow: Math.min(sourceBounds.startRow, focus.row),
    endRow: Math.max(sourceBounds.endRow, focus.row),
    startCol: Math.min(sourceBounds.startCol, focus.col),
    endCol: Math.max(sourceBounds.endCol, focus.col),
  };
  if (boundsEqual(sourceBounds, targetBounds)) {
    return {
      table,
      selection,
    };
  }

  const sourceMatrix = selectionToMatrix(table, selection);
  const sourceRowCount = Math.max(1, sourceMatrix.length);
  const sourceColumnCount = Math.max(1, Math.max(...sourceMatrix.map((row) => row.length), 1));
  let nextTable = table;

  for (let rowIndex = targetBounds.startRow; rowIndex <= targetBounds.endRow; rowIndex += 1) {
    for (let columnIndex = targetBounds.startCol; columnIndex <= targetBounds.endCol; columnIndex += 1) {
      if (isWithinBounds(rowIndex, columnIndex, sourceBounds)) continue;
      const sourceRowOffset = positiveModulo(rowIndex - sourceBounds.startRow, sourceRowCount);
      const sourceColumnOffset = positiveModulo(columnIndex - sourceBounds.startCol, sourceColumnCount);
      const nextValue = sourceMatrix[sourceRowOffset]?.[sourceColumnOffset] ?? "";
      nextTable = setCellValue(nextTable, rowIndex, columnIndex, nextValue);
    }
  }

  return {
    table: nextTable,
    selection: createRangeSelection(
      { row: targetBounds.startRow, col: targetBounds.startCol },
      { row: targetBounds.endRow, col: targetBounds.endCol },
    ),
  };
}

function encodeDelimitedCell(value, delimiter) {
  const text = String(value ?? "");
  if (!/["\r\n]/.test(text) && !text.includes(delimiter)) {
    return text;
  }
  return `"${text.replace(/"/g, "\"\"")}"`;
}

function trimTrailingBlankRow(rows, source) {
  if (!rows.length) return [];
  const endsWithLineBreak = /(?:\r\n|\r|\n)$/.test(source);
  if (!endsWithLineBreak) return rows;
  const lastRow = rows.at(-1) ?? [];
  const isBlank = lastRow.every((cell) => cell === "");
  return isBlank && rows.length > 1 ? rows.slice(0, -1) : rows;
}

function boundsEqual(left, right) {
  return left.startRow === right.startRow
    && left.endRow === right.endRow
    && left.startCol === right.startCol
    && left.endCol === right.endCol;
}

function isWithinBounds(rowIndex, columnIndex, bounds) {
  return rowIndex >= bounds.startRow
    && rowIndex <= bounds.endRow
    && columnIndex >= bounds.startCol
    && columnIndex <= bounds.endCol;
}

function positiveModulo(value, divisor) {
  return ((value % divisor) + divisor) % divisor;
}
