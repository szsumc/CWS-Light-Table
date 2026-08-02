export const DEFAULT_GRID_LIMIT_ROW_COUNT = 50;
export const DEFAULT_GRID_LIMIT_COLUMN_COUNT = 30;
export const DEFAULT_SHEET_NAME = "DataSheet";

export function createEmptyLightTable() {
  return {
    sourceName: "workbook.cws.html",
    sheetName: DEFAULT_SHEET_NAME,
    columns: [],
    leadingRows: [],
    rows: [],
    gridLimits: {
      maxRows: DEFAULT_GRID_LIMIT_ROW_COUNT,
      maxColumns: DEFAULT_GRID_LIMIT_COLUMN_COUNT,
    },
    pageSetup: {
      margins: {},
      paperSize: "",
      orientation: "",
      printArea: null,
      headerFooter: {},
      background: null,
    },
  };
}

export function cloneLightTable(table = {}) {
  return {
    sourceName: String(table?.sourceName || "workbook.cws.html"),
    sheetName: String(table?.sheetName || DEFAULT_SHEET_NAME),
    columns: Array.isArray(table?.columns)
      ? table.columns.map((column, index) => ({
        key: String(column?.key || columnKeyForIndex(index + 1)),
        label: String(column?.label ?? ""),
      }))
      : [],
    leadingRows: Array.isArray(table?.leadingRows)
      ? table.leadingRows.map((record) => ({ ...(record || {}) }))
      : [],
    rows: Array.isArray(table?.rows)
      ? table.rows.map((record) => ({ ...(record || {}) }))
      : [],
    gridLimits: normalizeGridLimits(table?.gridLimits),
    pageSetup: table?.pageSetup ? JSON.parse(JSON.stringify(table.pageSetup)) : {},
  };
}

export function columnKeyForIndex(columnIndex) {
  return `c${normalizePositiveIndex(columnIndex)}`;
}

export function getColumnCount(table) {
  return Math.max(0, Array.isArray(table?.columns) ? table.columns.length : 0);
}

export function getLeadingRowCount(table) {
  return Math.max(0, Array.isArray(table?.leadingRows) ? table.leadingRows.length : 0);
}

export function getHeaderRowIndex(table) {
  return getLeadingRowCount(table) + 1;
}

export function getSheetRowCount(table) {
  return getHeaderRowIndex(table) + Math.max(0, Array.isArray(table?.rows) ? table.rows.length : 0);
}

export function getGridLimits(table) {
  return normalizeGridLimits(table?.gridLimits);
}

export function isLeadingRow(table, rowIndex) {
  return normalizePositiveIndex(rowIndex) < getHeaderRowIndex(table);
}

export function isHeaderRow(table, rowIndex) {
  return normalizePositiveIndex(rowIndex) === getHeaderRowIndex(table);
}

export function isDataRow(table, rowIndex) {
  return normalizePositiveIndex(rowIndex) > getHeaderRowIndex(table);
}

export function getCellValue(table, rowIndex, columnIndex) {
  const rawValue = getCellStoredValue(table, rowIndex, columnIndex);
  return String(rawValue ?? "");
}

export function getCellStoredValue(table, rowIndex, columnIndex) {
  const row = normalizePositiveIndex(rowIndex);
  const col = normalizePositiveIndex(columnIndex);
  const key = columnKeyForIndex(col);
  const headerRowIndex = getHeaderRowIndex(table);

  if (row < headerRowIndex) {
    return table?.leadingRows?.[row - 1]?.[key] ?? "";
  }
  if (row === headerRowIndex) {
    return table?.columns?.[col - 1]?.label ?? "";
  }
  return table?.rows?.[row - headerRowIndex - 1]?.[key] ?? "";
}

export function getRowValues(table, rowIndex, columnCount = getColumnCount(table)) {
  const values = [];
  const safeColumnCount = Math.max(0, normalizeNonNegativeInteger(columnCount));
  for (let columnIndex = 1; columnIndex <= safeColumnCount; columnIndex += 1) {
    values.push(getCellValue(table, rowIndex, columnIndex));
  }
  return values;
}

export function setCellValue(table, rowIndex, columnIndex, value) {
  const row = normalizePositiveIndex(rowIndex);
  const col = normalizePositiveIndex(columnIndex);
  const nextValue = String(value ?? "");
  const nextTable = cloneLightTable(table);
  ensureColumnCount(nextTable, col);

  const headerRowIndex = getHeaderRowIndex(nextTable);
  const key = columnKeyForIndex(col);

  if (row < headerRowIndex) {
    ensureLeadingRowCount(nextTable, row);
    nextTable.leadingRows[row - 1][key] = nextValue;
    return nextTable;
  }

  if (row === headerRowIndex) {
    nextTable.columns[col - 1].label = nextValue;
    return nextTable;
  }

  const dataRowCount = row - headerRowIndex;
  ensureDataRowCount(nextTable, dataRowCount);
  nextTable.rows[dataRowCount - 1][key] = nextValue;
  return nextTable;
}

export function setHeaderRow(table, rowIndex) {
  const nextTable = cloneLightTable(table);
  const targetHeaderRow = normalizePositiveIndex(rowIndex);
  const columnCount = getColumnCount(nextTable);
  const rowCount = Math.max(getSheetRowCount(nextTable), targetHeaderRow);

  const leadingRows = [];
  for (let currentRow = 1; currentRow < targetHeaderRow; currentRow += 1) {
    leadingRows.push(buildRowRecord(nextTable, currentRow, columnCount));
  }

  const nextColumns = Array.from({ length: columnCount }, (_, offset) => ({
    key: columnKeyForIndex(offset + 1),
    label: getCellValue(nextTable, targetHeaderRow, offset + 1),
  }));

  const rows = [];
  for (let currentRow = targetHeaderRow + 1; currentRow <= rowCount; currentRow += 1) {
    rows.push(buildRowRecord(nextTable, currentRow, columnCount));
  }

  nextTable.leadingRows = leadingRows;
  nextTable.columns = nextColumns;
  nextTable.rows = rows;
  return nextTable;
}

export function setGridLimits(table, gridLimits) {
  const nextTable = cloneLightTable(table);
  const nextLimits = normalizeGridLimits(gridLimits);
  const headerRowIndex = Math.min(getHeaderRowIndex(nextTable), nextLimits.maxRows);
  const matrix = buildSheetMatrix(nextTable, nextLimits.maxRows, nextLimits.maxColumns);
  return rebuildTableFromMatrix(nextTable, matrix, headerRowIndex, nextLimits);
}

export function expandGridLimitsToInclude(table, rowIndex, columnIndex) {
  const nextTable = cloneLightTable(table);
  const nextLimits = normalizeGridLimits({
    maxRows: Math.max(getGridLimits(nextTable).maxRows, normalizePositiveIndex(rowIndex)),
    maxColumns: Math.max(getGridLimits(nextTable).maxColumns, normalizePositiveIndex(columnIndex)),
  });
  nextTable.gridLimits = nextLimits;
  return nextTable;
}

export function applyMatrixToLightTable(table, startCell, matrix) {
  const originRow = normalizePositiveIndex(startCell?.row);
  const originCol = normalizePositiveIndex(startCell?.col);
  const safeMatrix = Array.isArray(matrix) ? matrix : [];
  const rowCount = safeMatrix.length;
  const maxColumnSpan = safeMatrix.reduce(
    (max, row) => Math.max(max, Array.isArray(row) ? row.length : 0),
    0,
  );
  const endRow = rowCount ? originRow + rowCount - 1 : originRow;
  const endCol = maxColumnSpan ? originCol + maxColumnSpan - 1 : originCol;

  if (!rowCount || !maxColumnSpan) {
    return {
      table,
      changed: false,
      endRow,
      endCol,
    };
  }

  const nextTable = cloneLightTable(table);
  let changed = false;
  const previousColumnCount = nextTable.columns.length;
  ensureColumnCount(nextTable, endCol);
  if (nextTable.columns.length !== previousColumnCount) {
    changed = true;
  }

  const previousLimits = getGridLimits(nextTable);
  const nextLimits = normalizeGridLimits({
    maxRows: Math.max(previousLimits.maxRows, endRow),
    maxColumns: Math.max(previousLimits.maxColumns, endCol),
  });
  if (nextLimits.maxRows !== previousLimits.maxRows || nextLimits.maxColumns !== previousLimits.maxColumns) {
    changed = true;
  }
  nextTable.gridLimits = nextLimits;

  const headerRowIndex = getHeaderRowIndex(nextTable);

  safeMatrix.forEach((row, rowOffset) => {
    if (!Array.isArray(row) || !row.length) return;
    const targetRow = originRow + rowOffset;
    row.forEach((value, columnOffset) => {
      const targetColumn = originCol + columnOffset;
      const nextValue = String(value ?? "");
      const key = columnKeyForIndex(targetColumn);

      if (targetRow < headerRowIndex) {
        ensureLeadingRowCount(nextTable, targetRow);
        const record = nextTable.leadingRows[targetRow - 1];
        const previousValue = String(record?.[key] ?? "");
        if (previousValue !== nextValue) {
          changed = true;
        }
        if (nextValue === "") {
          delete record[key];
        } else {
          record[key] = nextValue;
        }
        return;
      }

      if (targetRow === headerRowIndex) {
        const previousValue = String(nextTable.columns[targetColumn - 1]?.label ?? "");
        if (previousValue !== nextValue) {
          changed = true;
        }
        nextTable.columns[targetColumn - 1].label = nextValue;
        return;
      }

      const dataRowCount = targetRow - headerRowIndex;
      ensureDataRowCount(nextTable, dataRowCount);
      const record = nextTable.rows[dataRowCount - 1];
      const previousValue = String(record?.[key] ?? "");
      if (previousValue !== nextValue) {
        changed = true;
      }
      if (nextValue === "") {
        delete record[key];
      } else {
        record[key] = nextValue;
      }
    });
  });

  return {
    table: nextTable,
    changed,
    endRow,
    endCol,
  };
}

export function getLastUsedRowIndex(table) {
  const columnCount = getColumnCount(table);
  const rowCount = getSheetRowCount(table);
  let lastUsedRow = getHeaderRowIndex(table);
  for (let rowIndex = 1; rowIndex <= rowCount; rowIndex += 1) {
    for (let columnIndex = 1; columnIndex <= columnCount; columnIndex += 1) {
      if (getCellValue(table, rowIndex, columnIndex) !== "") {
        lastUsedRow = rowIndex;
      }
    }
  }
  return lastUsedRow;
}

export function getLastUsedColumnIndex(table) {
  const columnCount = getColumnCount(table);
  const rowCount = getSheetRowCount(table);
  let lastUsedColumn = 1;
  for (let rowIndex = 1; rowIndex <= rowCount; rowIndex += 1) {
    for (let columnIndex = 1; columnIndex <= columnCount; columnIndex += 1) {
      if (getCellValue(table, rowIndex, columnIndex) !== "") {
        lastUsedColumn = columnIndex;
      }
    }
  }
  return lastUsedColumn;
}

export function insertRows(table, rowIndex, count = 1, placement = "above") {
  const nextTable = cloneLightTable(table);
  const rowSpan = normalizePositiveIndex(count);
  const anchorRow = normalizePositiveIndex(rowIndex);
  const limits = getGridLimits(nextTable);
  const columnCount = Math.max(1, limits.maxColumns, getColumnCount(nextTable));
  const rowCount = Math.max(limits.maxRows, getSheetRowCount(nextTable), anchorRow);
  const matrix = buildSheetMatrix(nextTable, rowCount, columnCount);
  const insertAt = placement === "below" ? anchorRow : anchorRow - 1;
  const blankRows = Array.from({ length: rowSpan }, () => Array.from({ length: columnCount }, () => ""));
  matrix.splice(insertAt, 0, ...blankRows);
  const nextHeaderRowIndex = insertAt < getHeaderRowIndex(nextTable)
    ? getHeaderRowIndex(nextTable) + rowSpan
    : getHeaderRowIndex(nextTable);
  return rebuildTableFromMatrix(nextTable, matrix, nextHeaderRowIndex, {
    maxRows: Math.max(limits.maxRows, matrix.length),
    maxColumns: Math.max(limits.maxColumns, columnCount),
  });
}

export function insertColumns(table, columnIndex, count = 1, placement = "left") {
  const nextTable = cloneLightTable(table);
  const columnSpan = normalizePositiveIndex(count);
  const anchorColumn = normalizePositiveIndex(columnIndex);
  const limits = getGridLimits(nextTable);
  const columnCount = Math.max(limits.maxColumns, getColumnCount(nextTable), anchorColumn);
  const rowCount = Math.max(1, limits.maxRows, getSheetRowCount(nextTable));
  const matrix = buildSheetMatrix(nextTable, rowCount, columnCount);
  const insertAt = placement === "right" ? anchorColumn : anchorColumn - 1;
  matrix.forEach((row) => {
    row.splice(insertAt, 0, ...Array.from({ length: columnSpan }, () => ""));
  });
  return rebuildTableFromMatrix(nextTable, matrix, getHeaderRowIndex(nextTable), {
    maxRows: Math.max(limits.maxRows, rowCount),
    maxColumns: Math.max(limits.maxColumns, columnCount + columnSpan),
  });
}

export function deleteRows(table, rowIndex, count = 1) {
  const nextTable = cloneLightTable(table);
  const rowSpan = normalizePositiveIndex(count);
  const startRow = normalizePositiveIndex(rowIndex);
  const limits = getGridLimits(nextTable);
  const columnCount = Math.max(1, limits.maxColumns, getColumnCount(nextTable));
  const rowCount = Math.max(1, limits.maxRows, getSheetRowCount(nextTable));
  const matrix = buildSheetMatrix(nextTable, rowCount, columnCount);
  const deleteAt = Math.min(startRow - 1, Math.max(0, matrix.length - 1));
  const actualDeletedRowCount = Math.min(rowSpan, matrix.length - deleteAt);

  if (actualDeletedRowCount <= 0) {
    return nextTable;
  }

  matrix.splice(deleteAt, actualDeletedRowCount);
  if (!matrix.length) {
    matrix.push(Array.from({ length: columnCount }, () => ""));
  }

  const headerRowIndex = getHeaderRowIndex(nextTable);
  const deletedBeforeHeader = Math.max(0, Math.min(actualDeletedRowCount, headerRowIndex - startRow));
  const nextHeaderRowIndex = Math.max(1, Math.min(
    matrix.length,
    headerRowIndex - deletedBeforeHeader,
  ));

  return rebuildTableFromMatrix(nextTable, matrix, nextHeaderRowIndex, {
    maxRows: Math.max(1, limits.maxRows - actualDeletedRowCount),
    maxColumns: Math.max(limits.maxColumns, columnCount),
  });
}

export function deleteColumns(table, columnIndex, count = 1) {
  const nextTable = cloneLightTable(table);
  const columnSpan = normalizePositiveIndex(count);
  const startColumn = normalizePositiveIndex(columnIndex);
  const limits = getGridLimits(nextTable);
  const columnCount = Math.max(1, limits.maxColumns, getColumnCount(nextTable));
  const rowCount = Math.max(1, limits.maxRows, getSheetRowCount(nextTable));
  const matrix = buildSheetMatrix(nextTable, rowCount, columnCount);
  const deleteAt = Math.min(startColumn - 1, Math.max(0, columnCount - 1));
  const actualDeletedColumnCount = Math.min(columnSpan, columnCount - deleteAt);

  if (actualDeletedColumnCount <= 0) {
    return nextTable;
  }

  matrix.forEach((row) => {
    row.splice(deleteAt, actualDeletedColumnCount);
    if (!row.length) {
      row.push("");
    }
  });

  return rebuildTableFromMatrix(nextTable, matrix, getHeaderRowIndex(nextTable), {
    maxRows: Math.max(limits.maxRows, rowCount),
    maxColumns: Math.max(1, limits.maxColumns - actualDeletedColumnCount),
  });
}

export function insertCells(table, bounds, direction = "right") {
  const nextTable = cloneLightTable(table);
  const normalizedBounds = normalizeBounds(bounds);
  const width = normalizedBounds.endCol - normalizedBounds.startCol + 1;
  const height = normalizedBounds.endRow - normalizedBounds.startRow + 1;
  const limits = getGridLimits(nextTable);
  const columnCount = Math.max(limits.maxColumns, getColumnCount(nextTable), normalizedBounds.endCol);
  const rowCount = Math.max(limits.maxRows, getSheetRowCount(nextTable), normalizedBounds.endRow);
  const matrix = buildSheetMatrix(nextTable, rowCount, columnCount);

  if (direction === "down") {
    const shifted = matrix.map((row) => [...row]);
    for (let rowIndex = 0; rowIndex < height; rowIndex += 1) {
      shifted.push(Array.from({ length: columnCount }, () => ""));
    }
    for (let columnIndex = normalizedBounds.startCol; columnIndex <= normalizedBounds.endCol; columnIndex += 1) {
      for (let rowIndex = rowCount; rowIndex >= normalizedBounds.startRow; rowIndex -= 1) {
        shifted[rowIndex - 1 + height][columnIndex - 1] = matrix[rowIndex - 1][columnIndex - 1];
      }
      for (let rowIndex = normalizedBounds.startRow; rowIndex <= normalizedBounds.endRow; rowIndex += 1) {
        shifted[rowIndex - 1][columnIndex - 1] = "";
      }
    }
    return rebuildTableFromMatrix(nextTable, shifted, getHeaderRowIndex(nextTable), {
      maxRows: Math.max(limits.maxRows, rowCount + height),
      maxColumns: Math.max(limits.maxColumns, columnCount),
    });
  }

  const shifted = matrix.map((row) => [...row]);
  for (let rowIndex = normalizedBounds.startRow; rowIndex <= normalizedBounds.endRow; rowIndex += 1) {
    shifted[rowIndex - 1].splice(normalizedBounds.startCol - 1, 0, ...Array.from({ length: width }, () => ""));
  }
  return rebuildTableFromMatrix(nextTable, shifted, getHeaderRowIndex(nextTable), {
    maxRows: Math.max(limits.maxRows, rowCount),
    maxColumns: Math.max(limits.maxColumns, columnCount + width),
  });
}

export function setPageSetup(table, pageSetup) {
  const nextTable = cloneLightTable(table);
  nextTable.pageSetup = JSON.parse(JSON.stringify(pageSetup ?? {}));
  return nextTable;
}

function buildRowRecord(table, rowIndex, columnCount) {
  const record = {};
  for (let columnIndex = 1; columnIndex <= columnCount; columnIndex += 1) {
    record[columnKeyForIndex(columnIndex)] = getCellValue(table, rowIndex, columnIndex);
  }
  return record;
}

function ensureColumnCount(table, columnCount) {
  while (table.columns.length < columnCount) {
    const nextIndex = table.columns.length + 1;
    table.columns.push({
      key: columnKeyForIndex(nextIndex),
      label: "",
    });
  }
}

function ensureLeadingRowCount(table, rowCount) {
  while (table.leadingRows.length < rowCount) {
    table.leadingRows.push({});
  }
}

function ensureDataRowCount(table, rowCount) {
  while (table.rows.length < rowCount) {
    table.rows.push({});
  }
}

function buildSheetMatrix(table, rowCount, columnCount) {
  const rows = [];
  for (let rowIndex = 1; rowIndex <= rowCount; rowIndex += 1) {
    const row = [];
    for (let columnIndex = 1; columnIndex <= columnCount; columnIndex += 1) {
      row.push(getCellValue(table, rowIndex, columnIndex));
    }
    rows.push(row);
  }
  return rows;
}

function rebuildTableFromMatrix(table, matrix, headerRowIndex, gridLimits) {
  const nextTable = cloneLightTable(table);
  const safeHeaderRowIndex = Math.max(1, Math.min(matrix.length || 1, normalizePositiveIndex(headerRowIndex)));
  const materializedBounds = findMaterializedBounds(matrix, safeHeaderRowIndex);
  const columnCount = materializedBounds.lastUsedColumn;
  const rowCount = materializedBounds.lastUsedRow;

  nextTable.columns = Array.from({ length: columnCount }, (_, offset) => ({
    key: columnKeyForIndex(offset + 1),
    label: String(matrix[safeHeaderRowIndex - 1]?.[offset] ?? ""),
  }));
  nextTable.leadingRows = [];
  for (let rowIndex = 1; rowIndex < safeHeaderRowIndex; rowIndex += 1) {
    nextTable.leadingRows.push(buildRecordFromMatrixRow(matrix[rowIndex - 1], columnCount));
  }
  nextTable.rows = [];
  for (let rowIndex = safeHeaderRowIndex + 1; rowIndex <= rowCount; rowIndex += 1) {
    nextTable.rows.push(buildRecordFromMatrixRow(matrix[rowIndex - 1], columnCount));
  }
  nextTable.gridLimits = normalizeGridLimits({
    maxRows: Math.max(gridLimits?.maxRows ?? DEFAULT_GRID_LIMIT_ROW_COUNT, safeHeaderRowIndex),
    maxColumns: Math.max(gridLimits?.maxColumns ?? DEFAULT_GRID_LIMIT_COLUMN_COUNT, 1),
  });
  return nextTable;
}

function buildRecordFromMatrixRow(row, columnCount) {
  const record = {};
  for (let columnIndex = 1; columnIndex <= columnCount; columnIndex += 1) {
    const value = String(row?.[columnIndex - 1] ?? "");
    if (value !== "") {
      record[columnKeyForIndex(columnIndex)] = value;
    }
  }
  return record;
}

function findMaterializedBounds(matrix, headerRowIndex) {
  let lastUsedRow = Math.max(1, normalizePositiveIndex(headerRowIndex));
  let lastUsedColumn = 1;
  matrix.forEach((row, rowOffset) => {
    row.forEach((value, columnOffset) => {
      if (String(value ?? "") !== "") {
        lastUsedRow = Math.max(lastUsedRow, rowOffset + 1);
        lastUsedColumn = Math.max(lastUsedColumn, columnOffset + 1);
      }
    });
  });
  return {
    lastUsedRow,
    lastUsedColumn,
  };
}

function normalizeGridLimits(gridLimits) {
  const maxRows = Math.max(1, Math.trunc(Number(gridLimits?.maxRows) || DEFAULT_GRID_LIMIT_ROW_COUNT));
  const maxColumns = Math.max(1, Math.trunc(Number(gridLimits?.maxColumns) || DEFAULT_GRID_LIMIT_COLUMN_COUNT));
  return { maxRows, maxColumns };
}

function normalizeBounds(bounds) {
  const startRow = normalizePositiveIndex(bounds?.startRow);
  const endRow = normalizePositiveIndex(bounds?.endRow ?? startRow);
  const startCol = normalizePositiveIndex(bounds?.startCol);
  const endCol = normalizePositiveIndex(bounds?.endCol ?? startCol);
  return {
    startRow: Math.min(startRow, endRow),
    endRow: Math.max(startRow, endRow),
    startCol: Math.min(startCol, endCol),
    endCol: Math.max(startCol, endCol),
  };
}

function normalizePositiveIndex(value) {
  return Math.max(1, Math.trunc(Number(value) || 1));
}

function normalizeNonNegativeInteger(value) {
  return Math.max(0, Math.trunc(Number(value) || 0));
}
