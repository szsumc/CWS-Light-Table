import {
  cloneLightTable,
  deleteRows,
  getCellValue,
  setCellValue,
} from "./table-model.js";

const SPACE_PATTERN = /[ \t\u3000]/g;
const LEADING_SPACE_PATTERN = /^[ \t\u3000]+/;
const TRAILING_SPACE_PATTERN = /[ \t\u3000]+$/;

export function removeDuplicateRows(table, options = {}) {
  const rows = normalizeIndexList(options.rows);
  const columns = normalizeIndexList(options.columns);
  const seen = new Set();
  const duplicateRows = [];

  if (!rows.length || !columns.length) {
    return {
      table,
      changed: false,
      checkedRowCount: rows.length,
      removedRowCount: 0,
      duplicateRows,
    };
  }

  rows.forEach((rowIndex) => {
    const key = JSON.stringify(columns.map((columnIndex) => getCellValue(table, rowIndex, columnIndex)));
    if (seen.has(key)) {
      duplicateRows.push(rowIndex);
      return;
    }
    seen.add(key);
  });

  if (!duplicateRows.length) {
    return {
      table,
      changed: false,
      checkedRowCount: rows.length,
      removedRowCount: 0,
      duplicateRows,
    };
  }

  let nextTable = table;
  [...duplicateRows].sort((left, right) => right - left).forEach((rowIndex) => {
    nextTable = deleteRows(nextTable, rowIndex, 1);
  });

  return {
    table: nextTable,
    changed: true,
    checkedRowCount: rows.length,
    removedRowCount: duplicateRows.length,
    duplicateRows,
  };
}

export function previewDuplicateRows(table, options = {}) {
  const result = removeDuplicateRows(table, options);
  return {
    checkedRowCount: result.checkedRowCount,
    removedRowCount: result.removedRowCount,
    duplicateRows: result.duplicateRows,
  };
}

export function removeBlankRows(table, options = {}) {
  const rows = normalizeIndexList(options.rows);
  const columns = normalizeIndexList(options.columns);
  const blankRows = rows.filter((rowIndex) => isBlankRow(table, rowIndex, columns));

  if (!rows.length || !columns.length || !blankRows.length) {
    return {
      table,
      changed: false,
      checkedRowCount: rows.length,
      removedRowCount: 0,
      blankRows,
    };
  }

  let nextTable = table;
  [...blankRows].sort((left, right) => right - left).forEach((rowIndex) => {
    nextTable = deleteRows(nextTable, rowIndex, 1);
  });

  return {
    table: nextTable,
    changed: true,
    checkedRowCount: rows.length,
    removedRowCount: blankRows.length,
    blankRows,
  };
}

export function previewBlankRows(table, options = {}) {
  const result = removeBlankRows(table, options);
  return {
    checkedRowCount: result.checkedRowCount,
    removedRowCount: result.removedRowCount,
    blankRows: result.blankRows,
  };
}

export function trimSpacesInCells(table, options = {}) {
  const rows = normalizeIndexList(options.rows);
  const columns = normalizeIndexList(options.columns);
  const trimOptions = normalizeTrimOptions(options);
  let nextTable = cloneLightTable(table);
  let changedCellCount = 0;
  let checkedCellCount = 0;

  rows.forEach((rowIndex) => {
    columns.forEach((columnIndex) => {
      checkedCellCount += 1;
      const previousValue = getCellValue(nextTable, rowIndex, columnIndex);
      const nextValue = cleanCellSpaces(previousValue, trimOptions);
      if (nextValue === previousValue) return;
      nextTable = setCellValue(nextTable, rowIndex, columnIndex, nextValue);
      changedCellCount += 1;
    });
  });

  return {
    table: nextTable,
    changed: changedCellCount > 0,
    checkedCellCount,
    changedCellCount,
  };
}

function isBlankRow(table, rowIndex, columns) {
  if (!columns.length) return false;
  return columns.every((columnIndex) => getCellValue(table, rowIndex, columnIndex) === "");
}

export function previewTrimSpaces(table, options = {}) {
  const result = trimSpacesInCells(table, options);
  return {
    checkedCellCount: result.checkedCellCount,
    changedCellCount: result.changedCellCount,
  };
}

export function normalizeCaseInCells(table, options = {}) {
  const rows = normalizeIndexList(options.rows);
  const columns = normalizeIndexList(options.columns);
  const mode = normalizeCaseMode(options.mode);
  let nextTable = cloneLightTable(table);
  let changedCellCount = 0;
  let checkedCellCount = 0;

  rows.forEach((rowIndex) => {
    columns.forEach((columnIndex) => {
      checkedCellCount += 1;
      const previousValue = getCellValue(nextTable, rowIndex, columnIndex);
      const nextValue = normalizeCellCase(previousValue, mode);
      if (nextValue === previousValue) return;
      nextTable = setCellValue(nextTable, rowIndex, columnIndex, nextValue);
      changedCellCount += 1;
    });
  });

  return {
    table: nextTable,
    changed: changedCellCount > 0,
    checkedCellCount,
    changedCellCount,
  };
}

export function previewNormalizeCase(table, options = {}) {
  const result = normalizeCaseInCells(table, options);
  return {
    checkedCellCount: result.checkedCellCount,
    changedCellCount: result.changedCellCount,
  };
}

export function formatNumberCommasInCells(table, options = {}) {
  const rows = normalizeIndexList(options.rows);
  const columns = normalizeIndexList(options.columns);
  const mode = normalizeNumberCommaMode(options.mode);
  let nextTable = cloneLightTable(table);
  let changedCellCount = 0;
  let numericCellCount = 0;
  let checkedCellCount = 0;

  rows.forEach((rowIndex) => {
    columns.forEach((columnIndex) => {
      checkedCellCount += 1;
      const previousValue = getCellValue(nextTable, rowIndex, columnIndex);
      const nextValue = formatNumberCommas(previousValue, mode);
      if (nextValue == null) return;
      numericCellCount += 1;
      if (nextValue === previousValue) return;
      nextTable = setCellValue(nextTable, rowIndex, columnIndex, nextValue);
      changedCellCount += 1;
    });
  });

  return {
    table: nextTable,
    changed: changedCellCount > 0,
    checkedCellCount,
    numericCellCount,
    changedCellCount,
  };
}

export function previewNumberCommas(table, options = {}) {
  const result = formatNumberCommasInCells(table, options);
  return {
    checkedCellCount: result.checkedCellCount,
    numericCellCount: result.numericCellCount,
    changedCellCount: result.changedCellCount,
  };
}

export function findEncodingIssuesInCells(table, options = {}) {
  const rows = normalizeIndexList(options.rows);
  const columns = normalizeIndexList(options.columns);
  const limit = Math.max(1, Math.trunc(Number(options.limit) || 1000));
  const issues = [];
  let issueCount = 0;
  let checkedCellCount = 0;

  rows.forEach((rowIndex) => {
    columns.forEach((columnIndex) => {
      checkedCellCount += 1;
      const value = getCellValue(table, rowIndex, columnIndex);
      if (!value) return;
      const issueTypes = detectEncodingIssueTypes(value);
      if (!issueTypes.length) return;
      issueCount += 1;
      if (issues.length < limit) {
        issues.push({
          row: rowIndex,
          col: columnIndex,
          types: issueTypes,
          snippet: createIssueSnippet(value),
        });
      }
    });
  });

  return {
    checkedCellCount,
    issueCount,
    issues,
    truncated: issueCount > issues.length,
  };
}

export function cleanCellSpaces(value, options = {}) {
  const source = String(value ?? "");
  const trimOptions = normalizeTrimOptions(options);
  const leadingMatch = source.match(LEADING_SPACE_PATTERN);
  const trailingMatch = source.match(TRAILING_SPACE_PATTERN);
  const leading = leadingMatch?.[0] || "";
  const trailing = trailingMatch?.[0] || "";
  const coreEnd = trailing ? source.length - trailing.length : source.length;
  let core = source.slice(leading.length, coreEnd);

  if (trimOptions.inner) {
    core = trimOptions.innerMode === "remove"
      ? core.replace(SPACE_PATTERN, "")
      : core.replace(/[ \t\u3000]+/g, " ");
  }

  return `${trimOptions.leading ? "" : leading}${core}${trimOptions.trailing ? "" : trailing}`;
}

export function normalizeCellCase(value, mode = "lower") {
  const source = String(value ?? "");
  const normalizedMode = normalizeCaseMode(mode);
  if (normalizedMode === "upper") {
    return source.toUpperCase();
  }
  if (normalizedMode === "capitalize-words") {
    return capitalizeWords(source);
  }
  if (normalizedMode === "full-width") {
    return convertToFullWidth(source);
  }
  if (normalizedMode === "half-width") {
    return convertToHalfWidth(source);
  }
  return source.toLowerCase();
}

export function formatNumberCommas(value, mode = "add") {
  const source = String(value ?? "");
  const parsed = parseStrictNumericString(source);
  if (!parsed) {
    return null;
  }
  const normalizedMode = normalizeNumberCommaMode(mode);
  const integerWithoutCommas = parsed.integer.replace(/,/g, "");
  const integer = normalizedMode === "remove"
    ? integerWithoutCommas
    : addThousandsSeparators(integerWithoutCommas);
  return `${parsed.leading}${parsed.sign}${integer}${parsed.decimal}${parsed.trailing}`;
}

function capitalizeWords(value) {
  const lower = String(value ?? "").toLowerCase();
  let result = "";
  let shouldCapitalizeNextLetter = true;
  for (const char of lower) {
    if (isLetter(char)) {
      result += shouldCapitalizeNextLetter ? char.toUpperCase() : char;
      shouldCapitalizeNextLetter = false;
      continue;
    }
    result += char;
    if (isWordConnector(char)) {
      continue;
    }
    shouldCapitalizeNextLetter = !isDigit(char);
  }
  return result;
}

function isLetter(char) {
  return char.toLowerCase() !== char.toUpperCase();
}

function isDigit(char) {
  return /[0-9]/.test(char);
}

function isWordConnector(char) {
  return char === "'" || char === "’";
}

function convertToFullWidth(value) {
  let result = "";
  const text = String(value ?? "");
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const codePoint = char.codePointAt(0);
    if (codePoint === 0x20) {
      result += String.fromCodePoint(0x3000);
    } else if (codePoint >= 0x21 && codePoint <= 0x7e) {
      result += String.fromCodePoint(codePoint + 0xfee0);
    } else if (isHalfWidthKanaCodePoint(codePoint)) {
      let kanaSequence = char;
      const nextChar = text[index + 1] || "";
      if (isHalfWidthKanaMark(nextChar)) {
        kanaSequence += nextChar;
        index += 1;
      }
      result += kanaSequence.normalize("NFKC");
    } else {
      result += char;
    }
  }
  return result;
}

function convertToHalfWidth(value) {
  let result = "";
  for (const char of String(value ?? "")) {
    const codePoint = char.codePointAt(0);
    if (codePoint === 0x3000) {
      result += " ";
    } else if (codePoint >= 0xff01 && codePoint <= 0xff5e) {
      result += String.fromCodePoint(codePoint - 0xfee0);
    } else {
      result += convertKanaToHalfWidth(char);
    }
  }
  return result;
}

const HALF_WIDTH_KANA_BY_FULL_WIDTH = new Map([
  ["\u3002", "\uff61"],
  ["\u300c", "\uff62"],
  ["\u300d", "\uff63"],
  ["\u3001", "\uff64"],
  ["\u30fb", "\uff65"],
  ["\u30f2", "\uff66"],
  ["\u30a1", "\uff67"],
  ["\u30a3", "\uff68"],
  ["\u30a5", "\uff69"],
  ["\u30a7", "\uff6a"],
  ["\u30a9", "\uff6b"],
  ["\u30e3", "\uff6c"],
  ["\u30e5", "\uff6d"],
  ["\u30e7", "\uff6e"],
  ["\u30c3", "\uff6f"],
  ["\u30fc", "\uff70"],
  ["\u30a2", "\uff71"],
  ["\u30a4", "\uff72"],
  ["\u30a6", "\uff73"],
  ["\u30a8", "\uff74"],
  ["\u30aa", "\uff75"],
  ["\u30ab", "\uff76"],
  ["\u30ad", "\uff77"],
  ["\u30af", "\uff78"],
  ["\u30b1", "\uff79"],
  ["\u30b3", "\uff7a"],
  ["\u30b5", "\uff7b"],
  ["\u30b7", "\uff7c"],
  ["\u30b9", "\uff7d"],
  ["\u30bb", "\uff7e"],
  ["\u30bd", "\uff7f"],
  ["\u30bf", "\uff80"],
  ["\u30c1", "\uff81"],
  ["\u30c4", "\uff82"],
  ["\u30c6", "\uff83"],
  ["\u30c8", "\uff84"],
  ["\u30ca", "\uff85"],
  ["\u30cb", "\uff86"],
  ["\u30cc", "\uff87"],
  ["\u30cd", "\uff88"],
  ["\u30ce", "\uff89"],
  ["\u30cf", "\uff8a"],
  ["\u30d2", "\uff8b"],
  ["\u30d5", "\uff8c"],
  ["\u30d8", "\uff8d"],
  ["\u30db", "\uff8e"],
  ["\u30de", "\uff8f"],
  ["\u30df", "\uff90"],
  ["\u30e0", "\uff91"],
  ["\u30e1", "\uff92"],
  ["\u30e2", "\uff93"],
  ["\u30e4", "\uff94"],
  ["\u30e6", "\uff95"],
  ["\u30e8", "\uff96"],
  ["\u30e9", "\uff97"],
  ["\u30ea", "\uff98"],
  ["\u30eb", "\uff99"],
  ["\u30ec", "\uff9a"],
  ["\u30ed", "\uff9b"],
  ["\u30ef", "\uff9c"],
  ["\u30f3", "\uff9d"],
  ["\u309b", "\uff9e"],
  ["\u309c", "\uff9f"],
]);

function isHalfWidthKanaCodePoint(codePoint) {
  return codePoint >= 0xff61 && codePoint <= 0xff9f;
}

function isHalfWidthKanaMark(char) {
  return char === "\uff9e" || char === "\uff9f";
}

function convertKanaToHalfWidth(char) {
  const direct = HALF_WIDTH_KANA_BY_FULL_WIDTH.get(char);
  if (direct) {
    return direct;
  }
  const decomposed = char.normalize("NFD");
  if (decomposed.length > 1) {
    let converted = "";
    for (const part of decomposed) {
      if (part === "\u3099") {
        converted += "\uff9e";
      } else if (part === "\u309a") {
        converted += "\uff9f";
      } else {
        const halfWidth = HALF_WIDTH_KANA_BY_FULL_WIDTH.get(part);
        if (!halfWidth) return char;
        converted += halfWidth;
      }
    }
    return converted;
  }
  if (char === "\u3099") {
    return "\uff9e";
  }
  if (char === "\u309a") {
    return "\uff9f";
  }
  return char;
}

function normalizeTrimOptions(options = {}) {
  return {
    leading: Boolean(options.leading),
    trailing: Boolean(options.trailing),
    inner: Boolean(options.inner),
    innerMode: options.innerMode === "remove" ? "remove" : "collapse",
  };
}

function normalizeCaseMode(mode) {
  return ["lower", "upper", "capitalize-words", "full-width", "half-width"].includes(mode) ? mode : "lower";
}

function normalizeNumberCommaMode(mode) {
  return mode === "remove" ? "remove" : "add";
}

function parseStrictNumericString(value) {
  const source = String(value ?? "");
  const match = source.match(/^([ \t\u3000]*)([+-]?)([^ \t\u3000]+)([ \t\u3000]*)$/);
  if (!match) return null;
  const [, leading, sign, body, trailing] = match;
  const numericMatch = body.match(/^([^.,]+(?:,[^.,]+)*|\d+)(\.\d+)?$/);
  if (!numericMatch) return null;
  const integer = numericMatch[1];
  const decimal = numericMatch[2] || "";
  if (integer.includes(",")) {
    if (!/^[1-9]\d{0,2}(?:,\d{3})+$/.test(integer)) {
      return null;
    }
  } else if (!/^(?:0|[1-9]\d*)$/.test(integer)) {
    return null;
  }
  return {
    leading,
    sign,
    integer,
    decimal,
    trailing,
  };
}

function addThousandsSeparators(value) {
  return String(value || "0").replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function detectEncodingIssueTypes(value) {
  const source = String(value ?? "");
  const types = [];
  if (/\uFFFD/.test(source)) {
    types.push("Replacement character");
  }
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/.test(source)) {
    types.push("Control character");
  }
  if (/[\u0080-\u009F]/.test(source)
    || /(?:[\u00C2\u00C3][\u0080-\u00BF\u00A0-\u00FF]|[\u00E3\u00E6-\u00E9][\u0080-\u00BF])/.test(source)
    || /(?:\u7E3A|\u7E67|\u8B41|\u839F|\u8373|\u8708|\u9A65|\u9AEF)/.test(source)) {
    types.push("Suspicious mojibake sequence");
  }
  return types;
}

function createIssueSnippet(value) {
  const normalized = String(value ?? "").replace(/\s+/g, " ").trim();
  const characters = Array.from(normalized);
  if (characters.length <= 80) {
    return normalized;
  }
  return `${characters.slice(0, 77).join("")}...`;
}

function normalizeIndexList(values) {
  if (!Array.isArray(values)) return [];
  const unique = new Set();
  values.forEach((value) => {
    const normalized = Math.max(1, Math.trunc(Number(value) || 0));
    if (normalized > 0) {
      unique.add(normalized);
    }
  });
  return [...unique].sort((left, right) => left - right);
}
