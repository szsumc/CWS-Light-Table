import { parseDelimitedText } from "./clipboard.js";
import { getSelectionBounds } from "./editor-state.js";

const TEXT_ENCODINGS = new Map([
  ["UTF-8", "utf-8"],
  ["Shift-JIS", "shift_jis"],
  ["UTF-16LE", "utf-16le"],
]);

export function inferImportFileKind(fileName = "") {
  const normalized = String(fileName || "").trim().toLowerCase();
  if (normalized.endsWith(".xlsx")) return "xlsx";
  if (normalized.endsWith(".xlsm")) return "xlsm";
  if (normalized.endsWith(".xls")) return "xls";
  if (normalized.endsWith(".csv")) return "csv";
  if (normalized.endsWith(".tsv")) return "tsv";
  if (normalized.endsWith(".txt")) return "txt";
  if (normalized.endsWith(".json")) return "json";
  if (normalized.endsWith(".xml")) return "xml";
  return "";
}

export function isExcelImportKind(fileKind) {
  return fileKind === "xlsx" || fileKind === "xlsm" || fileKind === "xls";
}

export function isTextImportKind(fileKind) {
  return fileKind === "csv"
    || fileKind === "tsv"
    || fileKind === "txt"
    || fileKind === "json"
    || fileKind === "xml";
}

export function getSupportedTextImportEncodings() {
  return Array.from(TEXT_ENCODINGS.keys());
}

export function normalizeTextImportEncoding(encoding = "UTF-8") {
  const normalized = String(encoding || "").trim();
  return TEXT_ENCODINGS.get(normalized) || TEXT_ENCODINGS.get("UTF-8");
}

export function inferDefaultTextImportEncoding(bytes) {
  const source = toUint8Array(bytes);
  if (source[0] === 0xef && source[1] === 0xbb && source[2] === 0xbf) {
    return "UTF-8";
  }
  if (source[0] === 0xff && source[1] === 0xfe) {
    return "UTF-16LE";
  }
  return "UTF-8";
}

export function inferDefaultDelimiterMode(fileKind) {
  if (fileKind === "csv") return "comma";
  if (fileKind === "tsv") return "tab";
  return "line";
}

export function decodeTextImportBytes(bytes, encoding = "UTF-8") {
  const source = stripKnownBom(toUint8Array(bytes));
  const decoder = new TextDecoder(normalizeTextImportEncoding(encoding));
  return decoder.decode(source);
}

export function parseTextImportContent(options) {
  const fileName = String(options?.fileName || "");
  const fileKind = inferImportFileKind(fileName);
  if (!isTextImportKind(fileKind)) {
    throw new Error("The selected file is not a supported text or structured-data import type.");
  }
  const encoding = String(options?.encoding || inferDefaultTextImportEncoding(options?.bytes));
  const text = decodeTextImportBytes(options?.bytes, encoding);
  const delimiterMode = resolveDelimiterMode(fileKind, options?.delimiterMode);
  const matrix = parseTextImportText({
    text,
    fileKind,
    delimiterMode,
  });
  return {
    fileKind,
    encoding,
    delimiterMode,
    text,
    matrix,
  };
}

export function parseTextImportText(options) {
  const fileKind = String(options?.fileKind || "");
  const text = String(options?.text || "");
  const delimiterMode = resolveDelimiterMode(fileKind, options?.delimiterMode);

  if (fileKind === "json") {
    return jsonValueToMatrix(JSON.parse(text));
  }
  if (fileKind === "xml") {
    return xmlStringToMatrix(text);
  }
  if (delimiterMode === "line") {
    return parseLineBasedText(text);
  }
  return parseDelimitedText(text, {
    delimiter: delimiterMode === "comma" ? "," : "\t",
  });
}

export function parseLineBasedText(text) {
  const source = String(text || "");
  if (source === "") return [];
  const rows = source
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .map((line) => [line]);
  if (rows.length > 1 && rows.at(-1)?.[0] === "" && /\r?\n$/.test(source)) {
    return rows.slice(0, -1);
  }
  return rows;
}

export function jsonValueToMatrix(value) {
  const rows = [["Path", "Value"]];
  appendJsonRows("$", value, rows);
  if (rows.length === 1) {
    rows.push(["$", ""]);
  }
  return rows;
}

export function xmlStringToMatrix(source) {
  const text = String(source || "").replace(/^\uFEFF/, "");
  const tokenPattern = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<[^>]+>|[^<]+/g;
  const tokens = text.match(tokenPattern) || [];
  const rows = [["Path", "Value"]];
  const stack = [];

  for (const token of tokens) {
    if (!token || /^<\?/.test(token) || /^<!--/.test(token)) {
      continue;
    }

    if (/^<!\[CDATA\[/.test(token)) {
      const value = decodeXmlEntities(token.slice(9, -3));
      if (value.trim() !== "" && stack.length) {
        rows.push([buildXmlPath(stack), value]);
      }
      continue;
    }

    if (token.startsWith("</")) {
      const closingName = token.slice(2, -1).trim();
      const activeNode = stack.at(-1);
      if (!activeNode || activeNode.name !== closingName) {
        throw new Error(`The XML structure is invalid near </${closingName}>.`);
      }
      stack.pop();
      continue;
    }

    if (token.startsWith("<")) {
      const selfClosing = /\/>$/.test(token);
      const content = token.slice(1, selfClosing ? -2 : -1).trim();
      if (!content) continue;
      const spaceIndex = content.search(/\s/);
      const name = spaceIndex === -1 ? content : content.slice(0, spaceIndex);
      const attributeText = spaceIndex === -1 ? "" : content.slice(spaceIndex + 1);
      if (!name) {
        throw new Error("The XML structure is invalid.");
      }
      stack.push({ name });
      const path = buildXmlPath(stack);
      const attributes = parseXmlAttributes(attributeText);
      attributes.forEach(([attributeName, attributeValue]) => {
        rows.push([`${path}/@${attributeName}`, attributeValue]);
      });
      if (selfClosing) {
        if (!attributes.length) {
          rows.push([path, ""]);
        }
        stack.pop();
      }
      continue;
    }

    const value = decodeXmlEntities(token);
    if (value.trim() !== "" && stack.length) {
      rows.push([buildXmlPath(stack), value.trim()]);
    }
  }

  if (stack.length) {
    throw new Error(`The XML structure is invalid near <${stack.at(-1)?.name || ""}>.`);
  }
  if (rows.length === 1) {
    rows.push(["/", ""]);
  }
  return rows;
}

export function getImportAnchorCell(selection) {
  if (!selection?.anchor && !selection?.focus) {
    return { row: 1, col: 1 };
  }
  const bounds = getSelectionBounds(selection);
  return {
    row: bounds.startRow,
    col: bounds.startCol,
  };
}

export function listExcelWorkbookSheets(workbook) {
  return Array.from(workbook?.SheetNames || []).map((sheetName, index) => {
    const sheet = workbook?.Sheets?.[sheetName];
    const range = decodeExcelRange(sheet?.["!ref"] || "A1");
    return {
      index,
      name: sheetName,
      rowCount: range.endRow - range.startRow + 1,
      colCount: range.endCol - range.startCol + 1,
    };
  });
}

export function excelWorkbookToMatrix(workbook, sheetRef = undefined) {
  const sheetName = resolveExcelSheetName(workbook, sheetRef);
  const sheet = workbook?.Sheets?.[sheetName];
  if (!sheet) {
    throw new Error("The Excel workbook does not contain the requested sheet.");
  }
  const range = decodeExcelRange(sheet["!ref"] || "A1");
  const matrix = [];

  for (let rowIndex = range.startRow; rowIndex <= range.endRow; rowIndex += 1) {
    const row = [];
    for (let columnIndex = range.startCol; columnIndex <= range.endCol; columnIndex += 1) {
      row.push(readExcelSheetCell(sheet, rowIndex, columnIndex));
    }
    matrix.push(row);
  }

  return {
    sheetName,
    matrix,
    rowCount: matrix.length,
    colCount: matrix.reduce((max, row) => Math.max(max, row.length), 0),
  };
}

function resolveDelimiterMode(fileKind, delimiterMode) {
  const normalized = String(delimiterMode || "").trim();
  if (normalized === "comma" || normalized === "tab" || normalized === "line") {
    return normalized;
  }
  return inferDefaultDelimiterMode(fileKind);
}

function stripKnownBom(bytes) {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return bytes.slice(3);
  }
  if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    return bytes.slice(2);
  }
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    return bytes.slice(2);
  }
  return bytes;
}

function toUint8Array(bytes) {
  if (bytes instanceof Uint8Array) return bytes;
  if (bytes instanceof ArrayBuffer) return new Uint8Array(bytes);
  if (ArrayBuffer.isView(bytes)) return new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return new Uint8Array(0);
}

function appendJsonRows(path, value, rows) {
  if (Array.isArray(value)) {
    if (!value.length) {
      rows.push([path, "[]"]);
      return;
    }
    value.forEach((entry, index) => appendJsonRows(`${path}[${index}]`, entry, rows));
    return;
  }

  if (value && typeof value === "object") {
    const entries = Object.entries(value);
    if (!entries.length) {
      rows.push([path, "{}"]);
      return;
    }
    entries.forEach(([key, entry]) => appendJsonRows(`${path}.${key}`, entry, rows));
    return;
  }

  rows.push([path, formatPrimitiveValue(value)]);
}

function parseXmlAttributes(attributeText) {
  const attributes = [];
  const attributePattern = /([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
  let match;
  while ((match = attributePattern.exec(attributeText)) !== null) {
    attributes.push([match[1], decodeXmlEntities(match[3] ?? match[4] ?? "")]);
  }
  return attributes;
}

function buildXmlPath(stack) {
  return `/${stack.map((entry) => entry.name).join("/")}`;
}

function decodeXmlEntities(value) {
  return String(value || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function resolveExcelSheetName(workbook, sheetRef = undefined) {
  const names = Array.from(workbook?.SheetNames || []);
  if (!names.length) {
    throw new Error("The Excel workbook does not contain any sheets.");
  }
  if (typeof sheetRef === "number" && Number.isInteger(sheetRef)) {
    return names[sheetRef] || names[0];
  }
  if (typeof sheetRef === "string" && sheetRef.trim() !== "" && names.includes(sheetRef)) {
    return sheetRef;
  }
  return names[0];
}

function decodeExcelRange(rangeRef) {
  const rangeText = String(rangeRef || "A1");
  const [startRef, endRef] = rangeText.split(":");
  const start = decodeExcelCellAddress(startRef);
  const end = decodeExcelCellAddress(endRef || startRef);
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  };
}

function decodeExcelCellAddress(address) {
  const match = String(address || "A1").trim().match(/^([A-Z]+)(\d+)$/i);
  if (!match) {
    return { row: 1, col: 1 };
  }
  return {
    col: decodeExcelColumnLabel(match[1]),
    row: Math.max(1, Number(match[2]) || 1),
  };
}

function decodeExcelColumnLabel(label) {
  return String(label || "")
    .toUpperCase()
    .split("")
    .reduce((total, char) => (total * 26) + (char.charCodeAt(0) - 64), 0);
}

function encodeExcelCellAddress(rowIndex, columnIndex) {
  let col = Math.max(1, Number(columnIndex) || 1);
  let label = "";
  while (col > 0) {
    const remainder = (col - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    col = Math.floor((col - 1) / 26);
  }
  return `${label}${Math.max(1, Number(rowIndex) || 1)}`;
}

function readExcelSheetCell(sheet, rowIndex, columnIndex) {
  const cell = sheet?.[encodeExcelCellAddress(rowIndex, columnIndex)];
  if (!cell) return "";
  if (cell.t === "b") {
    return cell.v ? "true" : "false";
  }
  if (cell.w != null && cell.w !== "") return String(cell.w);
  if (cell.v == null) return "";
  return formatPrimitiveValue(cell.v);
}

function formatPrimitiveValue(value) {
  if (value == null) return "";
  return String(value);
}
