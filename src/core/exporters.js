import {
  getCellValue,
  getColumnCount,
  getSheetRowCount,
} from "./table-model.js";

export const TEXT_EXPORT_TARGETS = ["csv", "tsv", "txt"];
export const TEXT_EXPORT_ENCODINGS = ["UTF-8", "Shift-JIS", "UTF-16LE", "UTF-16BE", "EUC-JP"];
export const TEXT_EXPORT_ROW_LINE_ENDINGS = ["CR+LF", "LF", "CR"];

export function getSupportedTextExportTargets() {
  return [...TEXT_EXPORT_TARGETS];
}

export function getSupportedTextExportEncodings() {
  return [...TEXT_EXPORT_ENCODINGS];
}

export function getSupportedTextExportRowLineEndings() {
  return [...TEXT_EXPORT_ROW_LINE_ENDINGS];
}

export function createDefaultTextExportSettings(target = "csv") {
  return normalizeTextExportSettings({
    target,
    encoding: "UTF-8",
    rowLineEnding: "CR+LF",
    normalizeCellLineBreaksToLf: false,
    withBom: false,
    quoteAllCells: false,
    includeHiddenData: true,
  });
}

export function normalizeTextExportSettings(settings = {}) {
  const target = normalizeTextExportTarget(settings.target);
  const encoding = normalizeTextExportEncoding(settings.encoding);
  const rowLineEnding = normalizeTextExportRowLineEnding(settings.rowLineEnding);
  return {
    target,
    encoding,
    rowLineEnding,
    normalizeCellLineBreaksToLf: Boolean(settings.normalizeCellLineBreaksToLf),
    withBom: Boolean(settings.withBom),
    quoteAllCells: target === "csv" && Boolean(settings.quoteAllCells),
    includeHiddenData: settings.includeHiddenData !== false,
  };
}

export function normalizeTextExportTarget(target = "csv") {
  const normalized = String(target || "").trim().toLowerCase();
  return TEXT_EXPORT_TARGETS.includes(normalized) ? normalized : "csv";
}

export function normalizeTextExportEncoding(encoding = "UTF-8") {
  const normalized = String(encoding || "").trim().toUpperCase();
  const match = TEXT_EXPORT_ENCODINGS.find((entry) => entry.toUpperCase() === normalized);
  return match || "UTF-8";
}

export function normalizeTextExportRowLineEnding(rowLineEnding = "CR+LF") {
  const normalized = String(rowLineEnding || "").trim().toUpperCase();
  if (normalized === "LF" || normalized === "CR") {
    return normalized;
  }
  return "CR+LF";
}

export function shouldEnableBomToggle(encoding = "UTF-8") {
  return ["UTF-8", "UTF-16LE", "UTF-16BE"].includes(normalizeTextExportEncoding(encoding));
}

export function resolveTextExportDelimiter(target = "csv") {
  const normalizedTarget = normalizeTextExportTarget(target);
  if (normalizedTarget === "csv") return ",";
  return "\t";
}

export function resolveTextExportRowLineEndingSequence(rowLineEnding = "CR+LF") {
  const normalized = normalizeTextExportRowLineEnding(rowLineEnding);
  if (normalized === "LF") return "\n";
  if (normalized === "CR") return "\r";
  return "\r\n";
}

export function buildTextExportMatrix(table, options = {}) {
  const columnCount = getColumnCount(table);
  const rowCount = getSheetRowCount(table);
  const visibleRowSet = options.visibleRowSet instanceof Set ? options.visibleRowSet : null;
  const includeHiddenData = options.includeHiddenData !== false;
  const matrix = [];

  for (let rowIndex = 1; rowIndex <= rowCount; rowIndex += 1) {
    if (!includeHiddenData && visibleRowSet && !visibleRowSet.has(rowIndex)) {
      continue;
    }
    const row = [];
    for (let columnIndex = 1; columnIndex <= columnCount; columnIndex += 1) {
      row.push(getCellValue(table, rowIndex, columnIndex));
    }
    matrix.push(row);
  }

  return matrix;
}

export function serializeTextExportMatrix(matrix, options = {}) {
  const settings = normalizeTextExportSettings(options);
  const delimiter = resolveTextExportDelimiter(settings.target);
  const rowLineEnding = resolveTextExportRowLineEndingSequence(settings.rowLineEnding);
  return (Array.isArray(matrix) ? matrix : [])
    .map((row) => {
      const safeRow = Array.isArray(row) ? row : [];
      return safeRow.map((value) => encodeDelimitedExportCell(value, {
        delimiter,
        quoteAllCells: settings.quoteAllCells,
        normalizeCellLineBreaksToLf: settings.normalizeCellLineBreaksToLf,
      })).join(delimiter);
    })
    .join(rowLineEnding);
}

export function serializeTextExport(table, options = {}) {
  const settings = normalizeTextExportSettings(options);
  const matrix = buildTextExportMatrix(table, settings);
  return {
    matrix,
    text: serializeTextExportMatrix(matrix, settings),
    settings,
  };
}

export function encodeTextExportBytes(text, options = {}) {
  const settings = normalizeTextExportSettings(options);
  let bytes;

  if (settings.encoding === "UTF-8") {
    bytes = new TextEncoder().encode(text);
  } else if (settings.encoding === "UTF-16LE") {
    bytes = encodeUtf16(text, false);
  } else if (settings.encoding === "UTF-16BE") {
    bytes = encodeUtf16(text, true);
  } else {
    const legacyEncode = typeof options.legacyEncode === "function" ? options.legacyEncode : null;
    if (!legacyEncode) {
      throw new Error(`This browser build does not have a ${settings.encoding} export encoder loaded.`);
    }
    const encoded = legacyEncode(settings.encoding, text);
    bytes = encoded instanceof Uint8Array ? encoded : Uint8Array.from(encoded || []);
  }

  if (!settings.withBom || !shouldEnableBomToggle(settings.encoding)) {
    return bytes;
  }
  return prependBomBytes(bytes, settings.encoding);
}

export function formatTextExportSummary(options = {}) {
  const settings = normalizeTextExportSettings(options);
  const parts = [
    settings.target.toUpperCase(),
    settings.encoding,
    settings.rowLineEnding,
    settings.withBom && shouldEnableBomToggle(settings.encoding) ? "BOM on" : "BOM off",
  ];
  if (settings.target === "csv") {
    parts.push(settings.quoteAllCells ? "Quote all" : "Quote minimal");
  }
  parts.push(settings.includeHiddenData ? "Include hidden" : "Visible rows only");
  if (settings.normalizeCellLineBreaksToLf) {
    parts.push("Cell LF");
  }
  return parts.join(" / ");
}

function encodeDelimitedExportCell(value, options = {}) {
  let text = String(value ?? "");
  if (options.normalizeCellLineBreaksToLf) {
    text = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  }
  const escaped = text.replace(/"/g, "\"\"");
  if (options.quoteAllCells) {
    return `"${escaped}"`;
  }
  if (!/["\r\n]/.test(text) && !text.includes(options.delimiter || ",")) {
    return text;
  }
  return `"${escaped}"`;
}

function encodeUtf16(text, bigEndian) {
  const source = String(text ?? "");
  const bytes = new Uint8Array(source.length * 2);
  for (let index = 0; index < source.length; index += 1) {
    const codeUnit = source.charCodeAt(index);
    const offset = index * 2;
    if (bigEndian) {
      bytes[offset] = (codeUnit >> 8) & 0xff;
      bytes[offset + 1] = codeUnit & 0xff;
    } else {
      bytes[offset] = codeUnit & 0xff;
      bytes[offset + 1] = (codeUnit >> 8) & 0xff;
    }
  }
  return bytes;
}

function prependBomBytes(bytes, encoding) {
  let bom = [];
  if (encoding === "UTF-8") {
    bom = [0xef, 0xbb, 0xbf];
  } else if (encoding === "UTF-16LE") {
    bom = [0xff, 0xfe];
  } else if (encoding === "UTF-16BE") {
    bom = [0xfe, 0xff];
  }
  const prefixed = new Uint8Array(bom.length + bytes.length);
  prefixed.set(bom, 0);
  prefixed.set(bytes, bom.length);
  return prefixed;
}
