import { getSelectionBounds, normalizeCell } from "./editor-state.js";
import {
  cloneLightTable,
  columnKeyForIndex,
  getCellValue,
  getColumnCount,
  getHeaderRowIndex,
  getRowValues,
  getSheetRowCount,
  isDataRow,
  setCellValue,
  setPageSetup,
} from "./table-model.js";

export const MAX_ADVANCED_FILTER_CONDITIONS = 20;
export const COLUMN_FILTER_OPERATORS = ["contains", "equals", "is-empty", "is-not-empty"];

const ADVANCED_OPERATOR_SET = new Set(["contains", "equals", "is-empty", "is-not-empty"]);
const LOGIC_PRECEDENCE = {
  OR: 1,
  AND: 2,
};
const NATURAL_COLLATOR = new Intl.Collator(undefined, {
  numeric: true,
  sensitivity: "base",
});

export function getSearchScope(table, selection, options = {}) {
  const bounds = getSelectionBounds(selection);
  if (options.scope === "selection") {
    return normalizeSearchScopeBounds(options.scopeBounds) || bounds;
  }
  if (options.scope === "global") {
    return {
      startRow: 1,
      endRow: getSheetRowCount(table),
      startCol: 1,
      endCol: Math.max(1, getColumnCount(table)),
    };
  }
  const cellCount = (bounds.endRow - bounds.startRow + 1) * (bounds.endCol - bounds.startCol + 1);
  if (cellCount > 1) {
    return bounds;
  }
  return {
    startRow: 1,
    endRow: getSheetRowCount(table),
    startCol: 1,
    endCol: Math.max(1, getColumnCount(table)),
  };
}

function normalizeSearchScopeBounds(bounds) {
  if (!bounds) return null;
  const startRow = Math.max(1, Math.trunc(Number(bounds.startRow) || 1));
  const endRow = Math.max(startRow, Math.trunc(Number(bounds.endRow) || startRow));
  const startCol = Math.max(1, Math.trunc(Number(bounds.startCol) || 1));
  const endCol = Math.max(startCol, Math.trunc(Number(bounds.endCol) || startCol));
  return {
    startRow,
    endRow,
    startCol,
    endCol,
  };
}

export function findNextMatch(table, selection, startCell, options = {}) {
  const scope = getSearchScope(table, selection, options);
  const matches = collectMatchesInScope(table, scope, options);
  if (!matches.length) return null;
  const anchor = normalizeCell(startCell);
  const inclusive = Boolean(options.includeCurrent);
  const nextIndex = matches.findIndex((match) => {
    if (match.row > anchor.row) return true;
    if (match.row < anchor.row) return false;
    if (match.col > anchor.col) return true;
    if (match.col < anchor.col) return false;
    return inclusive;
  });
  return nextIndex >= 0 ? matches[nextIndex] : matches[0];
}

export function replaceAllMatches(table, selection, options = {}) {
  const scope = getSearchScope(table, selection, options);
  const matches = collectMatchesInScope(table, scope, options);
  if (!matches.length) {
    return {
      table,
      count: 0,
      matches: [],
    };
  }
  let nextTable = table;
  matches.forEach((match) => {
    nextTable = applyMatchReplacement(nextTable, match, options);
  });
  return {
    table: nextTable,
    count: matches.length,
    matches,
  };
}

export function buildDefaultAdvancedFilterLogic(conditionCount) {
  const count = Math.max(0, Math.min(MAX_ADVANCED_FILTER_CONDITIONS, Math.trunc(Number(conditionCount) || 0)));
  if (count <= 0) return "";
  return Array.from({ length: count }, (_, index) => String(index + 1)).join(" AND ");
}

export function validateAdvancedFilterDefinition(definition = {}) {
  const conditions = normalizeAdvancedConditions(definition.conditions || []);
  if (!conditions.length) {
    return {
      ok: true,
      errors: [],
      logic: "",
      conditions,
      postfix: [],
    };
  }

  const requiredValueError = findConditionValueError(conditions);
  if (requiredValueError) {
    return {
      ok: false,
      errors: [requiredValueError],
      logic: buildDefaultAdvancedFilterLogic(conditions.length),
      conditions,
      postfix: [],
    };
  }

  const rawLogic = String(definition.logic ?? "").trim();
  const logic = rawLogic || buildDefaultAdvancedFilterLogic(conditions.length);
  const parsed = parseLogicExpression(logic, conditions.length);
  return {
    ok: parsed.ok,
    errors: parsed.errors,
    logic,
    conditions,
    postfix: parsed.postfix,
  };
}

export function getVisibleRowSet(table, filterState = {}) {
  const rowCount = getSheetRowCount(table);
  const visibleRows = new Set();
  for (let rowIndex = 1; rowIndex <= rowCount; rowIndex += 1) {
    if (isRowVisible(table, rowIndex, filterState)) {
      visibleRows.add(rowIndex);
    }
  }
  return visibleRows;
}

export function isRowVisible(table, rowIndex, filterState = {}) {
  const row = Math.max(1, Math.trunc(Number(rowIndex) || 1));
  const headerRowIndex = getHeaderRowIndex(table);
  const searchText = String(filterState.globalSearch ?? "").trim();

  if (row === headerRowIndex) {
    return true;
  }

  if (row < headerRowIndex) {
    return searchText === "" ? true : rowMatchesQuery(table, row, searchText);
  }

  if (searchText !== "" && !rowMatchesQuery(table, row, searchText)) {
    return false;
  }
  if (!rowMatchesColumnFilters(table, row, filterState.columnFilters || {})) {
    return false;
  }
  if (!rowMatchesAdvancedFilter(table, row, filterState.advancedFilter || null)) {
    return false;
  }
  return true;
}

export function sortTableRows(table, sortRules = []) {
  const normalizedRules = normalizeSortRules(sortRules);
  if (!normalizedRules.length) return cloneLightTable(table);
  const nextTable = cloneLightTable(table);
  nextTable.rows = nextTable.rows
    .map((record, index) => ({ record, index }))
    .sort((left, right) => compareSortedRecords(left, right, normalizedRules))
    .map((entry) => entry.record);
  return nextTable;
}

export function applyPageSetup(table, pageSetup) {
  return setPageSetup(table, normalizePageSetup(pageSetup));
}

export function resolvePrintAreaBounds(table, selection, pageSetup = {}) {
  const mode = String(pageSetup?.printArea?.mode || "entire-sheet");
  const fullBounds = {
    startRow: 1,
    endRow: getSheetRowCount(table),
    startCol: 1,
    endCol: Math.max(1, getColumnCount(table)),
  };

  if (mode === "selection") {
    return expandBoundsToHeaderRow(table, getSelectionBounds(selection));
  }
  if (mode === "custom") {
    const parsed = parseCellRange(String(pageSetup?.printArea?.range || ""));
    return parsed ? expandBoundsToHeaderRow(table, parsed) : fullBounds;
  }
  return fullBounds;
}

export function parseCellRange(rangeText) {
  const text = String(rangeText || "").trim().toUpperCase();
  if (!text) return null;
  const parts = text.split(":").map((part) => part.trim()).filter(Boolean);
  if (parts.length < 1 || parts.length > 2) return null;
  const start = parseCellRef(parts[0]);
  const end = parseCellRef(parts[1] || parts[0]);
  if (!start || !end) return null;
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  };
}

export function isCellWithinBounds(rowIndex, columnIndex, bounds) {
  if (!bounds) return true;
  return rowIndex >= bounds.startRow
    && rowIndex <= bounds.endRow
    && columnIndex >= bounds.startCol
    && columnIndex <= bounds.endCol;
}

function collectMatchesInScope(table, scope, options = {}) {
  const needle = String(options.findText ?? "");
  if (!needle) return [];
  const matches = [];
  for (let rowIndex = scope.startRow; rowIndex <= scope.endRow; rowIndex += 1) {
    for (let columnIndex = scope.startCol; columnIndex <= scope.endCol; columnIndex += 1) {
      const value = getCellValue(table, rowIndex, columnIndex);
      if (cellMatchesSearch(value, needle, options)) {
        matches.push({
          row: rowIndex,
          col: columnIndex,
          value,
        });
      }
    }
  }
  return matches;
}

function applyMatchReplacement(table, match, options) {
  const currentValue = getCellValue(table, match.row, match.col);
  const replacement = buildReplacementValue(
    currentValue,
    String(options.findText ?? ""),
    String(options.replaceText ?? ""),
    options,
  );
  return setCellValue(table, match.row, match.col, replacement);
}

function cellMatchesSearch(value, needle, options = {}) {
  const haystack = String(value ?? "");
  const searchNeedle = String(needle ?? "");
  if (searchNeedle === "") return false;
  if (options.wholeCell) {
    return compareText(haystack, searchNeedle, options.caseSensitive);
  }
  if (options.caseSensitive) {
    return haystack.includes(searchNeedle);
  }
  return haystack.toLocaleLowerCase().includes(searchNeedle.toLocaleLowerCase());
}

function compareText(left, right, caseSensitive) {
  if (caseSensitive) {
    return String(left ?? "") === String(right ?? "");
  }
  return String(left ?? "").toLocaleLowerCase() === String(right ?? "").toLocaleLowerCase();
}

function buildReplacementValue(currentValue, findText, replaceText, options) {
  const text = String(currentValue ?? "");
  if (findText === "") return text;
  if (options.wholeCell) {
    return replaceText;
  }
  if (options.caseSensitive) {
    return text.split(findText).join(replaceText);
  }
  const pattern = escapeRegExp(findText);
  return text.replace(new RegExp(pattern, "gi"), replaceText);
}

function rowMatchesQuery(table, rowIndex, query) {
  const normalizedQuery = String(query ?? "").trim().toLocaleLowerCase();
  if (!normalizedQuery) return true;
  const values = getRowValues(table, rowIndex, Math.max(1, getColumnCount(table)));
  return values.some((value) => String(value ?? "").toLocaleLowerCase().includes(normalizedQuery));
}

function rowMatchesColumnFilters(table, rowIndex, columnFilters) {
  if (!isDataRow(table, rowIndex)) return true;
  return Object.entries(columnFilters || {}).every(([key, filter]) => {
    if (!isValidColumnFilter(filter)) return true;
    const columnIndex = columnIndexFromKey(key);
    if (!columnIndex) return true;
    return cellPassesFilter(getCellValue(table, rowIndex, columnIndex), filter);
  });
}

function rowMatchesAdvancedFilter(table, rowIndex, advancedFilter) {
  if (!isDataRow(table, rowIndex)) return true;
  if (!advancedFilter?.enabled) return true;
  const compiled = validateAdvancedFilterDefinition(advancedFilter);
  if (!compiled.ok || !compiled.conditions.length) return true;
  const results = compiled.conditions.map((condition) => {
    const columnIndex = columnIndexFromKey(condition.columnKey);
    const value = getCellValue(table, rowIndex, columnIndex);
    return cellPassesFilter(value, condition);
  });
  return evaluatePostfix(compiled.postfix, results);
}

function normalizeAdvancedConditions(conditions) {
  return conditions
    .slice(0, MAX_ADVANCED_FILTER_CONDITIONS)
    .map((condition, index) => ({
      index: index + 1,
      columnKey: String(condition?.columnKey || columnKeyForIndex(1)),
      operator: normalizeFilterOperator(condition?.operator),
      value: String(condition?.value ?? ""),
    }));
}

function normalizeSortRules(rules) {
  return (Array.isArray(rules) ? rules : [])
    .map((rule) => ({
      columnKey: String(rule?.columnKey || ""),
      direction: String(rule?.direction || "asc").toLowerCase() === "desc" ? "desc" : "asc",
    }))
    .filter((rule) => columnIndexFromKey(rule.columnKey));
}

function compareSortedRecords(left, right, rules) {
  for (const rule of rules) {
    const columnIndex = columnIndexFromKey(rule.columnKey);
    if (!columnIndex) continue;
    const leftValue = String(left.record?.[rule.columnKey] ?? "");
    const rightValue = String(right.record?.[rule.columnKey] ?? "");
    const compared = compareSortableValues(leftValue, rightValue);
    if (compared !== 0) {
      return rule.direction === "desc" ? -compared : compared;
    }
  }
  return left.index - right.index;
}

function compareSortableValues(left, right) {
  const leftBlank = left === "";
  const rightBlank = right === "";
  if (leftBlank && rightBlank) return 0;
  if (leftBlank) return 1;
  if (rightBlank) return -1;
  return NATURAL_COLLATOR.compare(left, right);
}

function escapeRegExp(text) {
  return String(text ?? "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function cellPassesFilter(value, filter) {
  const operator = normalizeFilterOperator(filter?.operator);
  const text = String(value ?? "");
  const needle = String(filter?.value ?? "");
  if (operator === "is-empty") {
    return text === "";
  }
  if (operator === "is-not-empty") {
    return text !== "";
  }
  if (operator === "equals") {
    return compareText(text, needle, false);
  }
  return text.toLocaleLowerCase().includes(needle.toLocaleLowerCase());
}

function isValidColumnFilter(filter) {
  if (!filter) return false;
  const operator = normalizeFilterOperator(filter.operator);
  if (operator === "is-empty" || operator === "is-not-empty") {
    return true;
  }
  return String(filter.value ?? "") !== "";
}

function normalizeFilterOperator(operator) {
  const normalized = String(operator || "contains").toLowerCase();
  return ADVANCED_OPERATOR_SET.has(normalized) ? normalized : "contains";
}

function findConditionValueError(conditions) {
  const missing = conditions.find((condition) => requiresFilterValue(condition.operator) && condition.value === "");
  return missing ? `Condition ${missing.index} needs a value.` : "";
}

function requiresFilterValue(operator) {
  return normalizeFilterOperator(operator) === "contains" || normalizeFilterOperator(operator) === "equals";
}

function parseLogicExpression(expression, conditionCount) {
  const tokens = tokenizeLogicExpression(expression);
  if (!tokens.ok) {
    return {
      ok: false,
      errors: tokens.errors,
      postfix: [],
    };
  }

  const numbers = tokens.tokens.filter((token) => token.type === "number").map((token) => token.value);
  const outOfRange = numbers.find((value) => value < 1 || value > conditionCount);
  if (outOfRange) {
    return {
      ok: false,
      errors: ["Condition number is outside the current range."],
      postfix: [],
    };
  }

  const repeated = findRepeatedNumber(numbers);
  if (repeated) {
    return {
      ok: false,
      errors: [`Condition ${repeated} is repeated.`],
      postfix: [],
    };
  }

  const missing = findMissingNumbers(numbers, conditionCount);
  if (missing.length) {
    return {
      ok: false,
      errors: ["Every defined condition number must appear exactly once."],
      postfix: [],
    };
  }

  return buildPostfix(tokens.tokens);
}

function tokenizeLogicExpression(expression) {
  const source = String(expression ?? "");
  const tokens = [];
  let index = 0;
  while (index < source.length) {
    const char = source[index];
    if (/\s/.test(char)) {
      index += 1;
      continue;
    }
    if (char === "(") {
      tokens.push({ type: "lparen" });
      index += 1;
      continue;
    }
    if (char === ")") {
      tokens.push({ type: "rparen" });
      index += 1;
      continue;
    }
    if (/\d/.test(char)) {
      const match = source.slice(index).match(/^\d+/);
      tokens.push({ type: "number", value: Number(match?.[0] || 0) });
      index += match?.[0]?.length || 1;
      continue;
    }
    const operatorMatch = source.slice(index).match(/^(AND|OR)\b/i);
    if (operatorMatch) {
      tokens.push({ type: "operator", value: operatorMatch[1].toUpperCase() });
      index += operatorMatch[0].length;
      continue;
    }
    return {
      ok: false,
      errors: ["The logic expression contains unknown characters."],
      tokens: [],
    };
  }
  return {
    ok: true,
    errors: [],
    tokens,
  };
}

function buildPostfix(tokens) {
  const output = [];
  const operators = [];
  let expected = "operand";
  let balance = 0;

  for (const token of tokens) {
    if (expected === "operand") {
      if (token.type === "number") {
        output.push(token);
        expected = "operator";
        continue;
      }
      if (token.type === "lparen") {
        operators.push(token);
        balance += 1;
        continue;
      }
      return {
        ok: false,
        errors: ["The logic expression has invalid operator placement."],
        postfix: [],
      };
    }

    if (token.type === "operator") {
      while (operators.length) {
        const top = operators.at(-1);
        if (top?.type !== "operator") break;
        if (LOGIC_PRECEDENCE[top.value] < LOGIC_PRECEDENCE[token.value]) break;
        output.push(operators.pop());
      }
      operators.push(token);
      expected = "operand";
      continue;
    }

    if (token.type === "rparen") {
      if (balance <= 0) {
        return {
          ok: false,
          errors: ["The logic expression has unbalanced parentheses."],
          postfix: [],
        };
      }
      while (operators.length && operators.at(-1)?.type !== "lparen") {
        output.push(operators.pop());
      }
      operators.pop();
      balance -= 1;
      continue;
    }

    return {
      ok: false,
      errors: ["The logic expression has invalid operator placement."],
      postfix: [],
    };
  }

  if (expected === "operand") {
    return {
      ok: false,
      errors: ["The logic expression has invalid operator placement."],
      postfix: [],
    };
  }

  while (operators.length) {
    const token = operators.pop();
    if (token?.type === "lparen" || token?.type === "rparen" || balance > 0) {
      return {
        ok: false,
        errors: ["The logic expression has unbalanced parentheses."],
        postfix: [],
      };
    }
    output.push(token);
  }

  return {
    ok: true,
    errors: [],
    postfix: output,
  };
}

function evaluatePostfix(postfix, conditionResults) {
  const stack = [];
  postfix.forEach((token) => {
    if (token.type === "number") {
      stack.push(Boolean(conditionResults[token.value - 1]));
      return;
    }
    const right = Boolean(stack.pop());
    const left = Boolean(stack.pop());
    stack.push(token.value === "AND" ? left && right : left || right);
  });
  return Boolean(stack.pop());
}

function findRepeatedNumber(numbers) {
  const seen = new Set();
  for (const number of numbers) {
    if (seen.has(number)) return number;
    seen.add(number);
  }
  return 0;
}

function findMissingNumbers(numbers, conditionCount) {
  const seen = new Set(numbers);
  const missing = [];
  for (let index = 1; index <= conditionCount; index += 1) {
    if (!seen.has(index)) {
      missing.push(index);
    }
  }
  return missing;
}

function normalizePageSetup(pageSetup = {}) {
  const margins = pageSetup?.margins && typeof pageSetup.margins === "object" ? pageSetup.margins : {};
  const headerFooter = pageSetup?.headerFooter && typeof pageSetup.headerFooter === "object" ? pageSetup.headerFooter : {};
  const printArea = pageSetup?.printArea && typeof pageSetup.printArea === "object" ? pageSetup.printArea : {};
  return {
    margins: {
      top: normalizeMarginValue(margins.top),
      right: normalizeMarginValue(margins.right),
      bottom: normalizeMarginValue(margins.bottom),
      left: normalizeMarginValue(margins.left),
    },
    paperSize: normalizePaperSize(pageSetup?.paperSize),
    orientation: normalizeOrientation(pageSetup?.orientation),
    printArea: {
      mode: normalizePrintAreaMode(printArea.mode),
      range: String(printArea.range || ""),
    },
    headerFooter: {
      header: String(headerFooter.header || ""),
      footer: String(headerFooter.footer || ""),
    },
    background: pageSetup?.background
      ? {
        mode: String(pageSetup.background.mode || "solid-color"),
        color: String(pageSetup.background.color || ""),
      }
      : null,
  };
}

function normalizeMarginValue(value) {
  const number = Number(value);
  return Number.isFinite(number) && number >= 0 ? number : 12;
}

function normalizePaperSize(value) {
  const safe = String(value || "A4");
  return ["A4", "A3", "Letter"].includes(safe) ? safe : "A4";
}

function normalizeOrientation(value) {
  return String(value || "portrait").toLowerCase() === "landscape" ? "landscape" : "portrait";
}

function normalizePrintAreaMode(value) {
  const safe = String(value || "entire-sheet");
  return ["entire-sheet", "selection", "custom"].includes(safe) ? safe : "entire-sheet";
}

function expandBoundsToHeaderRow(table, bounds) {
  if (!bounds) return null;
  const headerRowIndex = getHeaderRowIndex(table);
  if (bounds.endRow < headerRowIndex) {
    return bounds;
  }
  return {
    ...bounds,
    startRow: Math.min(bounds.startRow, headerRowIndex),
  };
}

function parseCellRef(text) {
  const match = String(text || "").trim().toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;
  return {
    col: columnIndexFromLabel(match[1]),
    row: Number(match[2]),
  };
}

function columnIndexFromLabel(label) {
  let index = 0;
  for (const char of String(label || "")) {
    const code = char.charCodeAt(0);
    if (code < 65 || code > 90) return 0;
    index = (index * 26) + (code - 64);
  }
  return index;
}

function columnIndexFromKey(key) {
  const match = String(key || "").match(/^c(\d+)$/i);
  return match ? Number(match[1]) : 0;
}
