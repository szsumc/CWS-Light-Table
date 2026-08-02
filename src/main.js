import {
  listWorkbookSheets,
  parseCwsHtmlWorkbook,
  serializeLightTableToCwsHtml,
  workbookToLightTable,
} from "./core/cws.js";
import {
  applyMatrixToTable,
  fillSelectionByRepeat,
  parseDelimitedText,
  selectionFromMatrix,
  selectionToDelimitedText,
} from "./core/clipboard.js";
import {
  findEncodingIssuesInCells,
  formatNumberCommasInCells,
  normalizeCaseInCells,
  previewBlankRows,
  previewDuplicateRows,
  previewNormalizeCase,
  previewNumberCommas,
  previewTrimSpaces,
  removeBlankRows,
  removeDuplicateRows,
  trimSpacesInCells,
} from "./core/data-cleaning.js";
import {
  applyKeyboardEditIntent,
  createCellSelection,
  createColumnSelection,
  createRangeSelection,
  createRowSelection,
  getKeyboardEditIntent,
  getSelectionBounds,
  isCellSelected,
  isColumnHeaderSelected,
  isRowHeaderSelected,
  moveActiveCell,
  normalizeCell,
} from "./core/editor-state.js";
import {
  canRedo,
  canUndo,
  createHistoryState,
  pushHistorySnapshot,
  redoHistory,
  undoHistory,
} from "./core/history.js";
import {
  applyPageSetup,
  buildDefaultAdvancedFilterLogic,
  findNextMatch,
  getVisibleRowSet,
  parseCellRange,
  replaceAllMatches,
  resolvePrintAreaBounds,
  sortTableRows,
  validateAdvancedFilterDefinition,
  isCellWithinBounds,
} from "./core/table-operations.js";
import {
  columnKeyForIndex,
  createEmptyLightTable,
  DEFAULT_SHEET_NAME,
  getCellValue,
  getColumnCount,
  getGridLimits,
  getHeaderRowIndex,
  getLastUsedColumnIndex,
  getLastUsedRowIndex,
  getSheetRowCount,
  deleteColumns,
  deleteRows,
  insertCells,
  insertColumns,
  insertRows,
  setGridLimits,
  setCellValue,
  setHeaderRow,
} from "./core/table-model.js";
import {
  buildSheetGridView,
  columnLabelFromIndex,
  gridCellName,
  isViewportPaddingRow,
} from "./core/grid.js";
import {
  buildGridMetrics,
  estimateVisibleColumnCount,
  estimateVisibleRowCount,
} from "./core/grid-metrics.js";
import {
  calculateVisibleRange,
  visibleRangeSignature,
} from "./core/visible-range.js";
import {
  createViewportState,
  updateViewportState,
} from "./core/viewport-state.js";
import {
  excelWorkbookToMatrix,
  getImportAnchorCell,
  getSupportedTextImportEncodings,
  inferDefaultDelimiterMode,
  inferDefaultTextImportEncoding,
  inferImportFileKind,
  isExcelImportKind,
  isTextImportKind,
  listExcelWorkbookSheets,
  parseTextImportContent,
} from "./core/importers.js";
import {
  buildTextExportMatrix,
  createDefaultTextExportSettings,
  encodeTextExportBytes,
  formatTextExportSummary,
  getSupportedTextExportEncodings,
  getSupportedTextExportRowLineEndings,
  getSupportedTextExportTargets,
  normalizeTextExportEncoding,
  normalizeTextExportRowLineEnding,
  normalizeTextExportSettings,
  normalizeTextExportTarget,
  resolveTextExportDelimiter,
  serializeTextExportMatrix,
  shouldEnableBomToggle,
} from "./core/exporters.js";

const state = {
  workbook: null,
  table: createEmptyLightTable(),
  sourceFileName: "",
  saveFileHandle: null,
  activeCell: { row: 1, col: 1 },
  selection: createCellSelection({ row: 1, col: 1 }),
  copiedSelection: null,
  openMenuKey: null,
  openSubmenuKey: "",
  editing: null,
  history: createHistoryState(),
  fillDrag: null,
  selectionDrag: null,
  suppressNextGridClick: false,
  globalSearch: "",
  columnFilters: {},
  advancedFilter: {
    enabled: false,
    conditions: [],
    logic: "",
  },
  findReplace: {
    findText: "",
    replaceText: "",
    scope: "global",
    scopeBounds: null,
    caseSensitive: false,
    wholeCell: false,
    currentMatch: null,
  },
  advancedFilterDraft: null,
  removeDuplicatesDraft: null,
  removeBlankRowsDraft: null,
  trimSpacesDraft: null,
  normalizeCaseDraft: null,
  numberCommasDraft: null,
  encodingIssuesDraft: null,
  sortDraft: null,
  pageSetupDraft: null,
  gridLimitsDraft: null,
  openPanel: null,
  headerMenu: null,
  panelError: "",
  importDraft: null,
  exportDraft: null,
  exportSettings: createDefaultTextExportSettings("csv"),
  importFileMode: "",
  pendingFocusId: "",
  pendingScrollCell: null,
  language: "en",
  statusMessageCanonical: "Blank sheet ready. Open a CWS HTML file or save a new empty one.",
  viewportState: createViewportState(),
  logicalGridMetrics: buildGridMetrics(createEmptyLightTable()),
  gridMetrics: buildGridMetrics(createEmptyLightTable()),
  rowVirtualizationEnabled: false,
  columnVirtualizationEnabled: false,
  visibleRowLayout: createEmptyVisibleRowLayout(),
  rowViewport: createEmptyRowViewport(),
  columnViewport: createEmptyColumnViewport(),
  visibleRowsCache: null,
  visibleRange: calculateVisibleRange(buildGridMetrics(createEmptyLightTable()), createViewportState()),
  visibleRangeSignature: "",
  printing: false,
};

const MIN_VIRTUALIZED_ROW_COUNT = 200;
const MIN_VIRTUALIZED_COLUMN_COUNT = 60;
const MAX_BROWSER_PRINT_ROWS = 500;
const MAX_BROWSER_PRINT_COLUMNS = 30;
const MAX_BROWSER_PRINT_CELLS = 15000;
const APP_VERSION = "1.1.1";
const HELP_SITE_URL = "http://cws-light-table-docs.macsocloud.com/";
const VENDOR_SCRIPT_PATHS = {
  xlsx: "./vendor/xlsx/dist/xlsx.full.min.js",
  codepageBase: "./vendor/codepage/dist/cpexcel.full.js",
  codepageEucJp: "./vendor/codepage/bits/51932.js",
  codepageUtils: "./vendor/codepage/cputils.js",
};
const DEFAULT_LANGUAGE = "en";
const LANGUAGE_STORAGE_KEY = "cws-light-table.language";
const SUPPORTED_LANGUAGES = ["en", "zh-CN", "ja"];
const LANGUAGE_LABELS = {
  en: "English",
  "zh-CN": "简体中文",
  ja: "日本語",
};
const LANGUAGE_HTML_LANG = {
  en: "en",
  "zh-CN": "zh-CN",
  ja: "ja",
};

const TEXT_TRANSLATIONS = {
  "Application menu": { "zh-CN": "应用菜单", ja: "アプリケーションメニュー" },
  "Quick actions": { "zh-CN": "快捷操作", ja: "クイック操作" },
  "Formula bar": { "zh-CN": "公式栏", ja: "数式バー" },
  Global: { "zh-CN": "全局", ja: "全体" },
  "Global search": { "zh-CN": "全局搜索", ja: "全体検索" },
  "Clear global search": { "zh-CN": "清除全局搜索", ja: "全体検索をクリア" },
  "Spreadsheet grid": { "zh-CN": "电子表格网格", ja: "スプレッドシートグリッド" },
  "Printable sheet": { "zh-CN": "可打印工作表", ja: "印刷用シート" },
  File: { "zh-CN": "文件", ja: "ファイル" },
  Edit: { "zh-CN": "编辑", ja: "編集" },
  Search: { "zh-CN": "搜索", ja: "検索" },
  Data: { "zh-CN": "数据", ja: "データ" },
  View: { "zh-CN": "视图", ja: "表示" },
  Option: { "zh-CN": "选项", ja: "オプション" },
  Help: { "zh-CN": "帮助", ja: "ヘルプ" },
  Language: { "zh-CN": "语言", ja: "言語" },
  Site: { "zh-CN": "说明网站", ja: "説明サイト" },
  "Open CWS HTML": { "zh-CN": "打开 CWS HTML", ja: "CWS HTML を開く" },
  "Save CWS HTML": { "zh-CN": "保存 CWS HTML", ja: "CWS HTML を保存" },
  Import: { "zh-CN": "导入", ja: "インポート" },
  Export: { "zh-CN": "导出", ja: "エクスポート" },
  Print: { "zh-CN": "打印", ja: "印刷" },
  "Text / Structured Data": { "zh-CN": "文本 / 结构化数据", ja: "テキスト / 構造化データ" },
  "Find / Replace": { "zh-CN": "查找 / 替换", ja: "検索 / 置換" },
  "Set Header Row": { "zh-CN": "设置表头行", ja: "ヘッダー行を設定" },
  "Advanced Filter": { "zh-CN": "高级筛选", ja: "高度なフィルター" },
  "Clear Filters": { "zh-CN": "清除筛选", ja: "フィルターをクリア" },
  "Clean Data": { "zh-CN": "数据清理", ja: "データクリーニング" },
  "Remove Duplicates": { "zh-CN": "去除重复行", ja: "重複行を削除" },
  "Remove Blank Rows": { "zh-CN": "去除空白行", ja: "空白行を削除" },
  "Trim Spaces": { "zh-CN": "清理空格", ja: "空白を整理" },
  "Format Conversion": { "zh-CN": "格式转换", ja: "形式変換" },
  "Amount commas": { "zh-CN": "金额逗号", ja: "金額桁区切り" },
  "Comma Mode": { "zh-CN": "逗号模式", ja: "桁区切りモード" },
  "Add commas": { "zh-CN": "增加逗号", ja: "桁区切りを追加" },
  "Remove commas": { "zh-CN": "去除逗号", ja: "桁区切りを削除" },
  "Amount comma preview": { "zh-CN": "金额逗号预览", ja: "金額桁区切りプレビュー" },
  "Numeric cells": { "zh-CN": "数值单元格", ja: "数値セル" },
  "Only numeric strings are changed. Invalid numbers, IDs with leading zeros, currency symbols, and text are skipped.": {
    "zh-CN": "只修改数值字符串。非法数字、前导零编号、货币符号和文本会被跳过。",
    ja: "数値文字列のみ変更します。不正な数値、先頭ゼロの ID、通貨記号、テキストはスキップします。",
  },
  "Find Encoding Issues": { "zh-CN": "查找疑似乱码", ja: "文字化け候補を検索" },
  "Encoding Issue Preview": { "zh-CN": "疑似乱码预览", ja: "文字化け候補プレビュー" },
  "Suspect cells": { "zh-CN": "疑似单元格", ja: "候補セル" },
  "Issue results": { "zh-CN": "查找结果", ja: "検索結果" },
  "No suspicious encoding issues found.": { "zh-CN": "未找到疑似乱码。", ja: "文字化け候補は見つかりませんでした。" },
  "Click a result to jump to that cell. This check does not modify workbook data.": {
    "zh-CN": "点击结果可跳转到对应单元格。此检查不会修改工作簿数据。",
    ja: "結果をクリックするとそのセルへ移動します。このチェックはワークブックデータを変更しません。",
  },
  "Showing first 20 results.": { "zh-CN": "仅显示前 20 条结果。", ja: "先頭 20 件のみ表示しています。" },
  "Replacement character": { "zh-CN": "替换字符", ja: "置換文字" },
  "Control character": { "zh-CN": "控制字符", ja: "制御文字" },
  "Suspicious mojibake sequence": { "zh-CN": "疑似乱码序列", ja: "文字化けらしい並び" },
  Scope: { "zh-CN": "范围", ja: "範囲" },
  "Current selection": { "zh-CN": "当前选区", ja: "現在の選択範囲" },
  "Entire data area": { "zh-CN": "整个数据区域", ja: "データ領域全体" },
  "Columns To Compare": { "zh-CN": "用于判断重复的列", ja: "比較する列" },
  "Select All": { "zh-CN": "全选", ja: "すべて選択" },
  "Clear Selection": { "zh-CN": "清除选择", ja: "選択をクリア" },
  "Header row is not included when removing duplicates.": {
    "zh-CN": "去除重复行时不会处理表头行。",
    ja: "重複行の削除ではヘッダー行は処理されません。",
  },
  "Header row is not included when removing blank rows.": {
    "zh-CN": "去除空白行时不会处理表头行。",
    ja: "空白行の削除ではヘッダー行は処理されません。",
  },
  "A row is blank only when every cell in the row is empty.": {
    "zh-CN": "只有整行所有单元格都为空时，才会被视为空白行。",
    ja: "行内のすべてのセルが空の場合のみ空白行として扱います。",
  },
  "Entire data area means rows below the header row.": {
    "zh-CN": "整个数据区域指表头行下方的数据行。",
    ja: "データ領域全体はヘッダー行の下のデータ行を指します。",
  },
  "Duplicate Preview": { "zh-CN": "重复预览", ja: "重複プレビュー" },
  "Blank Row Preview": { "zh-CN": "空白行预览", ja: "空白行プレビュー" },
  "Trim Preview": { "zh-CN": "空格清理预览", ja: "空白整理プレビュー" },
  "Rows to check": { "zh-CN": "检查行数", ja: "確認する行数" },
  "Rows to remove": { "zh-CN": "将删除行数", ja: "削除する行数" },
  "Cells to check": { "zh-CN": "检查单元格数", ja: "確認するセル数" },
  "Cells to change": { "zh-CN": "将修改单元格数", ja: "変更するセル数" },
  "Conversion Preview": { "zh-CN": "转换预览", ja: "変換プレビュー" },
  "Conversion Mode": { "zh-CN": "转换模式", ja: "変換モード" },
  lowercase: { "zh-CN": "转为小写", ja: "小文字に変換" },
  UPPERCASE: { "zh-CN": "转为大写", ja: "大文字に変換" },
  "Capitalize Words": { "zh-CN": "单词首字母大写", ja: "単語の先頭を大文字" },
  "Full-width": { "zh-CN": "全角转换", ja: "全角に変換" },
  "Half-width": { "zh-CN": "半角转换", ja: "半角に変換" },
  "Spaces To Clean": { "zh-CN": "要清理的空格", ja: "整理する空白" },
  "Leading spaces": { "zh-CN": "前面空格", ja: "先頭の空白" },
  "Inner spaces": { "zh-CN": "中间空格", ja: "中間の空白" },
  "Trailing spaces": { "zh-CN": "末尾空格", ja: "末尾の空白" },
  "Inner Space Mode": { "zh-CN": "中间空格处理方式", ja: "中間空白の処理" },
  "Collapse consecutive spaces to one": { "zh-CN": "连续多个空格压缩为一个空格", ja: "連続する空白を1つに圧縮" },
  "Remove all inner spaces": { "zh-CN": "删除所有中间空格", ja: "すべての中間空白を削除" },
  "Spaces include half-width spaces, full-width spaces, and Tab.": {
    "zh-CN": "空格包括半角空格、全角空格和 Tab。",
    ja: "空白には半角スペース、全角スペース、Tab を含みます。",
  },
  "Insert Row Above": { "zh-CN": "在上方插入行", ja: "上に行を挿入" },
  "Insert Row Below": { "zh-CN": "在下方插入行", ja: "下に行を挿入" },
  "Insert Column Left": { "zh-CN": "在左侧插入列", ja: "左に列を挿入" },
  "Insert Column Right": { "zh-CN": "在右侧插入列", ja: "右に列を挿入" },
  "Insert Cells": { "zh-CN": "插入单元格", ja: "セルを挿入" },
  "Delete Rows": { "zh-CN": "删除行", ja: "行を削除" },
  "Delete Columns": { "zh-CN": "删除列", ja: "列を削除" },
  "Shift Right": { "zh-CN": "右移", ja: "右にシフト" },
  "Shift Down": { "zh-CN": "下移", ja: "下にシフト" },
  "Page Setup": { "zh-CN": "页面设置", ja: "ページ設定" },
  "Grid Limits": { "zh-CN": "网格限制", ja: "グリッド制限" },
  Version: { "zh-CN": "版本", ja: "バージョン" },
  Undo: { "zh-CN": "撤销", ja: "元に戻す" },
  Redo: { "zh-CN": "重做", ja: "やり直す" },
  "Copy Range": { "zh-CN": "复制区域", ja: "範囲をコピー" },
  "Paste Range": { "zh-CN": "粘贴区域", ja: "範囲を貼り付け" },
  "Clear": { "zh-CN": "清除", ja: "クリア" },
  "Selected cells were already blank.": { "zh-CN": "选中的单元格已经为空。", ja: "選択したセルはすでに空です。" },
  "Select at least one column.": { "zh-CN": "请至少选择一列。", ja: "少なくとも1列を選択してください。" },
  "Select at least one space cleanup option.": {
    "zh-CN": "请至少选择一种空格清理选项。",
    ja: "少なくとも1つの空白整理オプションを選択してください。",
  },
  "No data rows are available to clean.": {
    "zh-CN": "没有可清理的数据行。",
    ja: "整理できるデータ行がありません。",
  },
  "Opened the documentation site.": { "zh-CN": "已打开说明网站。", ja: "説明サイトを開きました。" },
  "The documentation site could not be opened automatically.": {
    "zh-CN": "无法自动打开说明网站。",
    ja: "説明サイトを自動で開けませんでした。",
  },
  "Load Selected Sheet": { "zh-CN": "加载所选工作表", ja: "選択したシートを読み込む" },
  "Blank sheet ready. Open a CWS HTML file or save a new empty one.": {
    "zh-CN": "空白工作表已就绪。打开 CWS HTML 文件，或保存新的空白文件。",
    ja: "空のシートを使用できます。CWS HTML ファイルを開くか、新しい空ファイルを保存してください。",
  },
  "Unsaved blank sheet": { "zh-CN": "未保存的空白工作表", ja: "未保存の空シート" },
  "This file has multiple sheets. Pick one sheet to load into the lightweight editor.": {
    "zh-CN": "此文件包含多个工作表。请选择一个工作表加载到轻量编辑器中。",
    ja: "このファイルには複数のシートがあります。軽量エディターに読み込むシートを選択してください。",
  },
  "Open header tools": { "zh-CN": "打开表头工具", ja: "ヘッダーツールを開く" },
  "Fill handle": { "zh-CN": "填充柄", ja: "フィルハンドル" },
  "Help panel": { "zh-CN": "帮助面板", ja: "ヘルプパネル" },
  "CWS Light Table Help": { "zh-CN": "CWS Light Table 帮助", ja: "CWS Light Table ヘルプ" },
  "Close": { "zh-CN": "关闭", ja: "閉じる" },
  "CWS Light Table is a lightweight single-sheet table editor for CWS HTML.": {
    "zh-CN": "CWS Light Table 是用于 CWS HTML 的轻量单工作表表格编辑器。",
    ja: "CWS Light Table は CWS HTML 用の軽量な単一シートテーブルエディターです。",
  },
  "Key features": { "zh-CN": "主要功能", ja: "主な機能" },
  "Key Features": { "zh-CN": "主要功能", ja: "主な機能" },
  "Data Cleaning": { "zh-CN": "数据清理", ja: "データクリーニング" },
  "Open and save CWS HTML files.": { "zh-CN": "打开和保存 CWS HTML 文件。", ja: "CWS HTML ファイルを開いて保存できます。" },
  "Import Excel and text or structured data into the grid.": { "zh-CN": "将 Excel、文本或结构化数据导入网格。", ja: "Excel、テキスト、構造化データをグリッドにインポートできます。" },
  "Edit cells with the formula bar, copy and paste, fill handle, and undo or redo.": {
    "zh-CN": "使用公式栏、复制粘贴、填充柄、撤销和重做来编辑单元格。",
    ja: "数式バー、コピーと貼り付け、フィルハンドル、元に戻す / やり直しでセルを編集できます。",
  },
  "Use find and replace, header rows, filters, print setup, and export tools.": {
    "zh-CN": "使用查找替换、表头行、筛选、打印设置和导出工具。",
    ja: "検索と置換、ヘッダー行、フィルター、印刷設定、エクスポートツールを使用できます。",
  },
  "Adjust grid size and insert rows, columns, or cells.": {
    "zh-CN": "调整网格大小，并插入行、列或单元格。",
    ja: "グリッドサイズを調整し、行、列、セルを挿入できます。",
  },
  "Remove duplicate rows by one selected column or by a multi-column comparison key.": {
    "zh-CN": "按单列或多列组合条件去除重复行。",
    ja: "1 列または複数列の比較キーで重複行を削除できます。",
  },
  "Remove fully blank rows from the current selection or the whole data area.": {
    "zh-CN": "从当前选区或整个数据区域去除完全空白行。",
    ja: "現在の選択範囲またはデータ領域全体から完全な空白行を削除できます。",
  },
  "Trim leading, trailing, and inner spaces with configurable inner-space handling.": {
    "zh-CN": "清理前置、末尾和中间空格，并可选择中间空格的处理方式。",
    ja: "先頭、末尾、中間の空白を整理し、中間空白の処理方法を選択できます。",
  },
  "Convert text case and width, including lowercase, uppercase, capitalized words, full-width, and half-width.": {
    "zh-CN": "进行格式转换，包括小写、大写、单词首字母大写、全角和半角。",
    ja: "小文字、大文字、単語先頭大文字、全角、半角の形式変換ができます。",
  },
  "Add or remove thousands separators for numeric strings while skipping non-numeric text.": {
    "zh-CN": "为数值字符串增加或去除千分位逗号，并跳过非数值文本。",
    ja: "数値文字列に桁区切りを追加または削除し、数値ではないテキストはスキップします。",
  },
  "Find suspicious encoding issues such as mojibake, replacement characters, and control characters without modifying data.": {
    "zh-CN": "查找疑似乱码、替换字符和控制字符，不修改数据。",
    ja: "文字化け候補、置換文字、制御文字をデータ変更なしで検索できます。",
  },
  "Data cleaning tools can run on the current selection or the whole data area depending on the tool.": {
    "zh-CN": "数据清理工具可根据功能作用于当前选区或整个数据区域。",
    ja: "データクリーニングツールは機能に応じて、現在の選択範囲またはデータ領域全体に適用できます。",
  },
  "Keyboard shortcuts": { "zh-CN": "快捷键", ja: "キーボードショートカット" },
  "Keyboard Shortcuts": { "zh-CN": "快捷键", ja: "キーボードショートカット" },
  "Ctrl+O: Open CWS HTML": { "zh-CN": "Ctrl+O：打开 CWS HTML", ja: "Ctrl+O: CWS HTML を開く" },
  "Ctrl+S: Save CWS HTML": { "zh-CN": "Ctrl+S：保存 CWS HTML", ja: "Ctrl+S: CWS HTML を保存" },
  "Ctrl+P: Print": { "zh-CN": "Ctrl+P：打印", ja: "Ctrl+P: 印刷" },
  "Ctrl+F / Ctrl+H: Find / Replace": { "zh-CN": "Ctrl+F / Ctrl+H：查找 / 替换", ja: "Ctrl+F / Ctrl+H: 検索 / 置換" },
  "Ctrl+Z / Ctrl+Y: Undo / Redo": { "zh-CN": "Ctrl+Z / Ctrl+Y：撤销 / 重做", ja: "Ctrl+Z / Ctrl+Y: 元に戻す / やり直す" },
  "Ctrl+A: Select all cells": { "zh-CN": "Ctrl+A：选择所有单元格", ja: "Ctrl+A: すべてのセルを選択" },
  Notes: { "zh-CN": "说明", ja: "メモ" },
  "This editor works with one sheet at a time.": { "zh-CN": "此编辑器一次处理一个工作表。", ja: "このエディターは一度に 1 つのシートを扱います。" },
  "Saved output stays in CWS HTML format.": { "zh-CN": "保存输出仍为 CWS HTML 格式。", ja: "保存される出力は CWS HTML 形式のままです。" },
  "Imported data is flattened to plain visible values.": { "zh-CN": "导入的数据会展平为普通可见值。", ja: "インポートしたデータは表示値としてフラット化されます。" },
  "Clipboard access can be limited in embedded or restricted browsers.": {
    "zh-CN": "嵌入式或受限浏览器中，剪贴板访问可能受限。",
    ja: "埋め込み環境や制限付きブラウザーでは、クリップボードアクセスが制限される場合があります。",
  },
  "Version panel": { "zh-CN": "版本面板", ja: "バージョンパネル" },
  "CWS Light Table Version": { "zh-CN": "CWS Light Table 版本", ja: "CWS Light Table バージョン" },
  "Version information": { "zh-CN": "版本信息", ja: "バージョン情報" },
  Product: { "zh-CN": "产品", ja: "製品" },
  Author: { "zh-CN": "作者", ja: "作成者" },
  "Lightweight CWS HTML single-sheet table editor with import, export, filtering, print, and structural editing tools.": {
    "zh-CN": "轻量 CWS HTML 单工作表表格编辑器，提供导入、导出、筛选、打印和结构编辑工具。",
    ja: "インポート、エクスポート、フィルター、印刷、構造編集ツールを備えた軽量 CWS HTML 単一シートテーブルエディターです。",
  },
  "Text import panel": { "zh-CN": "文本导入面板", ja: "テキストインポートパネル" },
  "Import Text / Structured Data": { "zh-CN": "导入文本 / 结构化数据", ja: "テキスト / 構造化データをインポート" },
  File: { "zh-CN": "文件", ja: "ファイル" },
  Type: { "zh-CN": "类型", ja: "種類" },
  Encoding: { "zh-CN": "编码", ja: "エンコーディング" },
  Delimiter: { "zh-CN": "分隔符", ja: "区切り文字" },
  Comma: { "zh-CN": "逗号", ja: "カンマ" },
  Tab: { "zh-CN": "制表符", ja: "タブ" },
  "Line-Based": { "zh-CN": "按行", ja: "行単位" },
  Preview: { "zh-CN": "预览", ja: "プレビュー" },
  "Re-read": { "zh-CN": "重新读取", ja: "再読み込み" },
  "Text export panel": { "zh-CN": "文本导出面板", ja: "テキストエクスポートパネル" },
  "Export Text Data": { "zh-CN": "导出文本数据", ja: "テキストデータをエクスポート" },
  Target: { "zh-CN": "目标", ja: "出力先" },
  "Row line ending": { "zh-CN": "行结束符", ja: "行末" },
  "Cell line breaks only LF": { "zh-CN": "单元格内换行仅使用 LF", ja: "セル内改行を LF のみにする" },
  "with BOM": { "zh-CN": "带 BOM", ja: "BOM 付き" },
  "Include hidden data": { "zh-CN": "包含隐藏数据", ja: "非表示データを含める" },
  "Quote all cells": { "zh-CN": "所有单元格加引号", ja: "すべてのセルを引用符で囲む" },
  "Excel import panel": { "zh-CN": "Excel 导入面板", ja: "Excel インポートパネル" },
  "Import Excel Sheet": { "zh-CN": "导入 Excel 工作表", ja: "Excel シートをインポート" },
  Workbook: { "zh-CN": "工作簿", ja: "ブック" },
  Sheet: { "zh-CN": "工作表", ja: "シート" },
  "Import Sheet": { "zh-CN": "导入工作表", ja: "シートをインポート" },
  "Find and replace panel": { "zh-CN": "查找和替换面板", ja: "検索と置換パネル" },
  "Find what": { "zh-CN": "查找内容", ja: "検索する文字列" },
  "Replace with": { "zh-CN": "替换为", ja: "置換後の文字列" },
  "Case Sensitive": { "zh-CN": "区分大小写", ja: "大文字と小文字を区別" },
  "Whole Cell": { "zh-CN": "匹配整个单元格", ja: "セル全体に一致" },
  Scope: { "zh-CN": "范围", ja: "範囲" },
  "Find Next": { "zh-CN": "查找下一个", ja: "次を検索" },
  Replace: { "zh-CN": "替换", ja: "置換" },
  "Replace All": { "zh-CN": "全部替换", ja: "すべて置換" },
  "Advanced filter panel": { "zh-CN": "高级筛选面板", ja: "高度なフィルターパネル" },
  Column: { "zh-CN": "列", ja: "列" },
  Operator: { "zh-CN": "操作符", ja: "演算子" },
  Value: { "zh-CN": "值", ja: "値" },
  Remove: { "zh-CN": "移除", ja: "削除" },
  "Add Condition": { "zh-CN": "添加条件", ja: "条件を追加" },
  "Reset to default logic": { "zh-CN": "重置为默认逻辑", ja: "既定のロジックに戻す" },
  "Logic expression": { "zh-CN": "逻辑表达式", ja: "ロジック式" },
  "Example:": { "zh-CN": "示例：", ja: "例:" },
  "Clear Advanced Filter": { "zh-CN": "清除高级筛选", ja: "高度なフィルターをクリア" },
  Apply: { "zh-CN": "应用", ja: "適用" },
  "Sort panel": { "zh-CN": "排序面板", ja: "並べ替えパネル" },
  Sort: { "zh-CN": "排序", ja: "並べ替え" },
  None: { "zh-CN": "无", ja: "なし" },
  "Primary column": { "zh-CN": "主要列", ja: "第 1 キー列" },
  Direction: { "zh-CN": "方向", ja: "方向" },
  "Secondary column": { "zh-CN": "次要列", ja: "第 2 キー列" },
  "Apply Sort": { "zh-CN": "应用排序", ja: "並べ替えを適用" },
  "Page setup panel": { "zh-CN": "页面设置面板", ja: "ページ設定パネル" },
  "Paper Size": { "zh-CN": "纸张大小", ja: "用紙サイズ" },
  Orientation: { "zh-CN": "方向", ja: "向き" },
  portrait: { "zh-CN": "纵向", ja: "縦" },
  landscape: { "zh-CN": "横向", ja: "横" },
  Top: { "zh-CN": "上", ja: "上" },
  Right: { "zh-CN": "右", ja: "右" },
  Bottom: { "zh-CN": "下", ja: "下" },
  Left: { "zh-CN": "左", ja: "左" },
  Header: { "zh-CN": "页眉", ja: "ヘッダー" },
  Footer: { "zh-CN": "页脚", ja: "フッター" },
  "Print Area": { "zh-CN": "打印区域", ja: "印刷範囲" },
  "Custom Range": { "zh-CN": "自定义范围", ja: "カスタム範囲" },
  Background: { "zh-CN": "背景", ja: "背景" },
  none: { "zh-CN": "无", ja: "なし" },
  "solid-color": { "zh-CN": "纯色", ja: "単色" },
  Color: { "zh-CN": "颜色", ja: "色" },
  "entire-sheet": { "zh-CN": "整个工作表", ja: "シート全体" },
  selection: { "zh-CN": "当前选择", ja: "選択範囲" },
  custom: { "zh-CN": "自定义", ja: "カスタム" },
  "Apply And Print": { "zh-CN": "应用并打印", ja: "適用して印刷" },
  "Grid limits panel": { "zh-CN": "网格限制面板", ja: "グリッド制限パネル" },
  "Maximum Rows": { "zh-CN": "最大行数", ja: "最大行数" },
  "Maximum Columns": { "zh-CN": "最大列数", ja: "最大列数" },
  "Header menu": { "zh-CN": "表头菜单", ja: "ヘッダーメニュー" },
  "Sort A -> Z": { "zh-CN": "按 A -> Z 排序", ja: "A -> Z で並べ替え" },
  "Sort Z -> A": { "zh-CN": "按 Z -> A 排序", ja: "Z -> A で並べ替え" },
  "Filter operator": { "zh-CN": "筛选操作符", ja: "フィルター演算子" },
  "Filter value": { "zh-CN": "筛选值", ja: "フィルター値" },
  "Clear Filter": { "zh-CN": "清除筛选", ja: "フィルターをクリア" },
  "Apply Filter": { "zh-CN": "应用筛选", ja: "フィルターを適用" },
  contains: { "zh-CN": "包含", ja: "含む" },
  equals: { "zh-CN": "等于", ja: "等しい" },
  "is empty": { "zh-CN": "为空", ja: "空" },
  "is not empty": { "zh-CN": "不为空", ja: "空ではない" },
  "Whole sheet": { "zh-CN": "整个工作表", ja: "シート全体" },
  "Column A": { "zh-CN": "A 列", ja: "列 A" },
  "No preview rows are available yet.": { "zh-CN": "暂无可预览的行。", ja: "プレビューできる行はまだありません。" },
  "A -> Z": { "zh-CN": "A -> Z", ja: "A -> Z" },
  "Z -> A": { "zh-CN": "Z -> A", ja: "Z -> A" },
  "BOM on": { "zh-CN": "BOM 开", ja: "BOM オン" },
  "BOM off": { "zh-CN": "BOM 关", ja: "BOM オフ" },
  "Quote all": { "zh-CN": "全部加引号", ja: "すべて引用" },
  "Quote minimal": { "zh-CN": "必要时加引号", ja: "必要最小限の引用" },
  "Include hidden": { "zh-CN": "包含隐藏行", ja: "非表示を含む" },
  "Visible rows only": { "zh-CN": "仅可见行", ja: "表示行のみ" },
  "Cell LF": { "zh-CN": "单元格 LF", ja: "セル LF" },
  "Save CWS HTML is unavailable when this file is opened directly from local file://. Use the browser download/open flow instead.": {
    "zh-CN": "通过本地 file:// 直接打开此文件时，无法使用保存 CWS HTML。请改用浏览器下载/打开流程。",
    ja: "このファイルをローカルの file:// から直接開いている場合、CWS HTML の保存は使用できません。ブラウザーのダウンロード / オープンの流れを使用してください。",
  },
  "Failed to load the selected file.": { "zh-CN": "无法加载所选文件。", ja: "選択したファイルを読み込めませんでした。" },
  "Failed to load the selected sheet.": { "zh-CN": "无法加载所选工作表。", ja: "選択したシートを読み込めませんでした。" },
  "The selected file is not recognized as CWS HTML.": { "zh-CN": "所选文件未被识别为 CWS HTML。", ja: "選択したファイルは CWS HTML として認識できません。" },
  "The CWS workbook JSON does not contain any sheets.": { "zh-CN": "CWS 工作簿 JSON 不包含任何工作表。", ja: "CWS ブック JSON にシートが含まれていません。" },
  "The CWS workbook model script was not found.": { "zh-CN": "未找到 CWS 工作簿模型脚本。", ja: "CWS ブックモデルスクリプトが見つかりません。" },
  "The workbook does not contain any sheets.": { "zh-CN": "工作簿不包含任何工作表。", ja: "ブックにシートが含まれていません。" },
  "The requested sheet could not be found.": { "zh-CN": "找不到请求的工作表。", ja: "指定されたシートが見つかりません。" },
  "The selected Excel workbook does not contain any sheets.": { "zh-CN": "所选 Excel 工作簿不包含任何工作表。", ja: "選択した Excel ブックにシートが含まれていません。" },
  "The selected file is not a supported text or structured-data import type.": {
    "zh-CN": "所选文件不是支持的文本或结构化数据导入类型。",
    ja: "選択したファイルは、サポートされているテキストまたは構造化データのインポート形式ではありません。",
  },
  "The selected file could not be parsed.": { "zh-CN": "无法解析所选文件。", ja: "選択したファイルを解析できませんでした。" },
  "Press Re-read after changing the encoding so the preview uses the original raw bytes.": {
    "zh-CN": "更改编码后请点击重新读取，以便预览使用原始字节。",
    ja: "エンコーディングを変更した後は、元の生バイトでプレビューするために再読み込みしてください。",
  },
  "There is no parsed data to import.": { "zh-CN": "没有可导入的解析数据。", ja: "インポートできる解析済みデータがありません。" },
  "The selected import did not contain any cells to insert.": { "zh-CN": "所选导入内容不包含可插入的单元格。", ja: "選択したインポート内容に挿入できるセルがありません。" },
  "Paste command is not available here. Use Ctrl+V / Cmd+V inside the grid.": {
    "zh-CN": "此处无法使用粘贴命令。请在网格内使用 Ctrl+V / Cmd+V。",
    ja: "ここでは貼り付けコマンドを使用できません。グリッド内で Ctrl+V / Cmd+V を使用してください。",
  },
  "Clipboard is empty.": { "zh-CN": "剪贴板为空。", ja: "クリップボードは空です。" },
  "Paste command is blocked here. Use Ctrl+V / Cmd+V inside the grid.": {
    "zh-CN": "此处粘贴命令被阻止。请在网格内使用 Ctrl+V / Cmd+V。",
    ja: "ここでは貼り付けコマンドがブロックされています。グリッド内で Ctrl+V / Cmd+V を使用してください。",
  },
  "Nothing to undo.": { "zh-CN": "没有可撤销的操作。", ja: "元に戻す操作はありません。" },
  "Nothing to redo.": { "zh-CN": "没有可重做的操作。", ja: "やり直す操作はありません。" },
  "Undid the last change.": { "zh-CN": "已撤销上一次更改。", ja: "直前の変更を元に戻しました。" },
  "Redid the last change.": { "zh-CN": "已重做上一次更改。", ja: "直前の変更をやり直しました。" },
  "Pasted clipboard range into the grid.": { "zh-CN": "已将剪贴板区域粘贴到网格中。", ja: "クリップボードの範囲をグリッドに貼り付けました。" },
  "Fill handle applied as copy-repeat.": { "zh-CN": "已按复制重复方式应用填充柄。", ja: "コピー繰り返しとしてフィルハンドルを適用しました。" },
  "Cleared the copied range highlight.": { "zh-CN": "已清除复制区域高亮。", ja: "コピー範囲のハイライトをクリアしました。" },
  "Advanced filter cleared.": { "zh-CN": "已清除高级筛选。", ja: "高度なフィルターをクリアしました。" },
  "Global search cleared.": { "zh-CN": "已清除全局搜索。", ja: "全体検索をクリアしました。" },
  "Applied the column filter.": { "zh-CN": "已应用列筛选。", ja: "列フィルターを適用しました。" },
  "Cleared the column filter.": { "zh-CN": "已清除列筛选。", ja: "列フィルターをクリアしました。" },
  "Cleared column and advanced filters.": { "zh-CN": "已清除列筛选和高级筛选。", ja: "列フィルターと高度なフィルターをクリアしました。" },
  "Applied the advanced filter.": { "zh-CN": "已应用高级筛选。", ja: "高度なフィルターを適用しました。" },
  "Applied the sort order.": { "zh-CN": "已应用排序顺序。", ja: "並べ替え順を適用しました。" },
  "Applied page setup settings.": { "zh-CN": "已应用页面设置。", ja: "ページ設定を適用しました。" },
  "Applied page setup and opened the browser print dialog.": {
    "zh-CN": "已应用页面设置，并打开浏览器打印对话框。",
    ja: "ページ設定を適用し、ブラウザーの印刷ダイアログを開きました。",
  },
  "Opened the browser print dialog.": { "zh-CN": "已打开浏览器打印对话框。", ja: "ブラウザーの印刷ダイアログを開きました。" },
  "Nothing to print from the current sheet.": { "zh-CN": "当前工作表没有可打印内容。", ja: "現在のシートには印刷できる内容がありません。" },
  "Printing canceled because the current print range is too large for stable browser printing.": {
    "zh-CN": "由于当前打印范围过大，无法稳定使用浏览器打印，已取消打印。",
    ja: "現在の印刷範囲が大きすぎてブラウザー印刷が安定しないため、印刷をキャンセルしました。",
  },
  "The current print range is large. Export to Excel before printing is recommended.": {
    "zh-CN": "当前打印范围较大，建议导出到 Excel 后打印。",
    ja: "現在の印刷範囲が大きいです。印刷前に Excel へエクスポートすることをおすすめします。",
  },
  "Enter text to find.": { "zh-CN": "请输入要查找的文本。", ja: "検索する文字列を入力してください。" },
  "No more matches were found in the current scope.": { "zh-CN": "当前范围内没有更多匹配项。", ja: "現在の範囲にこれ以上一致はありません。" },
  "No matches were found.": { "zh-CN": "未找到匹配项。", ja: "一致は見つかりませんでした。" },
  "Enter text to find before replacing.": { "zh-CN": "替换前请输入要查找的文本。", ja: "置換する前に検索する文字列を入力してください。" },
  "No match is available to replace.": { "zh-CN": "没有可替换的匹配项。", ja: "置換できる一致がありません。" },
  "No matches were found in the current scope.": { "zh-CN": "当前范围内未找到匹配项。", ja: "現在の範囲に一致は見つかりませんでした。" },
  "Enter a filter value or use Is Empty / Is Not Empty.": {
    "zh-CN": "请输入筛选值，或使用“为空 / 不为空”。",
    ja: "フィルター値を入力するか、「空 / 空ではない」を使用してください。",
  },
  "Choose a primary sort column.": { "zh-CN": "请选择主要排序列。", ja: "第 1 キー列を選択してください。" },
  "Enter a valid custom range such as A1:D20.": { "zh-CN": "请输入有效的自定义范围，例如 A1:D20。", ja: "A1:D20 のような有効なカスタム範囲を入力してください。" },
  "Maximum Rows must be an integer greater than or equal to 1.": {
    "zh-CN": "最大行数必须是大于等于 1 的整数。",
    ja: "最大行数は 1 以上の整数である必要があります。",
  },
  "Maximum Columns must be an integer greater than or equal to 1.": {
    "zh-CN": "最大列数必须是大于等于 1 的整数。",
    ja: "最大列数は 1 以上の整数である必要があります。",
  },
  "Enter a column count of 1 or greater.": { "zh-CN": "请输入大于等于 1 的列数。", ja: "1 以上の列数を入力してください。" },
  "Every defined condition number must appear exactly once.": {
    "zh-CN": "每个已定义的条件编号必须且只能出现一次。",
    ja: "定義済みの各条件番号は 1 回だけ指定する必要があります。",
  },
  "Condition number is outside the current range.": { "zh-CN": "条件编号超出当前范围。", ja: "条件番号が現在の範囲外です。" },
  "The logic expression contains unknown characters.": { "zh-CN": "逻辑表达式包含未知字符。", ja: "ロジック式に不明な文字が含まれています。" },
  "The logic expression has invalid operator placement.": { "zh-CN": "逻辑表达式的操作符位置无效。", ja: "ロジック式の演算子の位置が無効です。" },
  "The logic expression has unbalanced parentheses.": { "zh-CN": "逻辑表达式的括号不匹配。", ja: "ロジック式の括弧が対応していません。" },
  "Save canceled.": { "zh-CN": "保存已取消。", ja: "保存をキャンセルしました。" },
  "Export canceled.": { "zh-CN": "导出已取消。", ja: "エクスポートをキャンセルしました。" },
  "Save failed here because this browser does not support downloads. Open the app in a standard browser to save the CWS HTML file.": {
    "zh-CN": "保存失败，因为此浏览器不支持下载。请在标准浏览器中打开应用以保存 CWS HTML 文件。",
    ja: "このブラウザーはダウンロードに対応していないため保存できません。標準ブラウザーでアプリを開いて CWS HTML ファイルを保存してください。",
  },
  "Save failed because this browser blocked the file save.": {
    "zh-CN": "保存失败，因为浏览器阻止了文件保存。",
    ja: "ブラウザーがファイル保存をブロックしたため保存できませんでした。",
  },
  "Export failed here because this browser does not support downloads. Open the app in a standard browser to export the file.": {
    "zh-CN": "导出失败，因为此浏览器不支持下载。请在标准浏览器中打开应用以导出文件。",
    ja: "このブラウザーはダウンロードに対応していないためエクスポートできません。標準ブラウザーでアプリを開いてファイルをエクスポートしてください。",
  },
  "Export failed because this browser blocked the download.": {
    "zh-CN": "导出失败，因为浏览器阻止了下载。",
    ja: "ブラウザーがダウンロードをブロックしたためエクスポートできませんでした。",
  },
  "Export scope: the whole current sheet. Hidden-row handling is controlled by": {
    "zh-CN": "导出范围：整个当前工作表。隐藏行处理由",
    ja: "エクスポート範囲: 現在のシート全体。非表示行の扱いは",
  },
  "Hidden-row handling is controlled by": {
    "zh-CN": "隐藏行处理由",
    ja: "非表示行の扱いは",
  },
  "Press Re-read to refresh the preview from the original raw bytes using the selected encoding.": {
    "zh-CN": "点击重新读取，以使用所选编码从原始字节刷新预览。",
    ja: "選択したエンコーディングで元の生バイトからプレビューを更新するには、再読み込みを押してください。",
  },
  "Applies only to rows below the formal header row.": {
    "zh-CN": "仅应用于正式表头行下方的行。",
    ja: "正式なヘッダー行より下の行にのみ適用されます。",
  },
  "These values define the current editable grid size for this sheet.": {
    "zh-CN": "这些值定义当前工作表的可编辑网格大小。",
    ja: "これらの値は、このシートの現在の編集可能なグリッドサイズを定義します。",
  },
};

const MESSAGE_TRANSLATORS = [
  {
    pattern: /^(.+) \| (.+) \| header row (\d+) \| (\d+)\/(\d+) visible data rows \| (\d+) columns(.*)$/,
    build: ([, source, sheet, headerRow, visibleRows, totalRows, columns, suffix], language) => {
      const safeSource = source === "Unsaved blank sheet" ? translateText(source, language) : source;
      const tail = suffix ? translateMessage(suffix.trim().replace(/^\| /, ""), language) : "";
      const translated = language === "zh-CN"
        ? `${safeSource} | ${sheet} | 表头行 ${headerRow} | ${visibleRows}/${totalRows} 可见数据行 | ${columns} 列`
        : `${safeSource} | ${sheet} | ヘッダー行 ${headerRow} | 表示データ行 ${visibleRows}/${totalRows} | ${columns} 列`;
      return tail ? `${translated} | ${tail}` : translated;
    },
  },
  {
    pattern: /^viewport rows (\d+)-(\d+) \| viewport cols (\d+)-(\d+)$/,
    build: ([, r1, r2, c1, c2], language) => language === "zh-CN"
      ? `视口行 ${r1}-${r2} | 视口列 ${c1}-${c2}`
      : `表示行 ${r1}-${r2} | 表示列 ${c1}-${c2}`,
  },
  {
    pattern: /^(.*) Showing (\d+) of (\d+) data rows\.$/,
    build: ([, prefix, visible, total], language) => language === "zh-CN"
      ? `${translateMessage(prefix, language)} 正在显示 ${visible}/${total} 个数据行。`
      : `${translateMessage(prefix, language)} データ行 ${visible}/${total} 件を表示しています。`,
  },
  {
    pattern: /^Reading (.+)\.\.\.$/,
    build: ([, file], language) => language === "zh-CN" ? `正在读取 ${file}...` : `${file} を読み込んでいます...`,
  },
  {
    pattern: /^Global search filtered by "(.+)"\.$/,
    build: ([, query], language) => language === "zh-CN" ? `全局搜索已按“${query}”筛选。` : `全体検索を「${query}」で絞り込みました。`,
  },
  {
    pattern: /^Import anchor: (.+)\. Imported content always starts at the top-left cell of the current selection\.$/,
    build: ([, cell], language) => language === "zh-CN"
      ? `导入锚点：${cell}。导入内容始终从当前选择区域左上角单元格开始。`
      : `インポート基点: ${cell}。インポート内容は常に現在の選択範囲の左上セルから開始します。`,
  },
  {
    pattern: /^Import anchor: (.+)\. Selected sheet data will be written into the current document from that cell\.$/,
    build: ([, cell], language) => language === "zh-CN"
      ? `导入锚点：${cell}。所选工作表数据会从该单元格写入当前文档。`
      : `インポート基点: ${cell}。選択したシートデータはそのセルから現在のドキュメントに書き込まれます。`,
  },
  {
    pattern: /^Scope: (.+)$/,
    build: ([, scope], language) => language === "zh-CN" ? `范围：${translateMessage(scope, language)}` : `範囲: ${translateMessage(scope, language)}`,
  },
  {
    pattern: /^(.+) \/ first 6 rows$/,
    build: ([, target], language) => language === "zh-CN" ? `${target} / 前 6 行` : `${target} / 先頭 6 行`,
  },
  {
    pattern: /^Loaded (.+)\. Select one sheet to continue\.$/,
    build: ([, file], language) => language === "zh-CN" ? `已加载 ${file}。请选择一个工作表继续。` : `${file} を読み込みました。続行するシートを選択してください。`,
  },
  {
    pattern: /^Loaded (.+)\. Choose one sheet to import\.$/,
    build: ([, file], language) => language === "zh-CN" ? `已加载 ${file}。请选择一个工作表导入。` : `${file} を読み込みました。インポートするシートを選択してください。`,
  },
  {
    pattern: /^Loaded (.+)\. Review the preview and import it into the current sheet\.$/,
    build: ([, file], language) => language === "zh-CN" ? `已加载 ${file}。请检查预览并导入当前工作表。` : `${file} を読み込みました。プレビューを確認して現在のシートにインポートしてください。`,
  },
  {
    pattern: /^Loaded (.+) from (.+)\. The sheet is now shown in the grid\.$/,
    build: ([, sheet, file], language) => language === "zh-CN" ? `已从 ${file} 加载 ${sheet}。该工作表已显示在网格中。` : `${file} から ${sheet} を読み込みました。シートをグリッドに表示しています。`,
  },
  {
    pattern: /^Copied (\d+x\d+) range\.$/,
    build: ([, shape], language) => language === "zh-CN" ? `已复制 ${shape} 区域。` : `${shape} の範囲をコピーしました。`,
  },
  {
    pattern: /^Found a match at (.+)\.$/,
    build: ([, cell], language) => language === "zh-CN" ? `在 ${cell} 找到匹配项。` : `${cell} に一致が見つかりました。`,
  },
  {
    pattern: /^Replaced the match at (.+)\.$/,
    build: ([, cell], language) => language === "zh-CN" ? `已替换 ${cell} 的匹配项。` : `${cell} の一致を置換しました。`,
  },
  {
    pattern: /^Replaced (\d+) matches in the current scope\.$/,
    build: ([, count], language) => language === "zh-CN" ? `已在当前范围替换 ${count} 个匹配项。` : `現在の範囲で ${count} 件の一致を置換しました。`,
  },
  {
    pattern: /^Saved CWS HTML as (.+)\.$/,
    build: ([, file], language) => language === "zh-CN" ? `已保存 CWS HTML 为 ${file}。` : `CWS HTML を ${file} として保存しました。`,
  },
  {
    pattern: /^Started CWS HTML download as (.+)\. If no download appears, this browser may block file downloads\.$/,
    build: ([, file], language) => language === "zh-CN" ? `已开始下载 CWS HTML：${file}。如果没有下载出现，可能是浏览器阻止了文件下载。` : `CWS HTML のダウンロードを ${file} として開始しました。ダウンロードが表示されない場合、ブラウザーがファイルのダウンロードをブロックしている可能性があります。`,
  },
  {
    pattern: /^Exported (.+) as (.+)\.$/,
    build: ([, target, file], language) => language === "zh-CN" ? `已导出 ${target} 为 ${file}。` : `${target} を ${file} としてエクスポートしました。`,
  },
  {
    pattern: /^Started (.+) download as (.+)\. If no download appears, this browser may block file downloads\.$/,
    build: ([, target, file], language) => language === "zh-CN" ? `已开始下载 ${target}：${file}。如果没有下载出现，可能是浏览器阻止了文件下载。` : `${target} のダウンロードを ${file} として開始しました。ダウンロードが表示されない場合、ブラウザーがファイルのダウンロードをブロックしている可能性があります。`,
  },
  {
    pattern: /^Imported (.+) from (.+)\.$/,
    build: ([, sheet, file], language) => language === "zh-CN" ? `已从 ${file} 导入 ${sheet}。` : `${file} から ${sheet} をインポートしました。`,
  },
  {
    pattern: /^Imported (.+) data from (.+)\.$/,
    build: ([, kind, file], language) => language === "zh-CN" ? `已从 ${file} 导入 ${kind} 数据。` : `${file} から ${kind} データをインポートしました。`,
  },
  {
    pattern: /^(.*) Inserted (\d+x\d+) at (.+)\.$/,
    build: ([, prefix, shape, cell], language) => language === "zh-CN" ? `${translateMessage(prefix, language)} 已在 ${cell} 插入 ${shape}。` : `${translateMessage(prefix, language)} ${cell} に ${shape} を挿入しました。`,
  },
  {
    pattern: /^Set row (\d+) as the formal header row\.$/,
    build: ([, row], language) => language === "zh-CN" ? `已将第 ${row} 行设置为正式表头行。` : `${row} 行目を正式なヘッダー行に設定しました。`,
  },
  {
    pattern: /^Inserted (\d+) rows? (above|below) row (\d+)\.$/,
    build: ([, count, placement, row], language) => language === "zh-CN"
      ? `已在第 ${row} 行${placement === "above" ? "上方" : "下方"}插入 ${count} 行。`
      : `${row} 行目の${placement === "above" ? "上" : "下"}に ${count} 行を挿入しました。`,
  },
  {
    pattern: /^Inserted (\d+) columns? (left|right) column (.+)\.$/,
    build: ([, count, placement, column], language) => language === "zh-CN"
      ? `已在 ${column} 列${placement === "left" ? "左侧" : "右侧"}插入 ${count} 列。`
      : `${column} 列の${placement === "left" ? "左" : "右"}に ${count} 列を挿入しました。`,
  },
  {
    pattern: /^Inserted (\d+x\d+) blank area and shifted cells (right|down)\.$/,
    build: ([, shape, direction], language) => language === "zh-CN"
      ? `已插入 ${shape} 空白区域，并将单元格${direction === "right" ? "右移" : "下移"}。`
      : `${shape} の空白領域を挿入し、セルを${direction === "right" ? "右" : "下"}にシフトしました。`,
  },
  {
    pattern: /^Deleted (\d+) rows? starting at row (\d+)\.$/,
    build: ([, count, row], language) => language === "zh-CN" ? `已从第 ${row} 行开始删除 ${count} 行。` : `${row} 行目から ${count} 行を削除しました。`,
  },
  {
    pattern: /^Deleted (\d+) columns? starting at column (.+)\.$/,
    build: ([, count, column], language) => language === "zh-CN" ? `已从 ${column} 列开始删除 ${count} 列。` : `${column} 列から ${count} 列を削除しました。`,
  },
  {
    pattern: /^Cleared (\d+x\d+) selected cells\.$/,
    build: ([, shape], language) => language === "zh-CN" ? `已清空选中的 ${shape} 单元格。` : `選択した ${shape} セルをクリアしました。`,
  },
  {
    pattern: /^Removed (\d+) duplicate rows? from (\d+) checked rows\.$/,
    build: ([, removed, checked], language) => language === "zh-CN" ? `已从 ${checked} 行中删除 ${removed} 条重复行。` : `${checked} 行から ${removed} 件の重複行を削除しました。`,
  },
  {
    pattern: /^No duplicate rows found in (\d+) checked rows\.$/,
    build: ([, checked], language) => language === "zh-CN" ? `已检查 ${checked} 行，未发现重复行。` : `${checked} 行を確認しました。重複行はありません。`,
  },
  {
    pattern: /^Removed (\d+) blank rows? from (\d+) checked rows\.$/,
    build: ([, removed, checked], language) => language === "zh-CN" ? `已从 ${checked} 行中删除 ${removed} 条空白行。` : `${checked} 行から ${removed} 件の空白行を削除しました。`,
  },
  {
    pattern: /^No blank rows found in (\d+) checked rows\.$/,
    build: ([, checked], language) => language === "zh-CN" ? `已检查 ${checked} 行，未发现空白行。` : `${checked} 行を確認しました。空白行はありません。`,
  },
  {
    pattern: /^Cleaned spaces in (\d+) cells? from (\d+) checked cells\.$/,
    build: ([, changed, checked], language) => language === "zh-CN" ? `已检查 ${checked} 个单元格，清理 ${changed} 个单元格。` : `${checked} セルを確認し、${changed} セルの空白を整理しました。`,
  },
  {
    pattern: /^No spaces needed cleanup in (\d+) checked cells\.$/,
    build: ([, checked], language) => language === "zh-CN" ? `已检查 ${checked} 个单元格，没有需要清理的空格。` : `${checked} セルを確認しました。整理が必要な空白はありません。`,
  },
  {
    pattern: /^Converted text in (\d+) cells? from (\d+) checked cells\.$/,
    build: ([, changed, checked], language) => language === "zh-CN" ? `已检查 ${checked} 个单元格，转换 ${changed} 个单元格。` : `${checked} セルを確認し、${changed} セルを変換しました。`,
  },
  {
    pattern: /^No text conversion needed in (\d+) checked cells\.$/,
    build: ([, checked], language) => language === "zh-CN" ? `已检查 ${checked} 个单元格，没有需要转换的内容。` : `${checked} セルを確認しました。変換が必要な内容はありません。`,
  },
  {
    pattern: /^Updated number commas in (\d+) cells? from (\d+) numeric cells\.$/,
    build: ([, changed, numeric], language) => language === "zh-CN"
      ? `已从 ${numeric} 个数值单元格中更新 ${changed} 个单元格的逗号。`
      : `${numeric} 件の数値セルのうち ${changed} 件の桁区切りを更新しました。`,
  },
  {
    pattern: /^No number comma changes needed in (\d+) numeric cells\.$/,
    build: ([, numeric], language) => language === "zh-CN"
      ? `已检查 ${numeric} 个数值单元格，没有需要更新的逗号。`
      : `${numeric} 件の数値セルを確認しました。更新が必要な桁区切りはありません。`,
  },
  {
    pattern: /^Sorted (.+) (A -> Z|Z -> A)\.$/,
    build: ([, column, direction], language) => language === "zh-CN" ? `已按 ${column} ${direction} 排序。` : `${column} を ${direction} で並べ替えました。`,
  },
  {
    pattern: /^Updated the editable grid to (\d+) rows x (\d+) columns\.$/,
    build: ([, rows, columns], language) => language === "zh-CN" ? `已将可编辑网格更新为 ${rows} 行 x ${columns} 列。` : `編集可能なグリッドを ${rows} 行 x ${columns} 列に更新しました。`,
  },
  {
    pattern: /^Column label: (.+)$/,
    build: ([, label], language) => language === "zh-CN" ? `列标签：${label}` : `列ラベル: ${label}`,
  },
  {
    pattern: /^Selected sheet size: (\d+) rows x (\d+) columns\.$/,
    build: ([, rows, columns], language) => language === "zh-CN" ? `所选工作表大小：${rows} 行 x ${columns} 列。` : `選択したシートサイズ: ${rows} 行 x ${columns} 列。`,
  },
  {
    pattern: /^Summary: (.+)$/,
    build: ([, summary], language) => language === "zh-CN" ? `摘要：${translateExportSummary(summary, language)}` : `概要: ${translateExportSummary(summary, language)}`,
  },
  {
    pattern: /^Export failed: (.+)$/,
    build: ([, message], language) => language === "zh-CN" ? `导出失败：${translateMessage(message, language)}` : `エクスポートに失敗しました: ${translateMessage(message, language)}`,
  },
  {
    pattern: /^Save failed: (.+)$/,
    build: ([, message], language) => language === "zh-CN" ? `保存失败：${translateMessage(message, language)}` : `保存に失敗しました: ${translateMessage(message, language)}`,
  },
  {
    pattern: /^Condition (\d+) needs a value\.$/,
    build: ([, index], language) => language === "zh-CN" ? `条件 ${index} 需要填写值。` : `条件 ${index} には値が必要です。`,
  },
  {
    pattern: /^Condition (\d+) is repeated\.$/,
    build: ([, index], language) => language === "zh-CN" ? `条件 ${index} 重复出现。` : `条件 ${index} が重複しています。`,
  },
];

const MENU_DEFINITIONS = {
  file: [
    { label: "Open CWS HTML", action: "open", hint: "Ctrl+O" },
    { label: "Save CWS HTML", action: "save", hint: "Ctrl+S" },
    { label: "Import", submenuKey: "import" },
    { label: "Export", submenuKey: "export" },
    { label: "Print", action: "print", hint: "Ctrl+P" },
  ],
  import: [
    { label: "Excel", action: "import-excel" },
    { label: "Text / Structured Data", action: "import-text" },
  ],
  export: [
    { label: "CSV", action: "export-csv" },
    { label: "TSV", action: "export-tsv" },
    { label: "TXT", action: "export-txt" },
  ],
  search: [
    { label: "Find / Replace", action: "find-replace", hint: "Ctrl+F / Ctrl+H" },
  ],
  data: [
    { label: "Set Header Row", action: "set-header-row" },
    { label: "Advanced Filter", action: "advanced-filter" },
    { label: "Clear Filters", action: "clear-filters" },
    { label: "Clean Data", submenuKey: "clean-data" },
  ],
  "clean-data": [
    { label: "Remove Duplicates", action: "remove-duplicates" },
    { label: "Remove Blank Rows", action: "remove-blank-rows" },
    { label: "Trim Spaces", action: "trim-spaces" },
    { label: "Format Conversion", action: "normalize-case" },
    { label: "Amount commas", action: "number-commas" },
    { label: "Find Encoding Issues", action: "find-encoding-issues" },
  ],
  view: [
    { label: "Insert Row Above", action: "insert-row-above" },
    { label: "Insert Row Below", action: "insert-row-below" },
    { label: "Insert Column Left", action: "insert-column-left" },
    { label: "Insert Column Right", action: "insert-column-right" },
    { label: "Insert Cells", submenuKey: "insert-cells" },
    { label: "Delete Rows", action: "delete-rows" },
    { label: "Delete Columns", action: "delete-columns" },
  ],
  "insert-cells": [
    { label: "Shift Right", action: "insert-cells-right" },
    { label: "Shift Down", action: "insert-cells-down" },
  ],
  option: [
    { label: "Page Setup", action: "page-setup" },
    { label: "Grid Limits", action: "grid-limits" },
    { label: "Language", submenuKey: "language" },
  ],
  help: [
    { label: "Help", action: "help-page" },
    { label: "Version", action: "version-page" },
    { label: "Site", action: "help-site" },
  ],
};

const ICON_TOOLBAR_ITEMS = [
  { action: "open", label: "Open CWS HTML", icon: "open" },
  { action: "save", label: "Save CWS HTML", icon: "save" },
  { action: "print", label: "Print", icon: "print" },
  { action: "undo", label: "Undo", icon: "undo" },
  { action: "redo", label: "Redo", icon: "redo" },
  { action: "copy", label: "Copy Range", icon: "copy" },
  { action: "paste", label: "Paste Range", icon: "paste" },
  { action: "find-replace", label: "Find / Replace", icon: "find" },
  { action: "advanced-filter", label: "Advanced Filter", icon: "advanced-filter" },
  { action: "clear-filters", label: "Clear Filters", icon: "clear-filters" },
  { action: "grid-limits", label: "Grid Limits", icon: "grid-limits" },
];

const ICON_TOOLBAR_SVGS = {
  open: `
    <path d="M4 16.5V5.5a1 1 0 0 1 1-1h4.5l2 2H19a1 1 0 0 1 1 1v9a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1Z" />
    <path d="M9 12h6" />
    <path d="M12 9v6" />
  `,
  save: `
    <path d="M5 4.5h11l3 3V19a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1v-14a.5.5 0 0 1 .5-.5Z" />
    <path d="M8 4.5v5h7v-5" />
    <path d="M8 19v-5h8v5" />
  `,
  print: `
    <path d="M7 8.5v-4h10v4" />
    <path d="M7 15.5H6a2 2 0 0 1-2-2v-3a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v3a2 2 0 0 1-2 2h-1" />
    <path d="M7 13.5h10v6H7z" />
    <path d="M16.5 11h.01" />
  `,
  undo: `
    <path d="M9 8 5 12l4 4" />
    <path d="M6 12h8a5 5 0 1 1 0 10h-2" />
  `,
  redo: `
    <path d="m15 8 4 4-4 4" />
    <path d="M18 12h-8a5 5 0 1 0 0 10h2" />
  `,
  copy: `
    <rect x="9" y="9" width="9" height="11" rx="1.5" />
    <path d="M7 15H6a1 1 0 0 1-1-1V6.5A1.5 1.5 0 0 1 6.5 5H14a1 1 0 0 1 1 1v1" />
  `,
  paste: `
    <path d="M9 5.5h6" />
    <path d="M10 4h4a1 1 0 0 1 1 1v1h2a1 1 0 0 1 1 1v12a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h2V5a1 1 0 0 1 1-1Z" />
    <path d="M9 11.5h6" />
    <path d="M9 15.5h4" />
  `,
  find: `
    <circle cx="11" cy="11" r="5.5" />
    <path d="m15 15 4 4" />
  `,
  "advanced-filter": `
    <path fill="currentColor" stroke="none" d="M3.5 5.5A1.5 1.5 0 0 1 4.7 5h12.1a1.5 1.5 0 0 1 1.15 2.46L13.5 12.9V19a1.75 1.75 0 0 1-1.75 1.75h-.5A1.75 1.75 0 0 1 9.5 19v-6.1L3.05 7.46A1.5 1.5 0 0 1 3.5 5.5Z" />
    <circle cx="16.5" cy="12.5" r="4.25" fill="#fff" stroke="currentColor" stroke-width="1.6" />
    <text x="16.5" y="14.7" text-anchor="middle" font-size="7.5" font-weight="700" font-family="Segoe UI, Yu Gothic UI, sans-serif" fill="currentColor" stroke="none">A</text>
  `,
  "clear-filters": `
    <path fill="currentColor" stroke="none" d="M3.5 5.5A1.5 1.5 0 0 1 4.7 5h12.1a1.5 1.5 0 0 1 1.15 2.46L13.5 12.9V19a1.75 1.75 0 0 1-1.75 1.75h-.5A1.75 1.75 0 0 1 9.5 19v-6.1L3.05 7.46A1.5 1.5 0 0 1 3.5 5.5Z" />
    <circle cx="16.5" cy="12.5" r="4.25" fill="#fff" stroke="currentColor" stroke-width="1.6" />
    <path d="M14.2 12.5h4.6" />
  `,
  "grid-limits": `
    <path d="M8 4.5H4.5V8" />
    <path d="M16 4.5h3.5V8" />
    <path d="M8 19.5H4.5V16" />
    <path d="M16 19.5h3.5V16" />
    <path d="M9 9h6v6H9z" />
  `,
};

const refs = {
  menuButtons: Array.from(document.querySelectorAll("[data-menu-key]")),
  appMenu: document.getElementById("appMenu"),
  iconToolbar: document.getElementById("iconToolbar"),
  openFileInput: document.getElementById("openFileInput"),
  importFileInput: document.getElementById("importFileInput"),
  statusMessage: document.getElementById("statusMessage"),
  documentSummary: document.getElementById("documentSummary"),
  sheetPickerPanel: document.getElementById("sheetPickerPanel"),
  sheetSelect: document.getElementById("sheetSelect"),
  loadSheetButton: document.getElementById("loadSheetButton"),
  activeCellName: document.getElementById("activeCellName"),
  formulaBarInput: document.getElementById("formulaBarInput"),
  globalSearchInput: document.getElementById("globalSearchInput"),
  clearGlobalSearchButton: document.getElementById("clearGlobalSearchButton"),
  sheetGrid: document.getElementById("sheetGrid"),
  gridWrap: document.querySelector(".grid-wrap"),
  floatingLayer: document.getElementById("floatingLayer"),
  printHeader: document.getElementById("printHeader"),
  printFooter: document.getElementById("printFooter"),
};

let currentView = buildSheetGridView(state.table);
const debugGridEvents = new URLSearchParams(window.location.search).has("debug-grid");
let xlsxLoadPromise = null;
let codepageLoadPromise = null;

state.language = readStoredLanguage();

refs.menuButtons.forEach((button) => {
  button.addEventListener("click", handleMenuButtonClick);
});
refs.appMenu.addEventListener("click", handleMenuActionClick);
refs.iconToolbar?.addEventListener("click", handleIconToolbarClick);
refs.openFileInput.addEventListener("change", handleOpenFileChange);
refs.importFileInput.addEventListener("change", handleImportFileChange);
refs.loadSheetButton.addEventListener("click", () => {
  if (!state.workbook) return;
  loadSelectedSheet(Number(refs.sheetSelect.value || 0));
});
refs.sheetGrid.addEventListener("mousedown", handleGridMouseDown);
refs.sheetGrid.addEventListener("mousemove", handleGridPointerMove);
refs.sheetGrid.addEventListener("mouseover", handleGridMouseOver);
refs.sheetGrid.addEventListener("click", handleGridClick);
refs.sheetGrid.addEventListener("dblclick", handleGridDoubleClick);
refs.sheetGrid.addEventListener("input", handleGridInput);
refs.sheetGrid.addEventListener("keydown", handleGridEditorKeydown);
refs.gridWrap?.addEventListener("scroll", handleGridViewportScroll, { passive: true });
refs.formulaBarInput.addEventListener("focus", handleFormulaBarFocus);
refs.formulaBarInput.addEventListener("input", handleFormulaBarInput);
refs.formulaBarInput.addEventListener("keydown", handleFormulaBarKeydown);
refs.globalSearchInput.addEventListener("input", handleGlobalSearchInput);
refs.globalSearchInput.addEventListener("keydown", handleGlobalSearchKeydown);
refs.clearGlobalSearchButton.addEventListener("click", handleGlobalSearchClear);
refs.floatingLayer.addEventListener("click", handleFloatingLayerClick);
refs.floatingLayer.addEventListener("input", handleFloatingLayerInput);
refs.floatingLayer.addEventListener("change", handleFloatingLayerInput);
refs.floatingLayer.addEventListener("keydown", handleFloatingLayerKeydown);
document.addEventListener("click", handleDocumentClick);
document.addEventListener("mousemove", handleDocumentMouseMove);
document.addEventListener("mouseup", handleDocumentMouseUp);
document.addEventListener("copy", handleDocumentCopy);
document.addEventListener("paste", handleDocumentPaste);
document.addEventListener("keydown", handleDocumentKeydown);
window.addEventListener("keydown", handleGlobalUndoRedoKeydown, { capture: true });
window.addEventListener("resize", handleGridViewportResize);
window.addEventListener("afterprint", handleAfterPrint);

renderWorkspace();

function handleMenuButtonClick(event) {
  event.stopPropagation();
  const button = event.currentTarget;
  const menuKey = button.dataset.menuKey || "";
  const nextMenuKey = state.openMenuKey === menuKey ? null : menuKey;
  setMenuOpen(nextMenuKey);
}

async function handleMenuActionClick(event) {
  event.stopPropagation();
  const item = event.target.closest("[data-menu-action], [data-menu-submenu]");
  if (!item || item.getAttribute("aria-disabled") === "true") return;
  const submenuKey = item.dataset.menuSubmenu || "";
  if (submenuKey) {
    state.openSubmenuKey = state.openSubmenuKey === submenuKey ? "" : submenuKey;
    renderMenu(state.openMenuKey);
    return;
  }
  await executeMenuAction(item.dataset.menuAction || "");
}

async function handleIconToolbarClick(event) {
  event.stopPropagation();
  const button = event.target.closest("[data-toolbar-action]");
  if (!button || button.disabled) return;
  await executeMenuAction(button.dataset.toolbarAction || "");
}

async function executeMenuAction(action) {
  if (!action) return;
  if (state.openMenuKey) {
    setMenuOpen(null);
  }
  if (action.startsWith("language-")) {
    setLanguage(action.slice("language-".length));
  } else if (action === "open") {
    await handleOpenAction();
  } else if (action === "save") {
    if (isSaveActionDisabled()) {
      setStatus("Save CWS HTML is unavailable when this file is opened directly from local file://. Use the browser download/open flow instead.");
      return;
    }
    await handleSave();
  } else if (action === "import-excel") {
    openImportFilePicker("excel");
  } else if (action === "import-text") {
    openImportFilePicker("text");
  } else if (action === "export-csv") {
    openTextExportPanel("csv");
  } else if (action === "export-tsv") {
    openTextExportPanel("tsv");
  } else if (action === "export-txt") {
    openTextExportPanel("txt");
  } else if (action === "page-setup") {
    openPageSetupPanel();
  } else if (action === "grid-limits") {
    openGridLimitsPanel();
  } else if (action === "help-page") {
    openHelpPanel();
  } else if (action === "version-page") {
    openVersionPanel();
  } else if (action === "help-site") {
    openHelpSite();
  } else if (action === "print") {
    handlePrint();
  } else if (action === "undo") {
    handleUndo();
  } else if (action === "redo") {
    handleRedo();
  } else if (action === "copy") {
    await handleCopyCommand();
  } else if (action === "paste") {
    await handlePasteCommand();
  } else if (action === "find-replace") {
    openFindReplacePanel("find");
  } else if (action === "advanced-filter") {
    openAdvancedFilterPanel();
  } else if (action === "clear-filters") {
    handleClearFilters();
  } else if (action === "remove-duplicates") {
    openRemoveDuplicatesPanel();
  } else if (action === "remove-blank-rows") {
    openRemoveBlankRowsPanel();
  } else if (action === "trim-spaces") {
    openTrimSpacesPanel();
  } else if (action === "normalize-case") {
    openNormalizeCasePanel();
  } else if (action === "number-commas") {
    openNumberCommasPanel();
  } else if (action === "find-encoding-issues") {
    openFindEncodingIssuesPanel();
  } else if (action === "set-header-row") {
    handleSetHeaderRow();
  } else if (action === "insert-row-above") {
    handleInsertRows("above");
  } else if (action === "insert-row-below") {
    handleInsertRows("below");
  } else if (action === "insert-column-left") {
    handleInsertColumns("left");
  } else if (action === "insert-column-right") {
    handleInsertColumns("right");
  } else if (action === "insert-cells-right") {
    handleInsertCells("right");
  } else if (action === "insert-cells-down") {
    handleInsertCells("down");
  } else if (action === "delete-rows") {
    handleDeleteRows();
  } else if (action === "delete-columns") {
    handleDeleteColumns();
  } else if (action === "sort") {
    openSortPanel();
  }
}

async function handleOpenAction() {
  if (shouldUseFileHandleOpen()) {
    const didOpen = await openCwsHtmlWithFilePicker();
    if (didOpen) {
      return;
    }
  }
  refs.openFileInput.click();
}

async function handleOpenFileChange(event) {
  const [file] = event.currentTarget.files || [];
  event.currentTarget.value = "";
  if (!file) return;
  await loadCwsWorkbookFromFile(file);
}

async function openCwsHtmlWithFilePicker() {
  if (typeof window.showOpenFilePicker !== "function") {
    return false;
  }
  try {
    const [handle] = await window.showOpenFilePicker({
      multiple: false,
      types: [
        {
          description: "CWS HTML",
          accept: {
            "text/html": [".html", ".htm"],
          },
        },
      ],
    });
    if (!handle) {
      return false;
    }
    const file = await handle.getFile();
    await loadCwsWorkbookFromFile(file, { handle });
    return true;
  } catch (error) {
    if (error?.name === "AbortError") {
      return true;
    }
    return false;
  }
}

async function loadCwsWorkbookFromFile(file, { handle = null } = {}) {
  setMenuOpen(null);
  setStatus(`Reading ${file.name}...`);
  try {
    const source = await file.text();
    const workbook = parseCwsHtmlWorkbook(source, { fileName: file.name });
    state.workbook = workbook;
    state.sourceFileName = file.name;
    clearSaveFileHandle();
    const sheets = listWorkbookSheets(workbook);
    if (sheets.length > 1) {
      refs.sheetSelect.innerHTML = sheets
        .map((sheet) => `<option value="${sheet.index}">${escapeHtml(sheet.name)} (${sheet.rowCount}x${sheet.colCount})</option>`)
        .join("");
      refs.sheetPickerPanel.classList.remove("hidden");
      state.table = createEmptyLightTable();
      resetGridInteractionState();
      renderWorkspace();
      setStatus(`Loaded ${file.name}. Select one sheet to continue.`);
      return;
    }
    refs.sheetPickerPanel.classList.add("hidden");
    loadSelectedSheet(sheets[0]?.index ?? 0);
  } catch (error) {
    state.workbook = null;
    refs.sheetPickerPanel.classList.add("hidden");
    setStatus(error.message || "Failed to load the selected file.");
  }
}

async function handleImportFileChange(event) {
  const [file] = event.currentTarget.files || [];
  event.currentTarget.value = "";
  const importMode = state.importFileMode;
  state.importFileMode = "";
  if (!file || !importMode) return;
  setMenuOpen(null);
  try {
    if (importMode === "excel") {
      await beginExcelImport(file);
      return;
    }
    if (importMode === "text") {
      await beginTextImport(file);
    }
  } catch (error) {
    state.importDraft = null;
    state.openPanel = null;
    state.panelError = "";
    setStatus(error.message || `Failed to import ${file.name}.`);
    renderWorkspace();
  }
}

function loadSelectedSheet(sheetIndex) {
  try {
    const table = workbookToLightTable(state.workbook, sheetIndex);
    state.table = table;
    resetGridInteractionState();
    refs.sheetPickerPanel.classList.add("hidden");
    renderWorkspace();
    setStatus(`Loaded ${table.sheetName} from ${state.sourceFileName}. The sheet is now shown in the grid.`);
  } catch (error) {
    setStatus(error.message || "Failed to load the selected sheet.");
  }
}

function renderWorkspace() {
  rebuildGridRenderState({ honorPendingScroll: true });
  state.activeCell = clampActiveCell(state.activeCell, currentView);
  normalizeSelectionState();
  ensureActiveCellIsVisible();
  renderDocumentSummary();
  renderFormulaBar();
  renderGlobalSearch();
  renderIconToolbar();
  renderPrintChrome();
  renderGrid();
  renderFloatingLayer();
  if (state.openMenuKey) {
    renderMenu(state.openMenuKey);
  }
  syncLocalizedStaticUi();
  syncEditingFocus();
  updateGridViewportDebugData();
}

function readStoredLanguage() {
  try {
    const stored = window.localStorage?.getItem(LANGUAGE_STORAGE_KEY);
    return normalizeLanguage(stored);
  } catch (error) {
    return DEFAULT_LANGUAGE;
  }
}

function normalizeLanguage(language) {
  return SUPPORTED_LANGUAGES.includes(language) ? language : DEFAULT_LANGUAGE;
}

function currentLanguage() {
  return normalizeLanguage(state.language);
}

function translateText(text, language = currentLanguage()) {
  const source = String(text ?? "");
  if (language === DEFAULT_LANGUAGE) return source;
  return TEXT_TRANSLATIONS[source]?.[language] || source;
}

function translateMessage(message, language = currentLanguage()) {
  const source = String(message ?? "");
  if (language === DEFAULT_LANGUAGE || !source) return source;
  const exact = translateText(source, language);
  if (exact !== source) return exact;
  for (const translator of MESSAGE_TRANSLATORS) {
    const match = source.match(translator.pattern);
    if (match) {
      return translator.build(match, language);
    }
  }
  return source;
}

function translateExportSummary(summary, language = currentLanguage()) {
  return String(summary || "")
    .split(" / ")
    .map((part) => translateMessage(part, language))
    .join(" / ");
}

function localizeInlineUiText(text) {
  return escapeHtml(translateMessage(text));
}

function localizeAttrText(text) {
  return escapeAttr(translateMessage(text));
}

function setLanguage(language) {
  const nextLanguage = normalizeLanguage(language);
  if (state.language === nextLanguage) {
    setMenuOpen(null);
    return;
  }
  state.language = nextLanguage;
  try {
    window.localStorage?.setItem(LANGUAGE_STORAGE_KEY, nextLanguage);
  } catch (error) {
    // Ignore storage failures; the visible UI can still switch languages.
  }
  setMenuOpen(null);
  renderWorkspace();
  refs.statusMessage.textContent = translateMessage(state.statusMessageCanonical);
}

function syncLocalizedStaticUi() {
  document.documentElement.lang = LANGUAGE_HTML_LANG[currentLanguage()] || LANGUAGE_HTML_LANG.en;
  syncMenuButtonLabels();
  syncStaticElementText(refs.loadSheetButton, "Load Selected Sheet");
  syncStaticElementText(document.querySelector(".sheet-picker-bar .panel-note"), "This file has multiple sheets. Pick one sheet to load into the lightweight editor.");
  syncStaticElementText(document.querySelector(".quick-search-label"), "Search");
  syncStaticElementText(refs.clearGlobalSearchButton, "Clear");
  syncStaticElementAttr(document.querySelector(".menu-bar"), "aria-label", "Application menu");
  syncStaticElementAttr(refs.iconToolbar, "aria-label", "Quick actions");
  syncStaticElementAttr(refs.formulaBarInput, "aria-label", "Formula bar");
  syncStaticElementAttr(refs.globalSearchInput, "aria-label", "Global search");
  syncStaticElementAttr(refs.clearGlobalSearchButton, "aria-label", "Clear global search");
  syncStaticElementAttr(refs.sheetGrid, "aria-label", "Spreadsheet grid");
  localizeElement(refs.appMenu);
  localizeElement(refs.iconToolbar);
  localizeElement(refs.floatingLayer);
  refs.statusMessage.textContent = translateMessage(state.statusMessageCanonical);
}

function syncMenuButtonLabels() {
  refs.menuButtons.forEach((button) => {
    const menuKey = button.dataset.menuKey || "";
    const labels = {
      file: "File",
      edit: "Edit",
      search: "Search",
      data: "Data",
      view: "View",
      option: "Option",
      help: "Help",
    };
    if (labels[menuKey]) {
      button.textContent = translateText(labels[menuKey]);
    }
  });
}

function syncStaticElementText(element, sourceText) {
  if (element) {
    element.textContent = translateText(sourceText);
  }
}

function syncStaticElementAttr(element, attrName, sourceText) {
  if (element) {
    element.setAttribute(attrName, translateText(sourceText));
  }
}

function localizeElement(root) {
  if (!root || currentLanguage() === DEFAULT_LANGUAGE) return;
  const skipSelector = [
    "#sheetGrid",
    "#sheetSelect",
    "#excelImportSheetSelect",
    ".import-preview-table",
    ".text-export-preview",
    "[data-user-option='true']",
    "input",
    "textarea",
  ].join(", ");
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      const parent = node.parentElement;
      if (!parent || parent.closest(skipSelector)) {
        return NodeFilter.FILTER_REJECT;
      }
      return node.nodeValue.trim() ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT;
    },
  });
  const textNodes = [];
  while (walker.nextNode()) {
    textNodes.push(walker.currentNode);
  }
  textNodes.forEach((node) => {
    const source = node.nodeValue;
    const trimmed = source.trim();
    const translated = translateMessage(trimmed);
    node.nodeValue = translated === trimmed ? source : source.replace(trimmed, translated);
  });
  root.querySelectorAll("[aria-label], [title], [placeholder]").forEach((element) => {
    if (element.closest(skipSelector)) return;
    ["aria-label", "title", "placeholder"].forEach((attrName) => {
      if (element.hasAttribute(attrName)) {
        element.setAttribute(attrName, translateMessage(element.getAttribute(attrName)));
      }
    });
  });
}

function renderDocumentSummary() {
  const sourceLabel = state.sourceFileName || "Unsaved blank sheet";
  const sheetName = state.table?.sheetName || DEFAULT_SHEET_NAME;
  const columnCount = state.table?.columns?.length || 0;
  const rowCount = state.table?.rows?.length || 0;
  const visibleDataRows = countVisibleDataRows();
  const viewportSummary = debugGridEvents
    ? ` | viewport rows ${state.visibleRange.startRow}-${state.visibleRange.endRow} | viewport cols ${state.visibleRange.startColumn}-${state.visibleRange.endColumn}`
    : "";
  refs.documentSummary.textContent = translateMessage(`${sourceLabel} | ${sheetName} | header row ${getHeaderRowIndex(state.table)} | ${visibleDataRows}/${rowCount} visible data rows | ${columnCount} columns${viewportSummary}`);
  syncDocumentTitle();
}

function syncDocumentTitle() {
  document.title = state.sourceFileName || "CWS Light Table";
}

function renderFormulaBar() {
  refs.activeCellName.textContent = gridCellName(state.activeCell.row, state.activeCell.col);
  const formulaValue = state.editing ? state.editing.draft : getCellValue(state.table, state.activeCell.row, state.activeCell.col);
  if (refs.formulaBarInput.value !== formulaValue) {
    refs.formulaBarInput.value = formulaValue;
  }
  resizeTextarea(refs.formulaBarInput, 100);
}

function renderGlobalSearch() {
  if (refs.globalSearchInput.value !== state.globalSearch) {
    refs.globalSearchInput.value = state.globalSearch;
  }
}

function renderPrintChrome() {
  const pageSetup = state.table.pageSetup || {};
  refs.printHeader.textContent = pageSetup.headerFooter?.header || "";
  refs.printFooter.textContent = pageSetup.headerFooter?.footer || "";
  const backgroundColor = pageSetup.background?.mode === "solid-color" && pageSetup.background?.color
    ? pageSetup.background.color
    : "#fff";
  refs.gridWrap.style.background = backgroundColor;
}

function renderIconToolbar() {
  if (!refs.iconToolbar) return;
  refs.iconToolbar.innerHTML = ICON_TOOLBAR_ITEMS
    .map((item) => renderIconToolbarButton(item))
    .join("");
  localizeElement(refs.iconToolbar);
}

function renderIconToolbarButton(item) {
  const disabled = isToolbarActionDisabled(item.action);
  const iconMarkup = ICON_TOOLBAR_SVGS[item.icon] || "";
  return `
    <button
      class="icon-toolbar-button"
      type="button"
      data-toolbar-action="${escapeHtml(item.action)}"
      aria-label="${escapeHtml(item.label)}"
      title="${escapeHtml(item.label)}"
      ${disabled ? "disabled" : ""}
    >
      <svg class="icon-toolbar-svg" viewBox="0 0 24 24" aria-hidden="true" focusable="false">
        ${iconMarkup}
      </svg>
    </button>
  `;
}

function isToolbarActionDisabled(action) {
  if (action === "save") return isSaveActionDisabled();
  if (action === "undo") return !canUndo(state.history);
  if (action === "redo") return !canRedo(state.history);
  if (action === "copy" || action === "paste") return Boolean(state.editing);
  return false;
}

function renderGrid() {
  const printBounds = resolvePrintAreaBounds(state.table, state.selection, state.table.pageSetup || {});
  const copiedBounds = getCopiedSelectionBounds();
  const headerRowIndex = currentView.headerRowIndex;
  const tableRowCount = getSheetRowCount(state.table);
  const renderedColumnSpan = getRenderedGridColumnSpan();
  const topSpacerRow = renderVirtualSpacerRow(state.rowViewport.topSpacerHeight, renderedColumnSpan);
  const bottomSpacerRow = renderVirtualSpacerRow(state.rowViewport.bottomSpacerHeight, renderedColumnSpan);
  const leftHeaderSpacer = renderVirtualColumnSpacerCell(state.columnViewport.leftSpacerWidth, "th");
  const rightHeaderSpacer = renderVirtualColumnSpacerCell(state.columnViewport.rightSpacerWidth, "th");
  const headCells = currentView.columns
    .map((column) => {
      const isSelected = isColumnHeaderSelected(state.selection, column.index);
      return `<th scope="col" class="column-header${isSelected ? " is-selected" : ""}" data-col-header="true" data-col="${column.index}">${column.letter}</th>`;
    })
    .join("");
  const bodyRows = currentView.rows
    .map((row) => renderGridRow(row, printBounds, copiedBounds, headerRowIndex, tableRowCount))
    .join("");
  refs.sheetGrid.innerHTML = `
    <thead>
      <tr>
        <th class="corner-cell" aria-hidden="true"></th>
        ${leftHeaderSpacer}
        ${headCells}
        ${rightHeaderSpacer}
      </tr>
    </thead>
    <tbody>${topSpacerRow}${bodyRows}${bottomSpacerRow}</tbody>
  `;
}

function renderVirtualSpacerRow(height, columnCount) {
  const safeHeight = Math.max(0, Math.round(Number(height) || 0));
  if (!safeHeight) return "";
  return `<tr class="grid-spacer-row" data-print-hidden="true" aria-hidden="true"><td class="grid-spacer-cell" colspan="${Math.max(1, columnCount)}" style="height:${safeHeight}px"></td></tr>`;
}

function renderVirtualColumnSpacerCell(width, tagName = "td") {
  const safeWidth = Math.max(0, Math.round(Number(width) || 0));
  if (!safeWidth) return "";
  const safeTag = tagName === "th" ? "th" : "td";
  return `<${safeTag} class="grid-spacer-column" data-print-hidden="true" aria-hidden="true" style="width:${safeWidth}px;min-width:${safeWidth}px"></${safeTag}>`;
}

function renderGridRow(row, printBounds, copiedBounds, headerRowIndex, tableRowCount) {
  const isViewportPadding = isViewportPaddingRow(state.table, row.index) || row.index > tableRowCount;
  const isFormalHeader = row.index === headerRowIndex;
  const rowSelected = isRowHeaderSelected(state.selection, row.index);
  const printVisible = !isViewportPadding && rowWithinPrintArea(row.index, printBounds);
  const rowClasses = [
    isFormalHeader ? "print-table-header" : "",
  ].filter(Boolean).join(" ");
  const rowHeaderClass = [
    "row-header",
    rowSelected ? "is-selected" : "",
    isFormalHeader ? "is-formal-header" : "",
  ].filter(Boolean).join(" ");
  const leftColumnSpacer = renderVirtualColumnSpacerCell(state.columnViewport.leftSpacerWidth);
  const rightColumnSpacer = renderVirtualColumnSpacerCell(state.columnViewport.rightSpacerWidth);
  const cells = row.cells
    .map((cell) => renderGridCell(cell, isFormalHeader, printBounds, copiedBounds))
    .join("");
  return `<tr class="${rowClasses}"${!printVisible ? ' data-print-hidden="true"' : ""}><th scope="row" class="${rowHeaderClass}" data-row-header="true" data-row="${row.index}"${!printVisible ? ' data-print-hidden="true"' : ""}>${row.index}</th>${leftColumnSpacer}${cells}${rightColumnSpacer}</tr>`;
}

function renderGridCell(cell, isFormalHeader, printBounds, copiedBounds) {
  const isActive = cell.row === state.activeCell.row && cell.col === state.activeCell.col;
  const isSelected = isCellSelected(state.selection, cell.row, cell.col);
  const isEditingCell = Boolean(state.editing && state.editing.mode === "cell" && isActive);
  const isFillPreview = isFillPreviewCell(cell.row, cell.col);
  const copiedRangeClasses = getCopiedRangeCellClasses(copiedBounds, cell.row, cell.col);
  const copiedRangeOutline = copiedRangeClasses.length
    ? `<span class="copied-range-outline ${copiedRangeClasses.join(" ")}" aria-hidden="true"></span>`
    : "";
  const hasFillHandle = shouldRenderFillHandle(cell.row, cell.col);
  const isFindHit = Boolean(state.findReplace.currentMatch && sameCell(state.findReplace.currentMatch, cell));
  const printHidden = !isCellWithinBounds(cell.row, cell.col, printBounds);
  const className = [
    "grid-cell",
    isSelected ? "is-selected" : "",
    isFillPreview ? "is-fill-preview" : "",
    isActive ? "is-active" : "",
    isFindHit ? "is-find-hit" : "",
    isFormalHeader ? "is-formal-header" : "",
    ...copiedRangeClasses,
  ].filter(Boolean).join(" ");
  if (isEditingCell) {
    return `<td class="${className}" data-cell="true" data-row="${cell.row}" data-col="${cell.col}" aria-selected="${isSelected ? "true" : "false"}"${printHidden ? ' data-print-hidden="true"' : ""}>${copiedRangeOutline}<textarea id="activeCellEditor" class="cell-editor" rows="1" spellcheck="false">${escapeHtml(state.editing.draft)}</textarea></td>`;
  }
  const headerTools = isFormalHeader
    ? `<span class="header-tools"><button class="header-tools-button${state.columnFilters[columnKeyForIndex(cell.col)] ? " is-filtered" : ""}" type="button" data-header-menu-button="true" aria-label="Open header tools">▼</button></span>`
    : "";
  return `<td class="${className}" data-cell="true" data-row="${cell.row}" data-col="${cell.col}" aria-selected="${isSelected ? "true" : "false"}"${printHidden ? ' data-print-hidden="true"' : ""}>${copiedRangeOutline}${escapeHtml(cell.value)}${headerTools}${hasFillHandle ? '<button class="fill-handle" type="button" data-fill-handle="true" aria-label="Fill handle"></button>' : ""}</td>`;
}

function renderFloatingLayer() {
  const panels = [];
  if (state.openPanel === "find-replace") {
    panels.push(renderFindReplacePanel());
  }
  if (state.openPanel === "text-import") {
    panels.push(renderTextImportPanel());
  }
  if (state.openPanel === "text-export") {
    panels.push(renderTextExportPanel());
  }
  if (state.openPanel === "excel-import") {
    panels.push(renderExcelImportPanel());
  }
  if (state.openPanel === "advanced-filter") {
    panels.push(renderAdvancedFilterPanel());
  }
  if (state.openPanel === "remove-duplicates") {
    panels.push(renderRemoveDuplicatesPanel());
  }
  if (state.openPanel === "remove-blank-rows") {
    panels.push(renderRemoveBlankRowsPanel());
  }
  if (state.openPanel === "trim-spaces") {
    panels.push(renderTrimSpacesPanel());
  }
  if (state.openPanel === "normalize-case") {
    panels.push(renderNormalizeCasePanel());
  }
  if (state.openPanel === "number-commas") {
    panels.push(renderNumberCommasPanel());
  }
  if (state.openPanel === "find-encoding-issues") {
    panels.push(renderFindEncodingIssuesPanel());
  }
  if (state.openPanel === "sort") {
    panels.push(renderSortPanel());
  }
  if (state.openPanel === "page-setup") {
    panels.push(renderPageSetupPanel());
  }
  if (state.openPanel === "grid-limits") {
    panels.push(renderGridLimitsPanel());
  }
  if (state.openPanel === "help-page") {
    panels.push(renderHelpPanel());
  }
  if (state.openPanel === "version-page") {
    panels.push(renderVersionPanel());
  }
  if (state.headerMenu) {
    panels.push(renderHeaderMenu());
  }
  refs.floatingLayer.innerHTML = panels.join("");
  localizeElement(refs.floatingLayer);
}

function renderHelpPanel() {
  return `
    <section class="floating-panel floating-panel-wide" data-panel-type="help-page" role="dialog" aria-label="Help panel">
      <div class="floating-panel-header">
        <h2 class="floating-panel-title">CWS Light Table Help</h2>
        <button class="floating-panel-close" type="button" data-panel-action="close-panel">Close</button>
      </div>
      <p class="floating-panel-meta">CWS Light Table is a lightweight single-sheet table editor for CWS HTML.</p>
      <section class="help-panel-section" aria-label="Key features">
        <h3 class="help-panel-heading">Key Features</h3>
        <ul class="help-panel-list">
          <li>Open and save CWS HTML files.</li>
          <li>Import Excel and text or structured data into the grid.</li>
          <li>Edit cells with the formula bar, copy and paste, fill handle, and undo or redo.</li>
          <li>Use find and replace, header rows, filters, print setup, and export tools.</li>
          <li>Adjust grid size and insert rows, columns, or cells.</li>
        </ul>
      </section>
      <section class="help-panel-section" aria-label="Data Cleaning">
        <h3 class="help-panel-heading">Data Cleaning</h3>
        <ul class="help-panel-list">
          <li>Remove duplicate rows by one selected column or by a multi-column comparison key.</li>
          <li>Remove fully blank rows from the current selection or the whole data area.</li>
          <li>Trim leading, trailing, and inner spaces with configurable inner-space handling.</li>
          <li>Convert text case and width, including lowercase, uppercase, capitalized words, full-width, and half-width.</li>
          <li>Add or remove thousands separators for numeric strings while skipping non-numeric text.</li>
          <li>Find suspicious encoding issues such as mojibake, replacement characters, and control characters without modifying data.</li>
          <li>Data cleaning tools can run on the current selection or the whole data area depending on the tool.</li>
        </ul>
      </section>
      <section class="help-panel-section" aria-label="Keyboard shortcuts">
        <h3 class="help-panel-heading">Keyboard Shortcuts</h3>
        <ul class="help-panel-list">
          <li>Ctrl+O: Open CWS HTML</li>
          <li>Ctrl+S: Save CWS HTML</li>
          <li>Ctrl+P: Print</li>
          <li>Ctrl+F / Ctrl+H: Find / Replace</li>
          <li>Ctrl+Z / Ctrl+Y: Undo / Redo</li>
          <li>Ctrl+A: Select all cells</li>
        </ul>
      </section>
      <section class="help-panel-section" aria-label="Notes">
        <h3 class="help-panel-heading">Notes</h3>
        <ul class="help-panel-list">
          <li>This editor works with one sheet at a time.</li>
          <li>Saved output stays in CWS HTML format.</li>
          <li>Imported data is flattened to plain visible values.</li>
          <li>Clipboard access can be limited in embedded or restricted browsers.</li>
        </ul>
      </section>
    </section>
  `;
}

function renderVersionPanel() {
  return `
    <section class="floating-panel" data-panel-type="version-page" role="dialog" aria-label="Version panel">
      <div class="floating-panel-header">
        <h2 class="floating-panel-title">CWS Light Table Version</h2>
        <button class="floating-panel-close" type="button" data-panel-action="close-panel">Close</button>
      </div>
      <div class="help-panel-facts" aria-label="Version information">
        <div class="help-panel-fact-row">
          <span class="help-panel-fact-label">Product</span>
          <strong>CWS Light Table</strong>
        </div>
        <div class="help-panel-fact-row">
          <span class="help-panel-fact-label">Version</span>
          <strong>${APP_VERSION}</strong>
        </div>
        <div class="help-panel-fact-row">
          <span class="help-panel-fact-label">Author</span>
          <strong>Mac Su</strong>
        </div>
      </div>
      <p class="floating-panel-meta">Lightweight CWS HTML single-sheet table editor with import, export, filtering, print, and structural editing tools.</p>
    </section>
  `;
}

function renderRemoveDuplicatesPanel() {
  const draft = state.removeDuplicatesDraft || createRemoveDuplicatesDraft();
  const normalizedDraft = normalizeCleaningDraftColumns(draft);
  state.removeDuplicatesDraft = normalizedDraft;
  const rows = resolveCleaningRows(normalizedDraft.scope, { skipHeader: true });
  const columns = getSelectedCleaningColumns(normalizedDraft);
  const preview = previewDuplicateRows(state.table, { rows, columns });
  const disabled = !rows.length || !columns.length;
  return `
    <section class="floating-panel floating-panel-wide" data-panel-type="remove-duplicates" role="dialog" aria-label="Remove Duplicates">
      <div class="floating-panel-header">
        <h2 class="floating-panel-title">Remove Duplicates</h2>
        <button class="floating-panel-close" type="button" data-panel-action="close-panel">Close</button>
      </div>
      <label>
        <span>Scope</span>
        <select id="removeDuplicatesScope">
          ${renderCleaningScopeOptions(normalizedDraft.scope)}
        </select>
      </label>
      <p class="floating-panel-meta">Header row is not included when removing duplicates.</p>
      <p class="floating-panel-meta">Entire data area means rows below the header row.</p>
      <section class="help-panel-section" aria-label="Columns To Compare">
        <h3 class="help-panel-heading">Columns To Compare</h3>
        <div class="floating-panel-actions">
          <button class="button" type="button" data-panel-action="remove-duplicates-select-all">Select All</button>
          <button class="button" type="button" data-panel-action="remove-duplicates-clear-columns">Clear Selection</button>
        </div>
        <div class="floating-panel-checks">
          ${renderCleaningColumnCheckboxes(normalizedDraft, "removeDuplicatesColumn")}
        </div>
      </section>
      <section class="help-panel-section" aria-label="Duplicate Preview">
        <h3 class="help-panel-heading">Duplicate Preview</h3>
        <div class="help-panel-facts">
          <div class="help-panel-fact-row"><span class="help-panel-fact-label">Rows to check</span><strong>${preview.checkedRowCount}</strong></div>
          <div class="help-panel-fact-row"><span class="help-panel-fact-label">Rows to remove</span><strong>${preview.removedRowCount}</strong></div>
        </div>
      </section>
      ${state.panelError ? `<p class="floating-panel-error">${escapeHtml(state.panelError)}</p>` : ""}
      <div class="floating-panel-actions">
        <button class="button button-primary" type="button" data-panel-action="remove-duplicates-apply"${disabled ? ' aria-disabled="true"' : ""}>Apply</button>
      </div>
    </section>
  `;
}

function renderRemoveBlankRowsPanel() {
  const draft = state.removeBlankRowsDraft || createRemoveBlankRowsDraft();
  const normalizedDraft = normalizeRemoveBlankRowsDraft(draft);
  state.removeBlankRowsDraft = normalizedDraft;
  const rows = resolveCleaningRows(normalizedDraft.scope, { skipHeader: true });
  const columns = getBlankRowColumns();
  const preview = previewBlankRows(state.table, { rows, columns });
  const disabled = !rows.length || !columns.length;
  return `
    <section class="floating-panel floating-panel-wide" data-panel-type="remove-blank-rows" role="dialog" aria-label="Remove Blank Rows">
      <div class="floating-panel-header">
        <h2 class="floating-panel-title">Remove Blank Rows</h2>
        <button class="floating-panel-close" type="button" data-panel-action="close-panel">Close</button>
      </div>
      <label>
        <span>Scope</span>
        <select id="removeBlankRowsScope">
          ${renderCleaningScopeOptions(normalizedDraft.scope)}
        </select>
      </label>
      <p class="floating-panel-meta">Header row is not included when removing blank rows.</p>
      <p class="floating-panel-meta">Entire data area means rows below the header row.</p>
      <p class="floating-panel-meta">A row is blank only when every cell in the row is empty.</p>
      <section class="help-panel-section" aria-label="Blank Row Preview">
        <h3 class="help-panel-heading">Blank Row Preview</h3>
        <div class="help-panel-facts">
          <div class="help-panel-fact-row"><span class="help-panel-fact-label">Rows to check</span><strong>${preview.checkedRowCount}</strong></div>
          <div class="help-panel-fact-row"><span class="help-panel-fact-label">Rows to remove</span><strong>${preview.removedRowCount}</strong></div>
        </div>
      </section>
      ${state.panelError ? `<p class="floating-panel-error">${escapeHtml(state.panelError)}</p>` : ""}
      <div class="floating-panel-actions">
        <button class="button button-primary" type="button" data-panel-action="remove-blank-rows-apply"${disabled ? ' aria-disabled="true"' : ""}>Apply</button>
      </div>
    </section>
  `;
}

function renderTrimSpacesPanel() {
  const draft = state.trimSpacesDraft || createTrimSpacesDraft();
  const normalizedDraft = normalizeTrimSpacesDraft(draft);
  state.trimSpacesDraft = normalizedDraft;
  const rows = resolveCleaningRows(normalizedDraft.scope, { skipHeader: normalizedDraft.scope !== "selection" });
  const columns = getTrimSpacesColumns(normalizedDraft.scope);
  const preview = previewTrimSpaces(state.table, {
    rows,
    columns,
    leading: normalizedDraft.leading,
    trailing: normalizedDraft.trailing,
    inner: normalizedDraft.inner,
    innerMode: normalizedDraft.innerMode,
  });
  const disabled = !rows.length || !columns.length || !hasTrimSpaceOption(normalizedDraft);
  return `
    <section class="floating-panel floating-panel-wide" data-panel-type="trim-spaces" role="dialog" aria-label="Trim Spaces">
      <div class="floating-panel-header">
        <h2 class="floating-panel-title">Trim Spaces</h2>
        <button class="floating-panel-close" type="button" data-panel-action="close-panel">Close</button>
      </div>
      <label>
        <span>Scope</span>
        <select id="trimSpacesScope">
          ${renderCleaningScopeOptions(normalizedDraft.scope)}
        </select>
      </label>
      <p class="floating-panel-meta">Entire data area means rows below the header row.</p>
      <section class="help-panel-section" aria-label="Spaces To Clean">
        <h3 class="help-panel-heading">Spaces To Clean</h3>
        <div class="floating-panel-checks">
          <label><input id="trimSpacesLeadingInput" type="checkbox"${normalizedDraft.leading ? " checked" : ""} /> Leading spaces</label>
          <label><input id="trimSpacesInnerInput" type="checkbox"${normalizedDraft.inner ? " checked" : ""} /> Inner spaces</label>
          <label><input id="trimSpacesTrailingInput" type="checkbox"${normalizedDraft.trailing ? " checked" : ""} /> Trailing spaces</label>
        </div>
        <p class="floating-panel-meta">Spaces include half-width spaces, full-width spaces, and Tab.</p>
      </section>
      <section class="help-panel-section" aria-label="Inner Space Mode">
        <h3 class="help-panel-heading">Inner Space Mode</h3>
        <div class="floating-panel-checks">
          <label><input name="trimSpacesInnerMode" type="radio" value="collapse"${normalizedDraft.innerMode === "collapse" ? " checked" : ""}${normalizedDraft.inner ? "" : " disabled"} /> Collapse consecutive spaces to one</label>
          <label><input name="trimSpacesInnerMode" type="radio" value="remove"${normalizedDraft.innerMode === "remove" ? " checked" : ""}${normalizedDraft.inner ? "" : " disabled"} /> Remove all inner spaces</label>
        </div>
      </section>
      <section class="help-panel-section" aria-label="Trim Preview">
        <h3 class="help-panel-heading">Trim Preview</h3>
        <div class="help-panel-facts">
          <div class="help-panel-fact-row"><span class="help-panel-fact-label">Cells to check</span><strong>${preview.checkedCellCount}</strong></div>
          <div class="help-panel-fact-row"><span class="help-panel-fact-label">Cells to change</span><strong>${preview.changedCellCount}</strong></div>
        </div>
      </section>
      ${state.panelError ? `<p class="floating-panel-error">${escapeHtml(state.panelError)}</p>` : ""}
      <div class="floating-panel-actions">
        <button class="button button-primary" type="button" data-panel-action="trim-spaces-apply"${disabled ? ' aria-disabled="true"' : ""}>Apply</button>
      </div>
    </section>
  `;
}

function renderNormalizeCasePanel() {
  const draft = state.normalizeCaseDraft || createNormalizeCaseDraft();
  const normalizedDraft = normalizeNormalizeCaseDraft(draft);
  state.normalizeCaseDraft = normalizedDraft;
  const rows = resolveCleaningRows(normalizedDraft.scope, { skipHeader: normalizedDraft.scope !== "selection" });
  const columns = getTrimSpacesColumns(normalizedDraft.scope);
  const preview = previewNormalizeCase(state.table, {
    rows,
    columns,
    mode: normalizedDraft.mode,
  });
  const disabled = !rows.length || !columns.length;
  return `
    <section class="floating-panel floating-panel-wide" data-panel-type="normalize-case" role="dialog" aria-label="Format Conversion">
      <div class="floating-panel-header">
        <h2 class="floating-panel-title">Format Conversion</h2>
        <button class="floating-panel-close" type="button" data-panel-action="close-panel">Close</button>
      </div>
      <label>
        <span>Scope</span>
        <select id="normalizeCaseScope">
          ${renderCleaningScopeOptions(normalizedDraft.scope)}
        </select>
      </label>
      <p class="floating-panel-meta">Entire data area means rows below the header row.</p>
      <section class="help-panel-section" aria-label="Conversion Mode">
        <h3 class="help-panel-heading">Conversion Mode</h3>
        <div class="floating-panel-checks floating-panel-checks-stacked">
          <div class="floating-panel-check-row">
            <label><input name="normalizeCaseMode" type="radio" value="lower"${normalizedDraft.mode === "lower" ? " checked" : ""} /> lowercase</label>
            <label><input name="normalizeCaseMode" type="radio" value="upper"${normalizedDraft.mode === "upper" ? " checked" : ""} /> UPPERCASE</label>
            <label><input name="normalizeCaseMode" type="radio" value="capitalize-words"${normalizedDraft.mode === "capitalize-words" ? " checked" : ""} /> Capitalize Words</label>
          </div>
          <div class="floating-panel-check-row">
            <label><input name="normalizeCaseMode" type="radio" value="full-width"${normalizedDraft.mode === "full-width" ? " checked" : ""} /> Full-width</label>
            <label><input name="normalizeCaseMode" type="radio" value="half-width"${normalizedDraft.mode === "half-width" ? " checked" : ""} /> Half-width</label>
          </div>
        </div>
      </section>
      <section class="help-panel-section" aria-label="Conversion Preview">
        <h3 class="help-panel-heading">Conversion Preview</h3>
        <div class="help-panel-facts">
          <div class="help-panel-fact-row"><span class="help-panel-fact-label">Cells to check</span><strong>${preview.checkedCellCount}</strong></div>
          <div class="help-panel-fact-row"><span class="help-panel-fact-label">Cells to change</span><strong>${preview.changedCellCount}</strong></div>
        </div>
      </section>
      ${state.panelError ? `<p class="floating-panel-error">${escapeHtml(state.panelError)}</p>` : ""}
      <div class="floating-panel-actions">
        <button class="button button-primary" type="button" data-panel-action="normalize-case-apply"${disabled ? ' aria-disabled="true"' : ""}>Apply</button>
      </div>
    </section>
  `;
}

function renderNumberCommasPanel() {
  const draft = state.numberCommasDraft || createNumberCommasDraft();
  const normalizedDraft = normalizeNumberCommasDraft(draft);
  state.numberCommasDraft = normalizedDraft;
  const rows = resolveCleaningRows(normalizedDraft.scope, { skipHeader: normalizedDraft.scope !== "selection" });
  const columns = getTrimSpacesColumns(normalizedDraft.scope);
  const preview = previewNumberCommas(state.table, {
    rows,
    columns,
    mode: normalizedDraft.mode,
  });
  const disabled = !rows.length || !columns.length;
  return `
    <section class="floating-panel floating-panel-wide" data-panel-type="number-commas" role="dialog" aria-label="Amount commas">
      <div class="floating-panel-header">
        <h2 class="floating-panel-title">Amount commas</h2>
        <button class="floating-panel-close" type="button" data-panel-action="close-panel">Close</button>
      </div>
      <label>
        <span>Scope</span>
        <select id="numberCommasScope">
          ${renderCleaningScopeOptions(normalizedDraft.scope)}
        </select>
      </label>
      <p class="floating-panel-meta">Entire data area means rows below the header row.</p>
      <p class="floating-panel-meta">Only numeric strings are changed. Invalid numbers, IDs with leading zeros, currency symbols, and text are skipped.</p>
      <section class="help-panel-section" aria-label="Comma Mode">
        <h3 class="help-panel-heading">Comma Mode</h3>
        <div class="floating-panel-checks">
          <label><input name="numberCommasMode" type="radio" value="add"${normalizedDraft.mode === "add" ? " checked" : ""} /> Add commas</label>
          <label><input name="numberCommasMode" type="radio" value="remove"${normalizedDraft.mode === "remove" ? " checked" : ""} /> Remove commas</label>
        </div>
      </section>
      <section class="help-panel-section" aria-label="Amount comma preview">
        <h3 class="help-panel-heading">Amount comma preview</h3>
        <div class="help-panel-facts">
          <div class="help-panel-fact-row"><span class="help-panel-fact-label">Cells to check</span><strong>${preview.checkedCellCount}</strong></div>
          <div class="help-panel-fact-row"><span class="help-panel-fact-label">Numeric cells</span><strong>${preview.numericCellCount}</strong></div>
          <div class="help-panel-fact-row"><span class="help-panel-fact-label">Cells to change</span><strong>${preview.changedCellCount}</strong></div>
        </div>
      </section>
      ${state.panelError ? `<p class="floating-panel-error">${escapeHtml(state.panelError)}</p>` : ""}
      <div class="floating-panel-actions">
        <button class="button button-primary" type="button" data-panel-action="number-commas-apply"${disabled ? ' aria-disabled="true"' : ""}>Apply</button>
      </div>
    </section>
  `;
}

function renderFindEncodingIssuesPanel() {
  const draft = state.encodingIssuesDraft || createFindEncodingIssuesDraft();
  const normalizedDraft = normalizeFindEncodingIssuesDraft(draft);
  state.encodingIssuesDraft = normalizedDraft;
  const rows = resolveCleaningRows(normalizedDraft.scope, { skipHeader: normalizedDraft.scope !== "selection" });
  const columns = getTrimSpacesColumns(normalizedDraft.scope);
  const preview = findEncodingIssuesInCells(state.table, {
    rows,
    columns,
    limit: 20,
  });
  return `
    <section class="floating-panel floating-panel-wide" data-panel-type="find-encoding-issues" role="dialog" aria-label="Find Encoding Issues">
      <div class="floating-panel-header">
        <h2 class="floating-panel-title">Find Encoding Issues</h2>
        <button class="floating-panel-close" type="button" data-panel-action="close-panel">Close</button>
      </div>
      <label>
        <span>Scope</span>
        <select id="encodingIssuesScope">
          ${renderCleaningScopeOptions(normalizedDraft.scope)}
        </select>
      </label>
      <p class="floating-panel-meta">Entire data area means rows below the header row.</p>
      <p class="floating-panel-meta">Click a result to jump to that cell. This check does not modify workbook data.</p>
      <section class="help-panel-section" aria-label="Encoding Issue Preview">
        <h3 class="help-panel-heading">Encoding Issue Preview</h3>
        <div class="help-panel-facts">
          <div class="help-panel-fact-row"><span class="help-panel-fact-label">Cells to check</span><strong>${preview.checkedCellCount}</strong></div>
          <div class="help-panel-fact-row"><span class="help-panel-fact-label">Suspect cells</span><strong>${preview.issueCount}</strong></div>
        </div>
      </section>
      <section class="help-panel-section" aria-label="Issue results">
        <h3 class="help-panel-heading">Issue results</h3>
        ${renderEncodingIssueResults(preview)}
      </section>
      ${state.panelError ? `<p class="floating-panel-error">${escapeHtml(state.panelError)}</p>` : ""}
    </section>
  `;
}

function createRemoveDuplicatesDraft() {
  const scope = "selection";
  return {
    kind: "remove-duplicates",
    scope,
    selectedColumns: getDefaultCleaningColumns(scope),
  };
}

function createRemoveBlankRowsDraft() {
  return {
    kind: "remove-blank-rows",
    scope: "selection",
  };
}

function createTrimSpacesDraft() {
  const scope = "selection";
  return {
    kind: "trim-spaces",
    scope,
    leading: true,
    trailing: true,
    inner: false,
    innerMode: "collapse",
  };
}

function createNormalizeCaseDraft() {
  return {
    kind: "normalize-case",
    scope: "selection",
    mode: "lower",
  };
}

function createNumberCommasDraft() {
  return {
    kind: "number-commas",
    scope: "selection",
    mode: "add",
  };
}

function createFindEncodingIssuesDraft() {
  return {
    kind: "find-encoding-issues",
    scope: "selection",
  };
}

function renderEncodingIssueResults(preview) {
  if (!preview.issues.length) {
    return `<p class="floating-panel-meta">No suspicious encoding issues found.</p>`;
  }
  const results = preview.issues
    .map((issue) => {
      const cellName = gridCellName(issue.row, issue.col);
      const types = issue.types.map((type) => translateMessage(type)).join(", ");
      return `
        <button class="encoding-issue-result" type="button" data-panel-action="encoding-issue-jump" data-row="${issue.row}" data-col="${issue.col}">
          <span class="encoding-issue-cell">${escapeHtml(cellName)}</span>
          <span class="encoding-issue-type">${escapeHtml(types)}</span>
          <span class="encoding-issue-snippet" data-user-option="true">${escapeHtml(issue.snippet)}</span>
        </button>
      `;
    })
    .join("");
  const truncatedNote = preview.truncated ? '<p class="floating-panel-meta">Showing first 20 results.</p>' : "";
  return `<div class="encoding-issue-results">${results}</div>${truncatedNote}`;
}

function renderCleaningScopeOptions(selectedScope) {
  return ["selection", "data-area"]
    .map((scope) => {
      const label = scope === "selection" ? "Current selection" : "Entire data area";
      return `<option value="${scope}"${scope === selectedScope ? " selected" : ""}>${localizeInlineUiText(label)}</option>`;
    })
    .join("");
}

function renderCleaningColumnCheckboxes(draft, inputName) {
  const columns = getCleaningColumnOptions(draft.scope);
  const selectedColumns = new Set(getSelectedCleaningColumns(draft));
  if (!columns.length) {
    return `<p class="floating-panel-meta">Select at least one column.</p>`;
  }
  return columns
    .map((column) => `
      <label>
        <input name="${inputName}" type="checkbox" value="${column.index}"${selectedColumns.has(column.index) ? " checked" : ""} />
        ${escapeHtml(column.label)}
      </label>
    `)
    .join("");
}

function normalizeCleaningDraftColumns(draft) {
  const columns = getCleaningColumnOptions(draft.scope).map((column) => column.index);
  const validColumns = new Set(columns);
  const selectedColumns = Array.isArray(draft?.selectedColumns)
    ? getSelectedCleaningColumns(draft).filter((columnIndex) => validColumns.has(columnIndex))
    : columns;
  return {
    ...draft,
    selectedColumns,
  };
}

function normalizeTrimSpacesDraft(draft) {
  return {
    ...draft,
    scope: normalizeCleaningScope(draft?.scope),
    leading: draft?.leading !== false,
    trailing: draft?.trailing !== false,
    inner: Boolean(draft?.inner),
    innerMode: draft?.innerMode === "remove" ? "remove" : "collapse",
  };
}

function normalizeRemoveBlankRowsDraft(draft) {
  return {
    ...draft,
    scope: normalizeCleaningScope(draft?.scope),
  };
}

function normalizeNormalizeCaseDraft(draft) {
  const mode = isNormalizeCaseMode(draft?.mode) ? draft.mode : "lower";
  return {
    ...draft,
    scope: normalizeCleaningScope(draft?.scope),
    mode,
  };
}

function normalizeNumberCommasDraft(draft) {
  return {
    ...draft,
    scope: normalizeCleaningScope(draft?.scope),
    mode: draft?.mode === "remove" ? "remove" : "add",
  };
}

function normalizeFindEncodingIssuesDraft(draft) {
  return {
    ...draft,
    scope: normalizeCleaningScope(draft?.scope),
  };
}

function isNormalizeCaseMode(mode) {
  return ["lower", "upper", "capitalize-words", "full-width", "half-width"].includes(mode);
}

function getSelectedCleaningColumns(draft) {
  return Array.isArray(draft?.selectedColumns)
    ? draft.selectedColumns.map((columnIndex) => Math.max(1, Math.trunc(Number(columnIndex) || 1)))
    : [];
}

function getDefaultCleaningColumns(scope) {
  return getCleaningColumnOptions(scope).map((column) => column.index);
}

function getCleaningColumnOptions(scope) {
  const bounds = getSelectionBounds(state.selection);
  const maxUsedColumn = Math.max(1, getColumnCount(state.table), getLastUsedColumnIndex(state.table));
  const startColumn = scope === "selection" ? bounds.startCol : 1;
  const endColumn = scope === "selection" ? Math.max(bounds.endCol, bounds.startCol) : maxUsedColumn;
  return createIndexRange(startColumn, endColumn)
    .filter((columnIndex) => scope === "selection" || columnIndex <= maxUsedColumn)
    .map((columnIndex) => ({
      index: columnIndex,
      label: formatCleaningColumnLabel(columnIndex),
    }));
}

function getTrimSpacesColumns(scope) {
  const bounds = getSelectionBounds(state.selection);
  const maxUsedColumn = Math.max(1, getColumnCount(state.table), getLastUsedColumnIndex(state.table));
  const startColumn = scope === "selection" ? bounds.startCol : 1;
  const endColumn = scope === "selection" ? bounds.endCol : maxUsedColumn;
  return createIndexRange(startColumn, endColumn)
    .filter((columnIndex) => scope === "selection" || columnIndex <= maxUsedColumn);
}

function getBlankRowColumns() {
  const maxUsedColumn = Math.max(1, getColumnCount(state.table), getLastUsedColumnIndex(state.table));
  return createIndexRange(1, maxUsedColumn);
}

function formatCleaningColumnLabel(columnIndex) {
  const letter = columnLabelFromIndex(columnIndex);
  const header = getCellValue(state.table, getHeaderRowIndex(state.table), columnIndex);
  return header ? `${letter} ${header}` : `Column ${letter}`;
}

function resolveCleaningRows(scope, options = {}) {
  const skipHeader = Boolean(options.skipHeader);
  const headerRowIndex = getHeaderRowIndex(state.table);
  const lastUsedRow = Math.max(headerRowIndex, getLastUsedRowIndex(state.table));
  const bounds = scope === "selection"
    ? getSelectionBounds(state.selection)
    : {
      startRow: headerRowIndex + 1,
      endRow: lastUsedRow,
    };
  const startRow = Math.max(1, bounds.startRow);
  const endRow = Math.min(Math.max(startRow, bounds.endRow), Math.max(lastUsedRow, startRow));
  return createIndexRange(startRow, endRow)
    .filter((rowIndex) => (!skipHeader || rowIndex > headerRowIndex) && rowIndex <= lastUsedRow);
}

function hasTrimSpaceOption(draft) {
  return Boolean(draft?.leading || draft?.trailing || draft?.inner);
}

function normalizeCleaningScope(scope) {
  return scope === "data-area" ? "data-area" : "selection";
}

function updateCleaningColumnSelection(draft, rawColumnIndex, checked) {
  if (!draft) return;
  const columnIndex = Math.max(1, Math.trunc(Number(rawColumnIndex) || 1));
  const selectedColumns = new Set(getSelectedCleaningColumns(draft));
  if (checked) {
    selectedColumns.add(columnIndex);
  } else {
    selectedColumns.delete(columnIndex);
  }
  draft.selectedColumns = [...selectedColumns].sort((left, right) => left - right);
}

function selectAllCleaningColumns(draft) {
  if (!draft) return;
  draft.selectedColumns = getDefaultCleaningColumns(draft.scope);
  state.panelError = "";
}

function clearCleaningColumns(draft) {
  if (!draft) return;
  draft.selectedColumns = [];
  state.panelError = "";
}

function renderTextImportPanel() {
  const draft = state.importDraft;
  if (!draft || draft.kind !== "text") return "";
  const anchor = getImportAnchorCell(state.selection);
  const previewHtml = renderImportPreviewTable(draft.matrix || []);
  const delimiterOptions = draft.fileKind === "csv" || draft.fileKind === "tsv" || draft.fileKind === "txt"
    ? `
      <label>
        <span>Delimiter</span>
        <select id="textImportDelimiterSelect">
          <option value="comma"${draft.delimiterMode === "comma" ? " selected" : ""}>${localizeInlineUiText("Comma")}</option>
          <option value="tab"${draft.delimiterMode === "tab" ? " selected" : ""}>${localizeInlineUiText("Tab")}</option>
          <option value="line"${draft.delimiterMode === "line" ? " selected" : ""}>${localizeInlineUiText("Line-Based")}</option>
        </select>
      </label>
    `
    : "";
  return `
    <section class="floating-panel floating-panel-wide" data-panel-type="text-import" role="dialog" aria-label="Text import panel">
      <div class="floating-panel-header">
        <h2 class="floating-panel-title">Import Text / Structured Data</h2>
        <button class="floating-panel-close" type="button" data-panel-action="close-panel">Close</button>
      </div>
      <div class="page-setup-grid">
        <label class="grow">
          <span>File</span>
          <input type="text" value="${escapeAttr(draft.fileName)}" readonly />
        </label>
        <label>
          <span>Type</span>
          <input type="text" value="${escapeAttr(draft.fileKind.toUpperCase())}" readonly />
        </label>
      </div>
      <div class="page-setup-grid">
        <label>
          <span>Encoding</span>
          <select id="textImportEncodingSelect">
            ${getSupportedTextImportEncodings().map((encoding) => `<option value="${escapeAttr(encoding)}"${draft.encoding === encoding ? " selected" : ""}>${escapeHtml(encoding)}</option>`).join("")}
          </select>
        </label>
        ${delimiterOptions}
      </div>
      <p class="floating-panel-meta">Import anchor: ${escapeHtml(gridCellName(anchor.row, anchor.col))}. Imported content always starts at the top-left cell of the current selection.</p>
      ${draft.needsReread ? '<p class="floating-panel-meta">Press Re-read to refresh the preview from the original raw bytes using the selected encoding.</p>' : ""}
      ${draft.error ? `<p class="floating-panel-error">${escapeHtml(draft.error)}</p>` : ""}
      <div class="import-preview">
        <div class="import-preview-header">
          <span>Preview</span>
          <span class="menu-item-hint">${escapeHtml(formatMatrixShape(draft.matrix || []))}</span>
        </div>
        ${previewHtml}
      </div>
      <div class="floating-panel-actions">
        <button class="button" type="button" data-panel-action="text-import-reread">Re-read</button>
        <button class="button button-primary" type="button" data-panel-action="text-import-apply"${draft.error ? ' aria-disabled="true"' : ""}>Import</button>
      </div>
    </section>
  `;
}

function renderTextExportPanel() {
  const draft = state.exportDraft;
  if (!draft) return "";
  const summary = formatTextExportSummary(draft);
  const bomEnabled = shouldEnableBomToggle(draft.encoding);
  const targetOptions = getSupportedTextExportTargets()
    .map((target) => `<option value="${escapeAttr(target)}"${draft.target === target ? " selected" : ""}>${escapeHtml(target.toUpperCase())}</option>`)
    .join("");
  const encodingOptions = getSupportedTextExportEncodings()
    .map((encoding) => `<option value="${escapeAttr(encoding)}"${draft.encoding === encoding ? " selected" : ""}>${escapeHtml(encoding)}</option>`)
    .join("");
  const rowLineEndingOptions = getSupportedTextExportRowLineEndings()
    .map((rowLineEnding) => `<option value="${escapeAttr(rowLineEnding)}"${draft.rowLineEnding === rowLineEnding ? " selected" : ""}>${escapeHtml(rowLineEnding)}</option>`)
    .join("");
  const previewMatrix = buildTextExportMatrix(state.table, {
    includeHiddenData: draft.includeHiddenData,
    visibleRowSet: getVisibleRowInfo().visibleRowSet,
  }).slice(0, 6);
  const previewText = serializeTextExportMatrix(previewMatrix, draft);
  return `
    <section class="floating-panel floating-panel-wide" data-panel-type="text-export" role="dialog" aria-label="Text export panel">
      <div class="floating-panel-header">
        <h2 class="floating-panel-title">Export Text Data</h2>
        <button class="floating-panel-close" type="button" data-panel-action="close-panel">Close</button>
      </div>
      <div class="page-setup-grid">
        <label>
          <span>Target</span>
          <select id="textExportTargetSelect">${targetOptions}</select>
        </label>
        <label>
          <span>Encoding</span>
          <select id="textExportEncodingSelect">${encodingOptions}</select>
        </label>
        <label>
          <span>Row line ending</span>
          <select id="textExportRowLineEndingSelect">${rowLineEndingOptions}</select>
        </label>
      </div>
      <div class="floating-panel-checks">
        <label><input id="textExportNormalizeLfInput" type="checkbox"${draft.normalizeCellLineBreaksToLf ? " checked" : ""} /> Cell line breaks only LF</label>
        <label><input id="textExportBomInput" type="checkbox"${draft.withBom && bomEnabled ? " checked" : ""}${bomEnabled ? "" : " disabled"} /> with BOM</label>
        <label><input id="textExportIncludeHiddenInput" type="checkbox"${draft.includeHiddenData ? " checked" : ""} /> Include hidden data</label>
        ${draft.target === "csv" ? `<label><input id="textExportQuoteAllInput" type="checkbox"${draft.quoteAllCells ? " checked" : ""} /> Quote all cells</label>` : ""}
      </div>
      <p class="floating-panel-meta">Export scope: the whole current sheet. Hidden-row handling is controlled by <strong>Include hidden data</strong>.</p>
      <p class="floating-panel-meta">Summary: ${escapeHtml(summary)}</p>
      ${state.panelError ? `<p class="floating-panel-error">${escapeHtml(state.panelError)}</p>` : ""}
      <div class="import-preview">
        <div class="import-preview-header">
          <span>Preview</span>
          <span class="menu-item-hint">${escapeHtml(`${draft.target.toUpperCase()} / first 6 rows`)}</span>
        </div>
        <div class="import-preview-table-wrap"><pre class="text-export-preview">${escapeHtml(previewText)}</pre></div>
      </div>
      <div class="floating-panel-actions">
        <button class="button button-primary" type="button" data-panel-action="text-export-apply">Export</button>
      </div>
    </section>
  `;
}

function renderExcelImportPanel() {
  const draft = state.importDraft;
  if (!draft || draft.kind !== "excel") return "";
  const anchor = getImportAnchorCell(state.selection);
  const selected = draft.sheetOptions.find((entry) => entry.index === draft.selectedSheetIndex) || draft.sheetOptions[0];
  return `
    <section class="floating-panel" data-panel-type="excel-import" role="dialog" aria-label="Excel import panel">
      <div class="floating-panel-header">
        <h2 class="floating-panel-title">Import Excel Sheet</h2>
        <button class="floating-panel-close" type="button" data-panel-action="close-panel">Close</button>
      </div>
      <label>
        <span>Workbook</span>
        <input type="text" value="${escapeAttr(draft.fileName)}" readonly />
      </label>
      <label>
        <span>Sheet</span>
        <select id="excelImportSheetSelect">
          ${draft.sheetOptions.map((sheet) => `<option value="${sheet.index}"${sheet.index === draft.selectedSheetIndex ? " selected" : ""}>${escapeHtml(sheet.name)} (${sheet.rowCount}x${sheet.colCount})</option>`).join("")}
        </select>
      </label>
      <p class="floating-panel-meta">Import anchor: ${escapeHtml(gridCellName(anchor.row, anchor.col))}. Selected sheet data will be written into the current document from that cell.</p>
      ${selected ? `<p class="floating-panel-meta">Selected sheet size: ${selected.rowCount} rows x ${selected.colCount} columns.</p>` : ""}
      <div class="floating-panel-actions">
        <button class="button button-primary" type="button" data-panel-action="excel-import-apply">Import Sheet</button>
      </div>
    </section>
  `;
}

function renderFindReplacePanel() {
  const scopeLabel = formatScopeLabel();
  const scope = normalizeFindReplaceScope(state.findReplace.scope);
  return `
    <section class="floating-panel" data-panel-type="find-replace" role="dialog" aria-label="Find and replace panel">
      <div class="floating-panel-header">
        <h2 class="floating-panel-title">Find / Replace</h2>
        <button class="floating-panel-close" type="button" data-panel-action="close-panel">Close</button>
      </div>
      <label>
        <span>Scope</span>
        <select id="findReplaceScopeSelect">
          <option value="global"${scope === "global" ? " selected" : ""}>Global</option>
          <option value="selection"${scope === "selection" ? " selected" : ""}>Current selection</option>
        </select>
      </label>
      <label>
        <span>Find what</span>
        <input id="findTextInput" type="text" value="${escapeAttr(state.findReplace.findText)}" autocomplete="off" spellcheck="false" />
      </label>
      <label>
        <span>Replace with</span>
        <input id="replaceTextInput" type="text" value="${escapeAttr(state.findReplace.replaceText)}" autocomplete="off" spellcheck="false" />
      </label>
      <div class="floating-panel-checks">
        <label><input id="findCaseSensitiveInput" type="checkbox"${state.findReplace.caseSensitive ? " checked" : ""} /> Case Sensitive</label>
        <label><input id="findWholeCellInput" type="checkbox"${state.findReplace.wholeCell ? " checked" : ""} /> Whole Cell</label>
      </div>
      <p class="floating-panel-meta">Scope: ${escapeHtml(scopeLabel)}</p>
      ${state.panelError ? `<p class="floating-panel-error">${escapeHtml(state.panelError)}</p>` : ""}
      <div class="floating-panel-actions">
        <button class="button" type="button" data-panel-action="find-next">Find Next</button>
        <button class="button" type="button" data-panel-action="replace-one">Replace</button>
        <button class="button button-primary" type="button" data-panel-action="replace-all">Replace All</button>
      </div>
    </section>
  `;
}

function renderAdvancedFilterPanel() {
  const draft = state.advancedFilterDraft || createAdvancedFilterDraft();
  const columnOptions = renderColumnOptions();
  return `
    <section class="floating-panel" data-panel-type="advanced-filter" role="dialog" aria-label="Advanced filter panel">
      <div class="floating-panel-header">
        <h2 class="floating-panel-title">Advanced Filter</h2>
        <button class="floating-panel-close" type="button" data-panel-action="close-panel">Close</button>
      </div>
      <div class="advanced-filter-list">
        ${draft.conditions.map((condition, index) => `
          <div class="advanced-filter-condition">
            <div class="advanced-filter-number">${index + 1}</div>
            <label class="grow">
              <span>Column</span>
              <select data-condition-index="${index}" data-condition-field="columnKey">
                ${columnOptions.replace(`value="${escapeAttr(condition.columnKey)}"`, `value="${escapeAttr(condition.columnKey)}" selected`)}
              </select>
            </label>
            <label class="grow">
              <span>Operator</span>
              <select data-condition-index="${index}" data-condition-field="operator">
                ${renderOperatorOptions(condition.operator)}
              </select>
            </label>
            <label class="grow">
              <span>Value</span>
              <input type="text" value="${escapeAttr(condition.value)}" data-condition-index="${index}" data-condition-field="value" spellcheck="false" />
            </label>
            <button class="link-button advanced-filter-remove" type="button" data-panel-action="advanced-remove-condition" data-condition-index="${index}"${draft.conditions.length <= 1 ? " aria-disabled=\"true\"" : ""}>Remove</button>
          </div>
        `).join("")}
      </div>
      <div class="inline-row">
        <button class="link-button" type="button" data-panel-action="advanced-add-condition"${draft.conditions.length >= 20 ? " aria-disabled=\"true\"" : ""}>Add Condition</button>
        <button class="link-button" type="button" data-panel-action="advanced-reset-logic">Reset to default logic</button>
      </div>
      <label>
        <span>Logic expression</span>
        <input id="advancedFilterLogicInput" type="text" value="${escapeAttr(draft.logic)}" spellcheck="false" />
      </label>
      <p class="floating-panel-example">Example: <code>1 AND (2 OR 3)</code></p>
      ${state.panelError ? `<p class="floating-panel-error">${escapeHtml(state.panelError)}</p>` : ""}
      <div class="floating-panel-actions">
        <button class="button" type="button" data-panel-action="advanced-clear">Clear Advanced Filter</button>
        <button class="button button-primary" type="button" data-panel-action="advanced-apply">Apply</button>
      </div>
    </section>
  `;
}

function renderSortPanel() {
  const draft = state.sortDraft || createSortDraft();
  const columnOptions = `<option value="">${localizeInlineUiText("None")}</option>${renderColumnOptions()}`;
  return `
    <section class="floating-panel" data-panel-type="sort" role="dialog" aria-label="Sort panel">
      <div class="floating-panel-header">
        <h2 class="floating-panel-title">Sort</h2>
        <button class="floating-panel-close" type="button" data-panel-action="close-panel">Close</button>
      </div>
      <div class="page-setup-grid">
        <label class="grow">
          <span>Primary column</span>
          <select id="sortPrimaryColumn">${columnOptions.replace(`value="${escapeAttr(draft.primaryColumn)}"`, `value="${escapeAttr(draft.primaryColumn)}" selected`)}</select>
        </label>
        <label>
          <span>Direction</span>
          <select id="sortPrimaryDirection">
            <option value="asc"${draft.primaryDirection === "asc" ? " selected" : ""}>${localizeInlineUiText("A -> Z")}</option>
            <option value="desc"${draft.primaryDirection === "desc" ? " selected" : ""}>${localizeInlineUiText("Z -> A")}</option>
          </select>
        </label>
        <label class="grow">
          <span>Secondary column</span>
          <select id="sortSecondaryColumn">${columnOptions.replace(`value="${escapeAttr(draft.secondaryColumn)}"`, `value="${escapeAttr(draft.secondaryColumn)}" selected`)}</select>
        </label>
        <label>
          <span>Direction</span>
          <select id="sortSecondaryDirection">
            <option value="asc"${draft.secondaryDirection === "asc" ? " selected" : ""}>${localizeInlineUiText("A -> Z")}</option>
            <option value="desc"${draft.secondaryDirection === "desc" ? " selected" : ""}>${localizeInlineUiText("Z -> A")}</option>
          </select>
        </label>
      </div>
      ${state.panelError ? `<p class="floating-panel-error">${escapeHtml(state.panelError)}</p>` : ""}
      <div class="floating-panel-actions">
        <button class="button button-primary" type="button" data-panel-action="sort-apply">Apply Sort</button>
      </div>
    </section>
  `;
}

function renderPageSetupPanel() {
  const draft = state.pageSetupDraft || createPageSetupDraft(state.table.pageSetup);
  return `
    <section class="floating-panel" data-panel-type="page-setup" role="dialog" aria-label="Page setup panel">
      <div class="floating-panel-header">
        <h2 class="floating-panel-title">Page Setup</h2>
        <button class="floating-panel-close" type="button" data-panel-action="close-panel">Close</button>
      </div>
      <div class="page-setup-grid page-setup-basics">
        <label>
          <span>Paper Size</span>
          <select id="pageSetupPaperSize">
            ${renderFixedOptions(["A4", "A3", "Letter"], draft.paperSize)}
          </select>
        </label>
        <label>
          <span>Orientation</span>
          <select id="pageSetupOrientation">
            ${renderFixedOptions(["portrait", "landscape"], draft.orientation)}
          </select>
        </label>
      </div>
      <div class="page-setup-grid page-setup-margins">
        <label>
          <span>Top</span>
          <input id="pageSetupMarginTop" type="number" min="0" step="1" value="${escapeAttr(draft.margins.top)}" />
        </label>
        <label>
          <span>Right</span>
          <input id="pageSetupMarginRight" type="number" min="0" step="1" value="${escapeAttr(draft.margins.right)}" />
        </label>
        <label>
          <span>Bottom</span>
          <input id="pageSetupMarginBottom" type="number" min="0" step="1" value="${escapeAttr(draft.margins.bottom)}" />
        </label>
        <label>
          <span>Left</span>
          <input id="pageSetupMarginLeft" type="number" min="0" step="1" value="${escapeAttr(draft.margins.left)}" />
        </label>
      </div>
      <label>
        <span>Header</span>
        <input id="pageSetupHeader" type="text" value="${escapeAttr(draft.headerFooter.header)}" spellcheck="false" />
      </label>
      <label>
        <span>Footer</span>
        <input id="pageSetupFooter" type="text" value="${escapeAttr(draft.headerFooter.footer)}" spellcheck="false" />
      </label>
      <div class="page-setup-grid">
        <label class="grow">
          <span>Print Area</span>
          <select id="pageSetupPrintAreaMode">
            ${renderFixedOptions(["entire-sheet", "selection", "custom"], draft.printArea.mode)}
          </select>
        </label>
        <label class="grow">
          <span>Custom Range</span>
          <input id="pageSetupPrintAreaRange" type="text" value="${escapeAttr(draft.printArea.range)}" placeholder="A1:D20" spellcheck="false" />
        </label>
      </div>
      <div class="page-setup-grid">
        <label class="grow">
          <span>Background</span>
          <select id="pageSetupBackgroundMode">
            ${renderFixedOptions(["none", "solid-color"], draft.background?.mode || "none")}
          </select>
        </label>
        <label class="grow">
          <span>Color</span>
          <input id="pageSetupBackgroundColor" type="color" value="${escapeAttr(draft.background?.color || "#ffffff")}" />
        </label>
      </div>
      ${state.panelError ? `<p class="floating-panel-error">${escapeHtml(state.panelError)}</p>` : ""}
      <div class="floating-panel-actions">
        <button class="button" type="button" data-panel-action="page-setup-apply">Apply</button>
        <button class="button button-primary" type="button" data-panel-action="page-setup-print">Apply And Print</button>
      </div>
    </section>
  `;
}

function renderGridLimitsPanel() {
  const draft = state.gridLimitsDraft || createGridLimitsDraft();
  return `
    <section class="floating-panel" data-panel-type="grid-limits" role="dialog" aria-label="Grid limits panel">
      <div class="floating-panel-header">
        <h2 class="floating-panel-title">Grid Limits</h2>
        <button class="floating-panel-close" type="button" data-panel-action="close-panel">Close</button>
      </div>
      <div class="page-setup-grid">
        <label>
          <span>Maximum Rows</span>
          <input id="gridLimitsRowsInput" type="number" min="1" step="1" value="${escapeAttr(draft.maxRows)}" />
        </label>
        <label>
          <span>Maximum Columns</span>
          <input id="gridLimitsColumnsInput" type="number" min="1" step="1" value="${escapeAttr(draft.maxColumns)}" />
        </label>
      </div>
      <p id="gridLimitsColumnFeedback" class="floating-panel-meta">Column label: ${escapeHtml(formatGridLimitColumnFeedback(draft.maxColumns))}</p>
      <p class="floating-panel-meta">These values define the current editable grid size for this sheet.</p>
      ${state.panelError ? `<p class="floating-panel-error">${escapeHtml(state.panelError)}</p>` : ""}
      <div class="floating-panel-actions">
        <button class="button button-primary" type="button" data-panel-action="grid-limits-apply">Apply</button>
      </div>
    </section>
  `;
}

function renderImportPreviewTable(matrix) {
  if (!matrix.length) {
    return `<p class="floating-panel-meta">No preview rows are available yet.</p>`;
  }
  const previewRows = matrix.slice(0, 8);
  const previewColumnCount = Math.max(1, Math.min(6, previewRows.reduce((max, row) => Math.max(max, row.length), 0)));
  const body = previewRows
    .map((row) => `<tr>${Array.from({ length: previewColumnCount }, (_, index) => `<td>${escapeHtml(String(row[index] ?? ""))}</td>`).join("")}</tr>`)
    .join("");
  return `
    <div class="import-preview-table-wrap">
      <table class="import-preview-table">
        <tbody>${body}</tbody>
      </table>
    </div>
  `;
}

function renderHeaderMenu() {
  const headerMenu = state.headerMenu;
  const columnLabel = state.table.columns?.[headerMenu.columnIndex - 1]?.label || columnLabelFromIndex(headerMenu.columnIndex);
  return `
    <section class="header-menu" style="left:${headerMenu.left}px;top:${headerMenu.top}px;" role="dialog" aria-label="Header menu">
      <strong>${escapeHtml(columnLabel)}</strong>
      <div class="header-menu-actions">
        <button class="button" type="button" data-panel-action="header-sort-asc">Sort A -> Z</button>
        <button class="button" type="button" data-panel-action="header-sort-desc">Sort Z -> A</button>
      </div>
      <label>
        <span>Filter operator</span>
        <select id="headerMenuOperator">
          ${renderOperatorOptions(headerMenu.operator)}
        </select>
      </label>
      <label>
        <span>Filter value</span>
        <input id="headerMenuValue" type="text" value="${escapeAttr(headerMenu.value)}" spellcheck="false" />
      </label>
      <p class="header-menu-note">Applies only to rows below the formal header row.</p>
      ${headerMenu.error ? `<p class="floating-panel-error">${escapeHtml(headerMenu.error)}</p>` : ""}
      <div class="header-menu-actions">
        <button class="button" type="button" data-panel-action="header-filter-clear">Clear Filter</button>
        <button class="button button-primary" type="button" data-panel-action="header-filter-apply">Apply Filter</button>
      </div>
    </section>
  `;
}

function handleGridClick(event) {
  debugGridLog("grid-click:start", { suppressNextGridClick: state.suppressNextGridClick, target: describeEventTarget(event.target) });
  if (state.suppressNextGridClick) {
    state.suppressNextGridClick = false;
    debugGridLog("grid-click:suppressed", { target: describeEventTarget(event.target) });
    return;
  }
  if (event.target.closest("[data-fill-handle='true']")) return;
  if (event.target.closest("[data-header-menu-button='true']")) {
    const cellElement = event.target.closest("[data-cell='true']");
    const cell = cellFromDataset(cellElement);
    selectSingleCell(cell);
    openHeaderMenu(cell.col, event.target.closest("[data-header-menu-button='true']"));
    return;
  }
  if (event.target.closest("#activeCellEditor")) return;
  const cellElement = event.target.closest("[data-cell='true']");
  const rowHeaderElement = event.target.closest("[data-row-header='true']");
  const columnHeaderElement = event.target.closest("[data-col-header='true']");

  if (cellElement) {
    commitEditing("stay");
    const cell = cellFromDataset(cellElement);
    debugGridLog("grid-click:cell", { cell, shiftKey: event.shiftKey, activeCell: state.activeCell });
    if (event.shiftKey) {
      const anchorCell = state.selection.anchor || state.activeCell;
      state.activeCell = cell;
      state.selection = createRangeSelection(anchorCell, cell);
      state.findReplace.currentMatch = null;
      renderWorkspace();
      return;
    }
    if (sameCell(cell, state.activeCell) && state.selection.mode === "cell" && !state.editing) {
      debugGridLog("grid-click:begin-edit", { cell });
      beginCellEdit();
      return;
    }
    selectSingleCell(cell);
    return;
  }

  if (rowHeaderElement) {
    commitEditing("stay");
    const row = Number(rowHeaderElement.dataset.row || 1);
    const previousActiveCell = { ...state.activeCell };
    debugGridLog("grid-click:row-header", { row, activeCell: state.activeCell, shiftKey: event.shiftKey });
    state.activeCell = clampActiveCell({ row, col: state.activeCell.col }, currentView);
    if (event.shiftKey) {
      const anchorRow = state.selection.mode === "row"
        ? state.selection.anchor?.row || previousActiveCell.row
        : previousActiveCell.row;
      state.selection = createRowSelection(anchorRow, currentView.columnCount, row);
    } else {
      state.selection = createRowSelection(row, currentView.columnCount);
    }
    state.findReplace.currentMatch = null;
    renderWorkspace();
    return;
  }

  if (columnHeaderElement) {
    commitEditing("stay");
    const col = Number(columnHeaderElement.dataset.col || 1);
    const previousActiveCell = { ...state.activeCell };
    debugGridLog("grid-click:col-header", { col, activeCell: state.activeCell, shiftKey: event.shiftKey });
    state.activeCell = clampActiveCell({ row: state.activeCell.row, col }, currentView);
    if (event.shiftKey) {
      const anchorColumn = state.selection.mode === "column"
        ? state.selection.anchor?.col || previousActiveCell.col
        : previousActiveCell.col;
      state.selection = createColumnSelection(anchorColumn, currentView.rowCount, col);
    } else {
      state.selection = createColumnSelection(col, currentView.rowCount);
    }
    state.findReplace.currentMatch = null;
    renderWorkspace();
  }
}

function handleGridMouseDown(event) {
  debugGridLog("grid-mousedown:start", { button: event.button, shiftKey: event.shiftKey, target: describeEventTarget(event.target) });
  if (event.button !== 0) return;
  if (event.target.closest("[data-header-menu-button='true']")) return;
  if (event.target.closest("[data-fill-handle='true']")) {
    event.preventDefault();
    event.stopPropagation();
    const sourceBounds = getSelectionBounds(state.selection);
    state.fillDrag = {
      sourceSelection: cloneJsonValue(state.selection),
      previewCell: { row: sourceBounds.endRow, col: sourceBounds.endCol },
    };
    state.selectionDrag = null;
    renderGrid();
    return;
  }
  if (event.target.closest("#activeCellEditor")) return;
  const cellElement = event.target.closest("[data-cell='true']");
  if (!cellElement) return;
  commitEditing("stay");
  const clickedCell = cellFromDataset(cellElement);
  const selectionAnchor = event.shiftKey
    ? (state.selection.anchor || state.activeCell)
    : clickedCell;
  if (!event.shiftKey && !sameCell(clickedCell, state.activeCell)) {
    state.activeCell = clickedCell;
    state.selection = createCellSelection(clickedCell);
    state.findReplace.currentMatch = null;
    renderSelectionState();
  }
  state.selectionDrag = {
    anchor: selectionAnchor,
    focus: clickedCell,
    moved: false,
    extendSelection: event.shiftKey,
    startedOnActiveCell: !event.shiftKey && sameCell(clickedCell, state.activeCell) && state.selection.mode === "cell" && !state.editing,
  };
  debugGridLog("grid-mousedown:selection-drag", { clickedCell, selectionAnchor, activeCell: state.activeCell, selection: state.selection });
}

function handleGridMouseOver(event) {
  handleGridPointerMove(event);
}

function handleGridPointerMove(event) {
  if (state.fillDrag) {
    const cell = resolvePointerCell(event);
    if (!cell) return;
    if (sameCell(cell, state.fillDrag.previewCell)) return;
    state.fillDrag.previewCell = cell;
    renderGrid();
    return;
  }
  if (!state.selectionDrag) return;
  const cell = resolvePointerCell(event);
  if (!cell) return;
  if (sameCell(cell, state.selectionDrag.focus)) return;
  state.selectionDrag.focus = cell;
  state.selectionDrag.moved = state.selectionDrag.moved || !sameCell(cell, state.selectionDrag.anchor);
  state.activeCell = cell;
  state.selection = createRangeSelection(state.selectionDrag.anchor, cell);
  state.findReplace.currentMatch = null;
  renderSelectionState();
}

function handleGridDoubleClick(event) {
  if (event.target.closest("[data-header-menu-button='true']")) return;
  const cellElement = event.target.closest("[data-cell='true']");
  if (!cellElement) return;
  commitEditing("stay");
  const cell = cellFromDataset(cellElement);
  selectSingleCell(cell);
  beginCellEdit();
}

function handleGridInput(event) {
  const editor = event.target.closest("#activeCellEditor");
  if (!editor || !state.editing || state.editing.mode !== "cell") return;
  state.editing.draft = editor.value;
  if (refs.formulaBarInput.value !== editor.value) {
    refs.formulaBarInput.value = editor.value;
  }
  resizeTextarea(editor, 180);
  resizeTextarea(refs.formulaBarInput, 100);
}

function handleGridEditorKeydown(event) {
  const editor = event.target.closest("#activeCellEditor");
  if (!editor || !state.editing || state.editing.mode !== "cell") return;
  if (event.key === "Escape") {
    event.preventDefault();
    cancelEditing();
    return;
  }
  if (event.key === "Enter" && event.altKey) {
    event.preventDefault();
    insertLineBreakAtCursor(editor, "cell");
    return;
  }
  if (event.key === "Enter" && !event.altKey) {
    event.preventDefault();
    commitEditing(event.shiftKey ? "shift+enter" : "enter");
    return;
  }
  if (event.key === "Tab") {
    event.preventDefault();
    commitEditing(event.shiftKey ? "shift+tab" : "tab");
  }
}

function handleFormulaBarFocus() {
  if (!state.editing) {
    state.editing = {
      mode: "formula",
      draft: getCellValue(state.table, state.activeCell.row, state.activeCell.col),
    };
    return;
  }
  if (state.editing.mode !== "formula") {
    state.editing.mode = "formula";
    renderWorkspace();
  }
}

function handleFormulaBarInput(event) {
  if (!state.editing || state.editing.mode !== "formula") {
    state.editing = {
      mode: "formula",
      draft: event.currentTarget.value,
    };
  } else {
    state.editing.draft = event.currentTarget.value;
  }
  resizeTextarea(refs.formulaBarInput, 100);
}

function handleFormulaBarKeydown(event) {
  if (event.key === "Escape") {
    event.preventDefault();
    cancelEditing();
    return;
  }
  if (event.key === "Enter" && event.altKey) {
    event.preventDefault();
    insertLineBreakAtCursor(event.currentTarget, "formula");
    return;
  }
  if (event.key === "Enter" && !event.altKey) {
    event.preventDefault();
    commitEditing(event.shiftKey ? "shift+enter" : "enter");
    return;
  }
  if (event.key === "Tab") {
    event.preventDefault();
    commitEditing(event.shiftKey ? "shift+tab" : "tab");
  }
}

function handleGlobalSearchInput(event) {
  state.globalSearch = event.currentTarget.value;
  state.findReplace.currentMatch = null;
  renderWorkspace();
  setStatus(buildVisibleRowStatus(state.globalSearch ? `Global search filtered by "${state.globalSearch}".` : "Global search cleared."));
}

function handleGlobalSearchKeydown(event) {
  if (event.key === "Escape") {
    event.preventDefault();
    handleGlobalSearchClear();
  }
}

function handleGlobalSearchClear() {
  if (!state.globalSearch) return;
  state.globalSearch = "";
  state.findReplace.currentMatch = null;
  renderWorkspace();
  setStatus(buildVisibleRowStatus("Global search cleared."));
}

function handleFloatingLayerClick(event) {
  const actionTarget = event.target.closest("[data-panel-action]");
  if (!actionTarget) return;
  const action = actionTarget.dataset.panelAction || "";
  if (actionTarget.getAttribute("aria-disabled") === "true") return;
  if (action === "close-panel") {
    closeOpenPanel();
    return;
  }
  if (action === "text-import-reread") {
    handleTextImportReread();
    return;
  }
  if (action === "text-import-apply") {
    handleApplyTextImport();
    return;
  }
  if (action === "text-export-apply") {
    handleApplyTextExport();
    return;
  }
  if (action === "excel-import-apply") {
    handleApplyExcelImport();
    return;
  }
  if (action === "find-next") {
    runFindNext();
    return;
  }
  if (action === "replace-one") {
    runReplaceOne();
    return;
  }
  if (action === "replace-all") {
    runReplaceAll();
    return;
  }
  if (action === "advanced-add-condition") {
    handleAddAdvancedCondition();
    return;
  }
  if (action === "advanced-remove-condition") {
    handleRemoveAdvancedCondition(Number(actionTarget.dataset.conditionIndex || 0));
    return;
  }
  if (action === "advanced-reset-logic") {
    handleResetAdvancedLogic();
    return;
  }
  if (action === "advanced-apply") {
    handleApplyAdvancedFilter();
    return;
  }
  if (action === "advanced-clear") {
    state.advancedFilter = { enabled: false, conditions: [], logic: "" };
    state.panelError = "";
    closeOpenPanel();
    renderWorkspace();
    setStatus(buildVisibleRowStatus("Advanced filter cleared."));
    return;
  }
  if (action === "remove-duplicates-select-all") {
    selectAllCleaningColumns(state.removeDuplicatesDraft);
    renderWorkspace();
    return;
  }
  if (action === "remove-duplicates-clear-columns") {
    clearCleaningColumns(state.removeDuplicatesDraft);
    renderWorkspace();
    return;
  }
  if (action === "remove-duplicates-apply") {
    handleApplyRemoveDuplicates();
    return;
  }
  if (action === "remove-blank-rows-apply") {
    handleApplyRemoveBlankRows();
    return;
  }
  if (action === "trim-spaces-apply") {
    handleApplyTrimSpaces();
    return;
  }
  if (action === "normalize-case-apply") {
    handleApplyNormalizeCase();
    return;
  }
  if (action === "number-commas-apply") {
    handleApplyNumberCommas();
    return;
  }
  if (action === "encoding-issue-jump") {
    jumpToEncodingIssue(Number(actionTarget.dataset.row || 0), Number(actionTarget.dataset.col || 0));
    return;
  }
  if (action === "sort-apply") {
    handleApplySort();
    return;
  }
  if (action === "page-setup-apply") {
    handleApplyPageSetup(false);
    return;
  }
  if (action === "page-setup-print") {
    handleApplyPageSetup(true);
    return;
  }
  if (action === "grid-limits-apply") {
    handleApplyGridLimits();
    return;
  }
  if (action === "header-filter-apply") {
    handleApplyHeaderFilter();
    return;
  }
  if (action === "header-filter-clear") {
    handleClearHeaderFilter();
    return;
  }
  if (action === "header-sort-asc") {
    handleQuickSort(state.headerMenu?.columnIndex, "asc");
    return;
  }
  if (action === "header-sort-desc") {
    handleQuickSort(state.headerMenu?.columnIndex, "desc");
  }
}

function handleFloatingLayerInput(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement || target instanceof HTMLTextAreaElement)) {
    return;
  }

  if (target.id === "findTextInput") {
    state.findReplace.findText = target.value;
    state.findReplace.currentMatch = null;
    state.panelError = "";
    return;
  }
  if (target.id === "textImportEncodingSelect" && state.importDraft?.kind === "text") {
    state.importDraft.encoding = target.value;
    state.importDraft.needsReread = true;
    state.importDraft.error = "";
    renderWorkspace();
    return;
  }
  if (target.id === "gridLimitsRowsInput" && state.gridLimitsDraft) {
    state.gridLimitsDraft.maxRows = target.value;
    state.panelError = "";
    return;
  }
  if (target.id === "gridLimitsColumnsInput" && state.gridLimitsDraft) {
    state.gridLimitsDraft.maxColumns = target.value;
    state.panelError = "";
    const feedback = document.getElementById("gridLimitsColumnFeedback");
    if (feedback) {
      feedback.textContent = translateMessage(`Column label: ${formatGridLimitColumnFeedback(target.value)}`);
    }
    return;
  }
  if (target.id === "textExportTargetSelect" && state.exportDraft) {
    state.exportDraft.target = normalizeTextExportTarget(target.value);
    state.exportDraft.quoteAllCells = state.exportDraft.target === "csv" && Boolean(state.exportDraft.quoteAllCells);
    state.panelError = "";
    rememberTextExportSettings();
    renderWorkspace();
    return;
  }
  if (target.id === "textExportEncodingSelect" && state.exportDraft) {
    state.exportDraft.encoding = normalizeTextExportEncoding(target.value);
    if (!shouldEnableBomToggle(state.exportDraft.encoding)) {
      state.exportDraft.withBom = false;
    }
    state.panelError = "";
    rememberTextExportSettings();
    renderWorkspace();
    return;
  }
  if (target.id === "textExportRowLineEndingSelect" && state.exportDraft) {
    state.exportDraft.rowLineEnding = normalizeTextExportRowLineEnding(target.value);
    state.panelError = "";
    rememberTextExportSettings();
    renderWorkspace();
    return;
  }
  if (target.id === "textExportNormalizeLfInput" && state.exportDraft) {
    state.exportDraft.normalizeCellLineBreaksToLf = target.checked;
    state.panelError = "";
    rememberTextExportSettings();
    renderWorkspace();
    return;
  }
  if (target.id === "textExportBomInput" && state.exportDraft) {
    state.exportDraft.withBom = shouldEnableBomToggle(state.exportDraft.encoding) && target.checked;
    state.panelError = "";
    rememberTextExportSettings();
    renderWorkspace();
    return;
  }
  if (target.id === "textExportIncludeHiddenInput" && state.exportDraft) {
    state.exportDraft.includeHiddenData = target.checked;
    state.panelError = "";
    rememberTextExportSettings();
    renderWorkspace();
    return;
  }
  if (target.id === "textExportQuoteAllInput" && state.exportDraft) {
    state.exportDraft.quoteAllCells = state.exportDraft.target === "csv" && target.checked;
    state.panelError = "";
    rememberTextExportSettings();
    renderWorkspace();
    return;
  }
  if (target.id === "textImportDelimiterSelect" && state.importDraft?.kind === "text") {
    state.importDraft.delimiterMode = target.value;
    state.importDraft.error = "";
    if (!state.importDraft.needsReread) {
      refreshTextImportDraft(state.importDraft);
    }
    renderWorkspace();
    return;
  }
  if (target.id === "excelImportSheetSelect" && state.importDraft?.kind === "excel") {
    state.importDraft.selectedSheetIndex = Number(target.value || 0);
    return;
  }
  if (target.id === "replaceTextInput") {
    state.findReplace.replaceText = target.value;
    return;
  }
  if (target.id === "findReplaceScopeSelect") {
    state.findReplace.scope = normalizeFindReplaceScope(target.value);
    if (state.findReplace.scope === "selection") {
      captureFindReplaceScopeBounds();
    } else {
      state.findReplace.scopeBounds = null;
    }
    state.findReplace.currentMatch = null;
    state.panelError = "";
    renderFloatingLayer();
    return;
  }
  if (target.id === "findCaseSensitiveInput") {
    state.findReplace.caseSensitive = target.checked;
    state.findReplace.currentMatch = null;
    return;
  }
  if (target.id === "findWholeCellInput") {
    state.findReplace.wholeCell = target.checked;
    state.findReplace.currentMatch = null;
    return;
  }
  if (target.dataset.conditionField) {
    const index = Number(target.dataset.conditionIndex || 0);
    const field = target.dataset.conditionField;
    if (!state.advancedFilterDraft?.conditions[index]) return;
    state.advancedFilterDraft.conditions[index][field] = target.value;
    state.panelError = "";
    return;
  }
  if (target.id === "advancedFilterLogicInput") {
    if (!state.advancedFilterDraft) return;
    state.advancedFilterDraft.logic = target.value;
    state.panelError = "";
    return;
  }
  if (target.id === "removeDuplicatesScope" && state.removeDuplicatesDraft) {
    state.removeDuplicatesDraft.scope = normalizeCleaningScope(target.value);
    state.removeDuplicatesDraft.selectedColumns = getDefaultCleaningColumns(state.removeDuplicatesDraft.scope);
    state.panelError = "";
    renderWorkspace();
    return;
  }
  if (target.name === "removeDuplicatesColumn" && state.removeDuplicatesDraft) {
    updateCleaningColumnSelection(state.removeDuplicatesDraft, target.value, target.checked);
    state.panelError = "";
    renderWorkspace();
    return;
  }
  if (target.id === "removeBlankRowsScope" && state.removeBlankRowsDraft) {
    state.removeBlankRowsDraft.scope = normalizeCleaningScope(target.value);
    state.panelError = "";
    renderWorkspace();
    return;
  }
  if (target.id === "trimSpacesScope" && state.trimSpacesDraft) {
    state.trimSpacesDraft.scope = normalizeCleaningScope(target.value);
    state.panelError = "";
    renderWorkspace();
    return;
  }
  if (target.id === "trimSpacesLeadingInput" && state.trimSpacesDraft) {
    state.trimSpacesDraft.leading = target.checked;
    state.panelError = "";
    renderWorkspace();
    return;
  }
  if (target.id === "trimSpacesInnerInput" && state.trimSpacesDraft) {
    state.trimSpacesDraft.inner = target.checked;
    state.panelError = "";
    renderWorkspace();
    return;
  }
  if (target.id === "trimSpacesTrailingInput" && state.trimSpacesDraft) {
    state.trimSpacesDraft.trailing = target.checked;
    state.panelError = "";
    renderWorkspace();
    return;
  }
  if (target.name === "trimSpacesInnerMode" && state.trimSpacesDraft) {
    state.trimSpacesDraft.innerMode = target.value === "remove" ? "remove" : "collapse";
    state.panelError = "";
    renderWorkspace();
    return;
  }
  if (target.id === "normalizeCaseScope" && state.normalizeCaseDraft) {
    state.normalizeCaseDraft.scope = normalizeCleaningScope(target.value);
    state.panelError = "";
    renderWorkspace();
    return;
  }
  if (target.name === "normalizeCaseMode" && state.normalizeCaseDraft) {
    state.normalizeCaseDraft.mode = isNormalizeCaseMode(target.value) ? target.value : "lower";
    state.panelError = "";
    renderWorkspace();
    return;
  }
  if (target.id === "numberCommasScope" && state.numberCommasDraft) {
    state.numberCommasDraft.scope = normalizeCleaningScope(target.value);
    state.panelError = "";
    renderWorkspace();
    return;
  }
  if (target.name === "numberCommasMode" && state.numberCommasDraft) {
    state.numberCommasDraft.mode = target.value === "remove" ? "remove" : "add";
    state.panelError = "";
    renderWorkspace();
    return;
  }
  if (target.id === "encodingIssuesScope" && state.encodingIssuesDraft) {
    state.encodingIssuesDraft.scope = normalizeCleaningScope(target.value);
    state.panelError = "";
    renderWorkspace();
    return;
  }
  if (target.id === "sortPrimaryColumn") {
    state.sortDraft.primaryColumn = target.value;
    return;
  }
  if (target.id === "sortPrimaryDirection") {
    state.sortDraft.primaryDirection = target.value;
    return;
  }
  if (target.id === "sortSecondaryColumn") {
    state.sortDraft.secondaryColumn = target.value;
    return;
  }
  if (target.id === "sortSecondaryDirection") {
    state.sortDraft.secondaryDirection = target.value;
    return;
  }
  if (target.id === "pageSetupPaperSize") {
    state.pageSetupDraft.paperSize = target.value;
    state.panelError = "";
    return;
  }
  if (target.id === "pageSetupOrientation") {
    state.pageSetupDraft.orientation = target.value;
    state.panelError = "";
    return;
  }
  if (target.id === "pageSetupMarginTop") {
    state.pageSetupDraft.margins.top = target.value;
    return;
  }
  if (target.id === "pageSetupMarginRight") {
    state.pageSetupDraft.margins.right = target.value;
    return;
  }
  if (target.id === "pageSetupMarginBottom") {
    state.pageSetupDraft.margins.bottom = target.value;
    return;
  }
  if (target.id === "pageSetupMarginLeft") {
    state.pageSetupDraft.margins.left = target.value;
    return;
  }
  if (target.id === "pageSetupHeader") {
    state.pageSetupDraft.headerFooter.header = target.value;
    return;
  }
  if (target.id === "pageSetupFooter") {
    state.pageSetupDraft.headerFooter.footer = target.value;
    return;
  }
  if (target.id === "pageSetupPrintAreaMode") {
    state.pageSetupDraft.printArea.mode = target.value;
    state.panelError = "";
    return;
  }
  if (target.id === "pageSetupPrintAreaRange") {
    state.pageSetupDraft.printArea.range = target.value;
    state.panelError = "";
    return;
  }
  if (target.id === "pageSetupBackgroundMode") {
    state.pageSetupDraft.background = target.value === "none"
      ? null
      : { mode: "solid-color", color: state.pageSetupDraft.background?.color || "#ffffff" };
    state.panelError = "";
    return;
  }
  if (target.id === "pageSetupBackgroundColor") {
    state.pageSetupDraft.background = {
      mode: "solid-color",
      color: target.value,
    };
    return;
  }
  if (target.id === "headerMenuOperator" && state.headerMenu) {
    state.headerMenu.operator = target.value;
    state.headerMenu.error = "";
    return;
  }
  if (target.id === "headerMenuValue" && state.headerMenu) {
    state.headerMenu.value = target.value;
    state.headerMenu.error = "";
  }
}

function handleFloatingLayerKeydown(event) {
  if (event.key === "Escape") {
    event.preventDefault();
    if (state.headerMenu) {
      state.headerMenu = null;
      renderFloatingLayer();
      return;
    }
    closeOpenPanel();
    return;
  }
  if (event.key === "Enter" && event.target instanceof HTMLInputElement && event.target.id === "findTextInput") {
    event.preventDefault();
    runFindNext();
  }
}

function handleDocumentClick(event) {
  const target = event.target;
  if (refs.appMenu.contains(target) || refs.menuButtons.some((button) => button.contains(target))) {
    return;
  }
  setMenuOpen(null);
  if (state.headerMenu && !target.closest(".header-menu") && !target.closest("[data-header-menu-button='true']")) {
    state.headerMenu = null;
    renderFloatingLayer();
  }
}

function handleDocumentMouseMove(event) {
  handleGridPointerMove(event);
}

function handleDocumentMouseUp(event) {
  debugGridLog("document-mouseup:start", { target: describeEventTarget(event.target), hasSelectionDrag: Boolean(state.selectionDrag), hasFillDrag: Boolean(state.fillDrag) });
  if (state.fillDrag) {
    const releaseCell = resolvePointerCell(event);
    if (releaseCell) {
      state.fillDrag.previewCell = releaseCell;
    }
    const previewCell = state.fillDrag.previewCell;
    const sourceSelection = state.fillDrag.sourceSelection;
    state.fillDrag = null;
    state.suppressNextGridClick = isPointerOverGrid(event);
    if (!previewCell || !sourceSelection) {
      renderGrid();
      return;
    }
    const beforeSnapshot = captureSnapshot();
    const result = fillSelectionByRepeat(state.table, sourceSelection, previewCell);
    applySnapshot(
      {
        table: result.table,
        activeCell: previewCell,
        selection: result.selection,
      },
      {
        beforeSnapshot,
        recordHistory: hasTableChanged(beforeSnapshot.table, result.table),
        statusMessage: "Fill handle applied as copy-repeat.",
      },
    );
    return;
  }
  if (!state.selectionDrag) return;
  const dragState = state.selectionDrag;
  state.selectionDrag = null;
  if (!dragState.moved) {
    const releaseCell = resolvePointerCell(event) || dragState.focus || dragState.anchor;
    debugGridLog("document-mouseup:no-move", { releaseCell, startedOnActiveCell: dragState.startedOnActiveCell });
    if (releaseCell && dragState.extendSelection) {
      state.activeCell = releaseCell;
      state.selection = createRangeSelection(dragState.anchor, releaseCell);
      state.findReplace.currentMatch = null;
      state.suppressNextGridClick = isPointerOverGrid(event);
      renderSelectionState();
      return;
    }
    if (releaseCell && !dragState.startedOnActiveCell) {
      state.activeCell = releaseCell;
      state.selection = createCellSelection(releaseCell);
      state.findReplace.currentMatch = null;
      state.suppressNextGridClick = isPointerOverGrid(event);
      renderSelectionState();
    }
    return;
  }
  const releaseCell = resolvePointerCell(event) || dragState.focus;
  if (releaseCell) {
    state.activeCell = releaseCell;
    state.selection = createRangeSelection(dragState.anchor, releaseCell);
  }
  state.suppressNextGridClick = isPointerOverGrid(event);
  renderSelectionState();
}

function handleDocumentCopy(event) {
  if (isEditingTarget(event.target)) return;
  const text = selectionToDelimitedText(state.table, state.selection);
  if (event.clipboardData) {
    event.preventDefault();
    event.clipboardData.setData("text/plain", text);
  }
  rememberCopiedSelection();
  const bounds = getSelectionBounds(state.selection);
  setStatus(`Copied ${bounds.endRow - bounds.startRow + 1}x${bounds.endCol - bounds.startCol + 1} range.`);
}

function handleDocumentPaste(event) {
  if (isEditingTarget(event.target)) return;
  const text = event.clipboardData?.getData("text/plain") ?? "";
  const matrix = parseDelimitedText(text, { preserveEmptyCell: true });
  if (!matrix.length) return;
  event.preventDefault();
  applyClipboardMatrix(matrix, "Pasted clipboard range into the grid.");
}

async function handleDocumentKeydown(event) {
  const isPrimaryShortcut = event.ctrlKey || event.metaKey;
  const normalizedKey = String(event.key || "").toLowerCase();
  if (isPrimaryShortcut && !event.shiftKey && normalizedKey === "o") {
    event.preventDefault();
    setMenuOpen(null);
    await handleOpenAction();
    return;
  }
  if (isPrimaryShortcut && !event.shiftKey && normalizedKey === "s") {
    event.preventDefault();
    setMenuOpen(null);
    if (isSaveActionDisabled()) {
      setStatus("Save CWS HTML is unavailable when this file is opened directly from local file://. Use the browser download/open flow instead.");
      return;
    }
    await handleSave();
    return;
  }
  if (isPrimaryShortcut && !event.shiftKey && normalizedKey === "p") {
    event.preventDefault();
    setMenuOpen(null);
    handlePrint();
    return;
  }
  if (isPrimaryShortcut && !event.shiftKey && normalizedKey === "f") {
    event.preventDefault();
    openFindReplacePanel("find");
    return;
  }
  if (isPrimaryShortcut && normalizedKey === "h") {
    event.preventDefault();
    openFindReplacePanel("replace");
    return;
  }
  if (isPrimaryShortcut && !event.shiftKey && normalizedKey === "a") {
    event.preventDefault();
    state.activeCell = { row: 1, col: 1 };
    state.selection = createRangeSelection(
      { row: 1, col: 1 },
      { row: currentView.rowCount, col: currentView.columnCount },
    );
    state.findReplace.currentMatch = null;
    renderWorkspace();
    return;
  }
  if (event.key === "Escape") {
    if (state.headerMenu) {
      state.headerMenu = null;
      renderFloatingLayer();
      return;
    }
    if (state.openPanel) {
      closeOpenPanel();
      return;
    }
    if (state.openMenuKey) {
      setMenuOpen(null);
      return;
    }
    if (clearCopiedSelection({ render: true })) {
      event.preventDefault();
      setStatus("Cleared the copied range highlight.");
      return;
    }
  }
  if (isEditingTarget(event.target)) return;
  if (event.key === "Delete" && !event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey) {
    event.preventDefault();
    clearSelectedCells();
    return;
  }
  const editIntent = getKeyboardEditIntent(event);
  if (editIntent) {
    event.preventDefault();
    beginCellEdit({
      draft: applyKeyboardEditIntent(
        getCellValue(state.table, state.activeCell.row, state.activeCell.col),
        editIntent,
      ),
    });
    return;
  }
  const direction = keyToDirection(event);
  if (!direction) return;
  event.preventDefault();
  moveGridSelection(direction, event.shiftKey);
}

function moveGridSelection(direction, extendSelection) {
  const nextCell = moveActiveCell(state.activeCell, direction, {
    rowCount: currentView.rowCount,
    columnCount: currentView.columnCount,
  });
  const anchorCell = state.selection.anchor || state.activeCell;
  state.activeCell = nextCell;
  state.selection = extendSelection
    ? createRangeSelection(anchorCell, nextCell)
    : createCellSelection(nextCell);
  state.findReplace.currentMatch = null;
  renderWorkspace();
}

function beginCellEdit(options = {}) {
  state.selection = createCellSelection(state.activeCell);
  state.editing = {
    mode: "cell",
    draft: options.draft ?? getCellValue(state.table, state.activeCell.row, state.activeCell.col),
  };
  renderWorkspace();
}

function commitEditing(direction = "stay") {
  if (!state.editing) return;
  const beforeSnapshot = captureSnapshot();
  const nextTable = setCellValue(state.table, state.activeCell.row, state.activeCell.col, state.editing.draft);
  const nextView = buildSheetGridView(nextTable);
  let nextActiveCell = clampActiveCell(state.activeCell, nextView);
  if (direction !== "stay") {
    nextActiveCell = moveActiveCell(nextActiveCell, direction, {
      rowCount: nextView.rowCount,
      columnCount: nextView.columnCount,
    });
  }
  applySnapshot(
    {
      table: nextTable,
      activeCell: nextActiveCell,
      selection: createCellSelection(nextActiveCell),
    },
    {
      beforeSnapshot,
      recordHistory: hasTableChanged(beforeSnapshot.table, nextTable),
    },
  );
}

function cancelEditing() {
  state.editing = null;
  renderWorkspace();
}

async function handleSave() {
  commitEditing("stay");
  const html = serializeLightTableToCwsHtml(state.table, {
    fileName: state.sourceFileName || state.table.sourceName,
  });
  const fileName = normalizeDownloadName(state.sourceFileName || state.table.sourceName);
  const blob = new Blob([html], { type: "text/html;charset=utf-8" });
  try {
    const saveResult = await saveBlobWithPreferredTarget(blob, fileName, [
      {
        description: "CWS HTML",
        accept: {
          "text/html": [".html", ".htm"],
        },
      },
    ]);
    if (saveResult.kind === "handle") {
      const savedName = normalizeDownloadName(saveResult.handle?.name || fileName);
      updateSaveNameFromHandle(saveResult.handle);
      setStatus(`Saved CWS HTML as ${state.sourceFileName || savedName}.`);
    } else {
      saveWithDownloadLink(blob, fileName);
      setStatus(`Started CWS HTML download as ${fileName}. If no download appears, this browser may block file downloads.`);
    }
  } catch (error) {
    setStatus(formatSaveErrorMessage(error));
  }
}

async function saveBlobWithPreferredTarget(blob, fileName, types = []) {
  if (!shouldUseLocalSaveAsMode() && state.saveFileHandle) {
    try {
      await ensureWritePermissionForFileHandle(state.saveFileHandle);
      await writeBlobToFileHandle(state.saveFileHandle, blob);
      return { kind: "handle", handle: state.saveFileHandle };
    } catch (error) {
      if (!shouldRetryWithNewSaveHandle(error)) {
        throw error;
      }
      clearSaveFileHandle();
    }
  }
  const handle = await trySaveWithFilePicker(blob, fileName, types);
  if (handle) {
    return { kind: "handle", handle };
  }
  return { kind: "download" };
}

function shouldRetryWithNewSaveHandle(error) {
  const name = String(error?.name || "");
  return name === "NotAllowedError"
    || name === "InvalidStateError"
    || name === "SecurityError";
}

function shouldPersistSaveTarget() {
  return typeof window !== "undefined" && window.location?.protocol === "file:";
}

function shouldUseLocalSaveAsMode() {
  return shouldPersistSaveTarget();
}

function shouldUseFileHandleOpen() {
  return shouldPersistSaveTarget() && typeof window.showOpenFilePicker === "function";
}

async function writeBlobToFileHandle(handle, blob) {
  const writable = await handle.createWritable();
  await writable.write(blob);
  await writable.close();
}

async function ensureWritePermissionForFileHandle(handle) {
  if (!handle || typeof handle.queryPermission !== "function") {
    return;
  }
  const query = await handle.queryPermission({ mode: "readwrite" });
  if (query === "granted") {
    return;
  }
  if (typeof handle.requestPermission === "function") {
    const request = await handle.requestPermission({ mode: "readwrite" });
    if (request === "granted") {
      return;
    }
  }
  const error = new Error("Write permission was not granted for the selected file.");
  error.name = "NotAllowedError";
  throw error;
}

async function trySaveWithFilePicker(blob, fileName, types = []) {
  if (typeof window.showSaveFilePicker !== "function") {
    return false;
  }
  try {
    const handle = await window.showSaveFilePicker({
      suggestedName: fileName,
      types,
    });
    await writeBlobToFileHandle(handle, blob);
    return handle;
  } catch (error) {
    if (error?.name === "AbortError") {
      throw new Error("Save canceled.");
    }
    return false;
  }
}

function clearSaveFileHandle() {
  state.saveFileHandle = null;
}

function updateSaveNameFromHandle(handle) {
  if (!handle) return;
  if (typeof handle.name === "string" && handle.name) {
    state.sourceFileName = normalizeDownloadName(handle.name);
    renderDocumentSummary();
  }
}

function saveWithDownloadLink(blob, fileName) {
  const link = document.createElement("a");
  const objectUrl = URL.createObjectURL(blob);
  link.href = objectUrl;
  link.download = fileName;
  link.hidden = true;
  document.body.append(link);
  try {
    link.click();
  } finally {
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
  }
}

function formatSaveErrorMessage(error) {
  const message = String(error?.message || "");
  if (message === "Save canceled.") {
    return message;
  }
  if (/downloads are not supported/i.test(message)) {
    return "Save failed here because this browser does not support downloads. Open the app in a standard browser to save the CWS HTML file.";
  }
  return message
    ? `Save failed: ${message}`
    : "Save failed because this browser blocked the file save.";
}

function formatExportErrorMessage(error) {
  const message = String(error?.message || "");
  if (message === "Save canceled.") {
    return "Export canceled.";
  }
  if (/downloads are not supported/i.test(message)) {
    return "Export failed here because this browser does not support downloads. Open the app in a standard browser to export the file.";
  }
  return message
    ? `Export failed: ${message}`
    : "Export failed because this browser blocked the download.";
}

function handlePrint() {
  commitEditing("stay");
  state.headerMenu = null;
  setMenuOpen(null);
  startPrintFlow("Opened the browser print dialog.");
}

function setStatus(message) {
  state.statusMessageCanonical = String(message ?? "");
  refs.statusMessage.textContent = translateMessage(state.statusMessageCanonical);
}

function setMenuOpen(menuKey) {
  const previousMenuKey = state.openMenuKey;
  state.openMenuKey = getMenuItems(menuKey).length ? menuKey : null;
  state.openSubmenuKey = state.openMenuKey && state.openMenuKey === previousMenuKey
    ? state.openSubmenuKey
    : "";
  refs.menuButtons.forEach((button) => {
    const isOpen = button.dataset.menuKey === state.openMenuKey;
    button.setAttribute("aria-expanded", isOpen ? "true" : "false");
    button.classList.toggle("is-active", isOpen);
  });
  if (!state.openMenuKey) {
    refs.appMenu.classList.add("hidden");
    refs.appMenu.style.left = "";
    refs.appMenu.innerHTML = "";
    return;
  }
  renderMenu(state.openMenuKey);
}

function renderMenu(menuKey) {
  const button = refs.menuButtons.find((entry) => entry.dataset.menuKey === menuKey);
  const items = getMenuItems(menuKey);
  const submenuItems = getMenuItems(state.openSubmenuKey);
  refs.appMenu.setAttribute("aria-labelledby", button?.id || "");
  refs.appMenu.innerHTML = `
    <div class="menu-dropdown-column">
      ${items.map((item) => renderMenuItem(item, state.openSubmenuKey)).join("")}
    </div>
    ${submenuItems.length ? `<div class="menu-dropdown-column menu-dropdown-submenu">${submenuItems.map((item) => renderMenuItem(item, "")).join("")}</div>` : ""}
  `;
  refs.appMenu.classList.remove("hidden");
  refs.appMenu.style.left = `${button?.offsetLeft || 0}px`;
  localizeElement(refs.appMenu);
}

function renderMenuItem(item, activeSubmenuKey) {
  const hint = item.submenuKey ? ">" : (item.hint || "");
  const activeClass = item.submenuKey && item.submenuKey === activeSubmenuKey ? " is-active" : "";
  if (item.disabled) {
    return `<div class="menu-item is-disabled${activeClass}" role="menuitem" aria-disabled="true"><span>${escapeHtml(item.label)}</span><span class="menu-item-hint">${escapeHtml(hint)}</span></div>`;
  }
  const actionAttr = item.action ? ` data-menu-action="${escapeHtml(item.action)}"` : "";
  const submenuAttr = item.submenuKey ? ` data-menu-submenu="${escapeHtml(item.submenuKey)}" aria-haspopup="true"` : "";
  return `<button class="menu-item${activeClass}" type="button" role="menuitem"${actionAttr}${submenuAttr}><span>${escapeHtml(item.label)}</span><span class="menu-item-hint">${escapeHtml(hint)}</span></button>`;
}

function getSelectionRowCount() {
  const bounds = getSelectionBounds(state.selection);
  return bounds.endRow - bounds.startRow + 1;
}

function getSelectionColumnCount() {
  const bounds = getSelectionBounds(state.selection);
  return bounds.endCol - bounds.startCol + 1;
}

function isWholeSheetSelection() {
  const bounds = getSelectionBounds(state.selection);
  return bounds.startRow === 1
    && bounds.startCol === 1
    && bounds.endRow === currentView.rowCount
    && bounds.endCol === currentView.columnCount;
}

function shouldDisableInsertCells() {
  return isWholeSheetSelection() || state.selection.mode === "row" || state.selection.mode === "column";
}

function resolveRowInsertTarget(placement) {
  const bounds = getSelectionBounds(state.selection);
  if (state.selection.mode === "column") {
    return {
      anchorRow: state.activeCell.row,
      insertStartRow: placement === "above" ? state.activeCell.row : state.activeCell.row + 1,
      count: 1,
    };
  }
  return {
    anchorRow: placement === "above" ? bounds.startRow : bounds.endRow,
    insertStartRow: placement === "above" ? bounds.startRow : bounds.endRow + 1,
    count: getSelectionRowCount(),
  };
}

function resolveColumnInsertTarget(placement) {
  const bounds = getSelectionBounds(state.selection);
  if (state.selection.mode === "row") {
    return {
      anchorColumn: state.activeCell.col,
      insertStartColumn: placement === "left" ? state.activeCell.col : state.activeCell.col + 1,
      count: 1,
    };
  }
  return {
    anchorColumn: placement === "left" ? bounds.startCol : bounds.endCol,
    insertStartColumn: placement === "left" ? bounds.startCol : bounds.endCol + 1,
    count: getSelectionColumnCount(),
  };
}

function resolveRowDeleteTarget() {
  const bounds = getSelectionBounds(state.selection);
  if (state.selection.mode === "column") {
    return {
      startRow: state.activeCell.row,
      count: 1,
    };
  }
  return {
    startRow: bounds.startRow,
    count: getSelectionRowCount(),
  };
}

function resolveColumnDeleteTarget() {
  const bounds = getSelectionBounds(state.selection);
  if (state.selection.mode === "row") {
    return {
      startColumn: state.activeCell.col,
      count: 1,
    };
  }
  return {
    startColumn: bounds.startCol,
    count: getSelectionColumnCount(),
  };
}

function normalizeSelectionState() {
  if (!state.selection) {
    state.selection = createCellSelection(state.activeCell);
    return;
  }
  const anchor = clampActiveCell(state.selection.anchor || state.activeCell, currentView);
  const focus = clampActiveCell(state.selection.focus || state.activeCell, currentView);
  state.selection = state.selection.mode === "row"
    ? createRowSelection(anchor.row, currentView.columnCount, focus.row)
    : state.selection.mode === "column"
      ? createColumnSelection(anchor.col, currentView.rowCount, focus.col)
      : state.selection.mode === "range"
        ? createRangeSelection(anchor, focus)
        : createCellSelection(focus);
}

function selectSingleCell(cell) {
  state.activeCell = normalizeCell(cell);
  state.selection = createCellSelection(state.activeCell);
  state.editing = null;
  state.findReplace.currentMatch = null;
  renderWorkspace();
}

function resetGridInteractionState() {
  state.activeCell = { row: 1, col: 1 };
  state.selection = createCellSelection(state.activeCell);
  state.copiedSelection = null;
  state.editing = null;
  state.history = createHistoryState();
  state.fillDrag = null;
  state.selectionDrag = null;
  state.suppressNextGridClick = false;
  state.globalSearch = "";
  state.columnFilters = {};
  state.advancedFilter = { enabled: false, conditions: [], logic: "" };
  state.findReplace.currentMatch = null;
  state.openPanel = null;
  state.headerMenu = null;
  state.panelError = "";
  state.advancedFilterDraft = null;
  state.removeDuplicatesDraft = null;
  state.removeBlankRowsDraft = null;
  state.trimSpacesDraft = null;
  state.normalizeCaseDraft = null;
  state.numberCommasDraft = null;
  state.encodingIssuesDraft = null;
  state.sortDraft = null;
  state.pageSetupDraft = createPageSetupDraft(state.table.pageSetup);
  state.gridLimitsDraft = null;
  state.exportDraft = null;
}

function clampActiveCell(activeCell, view) {
  const normalized = normalizeCell(activeCell);
  return {
    row: Math.max(1, Math.min(view.rowCount, normalized.row)),
    col: Math.max(1, Math.min(view.columnCount, normalized.col)),
  };
}

function ensureActiveCellIsVisible() {
  if (isViewportPaddingRow(state.table, state.activeCell.row)) return;
  const visibleRows = getVisibleRowInfo().visibleRowSet;
  if (visibleRows.has(state.activeCell.row)) return;
  const nextRow = Array.from(visibleRows.values()).sort((left, right) => left - right)[0] || getHeaderRowIndex(state.table);
  state.activeCell = clampActiveCell({ row: nextRow, col: state.activeCell.col }, currentView);
  state.selection = createCellSelection(state.activeCell);
  state.findReplace.currentMatch = null;
}

function syncEditingFocus() {
  if (state.editing?.mode === "cell") {
    const editor = document.getElementById("activeCellEditor");
    if (editor) {
      editor.focus();
      editor.selectionStart = editor.value.length;
      editor.selectionEnd = editor.value.length;
      resizeTextarea(editor, 180);
    }
  }
  if (state.pendingFocusId) {
    const target = document.getElementById(state.pendingFocusId);
    if (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement || target instanceof HTMLSelectElement) {
      target.focus();
      if ("select" in target && typeof target.select === "function") {
        target.select();
      }
    }
    state.pendingFocusId = "";
  }
  if (state.pendingScrollCell) {
    const matchElement = refs.sheetGrid.querySelector(`[data-cell='true'][data-row='${state.pendingScrollCell.row}'][data-col='${state.pendingScrollCell.col}']`);
    matchElement?.scrollIntoView?.({ block: "nearest", inline: "nearest" });
    state.pendingScrollCell = null;
  }
}

function keyToDirection(event) {
  if (event.key === "ArrowLeft") return "left";
  if (event.key === "ArrowRight") return "right";
  if (event.key === "ArrowUp") return "up";
  if (event.key === "ArrowDown") return "down";
  if (event.key === "Tab") return event.shiftKey ? "shift+tab" : "tab";
  if (event.key === "Enter") return event.shiftKey ? "shift+enter" : "enter";
  return "";
}

function resizeTextarea(element, maxHeight) {
  element.style.height = "0px";
  element.style.height = `${Math.min(element.scrollHeight, maxHeight)}px`;
}

function handleGridViewportScroll() {
  const rowWindowChanged = syncGridViewportState();
  if (!rowWindowChanged) return;
  renderGrid();
  syncEditingFocus();
}

function handleGridViewportResize() {
  const rowWindowChanged = syncGridViewportState({ recomputeMetrics: true });
  if (!rowWindowChanged) return;
  renderGrid();
  syncEditingFocus();
}

function syncGridViewportState(options = {}) {
  if (options.recomputeMetrics) {
    rebuildGridRenderState({ honorPendingScroll: false });
    updateGridViewportDebugData();
    if (debugGridEvents) {
      renderDocumentSummary();
    }
    return true;
  }
  const previousRowWindowSignature = state.rowViewport.signature;
  const previousColumnWindowSignature = state.columnViewport.signature;
  const metrics = state.gridMetrics || buildGridMetrics(state.table, resolveGridMetricOptions());
  state.viewportState = updateViewportState(state.viewportState, refs.gridWrap, {
    contentWidth: metrics.contentWidth,
    contentHeight: metrics.contentHeight,
  });
  state.visibleRange = calculateVisibleRange(metrics, state.viewportState);
  state.visibleRangeSignature = visibleRangeSignature(state.visibleRange);
  updateRowViewportFromVisibleRange();
  updateColumnViewportFromVisibleRange();
  rebuildCurrentView();
  updateGridViewportDebugData();
  if (debugGridEvents) {
    renderDocumentSummary();
  }
  return state.rowViewport.signature !== previousRowWindowSignature
    || state.columnViewport.signature !== previousColumnWindowSignature;
}

function resolveGridMetricOptions() {
  const viewportState = state.viewportState || createViewportState();
  const visibleColumnCount = viewportState.width > 0
    ? estimateVisibleColumnCount(viewportState.width, buildGridMetrics(state.table))
    : undefined;
  const visibleRowCount = viewportState.height > 0
    ? estimateVisibleRowCount(viewportState.height, buildGridMetrics(state.table))
    : undefined;
  return {
    minColumnCount: visibleColumnCount,
    minRowCount: visibleRowCount,
  };
}

function updateGridViewportDebugData() {
  if (!refs.gridWrap || !refs.sheetGrid) return;
  refs.gridWrap.dataset.viewportRows = `${state.visibleRange.startRow}-${state.visibleRange.endRow}`;
  refs.gridWrap.dataset.viewportColumns = `${state.visibleRange.startColumn}-${state.visibleRange.endColumn}`;
  refs.sheetGrid.dataset.visibleRange = state.visibleRangeSignature;
}

function rebuildGridRenderState(options = {}) {
  state.logicalGridMetrics = buildGridMetrics(state.table, resolveGridMetricOptions());
  state.visibleRowLayout = buildVisibleRowLayout(
    state.table,
    getVisibleRowInfo().visibleRowSet,
    state.logicalGridMetrics.rowCount,
  );
  state.rowVirtualizationEnabled = !state.printing
    && shouldUseRowVirtualization(state.visibleRowLayout.totalVisibleRowCount);
  state.columnVirtualizationEnabled = !state.printing
    && shouldUseColumnVirtualization(state.logicalGridMetrics.columnCount);
  state.gridMetrics = state.rowVirtualizationEnabled
    ? buildVirtualizedGridMetrics(state.logicalGridMetrics, state.visibleRowLayout.totalVisibleRowCount)
    : state.logicalGridMetrics;
  state.viewportState = updateViewportState(state.viewportState, refs.gridWrap, {
    contentWidth: state.gridMetrics.contentWidth,
    contentHeight: state.gridMetrics.contentHeight,
  });
  state.visibleRange = calculateVisibleRange(state.gridMetrics, state.viewportState);
  if (options.honorPendingScroll) {
    applyPendingScrollCellToViewport();
  }
  state.visibleRangeSignature = visibleRangeSignature(state.visibleRange);
  updateRowViewportFromVisibleRange();
  updateColumnViewportFromVisibleRange();
  rebuildCurrentView();
}

function buildVisibleRowLayout(table, visibleRows, fullRowCount) {
  const totalRowCount = Math.max(1, Math.trunc(Number(fullRowCount) || 1));
  const tableRowCount = getSheetRowCount(table);
  const logicalRows = [];
  const rowSlotByIndex = new Map();
  for (let rowIndex = 1; rowIndex <= totalRowCount; rowIndex += 1) {
    if (rowIndex > tableRowCount || visibleRows.has(rowIndex)) {
      logicalRows.push(rowIndex);
      rowSlotByIndex.set(rowIndex, logicalRows.length);
    }
  }
  if (!logicalRows.length) {
    logicalRows.push(1);
    rowSlotByIndex.set(1, 1);
  }
  return {
    fullRowCount: totalRowCount,
    logicalRows,
    rowSlotByIndex,
    totalVisibleRowCount: logicalRows.length,
  };
}

function buildVirtualizedGridMetrics(metrics, visibleRowCount) {
  const safeMetrics = metrics || buildGridMetrics(state.table, resolveGridMetricOptions());
  const rowCount = Math.max(1, Math.trunc(Number(visibleRowCount) || 1));
  return {
    ...safeMetrics,
    rowCount,
    contentHeight: safeMetrics.columnHeaderHeight + rowCount * safeMetrics.rowHeight,
  };
}

function updateRowViewportFromVisibleRange() {
  const layout = state.visibleRowLayout || createEmptyVisibleRowLayout();
  const metrics = state.gridMetrics || buildGridMetrics(state.table, resolveGridMetricOptions());
  if (!state.rowVirtualizationEnabled) {
    state.rowViewport = {
      startSlot: 1,
      endSlot: Math.max(1, layout.totalVisibleRowCount || 1),
      renderedRowIndices: [...layout.logicalRows],
      topSpacerHeight: 0,
      bottomSpacerHeight: 0,
      signature: `full:${layout.totalVisibleRowCount}`,
    };
    return;
  }
  const totalVisibleRows = Math.max(1, layout.totalVisibleRowCount || 1);
  const startSlot = clampNumber(state.visibleRange.startRow, 1, totalVisibleRows);
  const endSlot = clampNumber(state.visibleRange.endRow, startSlot, totalVisibleRows);
  const renderedRowIndices = layout.logicalRows.slice(startSlot - 1, endSlot);
  state.rowViewport = {
    startSlot,
    endSlot,
    renderedRowIndices,
    topSpacerHeight: (startSlot - 1) * metrics.rowHeight,
    bottomSpacerHeight: Math.max(0, totalVisibleRows - endSlot) * metrics.rowHeight,
    signature: `${startSlot}:${endSlot}:${layout.totalVisibleRowCount}`,
  };
}

function updateColumnViewportFromVisibleRange() {
  const metrics = state.logicalGridMetrics || state.gridMetrics || buildGridMetrics(state.table, resolveGridMetricOptions());
  const totalColumnCount = Math.max(1, metrics.columnCount || 1);
  if (!state.columnVirtualizationEnabled) {
    state.columnViewport = {
      startColumn: 1,
      endColumn: totalColumnCount,
      renderedColumnIndices: createIndexRange(1, totalColumnCount),
      leftSpacerWidth: 0,
      rightSpacerWidth: 0,
      signature: `full:${totalColumnCount}`,
    };
    return;
  }
  const startColumn = clampNumber(state.visibleRange.startColumn, 1, totalColumnCount);
  const endColumn = clampNumber(state.visibleRange.endColumn, startColumn, totalColumnCount);
  state.columnViewport = {
    startColumn,
    endColumn,
    renderedColumnIndices: createIndexRange(startColumn, endColumn),
    leftSpacerWidth: (startColumn - 1) * metrics.columnWidth,
    rightSpacerWidth: Math.max(0, totalColumnCount - endColumn) * metrics.columnWidth,
    signature: `${startColumn}:${endColumn}:${totalColumnCount}`,
  };
}

function rebuildCurrentView() {
  const metrics = state.gridMetrics || buildGridMetrics(state.table, resolveGridMetricOptions());
  const layout = state.visibleRowLayout || createEmptyVisibleRowLayout();
  currentView = buildSheetGridView(state.table, {
    minColumnCount: Math.max(1, state.logicalGridMetrics?.columnCount || metrics.columnCount),
    minRowCount: Math.max(1, layout.fullRowCount || state.logicalGridMetrics?.rowCount || metrics.rowCount),
    rowIndices: state.rowViewport.renderedRowIndices,
    columnIndices: state.columnViewport.renderedColumnIndices,
  });
}

function applyPendingScrollCellToViewport() {
  if (!state.pendingScrollCell || !refs.gridWrap) return;
  const rowSlot = state.visibleRowLayout.rowSlotByIndex.get(state.pendingScrollCell.row);
  const nextScrollTop = rowSlot
    ? resolveScrollTopForRowSlot(rowSlot, state.viewportState, state.gridMetrics)
    : state.viewportState.scrollTop;
  const nextScrollLeft = resolveScrollLeftForColumn(state.pendingScrollCell.col, state.viewportState, state.gridMetrics);
  refs.gridWrap.scrollTop = nextScrollTop;
  refs.gridWrap.scrollLeft = nextScrollLeft;
  state.viewportState = updateViewportState(state.viewportState, refs.gridWrap, {
    contentWidth: state.gridMetrics.contentWidth,
    contentHeight: state.gridMetrics.contentHeight,
  });
  state.visibleRange = calculateVisibleRange(state.gridMetrics, state.viewportState);
}

function resolveScrollTopForRowSlot(rowSlot, viewportState, metrics) {
  const safeSlot = Math.max(1, Math.trunc(Number(rowSlot) || 1));
  const bodyHeight = Math.max(0, viewportState.height - metrics.columnHeaderHeight);
  const rowTop = (safeSlot - 1) * metrics.rowHeight;
  const rowBottom = rowTop + metrics.rowHeight;
  if (rowTop < viewportState.scrollTop) {
    return rowTop;
  }
  if (rowBottom > viewportState.scrollTop + bodyHeight) {
    return Math.max(0, rowBottom - bodyHeight);
  }
  return viewportState.scrollTop;
}

function resolveScrollLeftForColumn(columnIndex, viewportState, metrics) {
  const safeColumn = Math.max(1, Math.trunc(Number(columnIndex) || 1));
  const bodyWidth = Math.max(0, viewportState.width - metrics.rowHeaderWidth);
  const columnLeft = (safeColumn - 1) * metrics.columnWidth;
  const columnRight = columnLeft + metrics.columnWidth;
  if (columnLeft < viewportState.scrollLeft) {
    return columnLeft;
  }
  if (columnRight > viewportState.scrollLeft + bodyWidth) {
    return Math.max(0, columnRight - bodyWidth);
  }
  return viewportState.scrollLeft;
}

function createEmptyVisibleRowLayout() {
  return {
    fullRowCount: 1,
    logicalRows: [1],
    rowSlotByIndex: new Map([[1, 1]]),
    totalVisibleRowCount: 1,
  };
}

function createEmptyRowViewport() {
  return {
    startSlot: 1,
    endSlot: 1,
    renderedRowIndices: [1],
    topSpacerHeight: 0,
    bottomSpacerHeight: 0,
    signature: "",
  };
}

function createEmptyColumnViewport() {
  return {
    startColumn: 1,
    endColumn: 1,
    renderedColumnIndices: [1],
    leftSpacerWidth: 0,
    rightSpacerWidth: 0,
    signature: "",
  };
}

function shouldUseRowVirtualization(totalVisibleRowCount) {
  return Math.max(1, Math.trunc(Number(totalVisibleRowCount) || 1)) > MIN_VIRTUALIZED_ROW_COUNT;
}

function shouldUseColumnVirtualization(totalColumnCount) {
  return Math.max(1, Math.trunc(Number(totalColumnCount) || 1)) > MIN_VIRTUALIZED_COLUMN_COUNT;
}

function createIndexRange(start, end) {
  const safeStart = Math.max(1, Math.trunc(Number(start) || 1));
  const safeEnd = Math.max(safeStart, Math.trunc(Number(end) || safeStart));
  return Array.from({ length: safeEnd - safeStart + 1 }, (_, offset) => safeStart + offset);
}

function getRenderedGridColumnSpan() {
  return 1
    + currentView.columns.length
    + (state.columnViewport.leftSpacerWidth > 0 ? 1 : 0)
    + (state.columnViewport.rightSpacerWidth > 0 ? 1 : 0);
}

function getMenuItems(menuKey) {
  if (menuKey === "file") {
    return MENU_DEFINITIONS.file.map((item) => (
      item.action === "save"
        ? { ...item, disabled: isSaveActionDisabled() }
        : item
    ));
  }
  if (menuKey === "edit") {
    return [
      { label: "Undo", action: "undo", disabled: !canUndo(state.history), hint: "Ctrl+Z" },
      { label: "Redo", action: "redo", disabled: !canRedo(state.history), hint: "Ctrl+Y" },
      { label: "Copy Range", action: "copy", disabled: Boolean(state.editing), hint: "Ctrl+C" },
      { label: "Paste Range", action: "paste", disabled: Boolean(state.editing), hint: "Ctrl+V" },
    ];
  }
  if (menuKey === "view") {
    return [
      { label: "Insert Row Above", action: "insert-row-above", disabled: isWholeSheetSelection() },
      { label: "Insert Row Below", action: "insert-row-below", disabled: isWholeSheetSelection() },
      { label: "Insert Column Left", action: "insert-column-left", disabled: isWholeSheetSelection() },
      { label: "Insert Column Right", action: "insert-column-right", disabled: isWholeSheetSelection() },
      { label: "Insert Cells", submenuKey: "insert-cells", disabled: shouldDisableInsertCells() },
      { label: "Delete Rows", action: "delete-rows", disabled: isWholeSheetSelection() },
      { label: "Delete Columns", action: "delete-columns", disabled: isWholeSheetSelection() },
    ];
  }
  if (menuKey === "insert-cells") {
    return [
      { label: "Shift Right", action: "insert-cells-right", disabled: shouldDisableInsertCells() },
      { label: "Shift Down", action: "insert-cells-down", disabled: shouldDisableInsertCells() },
    ];
  }
  if (menuKey === "language") {
    return SUPPORTED_LANGUAGES.map((language) => ({
      label: `${state.language === language ? "✓ " : ""}${LANGUAGE_LABELS[language]}`,
      action: `language-${language}`,
    }));
  }
  return MENU_DEFINITIONS[menuKey] || [];
}

function isSaveActionDisabled() {
  return shouldUseLocalSaveAsMode();
}

async function handleCopyCommand() {
  const text = selectionToDelimitedText(state.table, state.selection);
  const bounds = getSelectionBounds(state.selection);
  const copiedMessage = `Copied ${bounds.endRow - bounds.startRow + 1}x${bounds.endCol - bounds.startCol + 1} range.`;
  rememberCopiedSelection();
  if (!navigator.clipboard?.writeText) {
    setStatus(`${copiedMessage} Clipboard write is not available here; use Ctrl+C / Cmd+C inside the grid if you need the system clipboard too.`);
    return;
  }
  try {
    await navigator.clipboard.writeText(text);
    setStatus(copiedMessage);
  } catch (error) {
    const detail = String(error?.message || "").trim();
    setStatus(detail
      ? `${copiedMessage} Clipboard write failed here: ${detail}`
      : `${copiedMessage} Clipboard write is blocked here.`);
  }
}

async function handlePasteCommand() {
  if (!navigator.clipboard?.readText) {
    setStatus("Paste command is not available here. Use Ctrl+V / Cmd+V inside the grid.");
    return;
  }
  try {
    const text = await navigator.clipboard.readText();
    const matrix = parseDelimitedText(text, { preserveEmptyCell: true });
    if (!matrix.length) {
      setStatus("Clipboard is empty.");
      return;
    }
    applyClipboardMatrix(matrix, "Pasted clipboard range into the grid.");
  } catch (error) {
    setStatus(error?.message || "Paste command is blocked here. Use Ctrl+V / Cmd+V inside the grid.");
  }
}

function applyClipboardMatrix(matrix, statusMessage) {
  const beforeSnapshot = captureSnapshot();
  const writeResult = applyMatrixToTable(state.table, state.activeCell, matrix);
  applySnapshot(
    {
      table: writeResult.table,
      activeCell: state.activeCell,
      selection: selectionFromMatrix(state.activeCell, matrix),
    },
    {
      beforeSnapshot,
      recordHistory: writeResult.changed,
      statusMessage,
    },
  );
}

function clearSelectedCells() {
  commitEditing("stay");
  const bounds = getSelectionBounds(state.selection);
  const width = bounds.endCol - bounds.startCol + 1;
  const height = bounds.endRow - bounds.startRow + 1;
  const blankMatrix = Array.from(
    { length: height },
    () => Array.from({ length: width }, () => ""),
  );
  const beforeSnapshot = captureSnapshot();
  const writeResult = applyMatrixToTable(state.table, { row: bounds.startRow, col: bounds.startCol }, blankMatrix);
  clearCopiedSelection();
  applySnapshot(
    {
      table: writeResult.table,
      activeCell: state.activeCell,
      selection: cloneJsonValue(state.selection),
    },
    {
      beforeSnapshot,
      recordHistory: writeResult.changed,
      statusMessage: writeResult.changed
        ? `Cleared ${height}x${width} selected cells.`
        : "Selected cells were already blank.",
    },
  );
}

function handleApplyRemoveDuplicates() {
  commitEditing("stay");
  const draft = normalizeCleaningDraftColumns(state.removeDuplicatesDraft || createRemoveDuplicatesDraft());
  state.removeDuplicatesDraft = draft;
  const rows = resolveCleaningRows(draft.scope, { skipHeader: true });
  const columns = getSelectedCleaningColumns(draft);
  if (!rows.length) {
    state.panelError = "No data rows are available to clean.";
    renderWorkspace();
    return;
  }
  if (!columns.length) {
    state.panelError = "Select at least one column.";
    renderWorkspace();
    return;
  }
  const beforeSnapshot = captureSnapshot();
  const result = removeDuplicateRows(state.table, { rows, columns });
  const nextView = buildSheetGridView(result.table);
  const nextActiveCell = clampActiveCell(state.activeCell, nextView);
  clearCopiedSelection();
  applySnapshot(
    {
      table: result.table,
      activeCell: nextActiveCell,
      selection: createCellSelection(nextActiveCell),
    },
    {
      beforeSnapshot,
      recordHistory: result.changed,
      statusMessage: result.changed
        ? `Removed ${result.removedRowCount} duplicate row${result.removedRowCount === 1 ? "" : "s"} from ${result.checkedRowCount} checked rows.`
        : `No duplicate rows found in ${result.checkedRowCount} checked rows.`,
    },
  );
}

function handleApplyRemoveBlankRows() {
  commitEditing("stay");
  const draft = normalizeRemoveBlankRowsDraft(state.removeBlankRowsDraft || createRemoveBlankRowsDraft());
  state.removeBlankRowsDraft = draft;
  const rows = resolveCleaningRows(draft.scope, { skipHeader: true });
  const columns = getBlankRowColumns();
  if (!rows.length) {
    state.panelError = "No data rows are available to clean.";
    renderWorkspace();
    return;
  }
  const beforeSnapshot = captureSnapshot();
  const result = removeBlankRows(state.table, { rows, columns });
  const nextView = buildSheetGridView(result.table);
  const nextActiveCell = clampActiveCell(state.activeCell, nextView);
  clearCopiedSelection();
  applySnapshot(
    {
      table: result.table,
      activeCell: nextActiveCell,
      selection: createCellSelection(nextActiveCell),
    },
    {
      beforeSnapshot,
      recordHistory: result.changed,
      statusMessage: result.changed
        ? `Removed ${result.removedRowCount} blank row${result.removedRowCount === 1 ? "" : "s"} from ${result.checkedRowCount} checked rows.`
        : `No blank rows found in ${result.checkedRowCount} checked rows.`,
    },
  );
}

function handleApplyTrimSpaces() {
  commitEditing("stay");
  const draft = normalizeTrimSpacesDraft(state.trimSpacesDraft || createTrimSpacesDraft());
  state.trimSpacesDraft = draft;
  const rows = resolveCleaningRows(draft.scope, { skipHeader: draft.scope !== "selection" });
  const columns = getTrimSpacesColumns(draft.scope);
  if (!rows.length) {
    state.panelError = "No data rows are available to clean.";
    renderWorkspace();
    return;
  }
  if (!hasTrimSpaceOption(draft)) {
    state.panelError = "Select at least one space cleanup option.";
    renderWorkspace();
    return;
  }
  const beforeSnapshot = captureSnapshot();
  const result = trimSpacesInCells(state.table, {
    rows,
    columns,
    leading: draft.leading,
    trailing: draft.trailing,
    inner: draft.inner,
    innerMode: draft.innerMode,
  });
  clearCopiedSelection();
  applySnapshot(
    {
      table: result.table,
      activeCell: state.activeCell,
      selection: state.selection,
    },
    {
      beforeSnapshot,
      recordHistory: result.changed,
      statusMessage: result.changed
        ? `Cleaned spaces in ${result.changedCellCount} cell${result.changedCellCount === 1 ? "" : "s"} from ${result.checkedCellCount} checked cells.`
        : `No spaces needed cleanup in ${result.checkedCellCount} checked cells.`,
    },
  );
}

function handleApplyNormalizeCase() {
  commitEditing("stay");
  const draft = normalizeNormalizeCaseDraft(state.normalizeCaseDraft || createNormalizeCaseDraft());
  state.normalizeCaseDraft = draft;
  const rows = resolveCleaningRows(draft.scope, { skipHeader: draft.scope !== "selection" });
  const columns = getTrimSpacesColumns(draft.scope);
  if (!rows.length) {
    state.panelError = "No data rows are available to clean.";
    renderWorkspace();
    return;
  }
  const beforeSnapshot = captureSnapshot();
  const result = normalizeCaseInCells(state.table, {
    rows,
    columns,
    mode: draft.mode,
  });
  clearCopiedSelection();
  applySnapshot(
    {
      table: result.table,
      activeCell: state.activeCell,
      selection: state.selection,
    },
    {
      beforeSnapshot,
      recordHistory: result.changed,
      statusMessage: result.changed
        ? `Converted text in ${result.changedCellCount} cell${result.changedCellCount === 1 ? "" : "s"} from ${result.checkedCellCount} checked cells.`
        : `No text conversion needed in ${result.checkedCellCount} checked cells.`,
    },
  );
}

function handleApplyNumberCommas() {
  commitEditing("stay");
  const draft = normalizeNumberCommasDraft(state.numberCommasDraft || createNumberCommasDraft());
  state.numberCommasDraft = draft;
  const rows = resolveCleaningRows(draft.scope, { skipHeader: draft.scope !== "selection" });
  const columns = getTrimSpacesColumns(draft.scope);
  if (!rows.length) {
    state.panelError = "No data rows are available to clean.";
    renderWorkspace();
    return;
  }
  const beforeSnapshot = captureSnapshot();
  const result = formatNumberCommasInCells(state.table, {
    rows,
    columns,
    mode: draft.mode,
  });
  clearCopiedSelection();
  applySnapshot(
    {
      table: result.table,
      activeCell: state.activeCell,
      selection: state.selection,
    },
    {
      beforeSnapshot,
      recordHistory: result.changed,
      statusMessage: result.changed
        ? `Updated number commas in ${result.changedCellCount} cell${result.changedCellCount === 1 ? "" : "s"} from ${result.numericCellCount} numeric cells.`
        : `No number comma changes needed in ${result.numericCellCount} numeric cells.`,
    },
  );
}

function jumpToEncodingIssue(rowIndex, columnIndex) {
  const row = Math.max(1, Math.trunc(Number(rowIndex) || 1));
  const col = Math.max(1, Math.trunc(Number(columnIndex) || 1));
  commitEditing("stay");
  state.activeCell = { row, col };
  state.selection = createCellSelection(state.activeCell);
  state.findReplace.currentMatch = null;
  state.pendingScrollCell = state.activeCell;
  state.panelError = "";
  renderWorkspace();
  setStatus(`Selected ${gridCellName(row, col)} from encoding issue results.`);
}

function handleUndo() {
  commitEditing("stay");
  const currentSnapshot = captureSnapshot();
  const result = undoHistory(state.history, currentSnapshot);
  if (!result.snapshot) {
    setStatus("Nothing to undo.");
    return;
  }
  state.history = result.history;
  applySnapshot(result.snapshot, {
    statusMessage: "Undid the last change.",
  });
}

function handleRedo() {
  commitEditing("stay");
  const currentSnapshot = captureSnapshot();
  const result = redoHistory(state.history, currentSnapshot);
  if (!result.snapshot) {
    setStatus("Nothing to redo.");
    return;
  }
  state.history = result.history;
  applySnapshot(result.snapshot, {
    statusMessage: "Redid the last change.",
  });
}

function captureSnapshot() {
  return {
    table: state.table,
    activeCell: { ...state.activeCell },
    selection: cloneJsonValue(state.selection),
  };
}

function applySnapshot(snapshot, options = {}) {
  if (options.recordHistory && options.beforeSnapshot) {
    state.history = pushHistorySnapshot(state.history, options.beforeSnapshot);
  }
  state.table = snapshot.table;
  state.activeCell = normalizeCell(snapshot.activeCell);
  state.selection = cloneJsonValue(snapshot.selection) || createCellSelection(state.activeCell);
  state.editing = null;
  state.fillDrag = null;
  state.findReplace.currentMatch = null;
  renderWorkspace();
  if (options.statusMessage) {
    setStatus(options.statusMessage);
  }
}

function hasTableChanged(left, right) {
  return JSON.stringify(left) !== JSON.stringify(right);
}

function shouldRenderFillHandle(rowIndex, columnIndex) {
  if (state.editing || state.fillDrag || state.selectionDrag) return false;
  if (!["cell", "range"].includes(state.selection.mode)) return false;
  const bounds = getSelectionBounds(state.selection);
  return rowIndex === bounds.endRow && columnIndex === bounds.endCol;
}

function isFillPreviewCell(rowIndex, columnIndex) {
  if (!state.fillDrag) return false;
  const sourceBounds = getSelectionBounds(state.fillDrag.sourceSelection);
  const previewBounds = getFillPreviewBounds();
  if (!previewBounds) return false;
  return isCellInsideBounds(rowIndex, columnIndex, previewBounds)
    && !isCellInsideBounds(rowIndex, columnIndex, sourceBounds);
}

function getCopiedSelectionBounds() {
  return state.copiedSelection ? getSelectionBounds(state.copiedSelection) : null;
}

function getCopiedRangeCellClasses(copiedBounds, rowIndex, columnIndex) {
  if (!copiedBounds || !isCellInsideBounds(rowIndex, columnIndex, copiedBounds)) {
    return [];
  }
  const edgeClasses = [];
  if (rowIndex === copiedBounds.startRow) {
    edgeClasses.push("copy-edge-top");
  }
  if (rowIndex === copiedBounds.endRow) {
    edgeClasses.push("copy-edge-bottom");
  }
  if (columnIndex === copiedBounds.startCol) {
    edgeClasses.push("copy-edge-left");
  }
  if (columnIndex === copiedBounds.endCol) {
    edgeClasses.push("copy-edge-right");
  }
  return edgeClasses.length ? ["is-copied-range", ...edgeClasses] : [];
}

function rememberCopiedSelection(selection = state.selection) {
  state.copiedSelection = cloneJsonValue(selection);
  renderGrid();
}

function clearCopiedSelection(options = {}) {
  if (!state.copiedSelection) {
    return false;
  }
  state.copiedSelection = null;
  if (options.render) {
    renderGrid();
  }
  return true;
}

function getFillPreviewBounds() {
  if (!state.fillDrag?.previewCell) return null;
  const sourceBounds = getSelectionBounds(state.fillDrag.sourceSelection);
  return {
    startRow: Math.min(sourceBounds.startRow, state.fillDrag.previewCell.row),
    endRow: Math.max(sourceBounds.endRow, state.fillDrag.previewCell.row),
    startCol: Math.min(sourceBounds.startCol, state.fillDrag.previewCell.col),
    endCol: Math.max(sourceBounds.endCol, state.fillDrag.previewCell.col),
  };
}

function isCellInsideBounds(rowIndex, columnIndex, bounds) {
  return rowIndex >= bounds.startRow
    && rowIndex <= bounds.endRow
    && columnIndex >= bounds.startCol
    && columnIndex <= bounds.endCol;
}

function isUndoShortcut(event) {
  return (event.ctrlKey || event.metaKey)
    && !event.shiftKey
    && String(event.key || "").toLowerCase() === "z";
}

function isRedoShortcut(event) {
  const key = String(event.key || "").toLowerCase();
  return (event.ctrlKey || event.metaKey)
    && (key === "y" || (key === "z" && event.shiftKey));
}

function handleGlobalUndoRedoKeydown(event) {
  if (event.defaultPrevented || isEditingTarget(event.target)) return;
  if (isUndoShortcut(event)) {
    event.preventDefault();
    event.stopPropagation();
    handleUndo();
    return;
  }
  if (isRedoShortcut(event)) {
    event.preventDefault();
    event.stopPropagation();
    handleRedo();
  }
}

function renderSelectionState() {
  renderFormulaBar();
  renderGrid();
}

function resolvePointerCell(event) {
  const pointerElement = resolvePointerElement(event);
  const hoveredCell = pointerElement?.closest?.("[data-cell='true']");
  return hoveredCell ? cellFromDataset(hoveredCell) : null;
}

function resolvePointerElement(event) {
  if (!event) return null;
  if (event.target instanceof Element) {
    return event.target;
  }
  if (typeof event.clientX !== "number" || typeof event.clientY !== "number") {
    return null;
  }
  return document.elementFromPoint(event.clientX, event.clientY);
}

function isPointerOverGrid(event) {
  return Boolean(resolvePointerElement(event)?.closest?.("#sheetGrid"));
}

function insertLineBreakAtCursor(element, mode) {
  const selectionStart = element.selectionStart ?? element.value.length;
  const selectionEnd = element.selectionEnd ?? element.value.length;
  const nextValue = `${element.value.slice(0, selectionStart)}\n${element.value.slice(selectionEnd)}`;
  element.value = nextValue;
  element.selectionStart = selectionStart + 1;
  element.selectionEnd = selectionStart + 1;
  state.selection = createCellSelection(state.activeCell);
  state.editing = {
    mode,
    draft: nextValue,
  };
  if (mode === "cell" && refs.formulaBarInput.value !== nextValue) {
    refs.formulaBarInput.value = nextValue;
  }
  resizeTextarea(element, mode === "cell" ? 180 : 100);
  resizeTextarea(refs.formulaBarInput, 100);
}

function rememberTextExportSettings(nextSettings = state.exportDraft) {
  if (!nextSettings) return;
  state.exportSettings = normalizeTextExportSettings(nextSettings);
}

function buildFilterState() {
  return {
    globalSearch: state.globalSearch,
    columnFilters: state.columnFilters,
    advancedFilter: state.advancedFilter,
  };
}

function getVisibleRowInfo() {
  const filterState = buildFilterState();
  const filterSignature = buildFilterSignature(filterState);
  const cached = state.visibleRowsCache;
  if (cached && cached.table === state.table && cached.filterSignature === filterSignature) {
    return cached;
  }
  const visibleRowSet = getVisibleRowSet(state.table, filterState);
  const headerRowIndex = getHeaderRowIndex(state.table);
  let visibleDataRowCount = 0;
  visibleRowSet.forEach((rowIndex) => {
    if (rowIndex > headerRowIndex) {
      visibleDataRowCount += 1;
    }
  });
  const nextCache = {
    table: state.table,
    filterSignature,
    visibleRowSet,
    visibleDataRowCount,
  };
  state.visibleRowsCache = nextCache;
  return nextCache;
}

function buildFilterSignature(filterState) {
  return JSON.stringify({
    globalSearch: String(filterState?.globalSearch || ""),
    columnFilters: filterState?.columnFilters || {},
    advancedFilter: filterState?.advancedFilter || {},
  });
}

function countVisibleDataRows() {
  return getVisibleRowInfo().visibleDataRowCount;
}

function buildVisibleRowStatus(prefix) {
  const totalRows = state.table.rows.length;
  const visibleRows = countVisibleDataRows();
  return `${prefix} Showing ${visibleRows} of ${totalRows} data rows.`;
}

function rowWithinPrintArea(rowIndex, printBounds) {
  return rowIndex >= printBounds.startRow && rowIndex <= printBounds.endRow;
}

function openFindReplacePanel(mode) {
  state.openPanel = "find-replace";
  state.headerMenu = null;
  state.panelError = "";
  state.findReplace.scope = getDefaultFindReplaceScope();
  state.findReplace.scopeBounds = state.findReplace.scope === "selection" ? getSelectionBounds(state.selection) : null;
  state.findReplace.currentMatch = null;
  state.pendingFocusId = mode === "replace" ? "replaceTextInput" : "findTextInput";
  renderWorkspace();
}

function openAdvancedFilterPanel() {
  state.openPanel = "advanced-filter";
  state.headerMenu = null;
  state.panelError = "";
  state.advancedFilterDraft = createAdvancedFilterDraft();
  state.pendingFocusId = "advancedFilterLogicInput";
  renderWorkspace();
}

function openSortPanel() {
  state.openPanel = "sort";
  state.headerMenu = null;
  state.panelError = "";
  state.sortDraft = createSortDraft();
  state.pendingFocusId = "sortPrimaryColumn";
  renderWorkspace();
}

function openPageSetupPanel() {
  state.openPanel = "page-setup";
  state.headerMenu = null;
  state.panelError = "";
  state.pageSetupDraft = createPageSetupDraft(state.table.pageSetup);
  state.pendingFocusId = "pageSetupPaperSize";
  renderWorkspace();
}

function openGridLimitsPanel() {
  state.openPanel = "grid-limits";
  state.headerMenu = null;
  state.panelError = "";
  state.gridLimitsDraft = createGridLimitsDraft();
  state.pendingFocusId = "gridLimitsRowsInput";
  renderWorkspace();
}

function openRemoveDuplicatesPanel() {
  state.openPanel = "remove-duplicates";
  state.headerMenu = null;
  state.panelError = "";
  state.removeDuplicatesDraft = createRemoveDuplicatesDraft();
  state.pendingFocusId = "removeDuplicatesScope";
  renderWorkspace();
}

function openRemoveBlankRowsPanel() {
  state.openPanel = "remove-blank-rows";
  state.headerMenu = null;
  state.panelError = "";
  state.removeBlankRowsDraft = createRemoveBlankRowsDraft();
  state.pendingFocusId = "removeBlankRowsScope";
  renderWorkspace();
}

function openTrimSpacesPanel() {
  state.openPanel = "trim-spaces";
  state.headerMenu = null;
  state.panelError = "";
  state.trimSpacesDraft = createTrimSpacesDraft();
  state.pendingFocusId = "trimSpacesScope";
  renderWorkspace();
}

function openNormalizeCasePanel() {
  state.openPanel = "normalize-case";
  state.headerMenu = null;
  state.panelError = "";
  state.normalizeCaseDraft = createNormalizeCaseDraft();
  state.pendingFocusId = "normalizeCaseScope";
  renderWorkspace();
}

function openNumberCommasPanel() {
  state.openPanel = "number-commas";
  state.headerMenu = null;
  state.panelError = "";
  state.numberCommasDraft = createNumberCommasDraft();
  state.pendingFocusId = "numberCommasScope";
  renderWorkspace();
}

function openFindEncodingIssuesPanel() {
  state.openPanel = "find-encoding-issues";
  state.headerMenu = null;
  state.panelError = "";
  state.encodingIssuesDraft = createFindEncodingIssuesDraft();
  state.pendingFocusId = "encodingIssuesScope";
  renderWorkspace();
}

function openHelpPanel() {
  state.openPanel = "help-page";
  state.headerMenu = null;
  state.panelError = "";
  state.pendingFocusId = "";
  renderWorkspace();
}

function openVersionPanel() {
  state.openPanel = "version-page";
  state.headerMenu = null;
  state.panelError = "";
  state.pendingFocusId = "";
  renderWorkspace();
}

function openHelpSite() {
  const opened = window.open(HELP_SITE_URL, "_blank", "noopener");
  setStatus(opened ? "Opened the documentation site." : "The documentation site could not be opened automatically.");
}

function openTextExportPanel(target) {
  state.openPanel = "text-export";
  state.headerMenu = null;
  state.panelError = "";
  state.exportDraft = normalizeTextExportSettings({
    ...state.exportSettings,
    target,
  });
  state.pendingFocusId = "textExportEncodingSelect";
  renderWorkspace();
}

function openImportFilePicker(mode) {
  state.importFileMode = mode;
  refs.importFileInput.accept = mode === "excel"
    ? ".xlsx,.xlsm,.xls"
    : ".csv,.tsv,.txt,.json,.xml";
  refs.importFileInput.click();
}

async function beginExcelImport(file) {
  setStatus(`Reading ${file.name}...`);
  const xlsx = await ensureXlsxLibrary();
  const workbook = xlsx.read(await file.arrayBuffer(), {
    type: "array",
    cellHTML: false,
    cellStyles: false,
  });
  const sheetOptions = listExcelWorkbookSheets(workbook);
  if (!sheetOptions.length) {
    throw new Error("The selected Excel workbook does not contain any sheets.");
  }
  if (sheetOptions.length === 1) {
    const imported = excelWorkbookToMatrix(workbook, sheetOptions[0].index);
    applyImportedMatrix(imported.matrix, `Imported ${imported.sheetName} from ${file.name}.`);
    return;
  }
  state.importDraft = {
    kind: "excel",
    fileName: file.name,
    workbook,
    sheetOptions,
    selectedSheetIndex: sheetOptions[0].index,
  };
  state.openPanel = "excel-import";
  state.headerMenu = null;
  state.panelError = "";
  state.pendingFocusId = "excelImportSheetSelect";
  renderWorkspace();
  setStatus(`Loaded ${file.name}. Choose one sheet to import.`);
}

async function beginTextImport(file) {
  const fileKind = inferImportFileKind(file.name);
  if (!isTextImportKind(fileKind)) {
    throw new Error("The selected file is not a supported text or structured-data import type.");
  }
  const bytes = new Uint8Array(await file.arrayBuffer());
  const draft = {
    kind: "text",
    fileName: file.name,
    fileKind,
    bytes,
    encoding: inferDefaultTextImportEncoding(bytes),
    delimiterMode: inferDefaultDelimiterMode(fileKind),
    needsReread: false,
    error: "",
    text: "",
    matrix: [],
  };
  refreshTextImportDraft(draft);
  state.importDraft = draft;
  state.openPanel = "text-import";
  state.headerMenu = null;
  state.panelError = "";
  state.pendingFocusId = "textImportEncodingSelect";
  renderWorkspace();
  setStatus(`Loaded ${file.name}. Review the preview and import it into the current sheet.`);
}

function handleTextImportReread() {
  if (state.importDraft?.kind !== "text") return;
  refreshTextImportDraft(state.importDraft);
  renderWorkspace();
}

function handleApplyTextImport() {
  const draft = state.importDraft;
  if (!draft || draft.kind !== "text") return;
  if (draft.needsReread) {
    draft.error = "Press Re-read after changing the encoding so the preview uses the original raw bytes.";
    renderWorkspace();
    return;
  }
  if (draft.error) {
    renderWorkspace();
    return;
  }
  if (!draft.matrix.length) {
    draft.error = "There is no parsed data to import.";
    renderWorkspace();
    return;
  }
  applyImportedMatrix(draft.matrix, `Imported ${draft.fileKind.toUpperCase()} data from ${draft.fileName}.`);
}

function handleApplyExcelImport() {
  const draft = state.importDraft;
  if (!draft || draft.kind !== "excel") return;
  const imported = excelWorkbookToMatrix(draft.workbook, draft.selectedSheetIndex);
  applyImportedMatrix(imported.matrix, `Imported ${imported.sheetName} from ${draft.fileName}.`);
}

async function handleApplyTextExport() {
  const draft = state.exportDraft;
  if (!draft) return;
  commitEditing("stay");
  const settings = normalizeTextExportSettings(draft);
  const visibleRowSet = getVisibleRowInfo().visibleRowSet;
  const matrix = buildTextExportMatrix(state.table, {
    includeHiddenData: settings.includeHiddenData,
    visibleRowSet,
  });
  const text = serializeTextExportMatrix(matrix, settings);
  let legacyEncode = null;
  if (settings.encoding === "Shift-JIS" || settings.encoding === "EUC-JP") {
    legacyEncode = await ensureLegacyTextExportEncoder();
  }
  const bytes = encodeTextExportBytes(text, {
    ...settings,
    legacyEncode,
  });
  const fileName = buildTextExportFileName(state.sourceFileName || state.table.sourceName, settings.target);
  const blob = new Blob([bytes], { type: getTextExportMimeType(settings.target, settings.encoding) });
  rememberTextExportSettings(settings);
  try {
    const savedWithFilePicker = await trySaveWithFilePicker(blob, fileName, getTextExportFilePickerTypes(settings.target));
    if (savedWithFilePicker) {
      setStatus(`Exported ${settings.target.toUpperCase()} as ${fileName}.`);
    } else {
      saveWithDownloadLink(blob, fileName);
      setStatus(`Started ${settings.target.toUpperCase()} download as ${fileName}. If no download appears, this browser may block file downloads.`);
    }
    closeOpenPanel();
  } catch (error) {
    setStatus(formatExportErrorMessage(error));
  }
}

function closeOpenPanel() {
  if (!state.openPanel) return;
  if (state.openPanel === "text-import" || state.openPanel === "excel-import") {
    state.importDraft = null;
  }
  if (state.openPanel === "text-export") {
    state.exportDraft = null;
  }
  if (state.openPanel === "grid-limits") {
    state.gridLimitsDraft = null;
  }
  if (state.openPanel === "remove-duplicates") {
    state.removeDuplicatesDraft = null;
  }
  if (state.openPanel === "remove-blank-rows") {
    state.removeBlankRowsDraft = null;
  }
  if (state.openPanel === "trim-spaces") {
    state.trimSpacesDraft = null;
  }
  if (state.openPanel === "normalize-case") {
    state.normalizeCaseDraft = null;
  }
  if (state.openPanel === "number-commas") {
    state.numberCommasDraft = null;
  }
  if (state.openPanel === "find-encoding-issues") {
    state.encodingIssuesDraft = null;
  }
  state.openPanel = null;
  state.panelError = "";
  renderWorkspace();
}

function refreshTextImportDraft(draft) {
  try {
    const result = parseTextImportContent({
      fileName: draft.fileName,
      bytes: draft.bytes,
      encoding: draft.encoding,
      delimiterMode: draft.delimiterMode,
    });
    draft.text = result.text;
    draft.matrix = result.matrix;
    draft.error = "";
    draft.needsReread = false;
  } catch (error) {
    draft.text = "";
    draft.matrix = [];
    draft.error = error.message || "The selected file could not be parsed.";
    draft.needsReread = false;
  }
}

function applyImportedMatrix(matrix, prefix) {
  const safeMatrix = Array.isArray(matrix) ? matrix : [];
  if (!safeMatrix.length) {
    setStatus("The selected import did not contain any cells to insert.");
    closeOpenPanel();
    return;
  }
  commitEditing("stay");
  const startCell = getImportAnchorCell(state.selection);
  const beforeSnapshot = captureSnapshot();
  const writeResult = applyMatrixToTable(state.table, startCell, safeMatrix);
  const importedSelection = selectionFromMatrix(startCell, safeMatrix);
  clearCopiedSelection();
  closeOpenPanelWithoutRender();
  applySnapshot(
    {
      table: writeResult.table,
      activeCell: startCell,
      selection: importedSelection,
    },
    {
      beforeSnapshot,
      recordHistory: writeResult.changed,
      statusMessage: `${prefix} Inserted ${formatMatrixShape(safeMatrix)} at ${gridCellName(startCell.row, startCell.col)}.`,
    },
  );
}

function closeOpenPanelWithoutRender() {
  state.importDraft = null;
  state.exportDraft = null;
  state.gridLimitsDraft = null;
  state.removeDuplicatesDraft = null;
  state.removeBlankRowsDraft = null;
  state.trimSpacesDraft = null;
  state.normalizeCaseDraft = null;
  state.numberCommasDraft = null;
  state.encodingIssuesDraft = null;
  state.openPanel = null;
  state.panelError = "";
}

function formatMatrixShape(matrix) {
  const rowCount = Array.isArray(matrix) ? matrix.length : 0;
  const columnCount = rowCount
    ? matrix.reduce((max, row) => Math.max(max, Array.isArray(row) ? row.length : 0), 0)
    : 0;
  return `${rowCount}x${columnCount}`;
}

async function ensureXlsxLibrary() {
  if (window.XLSX) return window.XLSX;
  if (!xlsxLoadPromise) {
    xlsxLoadPromise = new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = resolveVendorScriptSource(VENDOR_SCRIPT_PATHS.xlsx);
      script.async = true;
      script.dataset.cwsXlsxLoader = "true";
      script.onload = () => {
        if (window.XLSX) {
          resolve(window.XLSX);
          return;
        }
        xlsxLoadPromise = null;
        reject(new Error("The Excel import library loaded but did not initialize correctly."));
      };
      script.onerror = () => {
        xlsxLoadPromise = null;
        reject(new Error("Failed to load the Excel import library."));
      };
      document.head.append(script);
    });
  }
  return xlsxLoadPromise;
}

async function ensureLegacyTextExportEncoder() {
  if (window.cptable?.utils?.encode) {
    return createLegacyTextExportEncoder();
  }
  if (!codepageLoadPromise) {
    codepageLoadPromise = (async () => {
      await loadScriptOnce(VENDOR_SCRIPT_PATHS.codepageBase, "cwsCodepageBase");
      await loadScriptOnce(VENDOR_SCRIPT_PATHS.codepageEucJp, "cwsCodepageEucJp");
      await loadScriptOnce(VENDOR_SCRIPT_PATHS.codepageUtils, "cwsCodepageUtils");
    })();
  }
  await codepageLoadPromise;
  return createLegacyTextExportEncoder();
}

function createLegacyTextExportEncoder() {
  return (encoding, text) => {
    const codepage = mapTextExportEncodingToCodepage(encoding);
    if (!codepage) {
      throw new Error(`No legacy encoder is configured for ${encoding}.`);
    }
    if (!window.cptable?.utils?.hascp?.(codepage)) {
      throw new Error(`${encoding} export is not available in this browser build.`);
    }
    const encoded = window.cptable.utils.encode(codepage, String(text ?? ""));
    return encoded instanceof Uint8Array ? encoded : Uint8Array.from(encoded || []);
  };
}

function mapTextExportEncodingToCodepage(encoding) {
  if (encoding === "Shift-JIS") return 932;
  if (encoding === "EUC-JP") return 51932;
  return 0;
}

async function loadScriptOnce(src, marker) {
  if (document.querySelector(`script[data-script-marker="${marker}"]`)) {
    return;
  }
  await new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = resolveVendorScriptSource(src);
    script.async = true;
    script.dataset.scriptMarker = marker;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error(`Failed to load ${src}.`));
    document.head.append(script);
  });
}

function resolveVendorScriptSource(src) {
  if (typeof STANDALONE_VENDOR_SCRIPT_PAYLOADS !== "undefined") {
    const payload = STANDALONE_VENDOR_SCRIPT_PAYLOADS[src];
    if (payload) {
      return `data:text/javascript;base64,${payload}`;
    }
  }
  return src;
}

function handleSetHeaderRow() {
  commitEditing("stay");
  const bounds = getSelectionBounds(state.selection);
  const targetRow = state.selection.mode === "row" && getSelectionRowCount() > 1
    ? bounds.startRow
    : state.activeCell.row;
  const beforeSnapshot = captureSnapshot();
  const nextTable = setHeaderRow(state.table, targetRow);
  clearCopiedSelection();
  applySnapshot(
    {
      table: nextTable,
      activeCell: { row: targetRow, col: state.activeCell.col },
      selection: createCellSelection({ row: targetRow, col: state.activeCell.col }),
    },
    {
      beforeSnapshot,
      recordHistory: hasTableChanged(beforeSnapshot.table, nextTable),
      statusMessage: `Set row ${targetRow} as the formal header row.`,
    },
  );
}

function handleInsertRows(placement) {
  if (isWholeSheetSelection()) return;
  commitEditing("stay");
  const target = resolveRowInsertTarget(placement);
  const beforeSnapshot = captureSnapshot();
  const nextTable = insertRows(state.table, target.anchorRow, target.count, placement);
  const nextView = buildSheetGridView(nextTable);
  clearCopiedSelection();
  applySnapshot(
    {
      table: nextTable,
      activeCell: clampActiveCell({ row: target.insertStartRow, col: state.activeCell.col }, nextView),
      selection: createRowSelection(target.insertStartRow, nextView.columnCount, target.insertStartRow + target.count - 1),
    },
    {
      beforeSnapshot,
      recordHistory: hasTableChanged(beforeSnapshot.table, nextTable),
      statusMessage: `Inserted ${target.count} row${target.count === 1 ? "" : "s"} ${placement} row ${target.anchorRow}.`,
    },
  );
}

function handleInsertColumns(placement) {
  if (isWholeSheetSelection()) return;
  commitEditing("stay");
  const target = resolveColumnInsertTarget(placement);
  const beforeSnapshot = captureSnapshot();
  const nextTable = insertColumns(state.table, target.anchorColumn, target.count, placement);
  const nextView = buildSheetGridView(nextTable);
  clearCopiedSelection();
  applySnapshot(
    {
      table: nextTable,
      activeCell: clampActiveCell({ row: state.activeCell.row, col: target.insertStartColumn }, nextView),
      selection: createColumnSelection(target.insertStartColumn, nextView.rowCount, target.insertStartColumn + target.count - 1),
    },
    {
      beforeSnapshot,
      recordHistory: hasTableChanged(beforeSnapshot.table, nextTable),
      statusMessage: `Inserted ${target.count} column${target.count === 1 ? "" : "s"} ${placement} column ${formatColumnName(target.anchorColumn)}.`,
    },
  );
}

function handleInsertCells(direction) {
  if (shouldDisableInsertCells()) return;
  commitEditing("stay");
  const beforeSnapshot = captureSnapshot();
  const bounds = getSelectionBounds(state.selection);
  const nextTable = insertCells(state.table, bounds, direction);
  clearCopiedSelection();
  applySnapshot(
    {
      table: nextTable,
      activeCell: { row: bounds.startRow, col: bounds.startCol },
      selection: createRangeSelection(
        { row: bounds.startRow, col: bounds.startCol },
        { row: bounds.endRow, col: bounds.endCol },
      ),
    },
    {
      beforeSnapshot,
      recordHistory: hasTableChanged(beforeSnapshot.table, nextTable),
      statusMessage: `Inserted ${getSelectionRowCount()}x${getSelectionColumnCount()} blank area and shifted cells ${direction}.`,
    },
  );
}

function handleDeleteRows() {
  if (isWholeSheetSelection()) return;
  commitEditing("stay");
  const target = resolveRowDeleteTarget();
  const beforeSnapshot = captureSnapshot();
  const nextTable = deleteRows(state.table, target.startRow, target.count);
  const nextView = buildSheetGridView(nextTable);
  const nextActiveRow = Math.min(target.startRow, nextView.rowCount);
  const nextSelectionEndRow = Math.min(nextView.rowCount, nextActiveRow + target.count - 1);
  clearCopiedSelection();
  applySnapshot(
    {
      table: nextTable,
      activeCell: clampActiveCell({ row: nextActiveRow, col: state.activeCell.col }, nextView),
      selection: createRowSelection(nextActiveRow, nextView.columnCount, nextSelectionEndRow),
    },
    {
      beforeSnapshot,
      recordHistory: hasTableChanged(beforeSnapshot.table, nextTable),
      statusMessage: `Deleted ${target.count} row${target.count === 1 ? "" : "s"} starting at row ${target.startRow}.`,
    },
  );
}

function handleDeleteColumns() {
  if (isWholeSheetSelection()) return;
  commitEditing("stay");
  const target = resolveColumnDeleteTarget();
  const beforeSnapshot = captureSnapshot();
  const nextTable = deleteColumns(state.table, target.startColumn, target.count);
  const nextView = buildSheetGridView(nextTable);
  const nextActiveColumn = Math.min(target.startColumn, nextView.columnCount);
  const nextSelectionEndColumn = Math.min(nextView.columnCount, nextActiveColumn + target.count - 1);
  clearCopiedSelection();
  applySnapshot(
    {
      table: nextTable,
      activeCell: clampActiveCell({ row: state.activeCell.row, col: nextActiveColumn }, nextView),
      selection: createColumnSelection(nextActiveColumn, nextView.rowCount, nextSelectionEndColumn),
    },
    {
      beforeSnapshot,
      recordHistory: hasTableChanged(beforeSnapshot.table, nextTable),
      statusMessage: `Deleted ${target.count} column${target.count === 1 ? "" : "s"} starting at column ${formatColumnName(target.startColumn)}.`,
    },
  );
}

function openHeaderMenu(columnIndex, button) {
  const existing = state.columnFilters[columnKeyForIndex(columnIndex)];
  const rect = button.getBoundingClientRect();
  state.headerMenu = {
    columnIndex,
    operator: existing?.operator || "contains",
    value: existing?.value || "",
    left: Math.max(8, rect.left - 8),
    top: rect.bottom + 8,
    error: "",
  };
  renderFloatingLayer();
}

function handleQuickSort(columnIndex, direction) {
  if (!columnIndex) return;
  commitEditing("stay");
  const beforeSnapshot = captureSnapshot();
  const nextTable = sortTableRows(state.table, [{ columnKey: columnKeyForIndex(columnIndex), direction }]);
  state.headerMenu = null;
  applySnapshot(
    {
      table: nextTable,
      activeCell: state.activeCell,
      selection: state.selection,
    },
    {
      beforeSnapshot,
      recordHistory: hasTableChanged(beforeSnapshot.table, nextTable),
      statusMessage: buildVisibleRowStatus(`Sorted ${state.table.columns?.[columnIndex - 1]?.label || gridCellName(getHeaderRowIndex(state.table), columnIndex)} ${direction === "asc" ? "A -> Z" : "Z -> A"}.`),
    },
  );
}

function handleApplyHeaderFilter() {
  if (!state.headerMenu) return;
  const operator = state.headerMenu.operator;
  const value = state.headerMenu.value;
  if ((operator === "contains" || operator === "equals") && value === "") {
    state.headerMenu.error = "Enter a filter value or use Is Empty / Is Not Empty.";
    renderFloatingLayer();
    return;
  }
  state.columnFilters = {
    ...state.columnFilters,
    [columnKeyForIndex(state.headerMenu.columnIndex)]: {
      operator,
      value,
    },
  };
  state.headerMenu = null;
  renderWorkspace();
  setStatus(buildVisibleRowStatus("Applied the column filter."));
}

function handleClearHeaderFilter() {
  if (!state.headerMenu) return;
  const nextFilters = { ...state.columnFilters };
  delete nextFilters[columnKeyForIndex(state.headerMenu.columnIndex)];
  state.columnFilters = nextFilters;
  state.headerMenu = null;
  renderWorkspace();
  setStatus(buildVisibleRowStatus("Cleared the column filter."));
}

function handleClearFilters() {
  state.columnFilters = {};
  state.advancedFilter = { enabled: false, conditions: [], logic: "" };
  state.headerMenu = null;
  renderWorkspace();
  setStatus(buildVisibleRowStatus("Cleared column and advanced filters."));
}

function runFindNext() {
  commitEditing("stay");
  if (!state.findReplace.findText) {
    state.panelError = "Enter text to find.";
    renderFloatingLayer();
    return;
  }
  const startCell = state.findReplace.currentMatch || state.activeCell;
  const includeCurrent = !state.findReplace.currentMatch;
  const match = findNextMatch(state.table, state.selection, startCell, {
    ...state.findReplace,
    includeCurrent,
  });
  if (!match) {
    state.findReplace.currentMatch = null;
    state.panelError = "No more matches were found in the current scope.";
    renderFloatingLayer();
    setStatus("No matches were found.");
    return;
  }
  state.findReplace.currentMatch = { row: match.row, col: match.col };
  state.panelError = "";
  state.activeCell = { row: match.row, col: match.col };
  state.selection = createCellSelection(state.activeCell);
  state.pendingScrollCell = state.activeCell;
  renderWorkspace();
  setStatus(`Found a match at ${gridCellName(match.row, match.col)}.`);
}

function runReplaceOne() {
  commitEditing("stay");
  if (!state.findReplace.findText) {
    state.panelError = "Enter text to find before replacing.";
    renderFloatingLayer();
    return;
  }
  const match = resolveReplaceTarget();
  if (!match) {
    state.panelError = "No match is available to replace.";
    renderFloatingLayer();
    setStatus("No match is available to replace.");
    return;
  }
  const beforeSnapshot = captureSnapshot();
  const currentValue = getCellValue(state.table, match.row, match.col);
  const nextValue = buildSingleReplacement(currentValue);
  const nextTable = setCellValue(state.table, match.row, match.col, nextValue);
  applySnapshot(
    {
      table: nextTable,
      activeCell: { row: match.row, col: match.col },
      selection: createCellSelection({ row: match.row, col: match.col }),
    },
    {
      beforeSnapshot,
      recordHistory: hasTableChanged(beforeSnapshot.table, nextTable),
      statusMessage: `Replaced the match at ${gridCellName(match.row, match.col)}.`,
    },
  );
  const nextMatch = findNextMatch(state.table, state.selection, match, {
    ...state.findReplace,
    includeCurrent: false,
  });
  if (nextMatch) {
    state.findReplace.currentMatch = { row: nextMatch.row, col: nextMatch.col };
    state.activeCell = { row: nextMatch.row, col: nextMatch.col };
    state.selection = createCellSelection(state.activeCell);
    state.pendingScrollCell = state.activeCell;
    renderWorkspace();
  }
  state.panelError = "";
  renderFloatingLayer();
}

function runReplaceAll() {
  commitEditing("stay");
  if (!state.findReplace.findText) {
    state.panelError = "Enter text to find before replacing.";
    renderFloatingLayer();
    return;
  }
  const beforeSnapshot = captureSnapshot();
  const result = replaceAllMatches(state.table, state.selection, state.findReplace);
  if (!result.count) {
    state.panelError = "No matches were found in the current scope.";
    renderFloatingLayer();
    setStatus("No matches were found.");
    return;
  }
  applySnapshot(
    {
      table: result.table,
      activeCell: state.activeCell,
      selection: state.selection,
    },
    {
      beforeSnapshot,
      recordHistory: hasTableChanged(beforeSnapshot.table, result.table),
      statusMessage: `Replaced ${result.count} matches in the current scope.`,
    },
  );
  state.panelError = "";
  renderFloatingLayer();
}

function resolveReplaceTarget() {
  const currentMatch = state.findReplace.currentMatch;
  if (currentMatch && currentCellMatchesFind(currentMatch.row, currentMatch.col)) {
    return currentMatch;
  }
  return findNextMatch(state.table, state.selection, state.activeCell, {
    ...state.findReplace,
    includeCurrent: true,
  });
}

function currentCellMatchesFind(row, col) {
  if (!isCellInFindReplaceScope(row, col)) return false;
  const value = getCellValue(state.table, row, col);
  return state.findReplace.wholeCell
    ? compareMaybeCaseSensitive(value, state.findReplace.findText, state.findReplace.caseSensitive)
    : includesMaybeCaseSensitive(value, state.findReplace.findText, state.findReplace.caseSensitive);
}

function buildSingleReplacement(currentValue) {
  const text = String(currentValue ?? "");
  const findText = state.findReplace.findText;
  const replaceText = state.findReplace.replaceText;
  if (state.findReplace.wholeCell) {
    return replaceText;
  }
  if (state.findReplace.caseSensitive) {
    return text.replace(findText, replaceText);
  }
  const index = text.toLocaleLowerCase().indexOf(findText.toLocaleLowerCase());
  if (index < 0) return text;
  return `${text.slice(0, index)}${replaceText}${text.slice(index + findText.length)}`;
}

function handleAddAdvancedCondition() {
  if (!state.advancedFilterDraft) return;
  if (state.advancedFilterDraft.conditions.length >= 20) return;
  state.advancedFilterDraft.conditions.push(createAdvancedCondition());
  if (!state.advancedFilterDraft.logic || state.advancedFilterDraft.logic === buildDefaultAdvancedFilterLogic(state.advancedFilterDraft.conditions.length - 1)) {
    state.advancedFilterDraft.logic = buildDefaultAdvancedFilterLogic(state.advancedFilterDraft.conditions.length);
  }
  state.panelError = "";
  renderFloatingLayer();
}

function handleRemoveAdvancedCondition(index) {
  if (!state.advancedFilterDraft) return;
  if (state.advancedFilterDraft.conditions.length <= 1) return;
  state.advancedFilterDraft.conditions.splice(index, 1);
  state.advancedFilterDraft.logic = buildDefaultAdvancedFilterLogic(state.advancedFilterDraft.conditions.length);
  state.panelError = "";
  renderFloatingLayer();
}

function handleResetAdvancedLogic() {
  if (!state.advancedFilterDraft) return;
  state.advancedFilterDraft.logic = buildDefaultAdvancedFilterLogic(state.advancedFilterDraft.conditions.length);
  state.panelError = "";
  renderFloatingLayer();
}

function handleApplyAdvancedFilter() {
  if (!state.advancedFilterDraft) return;
  const result = validateAdvancedFilterDefinition(state.advancedFilterDraft);
  if (!result.ok) {
    state.panelError = result.errors.join(" ");
    renderFloatingLayer();
    return;
  }
  state.advancedFilter = {
    enabled: result.conditions.length > 0,
    conditions: result.conditions,
    logic: result.logic,
  };
  state.panelError = "";
  state.openPanel = null;
  renderWorkspace();
  setStatus(buildVisibleRowStatus("Applied the advanced filter."));
}

function handleApplySort() {
  if (!state.sortDraft?.primaryColumn) {
    state.panelError = "Choose a primary sort column.";
    renderFloatingLayer();
    return;
  }
  const rules = [
    { columnKey: state.sortDraft.primaryColumn, direction: state.sortDraft.primaryDirection },
  ];
  if (state.sortDraft.secondaryColumn) {
    rules.push({ columnKey: state.sortDraft.secondaryColumn, direction: state.sortDraft.secondaryDirection });
  }
  const beforeSnapshot = captureSnapshot();
  const nextTable = sortTableRows(state.table, rules);
  state.openPanel = null;
  state.panelError = "";
  applySnapshot(
    {
      table: nextTable,
      activeCell: state.activeCell,
      selection: state.selection,
    },
    {
      beforeSnapshot,
      recordHistory: hasTableChanged(beforeSnapshot.table, nextTable),
      statusMessage: buildVisibleRowStatus("Applied the sort order."),
    },
  );
}

function handleApplyPageSetup(printAfter) {
  const validationError = validatePageSetupDraft(state.pageSetupDraft);
  if (validationError) {
    state.panelError = validationError;
    renderFloatingLayer();
    return;
  }
  const beforeSnapshot = captureSnapshot();
  const nextTable = applyPageSetup(state.table, state.pageSetupDraft);
  const changed = hasTableChanged(beforeSnapshot.table, nextTable);
  state.openPanel = null;
  state.panelError = "";
  applySnapshot(
    {
      table: nextTable,
      activeCell: state.activeCell,
      selection: state.selection,
    },
    {
      beforeSnapshot,
      recordHistory: changed,
      statusMessage: "Applied page setup settings.",
    },
  );
  if (printAfter) {
    startPrintFlow("Applied page setup and opened the browser print dialog.");
  }
}

function handleApplyGridLimits() {
  const validationError = validateGridLimitsDraft(state.gridLimitsDraft);
  if (validationError) {
    state.panelError = validationError;
    renderFloatingLayer();
    return;
  }
  const nextLimits = {
    maxRows: Number(state.gridLimitsDraft.maxRows),
    maxColumns: Number(state.gridLimitsDraft.maxColumns),
  };
  const beforeSnapshot = captureSnapshot();
  const nextTable = setGridLimits(state.table, nextLimits);
  state.gridLimitsDraft = null;
  state.openPanel = null;
  state.panelError = "";
  clearCopiedSelection();
  applySnapshot(
    {
      table: nextTable,
      activeCell: clampActiveCell(state.activeCell, buildSheetGridView(nextTable)),
      selection: state.selection,
    },
    {
      beforeSnapshot,
      recordHistory: hasTableChanged(beforeSnapshot.table, nextTable),
      statusMessage: `Updated the editable grid to ${nextLimits.maxRows} rows x ${nextLimits.maxColumns} columns.`,
    },
  );
}

function createAdvancedCondition() {
  return {
    columnKey: state.table.columns?.[0]?.key || columnKeyForIndex(1),
    operator: "contains",
    value: "",
  };
}

function createAdvancedFilterDraft() {
  const sourceConditions = state.advancedFilter.conditions.length
    ? cloneJsonValue(state.advancedFilter.conditions)
    : [createAdvancedCondition()];
  return {
    conditions: sourceConditions,
    logic: state.advancedFilter.logic || buildDefaultAdvancedFilterLogic(sourceConditions.length),
  };
}

function createSortDraft() {
  return {
    primaryColumn: state.table.columns?.[0]?.key || "",
    primaryDirection: "asc",
    secondaryColumn: "",
    secondaryDirection: "asc",
  };
}

function createPageSetupDraft(pageSetup) {
  const source = pageSetup || {};
  return {
    margins: {
      top: source.margins?.top ?? 12,
      right: source.margins?.right ?? 12,
      bottom: source.margins?.bottom ?? 12,
      left: source.margins?.left ?? 12,
    },
    paperSize: source.paperSize || "A4",
    orientation: source.orientation || "portrait",
    printArea: {
      mode: source.printArea?.mode || "entire-sheet",
      range: source.printArea?.range || "",
    },
    headerFooter: {
      header: source.headerFooter?.header || "",
      footer: source.headerFooter?.footer || "",
    },
    background: source.background
      ? {
        mode: source.background.mode || "solid-color",
        color: source.background.color || "#ffffff",
      }
      : null,
  };
}

function startPrintFlow(statusMessage) {
  const printJob = createPrintJob();
  if (!printJob) {
    setStatus("Nothing to print from the current sheet.");
    return;
  }
  if (shouldBlockBrowserPrint(printJob)) {
    window.alert(translateMessage("The current print range is large. Export to Excel before printing is recommended."));
    setStatus("Printing canceled because the current print range is too large for stable browser printing.");
    return;
  }
  mountPrintSurface(printJob);
  state.printing = true;
  document.body.classList.add("is-printing-light-table");
  window.requestAnimationFrame(() => {
    window.requestAnimationFrame(() => {
      window.print();
      setStatus(statusMessage);
    });
  });
}

function handleAfterPrint() {
  if (!state.printing) return;
  state.printing = false;
  document.body.classList.remove("is-printing-light-table");
  clearPrintSurface();
  renderWorkspace();
}

function createPrintJob() {
  const table = state.table;
  const pageSetup = table.pageSetup || {};
  const printBounds = resolvePrintAreaBounds(table, state.selection, pageSetup);
  const visibleRowSet = getVisibleRowInfo().visibleRowSet;
  const headerRowIndex = getHeaderRowIndex(table);
  const totalRowCount = getSheetRowCount(table);
  const totalColumnCount = Math.max(1, getColumnCount(table));
  const startCol = Math.max(1, Math.min(totalColumnCount, printBounds.startCol));
  const endCol = Math.max(startCol, Math.min(totalColumnCount, printBounds.endCol));
  const printableColumns = createIndexRange(startCol, endCol);
  const printableRows = [];

  for (let rowIndex = Math.max(1, printBounds.startRow); rowIndex <= Math.min(totalRowCount, printBounds.endRow); rowIndex += 1) {
    if (rowIndex === headerRowIndex || visibleRowSet.has(rowIndex)) {
      printableRows.push(rowIndex);
    }
  }

  if (!printableRows.length || !printableColumns.length) {
    return null;
  }

  const printableDataRowCount = printableRows.filter((rowIndex) => rowIndex !== headerRowIndex).length;
  return {
    table,
    pageSetup,
    printBounds,
    printableRows,
    printableColumns,
    headerRowIndex,
    printableDataRowCount,
    printableColumnCount: printableColumns.length,
    printableCellCount: printableDataRowCount * printableColumns.length,
    sourceLabel: state.sourceFileName || "Unsaved blank sheet",
    sheetName: table.sheetName || DEFAULT_SHEET_NAME,
  };
}

function shouldBlockBrowserPrint(printJob) {
  if (!printJob) return false;
  return printJob.printableDataRowCount > MAX_BROWSER_PRINT_ROWS
    || printJob.printableColumnCount > MAX_BROWSER_PRINT_COLUMNS
    || printJob.printableCellCount > MAX_BROWSER_PRINT_CELLS;
}

function mountPrintSurface(printJob) {
  const { surface, styleElement } = ensurePrintSurfaceElements();
  styleElement.textContent = buildPrintPageCss(printJob.pageSetup || {});
  surface.innerHTML = buildPrintSurfaceHtml(printJob);
}

function clearPrintSurface() {
  const surface = document.getElementById("printSurface");
  const styleElement = document.getElementById("printPageStyle");
  if (surface) {
    surface.innerHTML = "";
  }
  if (styleElement) {
    styleElement.textContent = "";
  }
}

function ensurePrintSurfaceElements() {
  let surface = document.getElementById("printSurface");
  if (!surface) {
    surface = document.createElement("div");
    surface.id = "printSurface";
    surface.className = "print-surface";
    document.body.append(surface);
  }
  let styleElement = document.getElementById("printPageStyle");
  if (!styleElement) {
    styleElement = document.createElement("style");
    styleElement.id = "printPageStyle";
    document.head.append(styleElement);
  }
  return { surface, styleElement };
}

function buildPrintPageCss(pageSetup = {}) {
  const paperSize = normalizePrintPaperSize(pageSetup.paperSize);
  const orientation = normalizePrintOrientation(pageSetup.orientation);
  const margins = normalizePrintMargins(pageSetup.margins);
  return `@page { size: ${paperSize} ${orientation}; margin: ${margins.top}mm ${margins.right}mm ${margins.bottom}mm ${margins.left}mm; }`;
}

function buildPrintSurfaceHtml(printJob) {
  const header = String(printJob.pageSetup?.headerFooter?.header || "").trim();
  const footer = String(printJob.pageSetup?.headerFooter?.footer || "").trim();
  const backgroundColor = printJob.pageSetup?.background?.mode === "solid-color" && printJob.pageSetup?.background?.color
    ? printJob.pageSetup.background.color
    : "#ffffff";
  return `
    <div class="print-document">
      ${header ? `<div class="print-document-header">${escapeHtml(header)}</div>` : ""}
      <section class="print-sheet" style="background:${escapeAttr(backgroundColor)}">
        <div class="print-sheet-meta">${escapeHtml(`${printJob.sourceLabel} | ${printJob.sheetName}`)}</div>
        ${buildPrintTableHtml(printJob)}
      </section>
      ${footer ? `<div class="print-document-footer">${escapeHtml(footer)}</div>` : ""}
    </div>
  `;
}

function buildPrintTableHtml(printJob) {
  const headRow = buildPrintTableHeadRow(printJob);
  const bodyRows = printJob.printableRows
    .filter((rowIndex) => rowIndex !== printJob.headerRowIndex)
    .map((rowIndex) => buildPrintTableBodyRow(printJob, rowIndex))
    .join("");
  return `
    <table class="print-table" aria-label="Printable sheet">
      <thead>${headRow}</thead>
      <tbody>${bodyRows}</tbody>
    </table>
  `;
}

function buildPrintTableHeadRow(printJob) {
  const headerCells = printJob.printableColumns
    .map((columnIndex) => {
      const value = getCellValue(printJob.table, printJob.headerRowIndex, columnIndex);
      return `<th class="print-header-cell">${escapeHtml(value)}</th>`;
    })
    .join("");
  return `<tr><th class="print-corner-cell">${printJob.headerRowIndex}</th>${headerCells}</tr>`;
}

function buildPrintTableBodyRow(printJob, rowIndex) {
  const cells = printJob.printableColumns
    .map((columnIndex) => `<td class="print-data-cell">${escapeHtml(getCellValue(printJob.table, rowIndex, columnIndex))}</td>`)
    .join("");
  return `<tr><th class="print-row-header">${rowIndex}</th>${cells}</tr>`;
}

function normalizePrintPaperSize(value) {
  const text = String(value || "A4").trim();
  if (["A4", "A3", "Letter"].includes(text)) {
    return text;
  }
  return "A4";
}

function normalizePrintOrientation(value) {
  return String(value || "portrait").toLowerCase() === "landscape" ? "landscape" : "portrait";
}

function normalizePrintMargins(margins = {}) {
  return {
    top: normalizePrintMarginValue(margins.top, 12),
    right: normalizePrintMarginValue(margins.right, 12),
    bottom: normalizePrintMarginValue(margins.bottom, 12),
    left: normalizePrintMarginValue(margins.left, 12),
  };
}

function normalizePrintMarginValue(value, fallback) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return fallback;
  }
  return Math.max(0, Math.round(numeric * 100) / 100);
}

function createGridLimitsDraft() {
  const limits = getGridLimits(state.table);
  return {
    maxRows: String(Math.max(currentView.rowCount, limits.maxRows)),
    maxColumns: String(Math.max(currentView.columnCount, limits.maxColumns)),
  };
}

function validatePageSetupDraft(draft) {
  if (draft.printArea.mode === "custom" && !parseCellRange(draft.printArea.range)) {
    return "Enter a valid custom range such as A1:D20.";
  }
  return "";
}

function validateGridLimitsDraft(draft) {
  const maxRows = parseStrictPositiveInteger(draft?.maxRows);
  if (!maxRows.ok) {
    return "Maximum Rows must be an integer greater than or equal to 1.";
  }
  const maxColumns = parseStrictPositiveInteger(draft?.maxColumns);
  if (!maxColumns.ok) {
    return "Maximum Columns must be an integer greater than or equal to 1.";
  }
  const headerRowIndex = getHeaderRowIndex(state.table);
  if (maxRows.value < headerRowIndex) {
    return `Maximum Rows cannot be smaller than the current formal header row ${headerRowIndex}.`;
  }
  const lastUsedRow = getLastUsedRowIndex(state.table);
  if (maxRows.value < lastUsedRow) {
    return `Maximum Rows cannot be smaller than the last used row ${lastUsedRow}.`;
  }
  const lastUsedColumn = getLastUsedColumnIndex(state.table);
  if (maxColumns.value < lastUsedColumn) {
    return `Maximum Columns cannot be smaller than the last used column ${formatGridLimitColumnFeedback(lastUsedColumn)}.`;
  }
  return "";
}

function formatScopeLabel() {
  if (normalizeFindReplaceScope(state.findReplace.scope) === "global") {
    return "Whole sheet";
  }
  const scope = getFindReplaceSelectionBounds();
  if (scope.startRow === scope.endRow && scope.startCol === scope.endCol) {
    return gridCellName(scope.startRow, scope.startCol);
  }
  return `${gridCellName(scope.startRow, scope.startCol)}:${gridCellName(scope.endRow, scope.endCol)}`;
}

function normalizeFindReplaceScope(scope) {
  return scope === "selection" ? "selection" : "global";
}

function getDefaultFindReplaceScope() {
  const scope = getSelectionBounds(state.selection);
  const cellCount = (scope.endRow - scope.startRow + 1) * (scope.endCol - scope.startCol + 1);
  return cellCount > 1 ? "selection" : "global";
}

function getFindReplaceSelectionBounds() {
  return state.findReplace.scopeBounds || getSelectionBounds(state.selection);
}

function captureFindReplaceScopeBounds() {
  state.findReplace.scopeBounds = getSelectionBounds(state.selection);
}

function isCellInFindReplaceScope(row, col) {
  if (normalizeFindReplaceScope(state.findReplace.scope) !== "selection") return true;
  const bounds = getFindReplaceSelectionBounds();
  return row >= bounds.startRow
    && row <= bounds.endRow
    && col >= bounds.startCol
    && col <= bounds.endCol;
}

function renderColumnOptions() {
  if (!state.table.columns.length) {
    return `<option value="c1">${localizeInlineUiText("Column A")}</option>`;
  }
  return state.table.columns
    .map((column, index) => `<option value="${escapeAttr(column.key)}" data-user-option="true">${escapeHtml(column.label || columnLabelFromIndex(index + 1))}</option>`)
    .join("");
}

function renderOperatorOptions(selected) {
  const options = [
    { value: "contains", label: "contains" },
    { value: "equals", label: "equals" },
    { value: "is-empty", label: "is empty" },
    { value: "is-not-empty", label: "is not empty" },
  ];
  return options
    .map((option) => `<option value="${option.value}"${option.value === selected ? " selected" : ""}>${localizeInlineUiText(option.label)}</option>`)
    .join("");
}

function renderFixedOptions(options, selected) {
  return options
    .map((option) => `<option value="${escapeAttr(option)}"${option === selected ? " selected" : ""}>${localizeInlineUiText(option)}</option>`)
    .join("");
}

function parseStrictPositiveInteger(value) {
  const text = String(value ?? "").trim();
  if (!/^[1-9]\d*$/.test(text)) {
    return { ok: false, value: 0 };
  }
  return {
    ok: true,
    value: Number(text),
  };
}

function formatGridLimitColumnFeedback(value) {
  const parsed = parseStrictPositiveInteger(value);
  if (!parsed.ok) return "Enter a column count of 1 or greater.";
  return `${parsed.value} = ${formatColumnName(parsed.value)}`;
}

function formatColumnName(columnIndex) {
  return gridCellName(1, columnIndex).replace(/\d+$/, "");
}

function clampNumber(value, minimum, maximum) {
  return Math.min(Math.max(Number(value) || minimum, minimum), maximum);
}

function cellFromDataset(element) {
  return {
    row: Number(element.dataset.row || 1),
    col: Number(element.dataset.col || 1),
  };
}

function sameCell(left, right) {
  return left?.row === right?.row && left?.col === right?.col;
}

function compareMaybeCaseSensitive(left, right, caseSensitive) {
  return caseSensitive
    ? String(left ?? "") === String(right ?? "")
    : String(left ?? "").toLocaleLowerCase() === String(right ?? "").toLocaleLowerCase();
}

function includesMaybeCaseSensitive(left, right, caseSensitive) {
  return caseSensitive
    ? String(left ?? "").includes(String(right ?? ""))
    : String(left ?? "").toLocaleLowerCase().includes(String(right ?? "").toLocaleLowerCase());
}

function isEditingTarget(target) {
  return Boolean(target?.closest?.("#activeCellEditor, #formulaBarInput, #globalSearchInput, #floatingLayer input, #floatingLayer textarea, #floatingLayer select"));
}

function normalizeDownloadName(value) {
  const text = String(value || "workbook.cws.html").trim() || "workbook.cws.html";
  return /\.html?$/i.test(text) ? text : `${text}.html`;
}

function buildTextExportFileName(sourceName, target) {
  const normalizedTarget = normalizeTextExportTarget(target);
  const rawName = String(sourceName || "workbook").trim() || "workbook";
  const withoutCwsSuffix = rawName.replace(/\.cws\.html?$/i, "");
  const withoutGenericSuffix = withoutCwsSuffix.replace(/\.(html|htm|csv|tsv|txt)$/i, "");
  const baseName = withoutGenericSuffix || "workbook";
  return `${baseName}.${normalizedTarget}`;
}

function getTextExportFilePickerTypes(target) {
  const normalizedTarget = normalizeTextExportTarget(target);
  const extensions = normalizedTarget === "csv"
    ? [".csv"]
    : normalizedTarget === "tsv"
      ? [".tsv"]
      : [".txt"];
  return [
    {
      description: `${normalizedTarget.toUpperCase()} text`,
      accept: {
        "text/plain": extensions,
      },
    },
  ];
}

function getTextExportMimeType(target, encoding) {
  const normalizedTarget = normalizeTextExportTarget(target);
  const normalizedEncoding = normalizeTextExportEncoding(encoding);
  const baseType = normalizedTarget === "csv" ? "text/csv" : "text/plain";
  return `${baseType};charset=${normalizedEncoding.toLowerCase()}`;
}

function cloneJsonValue(value) {
  return value ? JSON.parse(JSON.stringify(value)) : value;
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "\"": "&quot;",
    "'": "&#39;",
  })[char]);
}

function escapeAttr(value) {
  return escapeHtml(value).replace(/`/g, "&#96;");
}

function debugGridLog(label, payload) {
  if (!debugGridEvents) return;
  document.documentElement.dataset.cwsGridDebug = `${label} ${JSON.stringify(payload)}`;
  console.log(`[cws-grid] ${label}`, payload);
}

function describeEventTarget(target) {
  if (!(target instanceof Element)) {
    return { kind: typeof target };
  }
  const cell = target.closest("[data-cell='true']");
  const rowHeader = target.closest("[data-row-header='true']");
  const columnHeader = target.closest("[data-col-header='true']");
  return {
    tag: target.tagName,
    cell: cell ? { row: cell.dataset.row || "", col: cell.dataset.col || "" } : null,
    rowHeader: rowHeader ? { row: rowHeader.dataset.row || "" } : null,
    columnHeader: columnHeader ? { col: columnHeader.dataset.col || "" } : null,
  };
}
