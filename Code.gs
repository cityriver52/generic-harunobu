/**
 * generic-harunobu: Google Sheets machine-readability checker (MVP)
 * Inspired by the Digital Agency's published machine-readability rules.
 * This is an independent implementation, NOT the official Harunobu engine.
 * @OnlyCurrentDoc
 */

const GH_CONFIG = Object.freeze({
  REPORT_SHEET_NAME: '機械可読性_検査結果',
  HEADER_ROW: 1,
  MAX_CELLS_PER_SHEET: 100000,
  MAX_FINDINGS: 4000,
});

/** Adds a menu when the spreadsheet opens. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('機械可読性チェック')
    .addItem('現在のシートを検査', 'checkActiveSheet')
    .addItem('全シートを検査', 'checkAllSheets')
    .addSeparator()
    .addItem('使い方・注意事項', 'showHarunobuHelp')
    .addToUi();
}

/** Checks the currently selected data sheet. */
function checkActiveSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  if (activeSheet.getName() === GH_CONFIG.REPORT_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('検査結果シートは対象外です。データのシートを開いて実行してください。');
    return;
  }
  executeChecks_(spreadsheet, [activeSheet]);
}

/** Checks all sheets except the generated report. */
function checkAllSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const candidates = spreadsheet.getSheets().filter(function (sheet) {
    return sheet.getName() !== GH_CONFIG.REPORT_SHEET_NAME;
  });
  executeChecks_(spreadsheet, candidates);
}

/** Main Apps Script entry point. Original cells are only read. */
function executeChecks_(spreadsheet, sheets) {
  if (sheets.length === 0) {
    SpreadsheetApp.getUi().alert('検査対象のシートがありません。');
    return;
  }
  const findings = [];
  const skipped = [];
  let truncated = false;
  sheets.forEach(function (sheet) {
    const range = sheet.getDataRange();
    const rowCount = range.getNumRows();
    const colCount = range.getNumColumns();
    if (rowCount * colCount > GH_CONFIG.MAX_CELLS_PER_SHEET) {
      skipped.push(sheet.getName() + '（' + rowCount + '行×' + colCount + '列）');
      return;
    }
    const input = {
      sheetName: sheet.getName(),
      values: range.getValues(),
      formulas: range.getFormulas(),
      mergedA1: range.getMergedRanges().map(function (merged) {
        return merged.getA1Notation();
      }),
      gid: sheet.getSheetId(),
    };
    const result = inspectGrid_(input, GH_CONFIG.MAX_FINDINGS - findings.length);
    Array.prototype.push.apply(findings, result.findings);
    truncated = truncated || result.truncated;
  });
  writeReport_(spreadsheet, findings, skipped, truncated, sheets.length);
  spreadsheet.toast('検出 ' + findings.length + '件。結果シートに出力しました。', '機械可読性チェック', 7);
}

/**
 * Pure inspection function. Does not invoke Google Apps Script services.
 * @param {{sheetName:string,values:*,formulas:*,mergedA1:string[],gid:number}} input
 * @param {number} limit Maximum findings to retain
 * @return {{findings:Array, truncated:boolean}}
 */
function inspectGrid_(input, limit) {
  const values = input.values || [];
  const formulas = input.formulas || [];
  const height = values.length;
  const width = height ? values[0].length : 0;
  const findings = [];
  let truncated = false;
  const max = Math.max(0, limit);
  function add(id, severity, a1, reason, fix) {
    if (findings.length >= max) {
      truncated = true;
      return;
    }
    findings.push({
      ruleId: id,
      severity: severity,
      sheetName: input.sheetName,
      cell: a1,
      reason: reason,
      suggestion: fix,
      gid: input.gid,
    });
  }
  if (!height || !width) return { findings: findings, truncated: false };

  const occupiedRows = Array(height).fill(0);
  const occupiedCols = Array(width).fill(0);
  const colHasDataBelowHeader = Array(width).fill(false);
  for (let r = 0; r < height; r++) {
    for (let c = 0; c < width; c++) {
      if (hasContent_(values[r][c]) || !!getMatrixCell_(formulas, r, c)) {
        occupiedRows[r]++;
        occupiedCols[c]++;
        if (r > 0) colHasDataBelowHeader[c] = true;
      }
    }
  }
  if (occupiedRows.every(function (count) { return count === 0; })) {
    return { findings: [], truncated: false };
  }

  // L1-12: merged cells in the used range.
  (input.mergedA1 || []).forEach(function (a1) {
    add('L1-12', '要修正', a1, 'データ範囲に結合セルがあります。', '結合を解除し、各セルを独立した値として扱ってください。');
  });

  // L3-01: clearly blank first row or apparent title row.
  if (height > 1 && occupiedRows[0] === 0 && occupiedRows.slice(1).some(function (n) { return n > 0; })) {
    add('L3-01', '要確認', 'A1', '先頭行が空白です。', '公開・連携用データでは1行目から列見出しを配置してください。');
  } else if (height > 1 && occupiedRows[0] === 1 && occupiedRows[1] >= 2) {
    add('L3-01', '要確認', 'A1:' + columnA1_(width) + '2', '1行目が表題、2行目が列見出しである可能性があります。', '表題は別シートなどで管理し、データは1行目を列見出しにしてください。');
  }

  // L1-05 / L2-03: assumes a single-row header at row 1.
  const headers = new Map();
  for (let c = 0; c < width; c++) {
    if (!occupiedCols[c]) continue;
    const value = values[0][c];
    const header = String(value == null ? '' : value).trim();
    const a1 = columnA1_(c + 1) + '1';
    if (!header) {
      if (colHasDataBelowHeader[c]) {
        add('L1-05', '要修正', a1, 'データがある列に見出しがありません（1行目を見出しと仮定）。', 'その列の意味がわかる固有の列見出しを1行目に入力してください。');
      }
    } else {
      if (headers.has(header)) {
        add('L2-03', '要修正', a1, '列見出し「' + clipText_(header) + '」が ' + headers.get(header) + ' と重複しています。', '列ごとに一意で完結した見出しへ変更してください。');
      } else {
        headers.set(header, a1);
      }
    }
  }

  // L1-03: entire empty lines/columns between non-empty lines/columns.
  // Precompute bounds to avoid repeated O(n) slice scans on large sheets.
  const firstRow = occupiedRows.findIndex(Boolean);
  const lastRow = occupiedRows.length - 1 - occupiedRows.slice().reverse().findIndex(Boolean);
  const firstCol = occupiedCols.findIndex(Boolean);
  const lastCol = occupiedCols.length - 1 - occupiedCols.slice().reverse().findIndex(Boolean);
  for (let r = Math.max(1, firstRow + 1); r < lastRow; r++) {
    if (occupiedRows[r] === 0) {
      add('L1-03', '要確認', String(r + 1) + ':' + String(r + 1), 'データの途中に空白行があります。', '装飾・区切り用の行であれば削除し、1つの連続した表にしてください。');
    }
  }
  for (let c = Math.max(1, firstCol + 1); c < lastCol; c++) {
    if (occupiedCols[c] === 0) {
      const letter = columnA1_(c + 1);
      add('L1-03', '要確認', letter + ':' + letter, 'データの途中に空白列があります。', '装飾・区切り用の列であれば削除し、表を分断しないでください。');
    }
  }

  for (let r = 0; r < height; r++) {
    for (let c = 0; c < width; c++) {
      const a1 = columnA1_(c + 1) + String(r + 1);
      const formula = getMatrixCell_(formulas, r, c);
      const value = values[r][c];
      if (formula) {
        add('L2-06', '要確認', a1, '数式が残っています。', '外部連携・公開用の確定データでは、必要に応じて計算結果を値に置き換えてください。業務用の数式は自動削除しません。');
      }
      if (typeof value !== 'string' || value === '') continue;
      if (/^[ \t\u3000]|[ \t\u3000]$/.test(value)) {
        add('L1-06', '要確認', a1, 'セルの先頭または末尾に空白文字があります。', '意図しない余分な空白・タブを取り除いてください。');
      }
      if (/[\r\n]/.test(value)) {
        add('L1-07', '要確認', a1, 'セル内に改行があり、複数データが含まれる可能性があります。', '一つのセルに複数の独立値を入れていないか確認し、必要なら列や行に分けてください。');
      }
      if (r > 0 && looksLikeNumericText_(value)) {
        add('L2-01', '要確認', a1, '数値に見える値「' + clipText_(value) + '」が文字列で保存されています。', '数値を数値型で保存し、単位や注釈は別列または表示形式に分離してください。コード・識別子なら修正不要です。');
      }
      if (r > 0 && /^その他[（(][^\r\n）)]{1,}[）)]$/.test(value)) {
        add('L2-05', '要確認', a1, '「その他」と具体的な自由記述が同じセルに入っています。', 'カテゴリ列には「その他」のみを入れ、内容は別の詳細列へ分離することを検討してください。');
      }
    }
  }
  return { findings: findings, truncated: truncated };
}

function hasContent_(value) {
  return value !== null && value !== undefined && value !== '';
}

function getMatrixCell_(matrix, row, col) {
  return matrix[row] && matrix[row][col] || '';
}

function columnA1_(oneBasedColumn) {
  let n = oneBasedColumn;
  let text = '';
  while (n > 0) {
    n--;
    text = String.fromCharCode(65 + n % 26) + text;
    n = Math.floor(n / 26);
  }
  return text;
}

function clipText_(text) {
  const normalized = String(text).replace(/[\r\n\t]/g, ' ');
  return normalized.length > 55 ? normalized.slice(0, 55) + '…' : normalized;
}

/** Strictly match grouped numbers or number + units; do not flag IDs like 00123. */
function looksLikeNumericText_(text) {
  const value = text.trim();
  return /^[-+]?\d{1,3}(?:,\d{3})+(?:\.\d+)?$/.test(value) ||
    /^[-+]?(?:\d+(?:\.\d+)?|\d{1,3}(?:,\d{3})+(?:\.\d+)?)\s*(?:円|人|件|個|匹|冊|%|％|kg|㎏)$/.test(value);
}

/** Writes a fresh report. Only the report sheet is changed. */
function writeReport_(spreadsheet, findings, skipped, truncated, sheetsScanned) {
  let report = spreadsheet.getSheetByName(GH_CONFIG.REPORT_SHEET_NAME);
  if (!report) report = spreadsheet.insertSheet(GH_CONFIG.REPORT_SHEET_NAME);
  const filter = report.getFilter();
  if (filter) filter.remove();
  report.clear();
  const headers = ['検査日時', '重要度', 'ルールID', '対象シート', '問題セル', '理由', '修正案', '対象セルへ'];
  report.getRange(1, 1, 1, headers.length).setValues([headers]);
  const checkedAt = new Date();
  const rows = findings.map(function (item) {
    const url = 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(spreadsheet.getId()) + '/edit#gid=' + item.gid + '&range=' + encodeURIComponent(item.cell);
    return [checkedAt, item.severity, item.ruleId, item.sheetName, item.cell,
      item.reason, item.suggestion, '=HYPERLINK("' + url + '","開く")'];
  });
  if (skipped.length) rows.push([checkedAt, '未検査', 'SYSTEM-SKIP', '', '',
    'サイズ上限を超えたため検査しませんでした：' + skipped.join('、'), '範囲を分割するか、上限値を見直してください。', '']);
  if (truncated) rows.push([checkedAt, '未検査', 'SYSTEM-LIMIT', '', '',
    '検出件数が上限に達したため、結果が途中までしか記録されていません。', '対象を小さく分割して再検査してください。', '']);
  if (rows.length === 0) rows.push([checkedAt, '情報', '', '', '', '対応ルールの範囲で問題は検出されませんでした。', '公式Harunobuでの合格・満点を保証するものではありません。', '']);
  if (report.getMaxRows() < rows.length + 1) report.insertRowsAfter(report.getMaxRows(), rows.length + 1 - report.getMaxRows());
  if (report.getMaxColumns() < headers.length) report.insertColumnsAfter(report.getMaxColumns(), headers.length - report.getMaxColumns());
  report.getRange(2, 1, rows.length, headers.length).setValues(rows);
  report.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  report.getRange(1, 1, 1, headers.length).setBackground('#263c59').setFontColor('#ffffff').setFontWeight('bold');
  report.setFrozenRows(1);
  [160, 90, 110, 180, 110, 450, 490, 110].forEach(function (w, index) { report.setColumnWidth(index + 1, w); });
  report.getRange(2, 1, rows.length, headers.length).setVerticalAlignment('top').setWrap(true);
  report.getRange(1, 1, rows.length + 1, headers.length).createFilter();
  spreadsheet.setActiveSheet(report);
  report.setActiveSelection('A1');
  console.log('Checked ' + sheetsScanned + ' sheet(s); findings=' + findings.length + '; skipped=' + skipped.length + '; truncated=' + truncated);
}

function showHarunobuHelp() {
  SpreadsheetApp.getUi().alert(
    '機械可読性チェッカー（MVP）',
    '1行目を列見出しとするデータ表を想定しています。\n' +
    '「現在のシートを検査」または「全シートを検査」を選ぶと、別シートに結果が出ます。\n\n' +
    '原本データは修正しません。検査結果シートは実行のたびに上書きします。\n' +
    '「要確認」は誤検出を含みます。正式なHarunobuの採点と同じではありません。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}