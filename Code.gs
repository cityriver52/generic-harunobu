/** Public GAS entrypoints for the web app and optional bound-sheet menu. */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('generic-harunobu | 機械可読性チェッカー');
}

/** Check a Sheets URL and save to source or new spreadsheet. */
function checkSpreadsheetByUrl(url, outputMode) {
  const mode = outputMode === undefined ? 'source' : outputMode;
  if (mode !== 'source' && mode !== 'new') {
    throw new Error('結果の保存先が正しくありません。画面を再読み込みしてください。');
  }
  const id = parseSpreadsheetId_(url);
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(id);
  } catch (error) {
    throw new Error('ファイルを開けませんでした。URL、共有権限、Googleアカウントの認証を確認してください。');
  }
  const candidates = selectTargetSheets_(spreadsheet);
  const audit = collectChecks_(candidates);
  const report = { saved: false, url: '', message: '', mode: mode };
  if (!candidates.length) {
    report.message = '診断対象のシートがありません。';
  } else if (mode === 'new') {
    try {
      const created = writeStandaloneReport_(spreadsheet, audit);
      report.saved = true;
      report.url = spreadsheetCellUrl_(created.spreadsheet.getId(), created.sheet.getSheetId(), 'A1');
      report.message = '診断結果を新しいスプレッドシートに保存しました。';
    } catch (error) {
      console.warn('Standalone report failed: ' + error);
      report.message = '診断は完了しましたが、新しい結果ファイルを作成できませんでした。作成権限・保存容量をご確認ください。';
    }
  } else {
    const lock = LockService.getScriptLock();
    if (lock.tryLock(10000)) {
      try {
        const output = writeReport_(spreadsheet, audit.findings, audit.skipped, audit.truncated, audit.scanned);
        report.saved = true;
        report.url = spreadsheetCellUrl_(id, output.getSheetId(), 'A1');
        report.message = '診断元ファイルに結果シートを保存しました。';
      } catch (error) {
        console.warn('Report could not be saved: ' + error);
        report.message = '診断は完了しましたが、診断元ファイルに保存できませんでした。編集権限または同名シートを確認してください。';
      } finally {
        lock.releaseLock();
      }
    } else {
      report.message = '診断は完了しましたが、他の実行と競合したため結果シートの保存を見送りました。';
    }
  }
  return {
    spreadsheetName: spreadsheet.getName(),
    spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + id + '/edit',
    spreadsheetId: id,
    totalSheets: candidates.length,
    scannedSheets: audit.scanned,
    skipped: audit.skipped,
    coverageNotes: audit.coverageNotes,
    truncated: audit.truncated,
    count: audit.findings.length,
    counts: {
      fix: audit.findings.filter(function(x) { return x.severity === '要修正'; }).length,
      review: audit.findings.filter(function(x) { return x.severity === '要確認'; }).length
    },
    findings: audit.findings.slice(0, GH_CONFIG.WEB_PREVIEW_LIMIT).map(function(item) {
      return {
        ruleId: item.ruleId, severity: item.severity, sheetName: item.sheetName,
        cell: item.cell, reason: item.reason, suggestion: item.suggestion,
        url: spreadsheetCellUrl_(id, item.gid, item.cell)
      };
    }),
    previewLimit: GH_CONFIG.WEB_PREVIEW_LIMIT,
    report: report
  };
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('機械可読性チェック')
    .addItem('現在のシートを検査', 'checkActiveSheet')
    .addItem('全シートを検査', 'checkAllSheets')
    .addSeparator()
    .addItem('使い方・注意事項', 'showHarunobuHelp')
    .addToUi();
}

function checkActiveSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  if (isGeneratedReportSheet_(activeSheet)) {
    SpreadsheetApp.getUi().alert('検査結果シートは対象外です。データのシートを開いて実行してください。');
    return;
  }
  executeChecks_(spreadsheet, [activeSheet]);
}

function checkAllSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const candidates = selectTargetSheets_(spreadsheet);
  executeChecks_(spreadsheet, candidates);
}

function executeChecks_(spreadsheet, sheets) {
  if (sheets.length === 0) {
    SpreadsheetApp.getUi().alert('検査対象のシートがありません。');
    return;
  }
  const audit = collectChecks_(sheets);
  const report = writeReport_(spreadsheet, audit.findings, audit.skipped, audit.truncated, audit.scanned);
  spreadsheet.setActiveSheet(report);
  report.setActiveSelection('A1');
  spreadsheet.toast('検出 ' + audit.findings.length + '件。結果シートに出力しました。', '機械可読性チェック', 7);
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
