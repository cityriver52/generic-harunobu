/** Only this module edits report sheets; never the original data. */
function writeReport_(spreadsheet, findings, skipped, truncated, sheetsScanned, sourceSpreadsheetId) {
  let report = spreadsheet.getSheetByName(GH_CONFIG.REPORT_SHEET_NAME);
  if (!report) {
    report = spreadsheet.insertSheet(GH_CONFIG.REPORT_SHEET_NAME);
    report.addDeveloperMetadata('GENERIC_HARUNOBU_REPORT','1');
  } else if (!isGeneratedReportSheet_(report)) {
    throw new Error('同名のユーザーシートがあるため、上書きを中止しました。');
  }
  const filter = report.getFilter();
  if (filter) filter.remove();
  report.clear();
  const headers = ['検査日時', '重要度', 'ルールID', '対象シート', '問題セル', '理由', '修正案', '対象セルへ'];
  report.getRange(1, 1, 1, headers.length).setValues([headers]);
  const checkedAt = new Date();
  const rows = findings.map(function (item) {
    const url = spreadsheetCellUrl_(sourceSpreadsheetId || spreadsheet.getId(), item.gid, item.cell);
    return [checkedAt, item.severity, item.ruleId, safeReportText_(item.sheetName), item.cell,
      safeReportText_(item.reason), safeReportText_(item.suggestion), '=HYPERLINK("' + url + '","開く")'];
  });
  if (skipped.length) rows.push([checkedAt, '未検査', 'SYSTEM-SKIP', '', '',
    '未検査のシート：' + skipped.join('、'), '範囲を分割するか、上限値を見直してください。', '']);
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
  console.log('Checked ' + sheetsScanned + ' sheet(s); findings=' + findings.length + '; skipped=' + skipped.length + '; truncated=' + truncated);
  return report;
}

/**
 * Create a new spreadsheet in the executing user's Google Drive.
 * Only findings are written; the inspected spreadsheet is not changed.
 */
function writeStandaloneReport_(sourceSpreadsheet, audit) {
  const date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const title = '機械可読性_診断結果_' + String(sourceSpreadsheet.getName()).slice(0, 65) + '_' + date;
  const destination = SpreadsheetApp.create(title);
  const initialSheet = destination.getSheets()[0];
  initialSheet.setName(GH_CONFIG.REPORT_SHEET_NAME);
  initialSheet.addDeveloperMetadata('GENERIC_HARUNOBU_REPORT', '1');
  const resultSheet = writeReport_(
    destination, audit.findings, audit.skipped, audit.truncated, audit.scanned, sourceSpreadsheet.getId()
  );
  return {spreadsheet: destination, sheet: resultSheet};
}
