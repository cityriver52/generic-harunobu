/** Google Sheets data collection, independent per-sheet error handling. */
function isGeneratedReportSheet_(sheet) {
  if (sheet.getName() !== GH_CONFIG.REPORT_SHEET_NAME) return false;
  try {
    if (sheet.getDeveloperMetadata().some(m => m.getKey() === 'GENERIC_HARUNOBU_REPORT')) return true;
  } catch (error) { console.warn('Report metadata unavailable: ' + error); }
  try {
    if (sheet.getMaxColumns() < 3) return false;
    const header = sheet.getRange(1,1,1,3).getDisplayValues()[0];
    return header[0] === '検査日時' && header[1] === '重要度' && header[2] === 'ルールID';
  } catch (error) { return false; }
}

function selectTargetSheets_(spreadsheet) {
  return spreadsheet.getSheets().filter(sheet => !isGeneratedReportSheet_(sheet));
}

function collectChecks_(sheets) {
  const result = {findings:[], skipped:[], coverageNotes:[], truncated:false, scanned:0};
  for (const sheet of sheets) {
    const name=sheet.getName();
    try {
      const range=sheet.getDataRange(), rows=range.getNumRows(), cols=range.getNumColumns();
      if (rows*cols > GH_CONFIG.MAX_CELLS_PER_SHEET) {
        result.skipped.push(name+'（'+rows+'行×'+cols+'列：サイズ上限超過）');
        continue;
      }
      const metadata=sheetInspectionMetadata_(sheet,rows,cols);
      if (metadata.limitations.length) result.coverageNotes.push(name+'：'+metadata.limitations.join('、'));
      const scan=inspectGrid_({
        sheetName:name, gid:sheet.getSheetId(), values:range.getValues(), formulas:range.getFormulas(),
        mergedA1:range.getMergedRanges().map(r=>r.getA1Notation()), ...metadata
      }, Math.max(0,GH_CONFIG.MAX_FINDINGS-result.findings.length));
      result.findings.push(...scan.findings);
      result.truncated=result.truncated || scan.truncated;
      result.scanned++;
    } catch (error) {
      console.error('Sheet inspection failed: '+name+' '+error);
      result.skipped.push(name+'（読み取りエラー）');
    }
  }
  return result;
}

function sheetInspectionMetadata_(sheet, rows, cols) {
  const info = { hiddenRows: [], hiddenColumns: [], objects: [], limitations: [] };
  const maxRows = Math.min(rows, 120);
  const maxCols = Math.min(cols, 50);
  try {
    if (typeof sheet.isRowHiddenByUser === 'function') {
      for (let r = 1; r <= maxRows; r++) {
        if (sheet.isRowHiddenByUser(r)) info.hiddenRows.push(r);
      }
      if (rows > maxRows) info.limitations.push('L1-13：非表示行は先頭' + maxRows + '行のみ確認');
    } else info.limitations.push('L1-13：非表示行は取得不可');
  } catch (e) { info.limitations.push('L1-13：非表示行の取得失敗'); }
  try {
    if (typeof sheet.isColumnHiddenByUser === 'function') {
      for (let c = 1; c <= maxCols; c++) {
        if (sheet.isColumnHiddenByUser(c)) info.hiddenColumns.push(c);
      }
      if (cols > maxCols) info.limitations.push('L1-13：非表示列は先頭' + maxCols + '列のみ確認');
    } else info.limitations.push('L1-13：非表示列は取得不可');
  } catch (e) { info.limitations.push('L1-13：非表示列の取得失敗'); }

  const sources = [
    { method: 'getImages', type: '画像' },
    { method: 'getDrawings', type: '図形' },
    { method: 'getCharts', type: 'グラフ' }
  ];
  sources.forEach(function(source) {
    if (typeof sheet[source.method] !== 'function') {
      info.limitations.push('L1-10：' + source.type + 'の取得不可');
      return;
    }
    try {
      const objects = sheet[source.method]() || [];
      objects.slice(0, 100).forEach(function(object) {
        let cell = 'A1';
        try {
          if (typeof object.getAnchorCell === 'function' && object.getAnchorCell())
            cell = object.getAnchorCell().getA1Notation();
          else if (typeof object.getContainerInfo === 'function') {
            const anchor = object.getContainerInfo();
            if (anchor && anchor.getAnchorRow && anchor.getAnchorColumn) {
              cell = columnA1_(anchor.getAnchorColumn()) + anchor.getAnchorRow();
            }
          }
        } catch (e) { /* keep safe fallback */ }
        info.objects.push({ type: source.type, cell: cell });
      });
      if (objects.length > 100) info.limitations.push('L1-10：' + source.type + 'は100件まで確認');
    } catch (e) { info.limitations.push('L1-10：' + source.type + 'の取得失敗'); }
  });
  return info;
}
