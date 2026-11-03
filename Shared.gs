/** Pure URL and cell helpers (no Google API calls). */
function parseSpreadsheetId_(url) {
  if (typeof url !== 'string' || url.length > 2048) {
    throw new Error('GoogleスプレッドシートのURLを入力してください。');
  }
  const match = /^https:\/\/docs\.google\.com\/spreadsheets\/d\/([A-Za-z0-9_-]{20,})(?:\/[^\s?#]*)?(?:\?[^\s#]*)?(?:#[^\s]*)?$/i.exec(url.trim());
  if (!match) {
    throw new Error('URLの形式が正しくありません。https://docs.google.com/spreadsheets/d/… のURLを指定してください。');
  }
  return match[1];
}

function spreadsheetCellUrl_(id, gid, a1) {
  return 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(id) +
    '/edit#gid=' + encodeURIComponent(gid) + '&range=' + encodeURIComponent(a1);
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

function looksLikeNumericText_(text) {
  const value = text.trim();
  return /^[-+]?\d{1,3}(?:,\d{3})+(?:\.\d+)?$/.test(value) ||
    /^[-+]?(?:\d+(?:\.\d+)?|\d{1,3}(?:,\d{3})+(?:\.\d+)?)\s*(?:円|人|件|個|匹|冊|%|％|kg|㎏)$/.test(value);
}

/** Escape user-controlled text to prevent formula evaluation in result sheets. */
function safeReportText_(value) { return typeof value === 'string' && value.startsWith('=') ? "'" + value : value; }
