const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '../Code.gs'), 'utf8');
const sandbox = { console };
vm.createContext(sandbox);
vm.runInContext(source, sandbox);
vm.runInContext(fs.readFileSync(path.join(__dirname, '../ExtendedChecks.gs'), 'utf8'), sandbox);
const ID = '1ABcdEfGhIjKlMnOpQrStUvWxYz_98765';

test('extracts ID from a standard Google spreadsheet URL', () => {
  assert.equal(sandbox.parseSpreadsheetId_(
    'https://docs.google.com/spreadsheets/d/' + ID + '/edit?usp=sharing#gid=0'
  ), ID);
});
test('rejects untrusted URLs', () => {
  for (const value of [
    'https://docs.google.com.evil.com/spreadsheets/d/' + ID + '/edit',
    'https://example.com/spreadsheets/d/' + ID + '/edit',
    'https://docs.google.com/document/d/' + ID + '/edit',
    'javascript:alert(1)', 'not a url', null
  ]) {
    assert.throws(() => sandbox.parseSpreadsheetId_(value));
  }
});
test('constructs a sheet-and-range link', () => {
  assert.equal(
    sandbox.spreadsheetCellUrl_(ID, 35, 'B12:C15'),
    'https://docs.google.com/spreadsheets/d/' + ID + '/edit#gid=35&range=B12%3AC15'
  );
});
test('returns analysis even if a view-only sheet refuses report writes', () => {
  const range = {
    getNumRows: () => 2, getNumColumns: () => 2,
    getValues: () => [['ID', '名前'], [1, ' A ']],
    getFormulas: () => [['', ''], ['', '']],
    getMergedRanges: () => []
  };
  const sheet = {
    getName: () => 'データ', getSheetId: () => 11, getDataRange: () => range
  };
  sandbox.SpreadsheetApp = { openById: (id) => {
    assert.equal(id, ID);
    return {
      getName: () => '台帳', getSheets: () => [sheet], getId: () => ID,
      getSheetByName: () => null,
      insertSheet: () => { throw new Error('No edit permission'); }
    };
  }};
  sandbox.LockService = { getScriptLock: () => ({
    tryLock: () => true, releaseLock: () => {}
  }) };
  const output = sandbox.checkSpreadsheetByUrl(
    'https://docs.google.com/spreadsheets/d/' + ID + '/edit'
  );
  assert.equal(output.count, 1);
  assert.equal(output.findings[0].ruleId, 'L1-06');
  assert.equal(output.report.saved, false);
});
test('uses the broad spreadsheets OAuth scope', () => {
  const manifest = JSON.parse(fs.readFileSync(path.join(__dirname, '../appsscript.json'), 'utf8'));
  assert.deepEqual(manifest.oauthScopes, ['https://www.googleapis.com/auth/spreadsheets']);
});
test('includes the browser-side asynchronous request', () => {
  const html = fs.readFileSync(path.join(__dirname, '../Index.html'), 'utf8');
  assert.match(html, /google\.script\.run/);
  assert.match(html, /checkSpreadsheetByUrl\(url\)/);
});
