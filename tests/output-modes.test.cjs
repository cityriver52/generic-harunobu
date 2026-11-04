const test = require('node:test');
const assert = require('node:assert/strict');
const {load} = require('./helpers.cjs');

const ID = '1ABcdEfGhIjKlMnOpQrStUvWxYz_98765';
const NEW = '1AbcDeFgHiJkLmNoPqRsTuVwXyZ987654';
const url = 'https://docs.google.com/spreadsheets/d/' + ID + '/edit';

function setup() {
  const range = {
    getNumRows: () => 2,
    getNumColumns: () => 2,
    getValues: () => [['ID', '名前'], [1, ' 山田 ']],
    getFormulas: () => [['', ''], ['', '']],
    getMergedRanges: () => []
  };
  const sheet = {getName: () => '台帳', getSheetId: () => 1, getDataRange: () => range};
  const source = {getName: () => '診断元', getId: () => ID, getSheets: () => [sheet]};
  const resultSheet = {
    getSheetId: () => 99,
    setName: () => {},
    addDeveloperMetadata: () => {}
  };
  const newBook = {getId: () => NEW, getSheets: () => [resultSheet]};
  let created = 0, writerArgs = null;
  const sandbox = load({
    SpreadsheetApp: {
      openById: id => {
        assert.equal(id, ID);
        return source;
      },
      create: name => {
        assert.match(name, /機械可読性_診断結果/);
        created++;
        return newBook;
      }
    },
    Utilities: {formatDate: () => '20261104_123000'},
    LockService: {getScriptLock: () => ({
      tryLock: () => true, releaseLock: () => {}
    })}
  });
  sandbox.writeReport_ = (...args) => {
    writerArgs = args;
    return resultSheet;
  };
  return {sandbox, created: () => created, writerArgs: () => writerArgs, source};
}

test('new output creates separate spreadsheet, linking to original cells', () => {
  const x = setup();
  const result = x.sandbox.checkSpreadsheetByUrl(url, 'new');
  assert.equal(result.report.saved, true);
  assert.equal(result.report.mode, 'new');
  assert.ok(result.report.url.includes(NEW));
  assert.ok(result.findings[0].url.includes(ID));
  assert.equal(x.created(), 1);
  assert.equal(x.writerArgs()[0].getId(), NEW);
  assert.equal(x.writerArgs()[5], ID);
});

test('source output saves in original and creates no new file', () => {
  const x = setup();
  const result = x.sandbox.checkSpreadsheetByUrl(url, 'source');
  assert.equal(result.report.saved, true);
  assert.equal(result.report.mode, 'source');
  assert.ok(result.report.url.includes(ID));
  assert.equal(x.created(), 0);
  assert.equal(x.writerArgs()[0], x.source);
});

test('omitted mode keeps legacy source behavior', () => {
  const x = setup();
  assert.equal(x.sandbox.checkSpreadsheetByUrl(url).report.mode, 'source');
});

test('invalid destination mode is rejected before reading', () => {
  const x = setup();
  assert.throws(() => x.sandbox.checkSpreadsheetByUrl(url, 'random'));
  assert.equal(x.created(), 0);
});

test('new-file creation failure still returns preview results', () => {
  const x = setup();
  x.sandbox.SpreadsheetApp.create = () => {throw Error('quota exceeded');};
  const result = x.sandbox.checkSpreadsheetByUrl(url, 'new');
  assert.equal(result.report.saved, false);
  assert.equal(result.count, 1);
  assert.match(result.report.message, /作成できません/);
});
