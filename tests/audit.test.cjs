const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '../Code.gs'), 'utf8');
const sandbox = { console };
vm.createContext(sandbox);
vm.runInContext(source, sandbox);
const audit = (data, limit = 4000) => sandbox.inspectGrid_({
  sheetName: '台帳', gid: 123, mergedA1: [],
  formulas: data.values.map(row => row.map(() => '')),
  ...data,
}, limit);
const issues = (data) => audit(data).findings;
const has = (items, id, cell) => items.some(x => x.ruleId === id && x.cell === cell);

test('empty sheet: no findings', () => {
  assert.equal(issues({ values: [['']] }).length, 0);
});

test('blank header and duplicate names map to exact official rule IDs', () => {
  const result = issues({ values: [['ID', '', 'ID'], [1, '東京', 3]] });
  assert.ok(has(result, 'L1-05', 'B1'));
  assert.ok(has(result, 'L2-03', 'C1'));
});

test('merged ranges are reported', () => {
  assert.ok(has(issues({ values: [['ID', '年'], [1, 2026]], mergedA1: ['A1:B1'] }), 'L1-12', 'A1:B1'));
});

test('separator row and column are suspicious', () => {
  const result = issues({ values: [['ID', '', '金額'], ['', '', ''], ['1', '', 200]] });
  assert.ok(has(result, 'L1-03', '2:2'));
  assert.ok(has(result, 'L1-03', 'B:B'));
});

test('numeric strings and multiline values are flagged, identifiers are not', () => {
  const result = issues({ values: [['番号', '金額', '備考'], ['00123', '1,200円', 'A\nB'], ['00456', 1300, '問題なし']] });
  assert.ok(has(result, 'L2-01', 'B2'));
  assert.ok(has(result, 'L1-07', 'C2'));
  assert.ok(!has(result, 'L2-01', 'A2'));
});

test('formulas and extra whitespace are flagged', () => {
  const result = issues({
    values: [['ID', '合計'], [1, ' 山田 ']],
    formulas: [['', ''], ['', '=SUM(A2:A2)']],
  });
  assert.ok(has(result, 'L2-06', 'B2'));
  assert.ok(has(result, 'L1-06', 'B2'));
});

test('title row is flagged only as a possible layout issue', () => {
  assert.ok(has(issues({ values: [['台帳', ''], ['ID', '名前'], [1, 'A']] }), 'L3-01', 'A1:B2'));
});

test('other detail in category cell is flagged', () => {
  assert.ok(has(issues({ values: [['ID', '区分'], [1, 'その他（相談）']] }), 'L2-05', 'B2'));
});

test('findings limit is reported, never silently dropped', () => {
  const result = audit({ values: [['ID'], [' 1 '], [' 2 ']] }, 1);
  assert.equal(result.findings.length, 1);
  assert.equal(result.truncated, true);
});

test('column naming handles wide tables', () => {
  assert.equal(sandbox.columnA1_(1), 'A');
  assert.equal(sandbox.columnA1_(26), 'Z');
  assert.equal(sandbox.columnA1_(27), 'AA');
  assert.equal(sandbox.columnA1_(703), 'AAA');
});
