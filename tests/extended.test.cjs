const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const sandbox = { console };
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(__dirname, '../Code.gs'), 'utf8'), sandbox);
vm.runInContext(fs.readFileSync(path.join(__dirname, '../ExtendedChecks.gs'), 'utf8'), sandbox);
function findings(values, extra = {}) {
  const result = sandbox.inspectGrid_({
    sheetName: 'テスト', gid: 1, values,
    formulas: values.map(row => row.map(() => '')),
    mergedA1: [], ...extra
  }, 500);
  return result.findings;
}
function has(values, id, extra) {
  return findings(values, extra).some(f => f.ruleId === id);
}
test('L1-02 multiple tables: repeated header', () => {
  assert.equal(has([['ID','値'],[1,2],['ID','値'],[2,4]], 'L1-02'), true);
});
test('L1-04 isolated footnote', () => {
  assert.equal(has([['ID','値'],[1,2],['注：速報値','']], 'L1-04'), true);
});
test('L1-08 compatibility or device-dependent glyph', () => {
  assert.equal(has([['ID'],['①']], 'L1-08'), true);
});
test('L1-09 zero/blank combination: candidate only', () => {
  const got = findings([['ID','人数'],[1,0],[2,''],[3,4]]);
  assert.equal(got.some(f => f.ruleId === 'L1-09' && f.severity === '要確認'), true);
});
test('L1-10 objects from metadata', () => {
  assert.equal(has([['ID'],[1]], 'L1-10', {objects:[{type:'画像',cell:'C4'}]}), true);
});
test('L1-13 hidden row and column from metadata', () => {
  const out = findings([['ID','金額'],[1,200]], {hiddenRows:[2],hiddenColumns:[1]});
  assert.ok(out.some(f => f.ruleId === 'L1-13' && f.cell === '2:2'));
  assert.ok(out.some(f => f.ruleId === 'L1-13' && f.cell === 'A:A'));
});
test('L2-02 previous grouped value may be omitted', () => {
  assert.equal(has([['地域','人数'],['札幌',3],['',2]], 'L2-02'), true);
});
test('L2-04 orthographic variants after NFKC normalization', () => {
  assert.equal(has([['区分'],['Ａ'],['A']], 'L2-04'), true);
});
test('L3-03 disclosure marker in numeric column', () => {
  assert.equal(has([['人数'],[1],[2],['X']], 'L3-03'), true);
});
test('L3-04 units potentially missing from column heading', () => {
  assert.equal(has([['人数'],[1],[2]], 'L3-04'), true);
});
test('L3-05 Japanese era without Gregorian year', () => {
  assert.equal(has([['年度'],['R5']], 'L3-05'), true);
});
test('L3-06 shortened prefecture names in prefecture column', () => {
  assert.equal(has([['都道府県'],['東京']], 'L3-06'), true);
});
test('L3-09 horizontally arranged year columns', () => {
  assert.equal(has([['ID','2022年','2023年','2024年'],[1,2,3,4]], 'L3-09'), true);
});
test('normal small dataset should not trigger new heuristic rules', () => {
  const out = findings([['ID','名前','年齢'],[1,'山田',30],[2,'鈴木',20]]);
  assert.equal(out.length, 0);
});
test('metadata reader does not break when Apps Script API methods are missing', () => {
  const result = sandbox.sheetInspectionMetadata_({}, 400, 100);
  assert.equal(result.hiddenRows.length, 0);
  assert.ok(result.limitations.length >= 4);
});
test('metadata reader detects hidden rows and image anchors', () => {
  const sheet = {
    isRowHiddenByUser: x => x === 2,
    isColumnHiddenByUser: x => x === 1,
    getImages: () => [{getAnchorCell: () => ({getA1Notation:()=> 'B3'})}],
    getDrawings: () => [],
    getCharts: () => []
  };
  const result = sandbox.sheetInspectionMetadata_(sheet, 3, 3);
  assert.equal(result.hiddenRows[0], 2);
  assert.equal(result.hiddenColumns[0], 1);
  assert.equal(result.objects[0].cell, 'B3');
});
