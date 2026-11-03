const test=require('node:test'),assert=require('node:assert/strict');
const {load}=require('./helpers.cjs');
test('recognizes only own output sheet',()=>{
 const s=load();
 const userSheet={getName:()=> '機械可読性_検査結果', getDeveloperMetadata:()=>[],getMaxColumns:()=>3,getRange:()=>({getDisplayValues:()=>[['ID','値','説明']]})};
 const generated={getName:()=> '機械可読性_検査結果',getDeveloperMetadata:()=>[{getKey:()=> 'GENERIC_HARUNOBU_REPORT'}]};
 assert.deepEqual(s.selectTargetSheets_({getSheets:()=>[userSheet,generated]}),[userSheet]);
});
test('one unreadable sheet does not abort the audit',()=>{
 const s=load();
 const range={getNumRows:()=>2,getNumColumns:()=>1,getValues:()=>[['ID'],[' A ']],getFormulas:()=>[[''],['']],getMergedRanges:()=>[]};
 const broken={getName:()=> '壊れた表',getDataRange:()=>{throw Error('not accessible')}};
 const good={getName:()=> '台帳',getSheetId:()=>1,getDataRange:()=>range};
 const result=s.collectChecks_([broken,good]);
 assert.equal(result.scanned,1);
 assert.equal(result.skipped.length,1);
 assert.ok(result.findings.some(f=>f.ruleId==='L1-06'));
});
test('formula-like sheet names are escaped on report output',()=>{
 const s=load();
 assert.equal(s.safeReportText_('=1+2'),"'=1+2");
 assert.equal(s.safeReportText_('利用者'),'利用者');
});
test('empty value grid still inspects objects',()=>{
 const s=load();
 const result=s.inspectGrid_({sheetName:'データ',gid:1,values:[['']],formulas:[['']],objects:[{type:'画像',cell:'B2'}],mergedA1:[]},100);
 assert.ok(result.findings.some(f=>f.ruleId==='L1-10'));
});
