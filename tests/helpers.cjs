'use strict';
const fs=require('node:fs'), path=require('node:path'), vm=require('node:vm');
const FILES=['Config.gs','Shared.gs','RuleEngine.gs','RuleHeuristics.gs','SheetReader.gs','ReportWriter.gs','Code.gs'];
function load(services={}) {
  const sandbox={console,...services};
  vm.createContext(sandbox);
  for (const file of FILES) vm.runInContext(fs.readFileSync(path.join(__dirname,'..',file),'utf8'),sandbox,{filename:file});
  return sandbox;
}
module.exports={load,FILES};
