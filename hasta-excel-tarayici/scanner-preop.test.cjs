const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const s=fs.readFileSync(__dirname+'/scanner.js','utf8');
const c={upper:x=>String(x).toLocaleUpperCase('tr-TR')};vm.createContext(c);
vm.runInContext(s.slice(s.indexOf('  function priorAbdominalOperations('),s.indexOf('  function derive(')),c);
const result=c.priorAbdominalOperations([
 {date:new Date(2020,0,1),text:'Kolesistektomi'},
 {date:new Date(2021,0,1),text:'Low anterior rezeksiyon'},
 {date:new Date(2024,0,1),text:'İnsizyonel herni onarımı'},
 {date:new Date(2020,0,1),text:'Kolesistektomi planlandı'},
 {date:new Date(2020,0,1),text:'Diz artroskopisi'}
],new Date(2023,0,1));
assert.equal(result.length,2);assert.match(result[0].text,/Kolesistektomi/);
console.log('Preoperative history filtering tests passed.');
