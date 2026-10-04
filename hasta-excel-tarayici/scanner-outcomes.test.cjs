const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const code=fs.readFileSync(__dirname+'/scanner.js','utf8');
const start=code.indexOf('  const POSTOP_EVENT_RULES'),end=code.indexOf('  function derive(',start);
const context={
  norm:x=>String(x??'').replace(/\s+/g,' ').trim(),
  cleanText:x=>String(x??'').replace(/\s+/g,' ').trim(),
  upper:x=>String(x??'').replace(/\s+/g,' ').trim().toLocaleUpperCase('tr-TR'),
  parseDateTime:x=>{const m=String(x??'').match(/(\d{2})\.(\d{2})\.(\d{4})/);return m?new Date(+m[3],+m[2]-1,+m[1]):null;},
  dateDiffDays:(a,b)=>Math.round((b-a)/86400000),
  dateText:d=>d?`${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${d.getFullYear()}`:'',
  Map,Set
};
vm.createContext(context);vm.runInContext(code.slice(start,end),context);
const surgery=new Date(2025,0,1);
const result=context.classifyPostoperativeEvents(surgery,[
  {source:'Konsültasyon',date:new Date(2025,0,5),text:'Pnömoni saptandı, antibiyotik başlandı.'},
  {source:'Radyoloji',date:new Date(2025,0,6),text:'Pulmoner emboli izlenmedi.'},
  {source:'Ameliyat',date:new Date(2025,0,12),text:'Yara debridmanı için yeniden ameliyat edildi.'},
  {source:'Hasta kaydı',date:new Date(2025,1,10),text:'Eksitus.'}
]);
assert.equal(result.morbidity30,true);
assert.equal(result.mortality30,false);
assert.equal(result.mortality90,true);
assert.equal(result.maxClavien,'V');
assert.equal(result.major,true);
assert.equal(result.events.some(x=>x.type==='Tromboemboli'),false);
assert.equal(result.events.some(x=>x.type==='Pulmoner'),true);
assert.equal(result.events.some(x=>x.type==='Reoperasyon'),true);
const chronic=context.classifyPostoperativeEvents(surgery,[
  {source:'Özgeçmiş',date:new Date(2025,0,8),text:'Kalp yetmezliği öyküsü ve eski diyaliz kaydı mevcut.'},
  {source:'Başvuru',date:new Date(2025,0,9),text:'Geçmişte yatış öyküsü var.'}
]);
assert.equal(chronic.events.length,0);
console.log('Retrospective outcome tests passed: windows, negation, morbidity, reoperation and mortality.');
