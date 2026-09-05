const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const code=fs.readFileSync(__dirname+'/scanner.js','utf8');
const calls=[];
const records=[{raporId:101},{raporId:101},{raporId:102},{raporId:103},{raporId:104},{id:999}];
const context={state:{stopped:false,paused:false},Number,String,Map,encodeURIComponent,log:()=>{},sleep:async()=>{},cleanText:v=>String(v??''),objectText:JSON.stringify,uniq:v=>[...new Set(v)],apiJson:async path=>{
 calls.push(path);
 if(path.includes('getRisHizmetSonucInfoList'))return{data:records,totalCount:6};
 if(path.endsWith('/101'))return{data:{durum:1,onayTarihi:'01.01.2020',raporTextByRapor:'Umbilikal defekt 4x2 cm',bulgular:'Herni'}};
 if(path.endsWith('/102'))return{data:{durum:0,onayTarihi:'01.01.2020',raporTextByRapor:'İptal rapor'}};
 if(path.endsWith('/103'))throw new Error('timeout');
 if(path.endsWith('/104'))return{data:{durum:1,onayTarihi:'01.01.2020',raporTextByRapor:'İsteyen Branş Görebilir!'}};
 throw new Error('Unexpected path '+path);
}};
vm.createContext(context);
const start=code.indexOf('  const radiologyCache='),end=code.indexOf('  async function settled(',start);
vm.runInContext(code.slice(start,end),context);
(async()=>{
 const result=await context.radiologyHistory({name:'TEST'},777);
 assert.equal(result.audit.total,4);assert.equal(result.audit.read,1);
 assert.equal(result.audit.noReport,1);assert.equal(result.audit.excluded,1);
 assert.equal(result.audit.failures.length,2);
 assert.equal(calls.filter(x=>x.endsWith('/101')).length,1);
 assert.equal(calls.filter(x=>x.endsWith('/103')).length,2);
 const filter=JSON.parse(new URL('https://example.test'+calls[0]).searchParams.get('filter'));
 assert.equal(filter[0].property,'hastaId');assert.equal(filter[0].value,777);
 assert.equal(result.reports.length,1);
 assert.match(code,/length:3/);assert.match(code,/controller.abort\(\),15000/);
 console.log('Background report tests passed: patient filter, dedupe, retry, unavailable and restricted reports.');
})().catch(e=>{console.error(e);process.exitCode=1;});
