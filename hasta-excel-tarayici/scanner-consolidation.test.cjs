const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const code=fs.readFileSync(__dirname+'/scanner.js','utf8');
const headers=['name','tc','operationNo','surgeryDate','recurrence','opCount','followRecurrence','readmission','revision','malign','benign'];
const context={norm:x=>String(x??'').trim(),upper:x=>String(x??'').toLocaleUpperCase('tr-TR'),cleanText:x=>String(x??''),uniq:x=>[...new Set(x)],parseDateTime:x=>new Date(x),getCell:(row,k)=>row[headers.indexOf(k)],setIfFound:(row,k,v)=>{const i=headers.indexOf(k);if(i>=0)row[i]=v;},state:{mode:'fonet-list',headers,headerMap:new Map(),rows:[headers],patients:[]}};
context.headerAliases={preop:[]};
vm.createContext(context);
for(const [start,end] of [['function malignancyEvidence','function derive'],['function consolidateDuplicatePatients','async function run']])vm.runInContext(code.slice(code.indexOf('  '+start),code.indexOf('  '+end)),context);
const cancer=context.malignancyEvidence({history:['Malignite saptanmadı.'],imaging:['Malignite lehine kitle.'],pathology:['Adenokarsinom.']});
assert.equal(cancer.length,2);assert.match(cancer[0],/Şüpheli/);assert.match(cancer[1],/Patoloji/);
const st=context.state;
for(let i=1;i<=3;i++){st.rows.push(Array(headers.length).fill(''));st.patients.push({rowIndex:i,tc:'11111111111',name:'TEST',ameliyatId:i===1?10:20,gelisId:i===1?1:2,operationNo:i===1?'A':'B',surgeryDate:i===1?'2020-01-01':'2021-01-01'});}
st.rows[2][7]='Evet — 2021-01-01; insizyonel herni yatışı';
assert.equal(context.consolidateDuplicatePatients(),2);
assert.equal(st.rows.length,2);assert.equal(st.rows[1][4],'Evet');assert.equal(st.rows[1][5],2);
assert.match(st.rows[1][7],/insizyonel herni yatışı/);assert.match(st.rows[1][7],/2021-01-01/);
console.log('Consolidation and malignancy evidence tests passed.');
