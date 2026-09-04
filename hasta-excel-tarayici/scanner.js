(() => {
  'use strict';
  if (window.top !== window || window.__fonetExcelHastaTarayici) return;
  window.__fonetExcelHastaTarayici = true;

  const STORE_KEY = 'fonetExcelHastaTarayiciV1';
  const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
  const norm = value => String(value ?? '').replace(/\s+/g, ' ').trim();
  const upper = value => norm(value).toLocaleUpperCase('tr-TR');
  const visible = el => {
    if (!el) return false;
    const r = el.getBoundingClientRect();
    const s = el.ownerDocument.defaultView.getComputedStyle(el);
    return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
  };
  const click = (el, dbl = false) => {
    if (!el) return;
    el.scrollIntoView({block: 'center', inline: 'nearest'});
    ['mousedown', 'mouseup', 'click'].forEach(type => el.dispatchEvent(new MouseEvent(type, {bubbles:true, cancelable:true, view:el.ownerDocument.defaultView})));
    if (dbl) el.dispatchEvent(new MouseEvent('dblclick', {bubbles:true, cancelable:true, detail:2, view:el.ownerDocument.defaultView}));
  };
  const setValue = (input, value) => {
    input.focus();
    const proto = Object.getPrototypeOf(input);
    const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
    if (setter) setter.call(input, value); else input.value = value;
    input.dispatchEvent(new Event('input', {bubbles:true}));
    input.dispatchEvent(new Event('change', {bubbles:true}));
    input.dispatchEvent(new KeyboardEvent('keyup', {bubbles:true, key:'0'}));
  };
  const waitFor = async (fn, timeout = 10000, interval = 150) => {
    const end = Date.now() + timeout;
    while (Date.now() < end) {
      try { const value = fn(); if (value) return value; } catch {}
      await sleep(interval);
    }
    return null;
  };
  const allDocs = () => {
    const docs = [];
    const visit = doc => {
      if (!doc || docs.includes(doc)) return;
      docs.push(doc);
      for (const frame of doc.querySelectorAll('iframe,frame')) {
        try { visit(frame.contentDocument); } catch {}
      }
    };
    visit(document);
    return docs;
  };
  const elements = selector => allDocs().flatMap(doc => [...doc.querySelectorAll(selector)]);
  const exactButton = text => elements('button,a,[role="button"],.x-btn').find(el => visible(el) && upper(el.innerText) === upper(text));
  const closeWindow = win => { const x = win?.querySelector('.x-tool-close,[data-qtip="Kapat"]'); if (x) click(x); };
  const parseDate = text => {
    const m = norm(text).match(/(\d{2})\.(\d{2})\.(\d{4})/);
    return m ? new Date(+m[3], +m[2]-1, +m[1]) : null;
  };
  const dateText = d => d ? `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${d.getFullYear()}` : '';
  const dateDiffDays = (a,b) => Math.round((b-a)/86400000);
  const dateDiffHuman = (start, end = new Date()) => {
    let y=end.getFullYear()-start.getFullYear(), m=end.getMonth()-start.getMonth(), d=end.getDate()-start.getDate();
    if (d<0) { m--; d += new Date(end.getFullYear(),end.getMonth(),0).getDate(); }
    if (m<0) { y--; m+=12; }
    return `${y} yıl ${m} ay ${d} gün`;
  };
  const uniq = list => [...new Set(list.map(norm).filter(Boolean))];

  const state = {
    workbook: null, sheetName: '', rows: [], headers: [], headerMap: new Map(), patients: [],
    results: {}, mode: 'excel', running: false, paused: false, stopped: false, current: 0, errors: 0
  };

  const FONET_TEMPLATE_HEADERS = [
    'ADI','TC','İŞLEM NO','TELEFON','CİNSİYET','YAŞ','ASA','AMELİYAT TARİHİ','AMELİYAT SÜRESİ(DK)','DEFEKTİN BOYUTU','NÜKS HASTA MI','YATIŞ SÜRESİ','LOKALİZASYON','yeniden yatış','TAKİP SÜRESİ(EN AZ 12 AY)','TAKİPTE NÜKS VAR MI?','PREOPERATİF TANI (NEDEN İNSİZYONEL HERNİ OLMUŞ )','MALİGN ETYOLOJİ','BENİGN ETYOLOJİ','ONLAY ONARIM','SUBLAY ONARIM','İNLAY ONARIM','ONLAY+ABDOMİNOPLASTİ/PANNİKÜLEKTOMİ','SUBLAY+ABDOMİNOPLASTİ/PANNİKÜLEKTOMİ','İNLAY+ABDOMİNOPLASTİ/PANNİKÜLEKTOMİ','POSTOPERATİF SSO(SURGİCAL SİTE OCCURRENCE)','NEKROZ','VAC','SEROMA','REVİZYON','NÜKS HASTA MI','OP SAYI','YANDAŞ HT','YANDAŞ DM','YANDAŞ KOAH','YANDAŞ KALP YETMEZLİĞİ','YANDAŞ HT','YANDAŞ GUATR','SİGARA KULLANIMI','KOMBİNE  AMELİYAT','GREFT','MORBİDİTE','MORTALİTE','FITIK TİPİ','DREN SÜRES','BMİ','PREOP HERQLI','PREOP BODY-Q','POSTOP HERQLI','POSTOP BODY-Q','preop EuroHS QoL','Postop EuroHS QoL'
  ];

  const panel = document.createElement('div');
  panel.id = 'fonet-excel-panel';
  panel.innerHTML = `
    <style>
      #fonet-excel-panel{position:fixed;right:10px;top:10px;z-index:2147483647;width:410px;background:#fff;border:2px solid #145b91;border-radius:9px;box-shadow:0 6px 24px #0006;font:13px Arial;color:#17212b;padding:11px}
      #fonet-excel-panel *{box-sizing:border-box} #fonet-excel-panel .title{font-weight:700;font-size:16px;color:#124e7b;margin-bottom:8px}
      #fonet-excel-panel .note{font-size:11px;color:#52606b;margin:5px 0} #fonet-excel-panel button{border:0;border-radius:5px;padding:8px 10px;margin:3px 2px;background:#0878bd;color:#fff;cursor:pointer}
      #fonet-excel-panel button.alt{background:#596a78} #fonet-excel-panel button.stop{background:#b72a2a} #fonet-excel-panel button:disabled{opacity:.45;cursor:not-allowed}
      #fonet-excel-panel input[type=file]{width:100%;margin:4px 0 7px} #fonet-excel-status{white-space:pre-wrap;line-height:1.45;margin:7px 0}
      #fonet-excel-progress{height:9px;background:#dfe8ee;border-radius:8px;overflow:hidden} #fonet-excel-bar{height:100%;width:0;background:#18a058}
      #fonet-excel-log{margin-top:7px;background:#f3f6f8;max-height:160px;overflow:auto;padding:6px;font-size:11px;border-radius:4px}
      #fonet-excel-panel .ok{color:#087a36}.bad{color:#b42318}
    </style>
    <div class="title">FONET Hasta ve Excel Tarayıcı v1.1</div>
    <input id="fx-file" type="file" accept=".xlsx,.xls" />
    <div>
      <button id="fx-load">Excel Listesini Hazırla</button>
      <button id="fx-fonet-list">FONET Açık Listeyi Al</button>
      <button id="fx-start" disabled>Taramayı Başlat</button>
      <button id="fx-pause" class="alt" disabled>Duraklat</button>
      <button id="fx-stop" class="stop" disabled>Durdur</button>
      <button id="fx-export" class="alt" disabled>Güncellenmiş Excel'i İndir</button>
    </div>
    <div id="fonet-excel-status">Excel dosyasını seçin. FONET'te Ameliyat &gt; Ameliyat ekranı açık olmalıdır.</div>
    <div id="fonet-excel-progress"><div id="fonet-excel-bar"></div></div>
    <div class="note">Bulunamayan hücre korunur. Sonuçlar her hastadan sonra kaydedilir.</div>
    <div id="fonet-excel-log"></div>`;
  document.documentElement.appendChild(panel);
  const $ = sel => panel.querySelector(sel);
  const logBox = $('#fonet-excel-log');
  const log = (text, bad=false) => {
    const line=document.createElement('div'); line.className=bad?'bad':'ok'; line.textContent=`${new Date().toLocaleTimeString('tr-TR')} ${text}`;
    logBox.prepend(line); while(logBox.children.length>35) logBox.lastChild.remove();
  };
  const updateStatus = text => {
    $('#fonet-excel-status').textContent = text || `İşlenen: ${state.current}/${state.patients.length}\nBaşarılı: ${Object.values(state.results).filter(x=>x.status==='Tamamlandı').length} | Hata: ${state.errors}`;
    $('#fonet-excel-bar').style.width = state.patients.length ? `${Math.min(100,state.current/state.patients.length*100)}%` : '0%';
  };
  const persist = () => {
    const compact={sheetName:state.sheetName,patients:state.patients,results:state.results,current:state.current,headers:state.headers};
    try { localStorage.setItem(STORE_KEY,JSON.stringify(compact)); } catch {}
  };

  const headerAliases = {
    name:['ADI','HASTA','ADI SOYADI'], tc:['TC','TC KİMLİK NO','T.C.'], operationNo:['İŞLEM NO','ISLEM NO'], phone:['TELEFON','TELEFON NO'], sex:['CİNSİYET'], age:['YAŞ'], asa:['ASA'],
    surgeryDate:['AMELİYAT TARİHİ'], duration:['AMELİYAT SÜRESİ(DK)','AMELİYAT SÜRESİ'], defect:['DEFEKTİN BOYUTU'], recurrence:['NÜKS HASTA MI'], stay:['YATIŞ SÜRESİ'], location:['LOKALİZASYON'],
    readmission:['YENİDEN YATIŞ'], followup:['TAKİP SÜRESİ(EN AZ 12 AY)'], followRecurrence:['TAKİPTE NÜKS VAR MI?'], preop:['PREOPERATİF TANI (NEDEN İNSİZYONEL HERNİ OLMUŞ )'],
    malign:['MALİGN ETYOLOJİ'], benign:['BENİGN ETYOLOJİ'], onlay:['ONLAY ONARIM'], sublay:['SUBLAY ONARIM'], inlay:['İNLAY ONARIM'],
    onlayAbd:['ONLAY+ABDOMİNOPLASTİ/PANNİKÜLEKTOMİ'], sublayAbd:['SUBLAY+ABDOMİNOPLASTİ/PANNİKÜLEKTOMİ'], inlayAbd:['İNLAY+ABDOMİNOPLASTİ/PANNİKÜLEKTOMİ'],
    sso:['POSTOPERATİF SSO(SURGİCAL SİTE OCCURRENCE)'], necrosis:['NEKROZ'], vac:['VAC'], seroma:['SEROMA'], revision:['REVİZYON'], opCount:['OP SAYI'], ht:['YANDAŞ HT'], dm:['YANDAŞ DM'],
    copd:['YANDAŞ KOAH'], hf:['YANDAŞ KALP YETMEZLİĞİ'], goiter:['YANDAŞ GUATR'], smoking:['SİGARA KULLANIMI'], combined:['KOMBİNE AMELİYAT','KOMBİNE  AMELİYAT'], graft:['GREFT'],
    morbidity:['MORBİDİTE'], mortality:['MORTALİTE'], herniaType:['FITIK TİPİ'], drain:['DREN SÜRES','DREN SÜRESİ'], bmi:['BMİ']
  };
  const findHeader = key => {
    const aliases=headerAliases[key]||[];
    for(const alias of aliases){const ix=state.headers.findIndex(h=>upper(h)===upper(alias));if(ix>=0)return ix;}
    return -1;
  };
  const getCell = (row,key) => { const i=findHeader(key); return i>=0?row[i]:''; };
  const setIfFound = (row,key,value) => { const i=findHeader(key); if(i>=0 && value!=='' && value!=null) row[i]=value; };

  function searchControls(){
    for(const doc of allDocs()){
      const buttons=[...doc.querySelectorAll('button,a,[role="button"],.x-btn')].filter(visible);
      const query=buttons.find(x=>upper(x.innerText)==='SORGULA'); const clear=buttons.find(x=>upper(x.innerText)==='TEMİZLE');
      if(!query||!clear)continue;
      let inputs=[...doc.querySelectorAll('input')].filter(x=>visible(x)&&/Kimlik No/i.test(norm(x.closest('.x-field,.x-form-item,table')?.innerText)));
      if(!inputs.length) inputs=[...doc.querySelectorAll('input')].filter(visible);
      const tcInput=inputs.find(x=>!x.disabled && !x.readOnly && (!x.value || /^\d{0,11}$/.test(x.value)));
      if(tcInput)return{doc,query,clear,tcInput};
    }
    return null;
  }
  function operationRows(doc, patient){
    const candidates=[...doc.querySelectorAll('tr[data-recordindex],[role="row"]')].filter(row=>visible(row)&&!row.closest('.x-window'));
    const name=upper(patient.name), tc=norm(patient.tc);
    return candidates.filter(row=>{const t=upper(row.innerText);return (/\d{2}\.\d{2}\.\d{4}/.test(t))&&(t.includes(name)||t.includes(tc));});
  }
  function openListRows(){
    let best=[];
    for(const doc of allDocs()){
      const rows=[...doc.querySelectorAll('tr[data-recordindex],[role="row"]')].filter(row=>{
        if(!visible(row)||row.closest('.x-window'))return false;
        const t=norm(row.innerText);return /\d{2}\.\d{2}\.\d{4}/.test(t)&&/\b\d{6,}\b/.test(t);
      });
      if(rows.length>best.length)best=rows;
    }
    return best;
  }
  function extListSource(){
    let best=null;
    for(const doc of allDocs()){
      const Ext=doc.defaultView.Ext;if(!Ext?.ComponentQuery)continue;
      let grids=[];try{grids=Ext.ComponentQuery.query('gridpanel');}catch{}
      for(const grid of grids){
        const store=grid.getStore?.(),count=store?.getCount?.()||0,total=store?.getTotalCount?.()||count;
        if(count&&(!best||Math.max(count,total)>Math.max(best.count,best.total)))best={grid,store,count,total};
      }
    }
    return best;
  }
  function extRecordData(record,index){
    const data=record?.data||{},entries=Object.entries(data),values=entries.map(x=>norm(x[1])).filter(Boolean),line=values.join(' | ');
    const byKey=rx=>norm(entries.find(([k])=>rx.test(k))?.[1]);
    const operationNo=byKey(/islem.*no|işlem.*no|operation/i)||(line.match(/\b\d{6,}\b/g)||[]).at(-1)||'';
    const surgeryDate=byKey(/ameliyat.*tarih|istek.*tarih|işlem.*tarih/i)||(line.match(/\d{2}\.\d{2}\.\d{4}(?:\s+\d{2}:\d{2}(?::\d{2})?)?/)||[])[0]||'';
    const name=byKey(/adi.*soyadi|ad.*soyad|hasta.*ad/i)||values.find(x=>/^[A-ZÇĞİÖŞÜ][A-ZÇĞİÖŞÜ .'-]{3,}$/i.test(x)&&!/(GENEL|CERRAH|UZMAN|DOKTOR|AMELİYAT|ACİL)/i.test(x))||'';
    return{index,operationNo,surgeryDate,name};
  }
  function selectExtOperation(operationNo,listIndex){
    const source=extListSource();if(!source)return false;
    let record=Number.isInteger(listIndex)?source.store.getAt?.(listIndex):null;
    if(!record)source.store.each?.(r=>{if(!record&&Object.values(r.data||{}).some(v=>norm(v)===operationNo))record=r;});
    if(!record)return false;
    try{
      const view=source.grid.getView?.(),node=view?.getNode?.(record);
      source.grid.getSelectionModel?.().select(record);view?.focusRow?.(record);
      if(node){click(node,true);return true;}
      view?.fireEvent?.('itemdblclick',view,record,node,listIndex,new MouseEvent('dblclick',{bubbles:true}));
      source.grid.fireEvent?.('itemdblclick',view,record,node,listIndex,new MouseEvent('dblclick',{bubbles:true}));
      return true;
    }catch{return false;}
  }
  function listRowData(row,index){
    const cells=[...row.querySelectorAll('td,[role="gridcell"]')].map(x=>norm(x.innerText)).filter(Boolean);
    const line=norm(row.innerText);const operationNo=(line.match(/\b\d{6,}\b/g)||[]).at(-1)||'';
    const surgeryDate=(line.match(/\d{2}\.\d{2}\.\d{4}(?:\s+\d{2}:\d{2}(?::\d{2})?)?/)||[])[0]||'';
    const opIndex=cells.findIndex(x=>x.includes(operationNo));
    const after=opIndex>=0?cells.slice(opIndex+1):cells;
    const name=after.find(x=>/^[A-ZÇĞİÖŞÜ][A-ZÇĞİÖŞÜ .'-]{3,}$/i.test(x)&&!/(GENEL|CERRAH|UZMAN|DOKTOR|AMELİYAT|ACİL)/i.test(x))||'';
    return{row,index,operationNo,surgeryDate,name};
  }
  function chooseOperation(rows, targetDate){
    const target=parseDate(targetDate); let best=null,bestDiff=Infinity;
    for(const row of rows){const d=parseDate(row.innerText);if(!d)continue;const diff=target?Math.abs(d-target):0;if(diff<bestDiff){best=row;bestDiff=diff;}}
    return best||rows[0]||null;
  }
  function fieldValue(label){
    const key=upper(label);
    for(const doc of allDocs()){
      for(const box of doc.querySelectorAll('.x-field,.x-form-item,table')){
        if(!visible(box)||!upper(box.innerText).includes(key))continue;
        const input=box.querySelector('input,textarea,.x-form-display-field'); const value=norm(input?.value??input?.innerText); if(value)return value;
      }
    }
    return '';
  }
  function currentPatientName(){return fieldValue('Adı Soyadı:');}
  function currentTc(){
    const values=elements('input,.x-form-display-field').filter(visible).map(x=>norm(x.value??x.innerText));
    return values.find(x=>/^\d{11}$/.test(x))||fieldValue('Kimlik No:').replace(/\D/g,'').slice(0,11);
  }
  function collectVisibleText(scopeDocs=allDocs()){
    const out=[];
    for(const doc of scopeDocs){
      for(const el of doc.querySelectorAll('textarea,input,.x-form-display-field,tr[data-recordindex],[role="row"]')){
        if(!visible(el))continue; const v=norm(el.value??el.innerText); if(v&&v.length>2)out.push(v);
      }
    }
    return uniq(out);
  }
  function currentOperationRows(){
    return elements('tr[data-recordindex],[role="row"]').filter(row=>visible(row)&&!row.closest('.x-window')&&/\d{2}\.\d{2}\.\d{4}/.test(norm(row.innerText)));
  }
  async function readTab(tabName, wait=500){
    const btn=exactButton(tabName); if(!btn)return[]; click(btn); await sleep(wait); return collectVisibleText();
  }
  async function readHistory(tc){
    const button=exactButton('Hasta Geçmişi'); if(!button)throw new Error('Hasta Geçmişi düğmesi bulunamadı');
    click(button);
    const win=await waitFor(()=>elements('.x-window').find(w=>visible(w)&&/Hasta Geçmiş/i.test(norm(w.innerText).slice(0,200))),10000);
    if(!win)throw new Error('Hasta geçmişi açılmadı');
    await sleep(700);
    const rows=[...win.querySelectorAll('tr[data-recordindex],[role="row"]')].filter(visible).map(x=>norm(x.innerText)).filter(x=>x.includes(tc)||/\d{2}\.\d{2}\.\d{4}/.test(x));
    closeWindow(win); await sleep(250); return uniq(rows);
  }
  function quantityFromMaterialRow(text){
    const cells=norm(text).split(/\s+/); const nums=cells.map(x=>Number(String(x).replace(',','.'))).filter(x=>Number.isFinite(x)&&x>0&&x<=10&&Number.isInteger(x));
    return nums.length?nums[nums.length-1]:1;
  }
  function derive(patient, details){
    const surgeryDate=parseDate(patient.surgeryDate)||parseDate(details.selectedOperation);
    const note=details.note.join(' | '), allHistory=details.history.join(' | '), all=upper(`${note} ${allHistory}`);
    const opRows=details.surgeries.map(text=>({text,date:parseDate(text)})).filter(x=>x.date);
    const previousHernia=details.history.filter(x=>/\bK43(?:\.|-)|İNSİZYONEL HERNİ|VENTRAL HERNİ/i.test(x)&&parseDate(x)&&surgeryDate&&parseDate(x)<surgeryDate);
    const laterHerniaOps=opRows.filter(x=>surgeryDate&&x.date>surgeryDate&&/İNSİZYONEL HERNİ|VENTRAL HERNİ/i.test(x.text));
    const laterAdmissions=details.history.filter(x=>{const d=parseDate(x);return d&&surgeryDate&&dateDiffDays(surgeryDate,d)>2&&/\bYATIŞ\b/i.test(x);});
    const previousAbdominal=opRows.filter(x=>surgeryDate&&x.date<surgeryDate&&/ABDOM|LAPAROT|LAPAROSK|APPEN|KOLESİST|KOLEKT|REZEKS|GASTREK|HERNİ|FITIK|SEZARYEN|HİSTEREKT|OOFOREKT|SALPEN|KOLON|REKT|İLEOST|KOLOST|PANKREAT|SPLENEKT|BARSAK|BAĞIRSAK|UMBİLİK|MİDE|BYPASS|SLEEVE/i.test(x.text));
    const cancers=uniq(details.history.flatMap(x=>x.match(/\bC\d{2}(?:\.\d+)?-[^|]+/gi)||[]));
    const materialRows=details.materials.filter(x=>/PROLEN|PROLENE|MESH|MEŞ/i.test(x));
    const prolenRows=materialRows.filter(x=>/PROLEN|PROLENE/i.test(x));
    const prolenCount=prolenRows.reduce((sum,x)=>sum+quantityFromMaterialRow(x),0);
    const ageSex=details.fields.ageSex||''; const sex=/\((Kadın|Erkek)\)/i.exec(ageSex)?.[1]||''; const age=/^(\d+)/.exec(ageSex)?.[1]||'';
    const times=(details.fields.surgeryTimes||'').match(/\d{2}:\d{2}/g)||[]; let duration='';
    if(times.length>=2){const [h1,m1]=times[0].split(':').map(Number),[h2,m2]=times[times.length-1].split(':').map(Number);duration=(h2*60+m2)-(h1*60+m1);if(duration<0)duration+=1440;}
    const diagnoses=upper(allHistory);
    const complications=[]; if(/NEKROZ/.test(all))complications.push('Nekroz');if(/SEROMA/.test(all))complications.push('Seroma');if(/YARA ENFEKSİY|CERRAHİ ALAN ENFEKSİY|ENFEKSİYON/.test(all))complications.push('Enfeksiyon');if(/DEHİS|EVİSSER/.test(all))complications.push('Dehisens');
    return {surgeryDate,previousHernia,laterHerniaOps,laterAdmissions,previousAbdominal,cancers,materialRows,prolenCount,sex,age,duration,diagnoses,all,note,complications};
  }
  function applyResult(patient, details){
    const row=state.rows[patient.rowIndex], d=derive(patient,details);
    setIfFound(row,'operationNo',details.fields.operationNo);
    setIfFound(row,'phone',details.fields.phone);
    setIfFound(row,'sex',d.sex); setIfFound(row,'age',d.age?Number(d.age):''); setIfFound(row,'asa',details.fields.asa);
    setIfFound(row,'duration',d.duration||''); setIfFound(row,'followup',d.surgeryDate?dateDiffHuman(d.surgeryDate):'');
    setIfFound(row,'recurrence',d.previousHernia.length?1:0);
    setIfFound(row,'followRecurrence',d.laterHerniaOps.length?`Evet – ${d.laterHerniaOps.map(x=>`${dateText(x.date)} ${norm(x.text)}`).join('; ')}`:'Hayır');
    if(d.previousAbdominal.length)setIfFound(row,'preop',d.previousAbdominal.map(x=>`${dateText(x.date)} ${norm(x.text)}`).join('; '));
    setIfFound(row,'readmission',d.laterAdmissions.length?d.laterAdmissions.map(x=>norm(x)).join('; '):'YOK');
    if(d.cancers.length){setIfFound(row,'malign',`1 – ${d.cancers.join('; ')}`);setIfFound(row,'benign',0);}else{setIfFound(row,'malign',0);setIfFound(row,'benign',1);}
    const hasAbd=/ABDOMİNOPLAST|PANNİKÜLEKT/i.test(d.note);
    if(/\bONLAY\b/i.test(d.note)){setIfFound(row,'onlay',1);if(hasAbd)setIfFound(row,'onlayAbd',1);}
    if(/\bSUBLAY\b/i.test(d.note)){setIfFound(row,'sublay',1);if(hasAbd)setIfFound(row,'sublayAbd',1);}
    if(/\bINLAY\b|\bİNLAY\b/i.test(d.note)){setIfFound(row,'inlay',1);if(hasAbd)setIfFound(row,'inlayAbd',1);}
    if(d.complications.length)setIfFound(row,'sso',d.complications.join(', '));
    if(/NEKROZ/.test(d.all))setIfFound(row,'necrosis',1); if(/\bVAC\b/.test(d.all))setIfFound(row,'vac',1); if(/SEROMA/.test(d.all))setIfFound(row,'seroma',1);
    if(/REVİZYON|REVIZYON|REAKSPLORASYON/.test(d.all))setIfFound(row,'revision',1);
    if(/\bI10(?:\.|-)/.test(d.diagnoses))setIfFound(row,'ht',1); if(/\bE1[01](?:\.|-)/.test(d.diagnoses))setIfFound(row,'dm',1);
    if(/\bJ4[34](?:\.|-)/.test(d.diagnoses))setIfFound(row,'copd',1); if(/\bI50(?:\.|-)/.test(d.diagnoses))setIfFound(row,'hf',1);
    if(/\bE0[0-7](?:\.|-)|GUATR/.test(d.diagnoses))setIfFound(row,'goiter',1);
    if(/SİGARA.*(İÇİYOR|KULLANIYOR|AKTİF)|AKTİF SİGARA/.test(d.all))setIfFound(row,'smoking',1);
    if(d.materialRows.length)setIfFound(row,'graft',uniq(d.materialRows).join('; '));
    if(/DREN/.test(d.note)){const m=d.note.match(/DREN[^|]{0,60}?(\d+)\s*GÜN/i);if(m)setIfFound(row,'drain',`${m[1]} GÜN`);}
    row[state.headerMap.get('PROLEN MESH ADEDİ')] = d.prolenCount||0;
    row[state.headerMap.get('MALZEME KAYDI')] = d.materialRows.length?uniq(d.materialRows).join('; '):'FONET sarf kaydında mesh/prolen bulunmadı';
    row[state.headerMap.get('FONET TARAMA DURUMU')] = 'Tamamlandı';
    return d;
  }
  async function scanPatient(patient){
    let found,selected;
    if(patient.mode==='fonet-list'){
      found=openListRows();
      selected=found.find(r=>norm(r.innerText).includes(patient.operationNo))||chooseOperation(found.filter(r=>upper(r.innerText).includes(upper(patient.name))),patient.surgeryDate);
      if(selected)click(selected,true);else if(!selectExtOperation(patient.operationNo,patient.listIndex))throw new Error('Açık listedeki ameliyat satırı yeniden bulunamadı');
    }else{
      const controls=searchControls(); if(!controls)throw new Error('Ameliyat arama ekranı bulunamadı');
      click(controls.clear); await sleep(180); setValue(controls.tcInput,patient.tc); click(controls.query);
      found=await waitFor(()=>operationRows(controls.doc,patient),8000);
      if(!found||!found.length)throw new Error('TC için ameliyat kaydı bulunamadı');
      selected=chooseOperation(found,patient.surgeryDate); click(selected);
    }
    await sleep(350);
    const loaded=await waitFor(()=>upper(currentPatientName())===upper(patient.name)||(currentPatientName()&&currentTc()),10000);
    if(!loaded)throw new Error('Hasta bilgileri yüklenmedi');
    patient.name=currentPatientName()||patient.name;patient.tc=currentTc()||patient.tc;
    if(patient.mode==='fonet-list'){
      const row=state.rows[patient.rowIndex];setIfFound(row,'name',patient.name);setIfFound(row,'tc',patient.tc);setIfFound(row,'operationNo',patient.operationNo);setIfFound(row,'surgeryDate',patient.surgeryDate);
    }
    const fields={operationNo:fieldValue('İşlem No :')||patient.operationNo,phone:fieldValue('Telefon')||fieldValue('Cep Telefonu'),ageSex:fieldValue('Yaş / Cinsiyet / D.Tar:'),asa:fieldValue('Asa/Euro Score:'),surgeryTimes:`${fieldValue('Ameliyat Saati')} ${fieldValue('Post-Op Saati')}`};
    const surgeries=found.map(x=>norm(x.innerText));
    const selectedOperation=selected?norm(selected.innerText):`${patient.surgeryDate} ${patient.name} ${patient.operationNo}`;
    const note=await readTab('Ameliyat Notları',650);
    await readTab('Ameliyat Bilgileri',250);
    const materials=await readTab('İlaç Sarf',650);
    await readTab('Hizmet Listesi',250);
    const history=patient.tc?await readHistory(patient.tc):[];
    return{fields,surgeries,selectedOperation,note,materials,history};
  }
  async function run(){
    if(!state.patients.length||state.running)return;
    state.running=true;state.stopped=false;state.errors=0;$('#fx-start').disabled=true;$('#fx-pause').disabled=false;$('#fx-stop').disabled=false;
    for(let i=state.current;i<state.patients.length;i++){
      if(state.stopped)break; while(state.paused&&!state.stopped)await sleep(250); if(state.stopped)break;
      const p=state.patients[i];
      try{const details=await scanPatient(p);const derived=applyResult(p,details);state.results[p.tc||`${p.operationNo}|${p.surgeryDate}`]={status:'Tamamlandı',details,derived:{prolenCount:derived.prolenCount,readmissions:derived.laterAdmissions.length}};log(`${p.name}: tamamlandı, Prolen mesh ${derived.prolenCount}, yeniden yatış ${derived.laterAdmissions.length}`);}
      catch(error){state.errors++;state.results[p.tc||`${p.operationNo}|${p.surgeryDate}`]={status:'Hata',error:String(error.message||error)};const r=state.rows[p.rowIndex];r[state.headerMap.get('FONET TARAMA DURUMU')]=`Hata: ${error.message||error}`;log(`${p.name||p.operationNo}: ${error.message||error}`,true);}
      state.current=i+1;persist();updateStatus();await sleep(220);
    }
    state.running=false;$('#fx-start').disabled=false;$('#fx-pause').disabled=true;$('#fx-stop').disabled=true;$('#fx-export').disabled=false;
    updateStatus(state.stopped?'Tarama durduruldu. Sonuçlar kaydedildi.':'Tarama tamamlandı. Güncellenmiş Excel indirilebilir.');
  }
  function ensureOutputColumns(){
    for(const h of ['PROLEN MESH ADEDİ','MALZEME KAYDI','FONET TARAMA DURUMU']){
      let ix=state.headers.findIndex(x=>upper(x)===upper(h));
      if(ix<0){ix=state.headers.length;state.headers.push(h);state.rows[0][ix]=h;for(let r=1;r<state.rows.length;r++)if(state.rows[r][ix]===undefined)state.rows[r][ix]='';}
      state.headerMap.set(h,ix);
    }
  }
  async function loadExcel(){
    const file=$('#fx-file').files[0];if(!file){updateStatus('Önce Excel dosyasını seçin.');return;}
    try{
      const data=await file.arrayBuffer();state.mode='excel';state.workbook=XLSX.read(data,{type:'array',cellStyles:true,cellDates:false});state.sheetName=state.workbook.SheetNames[0];
      const ws=state.workbook.Sheets[state.sheetName];state.rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:''});state.headers=state.rows[0].map(norm);ensureOutputColumns();
      const nameIx=findHeader('name'),tcIx=findHeader('tc'),dateIx=findHeader('surgeryDate');if(nameIx<0||tcIx<0)throw new Error('ADI ve TC sütunları bulunamadı');
      state.patients=state.rows.slice(1).map((r,i)=>({rowIndex:i+1,name:norm(r[nameIx]),tc:norm(r[tcIx]).replace(/\.0$/,''),surgeryDate:dateIx>=0?(typeof r[dateIx]==='number'?dateText(new Date((r[dateIx]-25569)*86400000)):norm(r[dateIx])):''})).filter(x=>x.name&&/^\d{11}$/.test(x.tc));
      state.results={};state.current=0;state.errors=0;persist();$('#fx-start').disabled=false;$('#fx-export').disabled=false;updateStatus(`${state.patients.length} hasta hazırlandı. FONET Ameliyat ekranındayken Taramayı Başlat'a basın.`);log(`${file.name}: ${state.patients.length} hasta yüklendi`);
    }catch(error){updateStatus(`Excel okunamadı: ${error.message||error}`);log(String(error.message||error),true);}
  }
  function loadFonetList(){
    try{
      const source=extListSource();
      const parsed=source&&source.count>openListRows().length
        ? source.store.getRange().map(extRecordData).filter(x=>x.operationNo)
        : openListRows().map(listRowData).filter(x=>x.operationNo);
      const unique=[...new Map(parsed.map(x=>[`${x.operationNo}|${x.surgeryDate}`,x])).values()];
      if(!unique.length)throw new Error('Açık ameliyat listesinde hasta satırı bulunamadı. Liste ekranını açık ve yüklenmiş bırakın.');
      state.mode='fonet-list';state.workbook=XLSX.utils.book_new();state.sheetName='Hastalar';state.headers=[...FONET_TEMPLATE_HEADERS];
      state.rows=[state.headers,...unique.map(x=>{const r=Array(state.headers.length).fill('');r[0]=x.name;r[2]=x.operationNo;r[7]=x.surgeryDate;return r;})];
      ensureOutputColumns();XLSX.utils.book_append_sheet(state.workbook,XLSX.utils.aoa_to_sheet(state.rows),state.sheetName);
      state.patients=unique.map((x,i)=>({rowIndex:i+1,name:x.name,tc:'',operationNo:x.operationNo,surgeryDate:x.surgeryDate,listIndex:x.index,mode:'fonet-list'}));
      state.results={};state.current=0;state.errors=0;persist();$('#fx-start').disabled=false;$('#fx-export').disabled=false;
      updateStatus(`${state.patients.length} ameliyat FONET açık listesinden alındı. Taramayı Başlat'a basın.`);log(`FONET listesinden ${state.patients.length} kayıt hazırlandı`);
    }catch(error){updateStatus(error.message||String(error));log(error.message||String(error),true);}
  }
  function exportExcel(){
    if(!state.rows.length)return;
    const ws=XLSX.utils.aoa_to_sheet(state.rows);ws['!cols']=state.headers.map((h,i)=>({wch:Math.min(60,Math.max(12,Math.max(...state.rows.slice(0,46).map(r=>norm(r[i]).length))+2))}));
    state.workbook.Sheets[state.sheetName]=ws;
    const auditRows=[['TC','Hasta','Durum','Prolen mesh adedi','Yeniden yatış','Hata']];
    for(const p of state.patients){const rr=state.results[p.tc||`${p.operationNo}|${p.surgeryDate}`]||{};auditRows.push([p.tc,p.name,rr.status||'Taranmadı',rr.derived?.prolenCount??'',rr.derived?.readmissions??'',rr.error||'']);}
    state.workbook.Sheets['FONET Tarama Kaydı']=XLSX.utils.aoa_to_sheet(auditRows);if(!state.workbook.SheetNames.includes('FONET Tarama Kaydı'))state.workbook.SheetNames.push('FONET Tarama Kaydı');
    XLSX.writeFile(state.workbook,`FONET_TARANMIS_${new Date().toISOString().slice(0,10)}.xlsx`,{compression:true,cellStyles:true});
  }
  $('#fx-load').onclick=loadExcel;$('#fx-fonet-list').onclick=loadFonetList;$('#fx-start').onclick=run;$('#fx-pause').onclick=()=>{state.paused=!state.paused;$('#fx-pause').textContent=state.paused?'Devam Et':'Duraklat';updateStatus(state.paused?'Tarama duraklatıldı.':'Tarama sürüyor...');};
  $('#fx-stop').onclick=()=>{state.stopped=true;state.paused=false;};$('#fx-export').onclick=exportExcel;
})();
