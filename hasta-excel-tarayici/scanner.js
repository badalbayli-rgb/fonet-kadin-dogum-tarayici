(() => {
  'use strict';
  if (window.top !== window || window.__fonetExcelHastaTarayici) return;
  window.__fonetExcelHastaTarayici = true;

  const STORE_KEY = 'fonetExcelHastaTarayiciV1';
  const activeRequests=new Set();
  const abortRequests=()=>{for(const request of activeRequests)request.abort();};
  const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
  const norm = value => String(value ?? '').replace(/\s+/g, ' ').trim();
  const cleanText = value => norm(String(value ?? '').replace(/<br\s*\/?>/gi,' ').replace(/<[^>]+>/g,' ').replace(/&nbsp;|&#160;/gi,' ').replace(/&amp;/gi,'&'));
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
    const fire=(type,detail)=>el.dispatchEvent(new MouseEvent(type,{bubbles:true,cancelable:true,detail,view:el.ownerDocument.defaultView}));
    ['mousedown','mouseup','click'].forEach(type=>fire(type,1));
    if(dbl){['mousedown','mouseup','click'].forEach(type=>fire(type,2));fire('dblclick',2);}
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
  const parseDateTime = text => {
    const m=norm(text).match(/(\d{2})\.(\d{2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
    return m?new Date(+m[3],+m[2]-1,+m[1],+(m[4]||0),+(m[5]||0)):null;
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
    results: {}, mode: 'excel', running: false, paused: false, stopped: false, destroyRequested: false, current: 0, errors: 0
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
      #fonet-excel-panel .head{display:flex;align-items:center;justify-content:space-between;gap:8px} #fonet-excel-panel .head .title{margin-bottom:0}
      #fonet-excel-panel .close{background:#475569;padding:6px 10px} #fonet-excel-panel .ok{color:#087a36}.bad{color:#b42318}
    </style>
    <div class="head"><div class="title">FONET Hasta ve Excel Tarayıcı v1.9.1 — Arka plan</div><button id="fx-close" class="close">Kapat</button></div>
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
  const setIfFound = (row,key,value) => {
    if(value===''||value==null)return;
    const aliases=(headerAliases[key]||[]).map(upper);
    state.headers.forEach((header,index)=>{if(aliases.includes(upper(header)))row[index]=value;});
  };

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
        if(!count)continue;
        const keyList=Object.keys(store.getAt?.(0)?.data||{}),keys=keyList.join(' ');
        const normalizedKeys=keyList.map(k=>String(k).replace(/[^a-z0-9]/gi,'').toLowerCase());
        const isOperationGrid=['islemno','adisoyadi','gelisid','birimsevkid'].every(k=>normalizedKeys.includes(k));
        if(!isOperationGrid)continue;
        const sample=JSON.stringify(store.getAt?.(0)?.data||{}).slice(0,1800);
        let score=Math.min(count,2500)/30;
        score+=10000;
        if(/islem.?no|işlem.?no|ameliyat|hasta|adi.?soyadi|ad.?soyad/i.test(keys))score+=500;
        if(/insizyonel|herni|ameliyat|cerrahi/i.test(sample))score+=250;
        if(/hizmet.?list|laboratuvar|radyoloji|menü|menu/i.test(`${grid.title||''} ${grid.itemId||''}`))score-=1000;
        if(!best||score>best.score)best={grid,store,count,total,score};
      }
    }
    return best;
  }
  function extRecordData(record,index){
    const data=record?.data||{},entries=Object.entries(data),values=entries.map(x=>norm(x[1])).filter(Boolean),line=values.join(' | ');
    const nk=k=>String(k).replace(/[^a-zA-Z0-9çğıöşüÇĞİÖŞÜ]/g,'').toLocaleLowerCase('tr-TR');
    const exact=(keys)=>{for(const wanted of keys){const hit=entries.find(([k])=>nk(k)===wanted);if(hit&&norm(hit[1]))return norm(hit[1]);}return'';};
    const operationNo=exact(['islemno','ameliyatno','protokolno'])||(line.match(/\b\d{6,}\b/g)||[]).at(-1)||'';
    const surgeryDate=exact(['ameliyattarihi','islemtarihi','istemtarihi','tarih','baslamatarihi'])||(line.match(/\d{2}\.\d{2}\.\d{4}(?:\s+\d{2}:\d{2}(?::\d{2})?)?/)||[])[0]||'';
    let name=exact(['adisoyadi','adsoyad','hastaadisoyadi','hastaadsoyad']);
    if(!name)name=norm(entries.find(([k])=>/hasta/i.test(k)&&/ad|adi/i.test(k)&&!/doktor|personel/i.test(k))?.[1]);
    if(!name)name=values.find(x=>/^[A-ZÇĞİÖŞÜ][A-ZÇĞİÖŞÜ .'-]{3,}$/i.test(x)&&!/(GENEL|CERRAH|UZMAN|DOKTOR|AMELİYAT|ACİL|AMBULANS|ÜCRET|ALGOLOJİ)/i.test(x))||'';
    return{
      index,operationNo,surgeryDate,name,
      ameliyatId:norm(data.id),gelisId:norm(data.gelisId),birimSevkId:norm(data.birimSevkId),
      isteyenBirimSevkId:norm(data.isteyenBirimSevkId),kimlikId:norm(data.kimlikId),raw:data
    };
  }
  function selectExtOperation(operationNo,listIndex){
    const source=extListSource();if(!source)return false;
    let record=Number.isInteger(listIndex)?source.store.getAt?.(listIndex):null;
    if(!record)source.store.each?.(r=>{if(!record&&Object.values(r.data||{}).some(v=>norm(v)===operationNo))record=r;});
    if(!record)return false;
    try{
      const view=source.grid.getView?.(),node=view?.getNode?.(record);
      source.grid.getSelectionModel?.().select(record);view?.focusRow?.(record);
      if(node)click(node,true);
      const event=new MouseEvent('dblclick',{bubbles:true,cancelable:true,detail:2,view:record?.store?.proxy?.reader?.rawData?.defaultView||window});
      view?.fireEvent?.('itemdblclick',view,record,node,listIndex,event);
      source.grid.fireEvent?.('itemdblclick',view,record,node,listIndex,event);
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
    return fieldValues(label)[0]||'';
  }
  function fieldValues(label){
    const key=upper(label);
    const found=[];
    for(const doc of allDocs()){
      for(const box of doc.querySelectorAll('.x-form-item,.x-field')){
        if(!visible(box)||!upper(box.innerText).includes(key))continue;
        for(const input of box.querySelectorAll('input,textarea,.x-form-display-field')){
          const value=norm(input.value??input.innerText);if(value)found.push(value);
        }
      }
    }
    return uniq(found);
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
  async function readRadiology(){
    const btn=exactButton('Radyoloji');if(!btn)return[];click(btn);await sleep(260);
    const texts=[...collectVisibleText()];
    const rows=elements('tr[data-recordindex],[role="row"]').filter(r=>visible(r)&&/ABDOM|BATIN|TOMOGRAF|\bBT\b|ULTRASON|\bUSG\b|MRG|MANYETİK/i.test(norm(r.innerText))).slice(0,6);
    for(const row of rows){
      click(row);await sleep(100);texts.push(...collectVisibleText());
      const win=elements('.x-window').find(w=>visible(w)&&/RADYOLOJ|RAPOR|GÖRÜNTÜ/i.test(norm(w.innerText).slice(0,250)));
      if(win){texts.push(norm(win.innerText));closeWindow(win);await sleep(50);}
    }
    return uniq(texts);
  }
  async function readHistory(tc){
    const button=exactButton('Hasta Geçmişi'); if(!button)throw new Error('Hasta Geçmişi düğmesi bulunamadı');
    click(button);
    const win=await waitFor(()=>elements('.x-window').find(w=>visible(w)&&/Hasta Geçmiş/i.test(norm(w.innerText).slice(0,200))),10000);
    if(!win)throw new Error('Hasta geçmişi açılmadı');
    try{
      await sleep(250);
      return uniq([...win.querySelectorAll('tr[data-recordindex],[role="row"]')].filter(visible).map(x=>norm(x.innerText)).filter(x=>x.includes(tc)||/\d{2}\.\d{2}\.\d{4}/.test(x)));
    }finally{closeWindow(win);await sleep(80);}
  }
  function quantityFromMaterialRow(text){
    const labelled=norm(text).match(/Miktar\s*:\s*(\d+(?:[.,]\d+)?)/i);
    if(labelled)return Number(labelled[1].replace(',','.'))||0;
    const cells=norm(text).split(/\s+/); const nums=cells.map(x=>Number(String(x).replace(',','.'))).filter(x=>Number.isFinite(x)&&x>0&&x<=10&&Number.isInteger(x));
    return nums.length?nums[nums.length-1]:1;
  }
  function classifyEhsLocation(imagingText,noteText){
    const fold=text=>upper(text).normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/İ/g,'I');
    const sentences=fold([...(imagingText||[]),noteText||''].join(' | ')).split(/[|\n]|(?<=[.!?])\s+/).filter(s=>/HERNI|DEFEKT|FASYA|FITIK/.test(s)&&!/HERNI\s+(?:YOK|IZLENMEDI|SAPTANMADI)/.test(s));
    const codes=new Set();
    for(const s of sentences){
      if(/SUBKSIFOID|SUBXIFOID|SUBXIPHOID/.test(s))codes.add('M1');
      if(/EPIGASTR|SUPRAUMB|UST ORTA HAT/.test(s))codes.add('M2');
      if(/(?:^|[^A-Z])(?:PERI|PARA)?UMB(?:I)?LIKAL|GOBEK CEVRE|GOBEK HIZA/.test(s))codes.add('M3');
      if(/INFRAUMB|ALT ORTA HAT/.test(s))codes.add('M4');
      if(/SUPRAPUB/.test(s))codes.add('M5');
      const above=s.match(/GOBE(?:K|GI)(?:N)?\s+(\d+(?:[.,]\d+)?)\s*CM\s+(?:UST|YUKARI)/);
      const below=s.match(/GOBE(?:K|GI)(?:N)?\s+(\d+(?:[.,]\d+)?)\s*CM\s+(?:ALT|ASAGI)/);
      if(above)codes.add(Number(above[1].replace(',','.'))<=3?'M3':'M2');
      if(below)codes.add(Number(below[1].replace(',','.'))<=3?'M3':'M4');
    }
    if(codes.size)return [...codes].sort().join(', ');
    return '';
  }
  function derive(patient, details){
    const surgeryDate=parseDateTime(patient.surgeryDate)||parseDateTime(details.selectedOperation);
    const note=details.note.join(' | '), allHistory=details.history.join(' | '), all=upper(`${note} ${allHistory} ${(details.stay||[]).join(' | ')}`);
    const opRows=uniq(details.surgeries||[]).map(text=>({text,date:parseDateTime(text)})).filter(x=>x.date);
    const previousHernia=opRows.filter(x=>surgeryDate&&x.date<surgeryDate&&/603801|\bK43(?:\.|-)|İNSİZYONEL HERNİ|VENTRAL HERNİ/i.test(x.text)).map(x=>x.text);
    const laterHerniaOps=opRows.filter(x=>surgeryDate&&x.date>surgeryDate&&/603801|İNSİZYONEL HERNİ|VENTRAL HERNİ/i.test(x.text));
    const laterDebridement=opRows.filter(x=>surgeryDate&&x.date>surgeryDate&&/YARA[^|]{0,40}DEBRİDMAN|YARA[^|]{0,40}DEBRIDMAN|DEBRİDMAN|DEBRIDMAN/i.test(x.text));
    const laterAdmissions=details.history.filter(x=>{const d=parseDate(x);return d&&surgeryDate&&dateDiffDays(surgeryDate,d)>2&&/\bYATIŞ\b/i.test(x);});
    const previousAbdominal=opRows.filter(x=>surgeryDate&&x.date<surgeryDate&&/ABDOM|LAPAROT|LAPAROSK|APPEN|KOLESİST|KOLEKT|REZEKS|GASTREK|HERNİ|FITIK|SEZARYEN|HİSTEREKT|OOFOREKT|SALPEN|KOLON|REKT|İLEOST|KOLOST|PANKREAT|SPLENEKT|BARSAK|BAĞIRSAK|UMBİLİK|MİDE|BYPASS|SLEEVE/i.test(x.text));
    const malignancyPattern=/\bC\d{2}(?:\.\d+)?\b|MALIGN|KANSER|KARSINOM|ADENOKARSINOM|METASTA|610410|REKTUM TUMORUNDE LOW ANTERIOR REZEKSIYON/;
    const cancers=uniq([...details.history,...details.note,...(details.surgeries||[])].filter(x=>{
      const folded=upper(x).normalize('NFD').replace(/[\u0300-\u036f]/g,'');
      return malignancyPattern.test(folded)&&!/MALIGN(?:ITE)?\s+(?:YOK|IZLENMEDI|SAPTANMADI|NEGATIF)|MALIGNITE ACISINDAN SUPHE/.test(folded);
    }));
    const herniaOpRows=opRows.filter(x=>/603801|\bK43(?:\.|-)|İNSİZYONEL HERNİ|INSIZYONEL HERNI|VENTRAL HERNİ|VENTRAL HERNI/i.test(x.text));
    const herniaOperationDates=uniq([
      ...(surgeryDate?[dateText(surgeryDate)]:[]),
      ...herniaOpRows.map(x=>dateText(x.date))
    ]);
    const opCount=Math.max(1,herniaOperationDates.length);
    const materialRows=details.materials.filter(x=>/MESH|MEŞ|CERRAHİ YAMA|HERNİ YAMASI|YAMA KOMPOZİT|PROLEN(?:E)?\s+(?:MESH|MEŞ|YAMA)/i.test(x)&&!/SÜTÜR|SUTUR|İĞNE|IGNE|YUVARLAK|ÖRGÜLÜ|EMİLEBİLEN|EMILEBILEN/i.test(x));
    const prolenRows=materialRows.filter(x=>/PROLEN|PROLENE/i.test(x));
    const prolenCount=prolenRows.reduce((sum,x)=>sum+quantityFromMaterialRow(x),0);
    const meshCount=materialRows.reduce((sum,x)=>sum+quantityFromMaterialRow(x),0);
    const ageSex=details.fields.ageSex||''; const sex=details.fields.gender||/\((Kadın|Erkek)\)/i.exec(ageSex)?.[1]||''; const age=/^(\d+)/.exec(ageSex)?.[1]||'';
    const times=(details.fields.surgeryTimes||'').match(/\d{1,2}:\d{2}/g)||[]; let duration='';
    if(times.length===2){const [h1,m1]=times[0].split(':').map(Number),[h2,m2]=times[1].split(':').map(Number);duration=(h2*60+m2)-(h1*60+m1);if(duration<0)duration+=1440;if(duration<=0||duration>720)duration='';}
    const diagnoses=upper(allHistory);
    const complications=[]; if(/NEKROZ/.test(all))complications.push('Nekroz');if(/SEROMA/.test(all))complications.push('Seroma');if(/YARA ENFEKSİY|CERRAHİ ALAN ENFEKSİY|ENFEKSİYON/.test(all))complications.push('Enfeksiyon');if(/DEHİS|EVİSSER/.test(all))complications.push('Dehisens');
    const imagingAndNote=`${(details.imaging||[]).join(' | ')} | ${note}`;
    const focusedDefect=imagingAndNote.match(/(?:DEF[EİI]KT|HERNİ|HERNI|HERNIA|HERNİ KESESİ|HERNI KESESI|HERNIA SAC|FASYA)[^|.]{0,180}?(\d+(?:[.,]\d+)?)\s*(?:x|×|\*)\s*(\d+(?:[.,]\d+)?)\s*(?:cm|mm)/i)||imagingAndNote.match(/(\d+(?:[.,]\d+)?)\s*(?:x|×|\*)\s*(\d+(?:[.,]\d+)?)\s*(?:cm|mm)[^|.]{0,140}?(?:DEF[EİI]KT|HERNİ|HERNI|HERNIA|HERNİ KESESİ|HERNI KESESI|HERNIA SAC|FASYA)/i);
    const defectUnit=focusedDefect?.[0]?.match(/(cm|mm)\b/i)?.[1]?.toLowerCase()||'';
    const sacOnly=focusedDefect&&/HERN[İI]\s+KESE|HERNIA SAC/i.test(focusedDefect[0])&&!/DEFEKT|FASYA/i.test(focusedDefect[0]);
    const reportMeasurements=extractHerniaMeasurements(details.imaging||[]);
    const defect=reportMeasurements.length?reportMeasurements.join('; '):focusedDefect?`${focusedDefect[1]} x ${focusedDefect[2]}${defectUnit?` ${defectUnit}`:''}${sacOnly?' (herni kesesi; defekt ölçüsü belirtilmemiş)':''}`:'';
    const location=classifyEhsLocation(details.imaging,note);
    const dischargeDates=(details.dischargeFields||[]).map(parseDateTime).filter(x=>x&&surgeryDate&&x>=surgeryDate).sort((a,b)=>a-b);
    const discharge=dischargeDates[0]||null;
    const stayDays=discharge&&surgeryDate?Math.max(0,Math.round(((discharge-surgeryDate)/86400000)*10)/10):'';
    const sameDayHistoryOps=opRows.filter(x=>surgeryDate&&x.date.toDateString()===surgeryDate.toDateString()&&!norm(x.text).includes(norm(patient.operationNo))&&/\b\d{6}\s*-/i.test(x.text));
    const extraMacros=['makro2koduAdi','makro3koduAdi','makro4koduAdi','makro5koduAdi'].map(k=>norm(patient.raw?.[k])).filter(Boolean);
    const combinedOps=uniq([...extraMacros,...sameDayHistoryOps.map(x=>norm(x.text))]);
    const heightRaw=Number(String(details.fields.height||'').replace(',','.').match(/\d+(?:\.\d+)?/)?.[0]||0);
    const weight=Number(String(details.fields.weight||'').replace(',','.').match(/\d+(?:\.\d+)?/)?.[0]||0);
    const height=heightRaw>3?heightRaw/100:heightRaw;
    const bmi=height>=1&&height<=2.5&&weight>=20&&weight<=400?Math.round((weight/(height*height))*10)/10:'';
    const deathDates=uniq([details.fields.deathDate,...details.history.filter(x=>/ÖLÜM|OLUM|VEFAT|EXİTUS|EKSİTUS/i.test(x))]).map(x=>parseDateTime(x)).filter(Boolean);
    const mortality=deathDates.some(d=>surgeryDate&&d>=surgeryDate&&dateDiffDays(surgeryDate,d)<=60);
    return {surgeryDate,previousHernia,laterHerniaOps,laterDebridement,laterAdmissions,previousAbdominal,cancers,opCount,materialRows,prolenCount,meshCount,sex,age,duration,diagnoses,all,note,complications,defect,location,stayDays,combinedOps,bmi,mortality};
  }
  function applyResult(patient, details){
    const row=state.rows[patient.rowIndex], d=derive(patient,details);
    setIfFound(row,'operationNo',details.fields.operationNo);
    setIfFound(row,'phone',details.fields.phone);
    setIfFound(row,'sex',d.sex); setIfFound(row,'age',d.age?Number(d.age):''); setIfFound(row,'asa',details.fields.asa);
    setIfFound(row,'duration',d.duration||''); setIfFound(row,'followup',d.surgeryDate?dateDiffHuman(d.surgeryDate):'');
    setIfFound(row,'defect',d.defect);setIfFound(row,'location',d.location);setIfFound(row,'stay',d.stayDays);
    setIfFound(row,'recurrence',d.opCount>1?1:0);setIfFound(row,'opCount',d.opCount);
    setIfFound(row,'followRecurrence',d.laterHerniaOps.length?`Evet – ${d.laterHerniaOps.map(x=>`${dateText(x.date)} ${norm(x.text)}`).join('; ')}`:'Hayır');
    if(d.previousAbdominal.length)setIfFound(row,'preop',d.previousAbdominal.map(x=>`${dateText(x.date)} ${norm(x.text)}`).join('; '));
    setIfFound(row,'readmission',d.laterAdmissions.length?d.laterAdmissions.map(x=>norm(x)).join('; '):'YOK');
    if(d.cancers.length){setIfFound(row,'malign',1);setIfFound(row,'benign',0);row[state.headerMap.get('MALİGNİTE KAYNAĞI')]=d.cancers.join('; ');}
    if(d.meshCount>=2){setIfFound(row,'onlay',1);setIfFound(row,'inlay',0);}
    else if(d.meshCount===1){setIfFound(row,'inlay',1);setIfFound(row,'onlay',0);}
    const hasAbd=/ABDOMİNOPLAST|PANNİKÜLEKT/i.test(d.note);
    if(/\bONLAY\b/i.test(d.note)){setIfFound(row,'onlay',1);if(hasAbd)setIfFound(row,'onlayAbd',1);}
    if(/\bSUBLAY\b/i.test(d.note)){setIfFound(row,'sublay',1);if(hasAbd)setIfFound(row,'sublayAbd',1);}
    if(/\bINLAY\b|\bİNLAY\b/i.test(d.note)){setIfFound(row,'inlay',1);if(hasAbd)setIfFound(row,'inlayAbd',1);}
    if(d.complications.length)setIfFound(row,'sso',d.complications.join(', '));
    if(/NEKROZ/.test(d.all))setIfFound(row,'necrosis',1); if(/\bVAC\b/.test(d.all))setIfFound(row,'vac',1); if(/SEROMA/.test(d.all))setIfFound(row,'seroma',1);
    if(d.laterDebridement.length){setIfFound(row,'necrosis',1);setIfFound(row,'vac',1);}
    if(/REVİZYON|REVIZYON|REAKSPLORASYON/.test(d.all))setIfFound(row,'revision',1);
    if(d.laterHerniaOps.length)setIfFound(row,'revision',1);
    if(/\bI1[0-5](?:\.|-)|HİPERTANSİYON|HIPERTANSIYON|\bHT\b/.test(d.all))setIfFound(row,'ht',1);
    if(/\bE1[0-4](?:\.|-)|DİYABET|DIYABET|DIABETES|\bDM\b/.test(d.all))setIfFound(row,'dm',1);
    if(/\bJ4[34](?:\.|-)|\bKOAH\b|KRONİK OBSTRÜKTİF|KRONIK OBSTRUKTIF|AMFİZEM|AMFIZEM/.test(d.all))setIfFound(row,'copd',1);
    if(/\bI50(?:\.|-)|KALP YETMEZLİĞİ|KALP YETMEZLIGI|KARDİYAK YETMEZLİK|KARDIYAK YETMEZLIK/.test(d.all))setIfFound(row,'hf',1);
    if(/\bE0[4-5](?:\.|-)|GUATR|MULTİNODÜLER|MULTINODULER|TİROİD NODÜL|TIROID NODUL/.test(d.all))setIfFound(row,'goiter',1);
    if(/SİGARA.*(İÇİYOR|KULLANIYOR|AKTİF)|AKTİF SİGARA/.test(d.all))setIfFound(row,'smoking',1);
    if(d.combinedOps.length)setIfFound(row,'combined',d.combinedOps.join('; '));
    if(d.mortality)setIfFound(row,'mortality',1);
    if(d.bmi)setIfFound(row,'bmi',d.bmi);
    if(d.materialRows.length)setIfFound(row,'graft',uniq(d.materialRows).join('; '));
    if(/DREN/.test(d.note)){const m=d.note.match(/DREN[^|]{0,60}?(\d+)\s*GÜN/i);if(m)setIfFound(row,'drain',`${m[1]} GÜN`);}
    row[state.headerMap.get('PROLEN MESH ADEDİ')] = d.prolenCount||0;
    row[state.headerMap.get('MALZEME KAYDI')] = d.materialRows.length?uniq(d.materialRows).join('; '):'FONET sarf kaydında mesh/yama bulunmadı';
    row[state.headerMap.get('FONET TARAMA DURUMU')] = 'Tamamlandı';
    if(details.radiologyAudit){
      const a=details.radiologyAudit;
      row[state.headerMap.get('RADYOLOJİ OKUMA')]=`${a.read}/${a.total} rapor okundu; raporsuz tetkik: ${a.noReport||0}; iptal/onaysız: ${a.excluded||0}; hata: ${a.failures.length}`;
      row[state.headerMap.get('RADYOLOJİ KAYNAKLARI')]=details.imaging.join('\n').slice(0,32000);
      row[state.headerMap.get('FONET TARAMA DURUMU')]=a.failures.length?`Eksik: ${a.failures.length} radyoloji raporu okunamadı`:'Tamamlandı';
    }
    return d;
  }

  function apiBase(){return `${location.origin}/hbys-rs/hbys`;}
  async function apiJson(path){
    if(state.stopped)throw new Error('Tarama durduruldu');
    const sep=path.includes('?')?'&':'?';
    const controller=new AbortController();
    activeRequests.add(controller);
    const timer=setTimeout(()=>controller.abort(),15000);
    try{
    const response=await fetch(`${apiBase()}${path}${sep}_dc=${Date.now()}`,{signal:controller.signal,credentials:'include',headers:{Accept:'application/json, text/plain, */*'}});
    const body=await response.text();
    if(!response.ok)throw new Error(`HTTP ${response.status}`);
    try{return JSON.parse(body);}catch{throw new Error('FONET yanıtı okunamadı');}
    }catch(error){if(error.name==='AbortError')throw new Error('İstek 15 saniyede yanıt vermedi');throw error;}
    finally{clearTimeout(timer);activeRequests.delete(controller);}
  }
  function payloadRows(payload){
    if(Array.isArray(payload))return payload;
    for(const key of ['data','list','rows','result','content'])if(Array.isArray(payload?.[key]))return payload[key];
    return payload?.data&&typeof payload.data==='object'?[payload.data]:(payload&&typeof payload==='object'?[payload]:[]);
  }
  function objectText(value,depth=0,seen=new Set()){
    if(value==null||depth>7)return'';
    if(typeof value!=='object')return cleanText(value);
    if(seen.has(value))return'';seen.add(value);
    return uniq(Object.values(value).map(x=>objectText(x,depth+1,seen))).join(' | ');
  }
  function flatObject(value,prefix='',depth=0,out={}){
    if(!value||typeof value!=='object'||depth>8)return out;
    for(const [key,item] of Object.entries(value)){
      const path=prefix?`${prefix}.${key}`:key;
      if(item&&typeof item==='object')flatObject(item,path,depth+1,out);else if(item!=null&&norm(item))out[path]=item;
    }
    return out;
  }
  function deepValue(value,names){
    const wanted=names.map(x=>String(x).replace(/[^a-z0-9]/gi,'').toLowerCase());
    for(const [path,item] of Object.entries(flatObject(value))){
      const key=path.split('.').pop().replace(/[^a-z0-9]/gi,'').toLowerCase();
      if(wanted.includes(key)&&norm(item))return norm(item);
    }
    return'';
  }
  function deepPathValue(value,pathPattern){
    for(const [path,item] of Object.entries(flatObject(value)))if(pathPattern.test(path)&&norm(item))return cleanText(item);
    return'';
  }
  function deepValues(value,names){
    const wanted=names.map(x=>String(x).replace(/[^a-z0-9]/gi,'').toLowerCase());
    return Object.entries(flatObject(value)).filter(([path,item])=>{
      const key=path.split('.').pop().replace(/[^a-z0-9]/gi,'').toLowerCase();
      return wanted.includes(key)&&norm(item);
    }).map(([,item])=>norm(item));
  }
  function asaDisplay(detailRoot){
    const values=deepValues(detailRoot,['asaAdi','asaSkoruAdi','asaSinifi','asaSınıfı','asaEuroScore']);
    for(const value of values){
      const roman=upper(value).match(/(?:ASA\s*)?\b(V|IV|III|II|I)\b/)?.[1];
      if(roman)return roman;
    }
    const raw=deepValue(detailRoot,['asaSkoru','asa']);
    const roman=upper(raw).match(/(?:ASA\s*)?\b(V|IV|III|II|I)\b/)?.[1];
    if(roman)return roman;
    const numeric=Number(raw);
    return Number.isInteger(numeric)&&numeric>=0&&numeric<=4?['I','II','III','IV','V'][numeric]:'';
  }
  function surgeryTimePair(detailRoot){
    const start=deepValue(detailRoot,['ameliyatBaslangicSaati','ameliyatBaslangicSaat','ameliyatBasSaati','ameliyatBaslamaSaati','ameliyatBaslangicTarihi','ameliyatBaslamaTarihi']);
    const end=deepValue(detailRoot,['ameliyatBitisSaati','ameliyatBitisSaat','ameliyatSonSaati','ameliyatTamamlanmaSaati','ameliyatBitisTarihi','ameliyatTamamlanmaTarihi']);
    const time=x=>{const all=norm(x).match(/\b\d{1,2}:\d{2}\b/g)||[];return all.at(-1)||'';};
    if(time(start)&&time(end))return`${time(start)} ${time(end)}`;
    const entries=Object.entries(flatObject(detailRoot));
    const byPath=rx=>time(entries.find(([path,value])=>rx.test(path.replace(/[^a-z0-9]/gi,'').toLowerCase())&&time(value))?.[1]);
    const pathStart=byPath(/ameliyat.*(?:baslangic|baslama|bassaati|start)/);
    const pathEnd=byPath(/ameliyat.*(?:bitis|tamamlanma|sonsaati|end)/);
    if(pathStart&&pathEnd)return`${pathStart} ${pathEnd}`;
    const operationTimes=entries.filter(([path])=>/ameliyat.*(?:saat|time)/i.test(path.replace(/[^a-z0-9]/gi,''))).flatMap(([,value])=>norm(value).match(/\b\d{1,2}:\d{2}\b/g)||[]);
    if(operationTimes.length>=2)return operationTimes.slice(-2).join(' ');
    const ordered=entries.filter(([path])=>/saatbilgisi/i.test(path.replace(/[^a-z0-9]/gi,''))).flatMap(([,value])=>norm(value).match(/\b\d{1,2}:\d{2}\b/g)||[]);
    return ordered.length>=2?ordered.slice(-2).join(' '):'';
  }
  function materialPath(birimSevkId){
    const filter=encodeURIComponent(JSON.stringify([{property:'birimSevk.id',value:Number(birimSevkId)||birimSevkId,filterType:'kriterPanel',type:'Long',operator:'='}]));
    return`/Stok/StokCikisServisDepoHasta/getHastaCikisKayitList?start=0&limit=2000&page=1&filter=${filter}`;
  }
  function materialRecordText(record){
    const flat=flatObject(record),entries=Object.entries(flat);
    const named=(rx)=>norm(entries.find(([path,value])=>rx.test(path)&&norm(value))?.[1]);
    const name=named(/(?:^|\.)(?:malzeme|stok)(?:Adi|Adı|\.adi|\.adı|\.ad)$/i)||named(/(?:^|\.)(?:malzemeAdi|malzemeAdı|stokAdi|stokAdı|adi|adı)$/i);
    const code=named(/(?:^|\.)(?:malzeme|stok)(?:Kodu|\.kodu)$/i)||named(/(?:^|\.)(?:malzemeKodu|stokKodu|kodu)$/i);
    const quantity=named(/(?:^|\.)(?:miktar|adet)$/i);
    const unit=named(/(?:^|\.)(?:birimAdi|birimAdı|olcuBirimAdi|ölçüBirimAdı|birim\.adi)$/i);
    const date=named(/(?:^|\.)(?:tarih|islemTarihi|işlemTarihi)$/i);
    const compact=norm(`${code} ${name} | Miktar: ${quantity} ${unit} | ${date}`);
    return /MESH|MEŞ|CERRAHİ YAMA|HERNİ YAMASI|YAMA KOMPOZİT|PROLEN|PROLENE/i.test(compact)?compact:objectText(record);
  }
  async function patientHistory(patient){
    if(!patient.kimlikId)return[];
    const filter=encodeURIComponent(JSON.stringify([{property:'hastaGelis.hasta.id',value:Number(patient.kimlikId)||patient.kimlikId,filterType:'kriterPanel',type:'Long',operator:'='}]));
    const payload=await settled(`/Tibbi/HastaBirimSevk/getKayitList?start=0&limit=2000&page=1&filter=${filter}`);
    return payload.__error?[]:payloadRows(payload).map(x=>objectText(x));
  }
  function extStoreDescriptor(titlePattern){
    for(const doc of allDocs()){
      const Ext=doc.defaultView.Ext;if(!Ext?.ComponentQuery)continue;
      let grids=[];try{grids=Ext.ComponentQuery.query('gridpanel');}catch{}
      for(const grid of grids){
        const store=grid.getStore?.(),proxy=store?.getProxy?.();
        const label=norm(`${grid.title||''} ${grid.itemId||''} ${grid.id||''}`);
        if(titlePattern.test(label)&&norm(proxy?.url||proxy?.api?.read))return{url:norm(proxy.url||proxy.api.read),extra:{...(proxy.extraParams||{})}};
      }
    }
    return null;
  }
  function descriptorPath(descriptor,patient){
    if(!descriptor?.url)return'';
    const url=new URL(descriptor.url,location.origin);
    let patientScoped=false;
    for(const [key,raw] of Object.entries(descriptor.extra||{})){
      let value=raw;
      if(/hastaGelis/i.test(key)){value=patient.gelisId;patientScoped=true;}
      else if(/birimSevk/i.test(key)){value=patient.birimSevkId;patientScoped=true;}
      else if(/hasta.*(?:kimlik|id)/i.test(key)&&patient.kimlikId){value=patient.kimlikId;patientScoped=true;}
      else if(key==='filter'){
        try{
          const filters=typeof raw==='string'?JSON.parse(raw):raw;
          for(const item of filters||[]){
            const property=item.property||'';
            if(/hastaGelis/i.test(property)){item.value=Number(patient.gelisId)||patient.gelisId;patientScoped=true;}
            else if(/birimSevk/i.test(property)){item.value=Number(patient.birimSevkId)||patient.birimSevkId;patientScoped=true;}
            else if(/hasta.*(?:kimlik|id)/i.test(property)&&patient.kimlikId){item.value=Number(patient.kimlikId)||patient.kimlikId;patientScoped=true;}
          }
          value=JSON.stringify(filters);
        }catch{}
      }
      url.searchParams.set(key,String(value));
    }
    url.searchParams.set('start','0');url.searchParams.set('page','1');url.searchParams.set('limit','1000');
    return patientScoped?`${url.pathname.replace(/^\/hbys-rs\/hbys/i,'')}${url.search}`:'';
  }
  // FONET risSonucForm: HASTA -> hastaId; report body -> getRisRaporSonucByRaporId.
  const radiologyCache=new Map();
  async function radiologyHistory(patient,hastaId){
    if(!hastaId)throw new Error('Radyoloji için hasta dosya kimliği doğrulanamadı');
    const cacheKey=String(hastaId);
    if(radiologyCache.has(cacheKey))return radiologyCache.get(cacheKey);
    const task=(async()=>{
      const records=[],reports=[],failures=[];let expected=null,noReport=0,excluded=0;
      const filter=encodeURIComponent(JSON.stringify([{property:'hastaId',value:Number(hastaId),type:'Long',operator:'='}]));
      for(let start=0;start<10000;start+=500){
        if(state.stopped)throw new Error('Tarama durduruldu');
        const payload=await apiJson('/Ris/RisHizmetSonuc/getRisHizmetSonucInfoList?start='+start+'&limit=500&page='+(start/500+1)+'&filter='+filter);
        if(payload.success===false)throw new Error('Radyoloji listesi alınamadı');
        const rows=Array.isArray(payload.data)?payload.data:Array.isArray(payload.rows)?payload.rows:null;
        if(!rows)throw new Error('Radyoloji liste yanıtının yapısı doğrulanamadı');
        const total=payload.totalCount??payload.total;
        if(total!=null)expected=Number(total);
        records.push(...rows);
        if(expected!=null&&records.length>=expected)break;
        if(rows.length<500)break;
      }
      if(expected!=null&&records.length<expected)failures.push('Radyoloji listesi eksik: '+records.length+'/'+expected);
      if(records.length>=10000&&expected==null)failures.push('Radyoloji liste sınırına ulaşıldı; listenin tamamı doğrulanamadı');
      const uniqueReports=new Map();
      for(const record of records){
        if(!record.raporId){noReport++;continue;}
        uniqueReports.set(String(record.raporId),record);
      }
      let index=0;
      for(const [id,record] of uniqueReports){
        if(state.stopped)throw new Error('Tarama durduruldu');
        while(state.paused&&!state.stopped)await sleep(250);
        log(patient.name+': arka plan rapor '+(++index)+'/'+uniqueReports.size);
        let result;
        for(let attempt=0;attempt<2;attempt++){
          try{
            const payload=await apiJson('/Ris/RisHizmetSonuc/getRisRaporSonucByRaporId/'+encodeURIComponent(id));
            if(payload.success===false||!payload.data)throw new Error('Rapor yanıtı alınamadı');
            const data=payload.data;
            const text=uniq([data.raporTextByRapor,data.bulgular,data.cekimTeknigi,data.karsilastirma].map(cleanText).filter(Boolean)).join(' | ');
            if(/İsteyen Branş Görebilir|İsteyen Dr. Görebilir/i.test(text))throw new Error('Rapora erişim yetkisi yok');
            if(data.durum!=null&&Number(data.durum)!==1){excluded++;result=true;break;}
            if(!data.onayTarihi&&!data.asistanOnayTarihi){excluded++;result=true;break;}
            if(!text)throw new Error('Onaylı raporun metni boş');
            reports.push({source:'Rapor '+id+' | '+objectText(record)+' | '+(data.onayTarihi||data.asistanOnayTarihi),text});
            result=true;break;
          }catch(error){if(attempt===1)failures.push('Rapor '+id+': '+error.message);}
        }
      }
      return{reports,audit:{total:uniqueReports.size,read:reports.length,noReport,excluded,failures,examTotal:records.length}};
    })();
    radiologyCache.set(cacheKey,task);
    try{const result=await task;if(result.audit.failures.length)radiologyCache.delete(cacheKey);return result;}
    catch(error){radiologyCache.delete(cacheKey);throw error;}
  }
  async function settled(path){try{return await apiJson(path);}catch(error){return{__error:String(error.message||error)};}}
  async function operationHistory(patient){
    if(!patient.kimlikId)return[];
    const filter=encodeURIComponent(JSON.stringify([{index:1,property:'kimlikId',value:Number(patient.kimlikId)||patient.kimlikId,filterType:'kriterPanel',type:'Long',operator:'='}]));
    const payload=await settled(`/Ameliyat/Ameliyat/getKayitList?start=0&limit=1000&page=1&filter=${filter}`);
    return payload.__error?[]:payloadRows(payload).map(x=>objectText(x));
  }
  async function scanPatientBackground(patient){
    if(!patient.ameliyatId||!patient.gelisId||!patient.birimSevkId)throw new Error('Kayıt servis kimlikleri bulunamadı; listeyi yeniden alın');
    const detailPayload=await settled(`/Ameliyat/Ameliyat/getKayit/${encodeURIComponent(patient.birimSevkId)}`);
    const detailRoot=detailPayload.__error?{}:(detailPayload.data||detailPayload);
    const visit=detailRoot.birimSevk?.hastaGelis;
    const matchedPatient=visit?.hasta;
    const identity=matchedPatient?.kimlik;
    if(!identity||!matchedPatient.id||String(visit.id)!==String(patient.gelisId)||String(detailRoot.birimSevk.id)!==String(patient.birimSevkId))throw new Error('Hasta ve geliş kimliği servis yanıtıyla eşleşmedi');
    const verifiedName=norm(identity.adiSoyadi||[identity.adi,identity.soyadi].filter(Boolean).join(' '));
    if(verifiedName&&upper(verifiedName)!==upper(patient.name))throw new Error('Hasta adı servis yanıtıyla eşleşmedi; başka hastaya yazılmadı');
    const materialIds=uniq([
      patient.birimSevkId,patient.isteyenBirimSevkId,
      deepValue(detailRoot,['birimSevkId']),deepValue(detailRoot,['isteyenBirimSevkId']),
      ...Object.entries(flatObject(detailRoot)).filter(([path])=>/(?:^|\.)(?:birimSevk|ustBirimSevk)\.id$/i.test(path)).map(([,value])=>value)
    ]);
    const [notePayload,servicePayload,patientPayload,consultPayload,materialPayloads,history,visitHistory,radiology]=await Promise.all([
      settled(`/Ameliyat/Ameliyat/getAmeliyatPersonelList/${encodeURIComponent(patient.ameliyatId)}/-1`),
      settled(`/Tibbi/HastaHizmet/getHizmetList/${encodeURIComponent(patient.birimSevkId)}/${encodeURIComponent(patient.gelisId)}`),
      settled(`/Tibbi/HastaBirimSevk/getSevkUyariInfo/${encodeURIComponent(patient.birimSevkId)}`),
      settled(`/Poliklinik/Poliklinik/getHastaGelisKonsultasyonList/${encodeURIComponent(patient.gelisId)}/1`),
      Promise.all(materialIds.map(id=>settled(materialPath(id)))),
      operationHistory(patient),
      patientHistory(patient),
      Promise.resolve([]) // Raporlar aşağıda aktif hastanın Tüm Gelişler ekranından okunur.
    ]);
    const failed=[detailPayload,notePayload,patientPayload].filter(x=>x?.__error).length;
    if(failed===3)throw new Error('FONET arka plan servisleri yanıt vermedi');
    const noteRows=notePayload.__error?[]:payloadRows(notePayload);
    const serviceRows=servicePayload.__error?[]:payloadRows(servicePayload);
    const consultRows=consultPayload.__error?[]:payloadRows(consultPayload);
    const materialRecords=materialPayloads.flatMap(payload=>payload.__error?[]:payloadRows(payload));
    const patientRoot=patientPayload.__error?{}:patientPayload;
    const note=noteRows.map(x=>norm(x?.notu||x?.hizmet||objectText(x))).filter(Boolean);
    const services=serviceRows.map(x=>objectText(x));
    const consultations=consultRows.map(x=>objectText(x));
    const patientText=objectText(patientRoot);
    const combinedRoot={identity,detail:detailRoot,patient:patientRoot};
    const ageSexRaw=deepValue(combinedRoot,['yasCinsiyetDogumTarihi','yasCinsiyet'])||deepPathValue(combinedRoot,/(?:yas|yaş).*cinsiyet|cinsiyet.*(?:dogum|doğum)/i);
    const genderRaw=deepValue(combinedRoot,['cinsiyetAdi','cinsiyetAd','cinsiyetKodu','cinsiyet'])||deepPathValue(combinedRoot,/cinsiyet.*(?:\.adi|\.adı|\.ad|aciklama|açıklama|text|kodu)$/i)||ageSexRaw;
    const gender=Number(identity.cinsiyet)===2?'Kadın':Number(identity.cinsiyet)===1?'Erkek':/KADIN|\bK\b/i.test(genderRaw)?'Kadın':/ERKEK|\bE\b/i.test(genderRaw)?'Erkek':'';
    const birthDate=deepValue(combinedRoot,['dogumTarihi','doğumTarihi'])||deepPathValue(combinedRoot,/(?:dogum|doğum).*tarih/i);
    const birth=parseDateTime(birthDate),calculatedAge=birth?Math.max(0,new Date().getFullYear()-birth.getFullYear()-(new Date()<new Date(new Date().getFullYear(),birth.getMonth(),birth.getDate())?1:0)):'';
    const ageSex=ageSexRaw||deepValue(combinedRoot,['yas'])||(gender?`${calculatedAge||''}Yıl (${gender}) / ${birthDate}`:'')||patientText.match(/\d+\s*(?:Yıl|Yaş)[^|]{0,40}\((?:Kadın|Erkek)\)/i)?.[0]||'';
    const fields={
      operationNo:patient.operationNo,
      tc:norm(identity.tcKimlikNo||identity.kimlikNo||identity.tckn),
      name:patient.name||deepValue(combinedRoot,['hastaAdiSoyadi','hastaAdSoyad']),
      phone:deepValue(combinedRoot,['telefonGsm','cepTelefonu','cepTelefon','telefonNo','telefon','gsm','mobilTelefon']),
      ageSex,
      gender,
      asa:asaDisplay(detailRoot),
      surgeryTimes:surgeryTimePair(detailRoot),
      height:deepValue(combinedRoot,['boy','boyCm','vucutBoyu','uzunluk']),
      weight:deepValue(combinedRoot,['kilo','agirlik','ağırlık','vucutAgirligi']),
      deathDate:deepValue(combinedRoot,['olumTarihi','ölümTarihi','vefatTarihi'])
    };
    const dischargeFields=Object.entries(flatObject(patientRoot)).filter(([k])=>/taburcu|cikis|bitiş|bitis/i.test(k)).map(([,v])=>norm(v));
    let rad;
    try{rad=await radiologyHistory(patient,matchedPatient.id);}
    catch(error){rad={reports:[],audit:{total:0,read:0,failures:[error.message],noReport:0,excluded:0}};}
    return{
      fields,selectedOperation:objectText(patient.raw||{})||`${patient.surgeryDate} ${patient.name}`,
      surgeries:history.length?history:[objectText(patient.raw||{})],note,
      materials:uniq(materialRecords.map(materialRecordText).filter(x=>/MESH|MEŞ|CERRAHİ YAMA|HERNİ YAMASI|YAMA KOMPOZİT|PROLEN|PROLENE/i.test(x))),
      history:uniq([...history,...visitHistory,...consultations,...services]),stay:[patientText],imaging:rad.reports.map(r=>r.source+' | '+r.text),radiologyAudit:rad.audit,dischargeFields
    };
  }
  function uiComponents(selector){
    return allDocs().flatMap(doc=>{try{return doc.defaultView.Ext?.ComponentQuery.query(selector)||[];}catch{return[];}});
  }
  function extractHerniaMeasurements(reports){
    const found=[];
    for(const report of reports){
      const text=cleanText(report);
      for(const sentence of text.split(/\||(?<=[.!?])\s+/)){
        if(!/HERN[İI]|DEFEKT|FASYA|FITIK/i.test(sentence))continue;
        for(const m of sentence.matchAll(/(\d+(?:[.,]\d+)?)\s*(?:x|×|\*)\s*(\d+(?:[.,]\d+)?)(?:\s*(?:x|×|\*)\s*(\d+(?:[.,]\d+)?))?\s*(cm|mm)/gi)){
          const sac=/HERN[İI]\s+KESE|HERNIA SAC/i.test(sentence)&&!/DEFEKT|FASYA/i.test(sentence);
          const size=[m[1],m[2],m[3]].filter(Boolean).join(' x ')+' '+m[4].toLowerCase();
          found.push(size+(sac?' (herni kesesi; defekt ölçüsü belirtilmemiş)':'')+' — '+sentence.trim());
        }
      }
    }
    return uniq(found);
  }
  function uiShown(component){return component?.el?.dom&&visible(component.el.dom);}
  function radiologyRoot(){
    const grid=uiComponents('gridpanel').find(g=>uiShown(g)&&/İstem Tarihi/.test(g.el.dom.innerText)&&/Rapor Türü/.test(g.el.dom.innerText));
    if(!grid)return null;
    let root=grid;
    while(root&&!/Rapor Sonuç Bilgileri/.test(root.el?.dom?.innerText||''))root=root.ownerCt;
    return root?{grid,root}:null;
  }
  async function uiIdle(){
    await sleep(250);
    const ok=await waitFor(()=>!allDocs().some(doc=>doc.defaultView.Ext?.Ajax?.isLoading?.())&&!elements('.x-mask-msg').some(visible),20000,150);
    if(!ok)throw new Error('FONET yanıtı beklenirken süre doldu');
    await sleep(150);
  }
  function activateGridRecord(grid,record){
    const view=grid.getView(),index=grid.getStore().indexOf(record);
    grid.getSelectionModel().select(record);
    const node=view.getNode(record);
    if(node)click(node);else view.fireEvent('itemclick',view,record,null,index,{});
  }
  function reportText(root){
    const editors=root.query?.('htmleditor')||[];
    const values=editors.filter(uiShown).map(e=>cleanText(e.getValue?.()||''));
    for(const frame of root.el.dom.querySelectorAll('iframe')){
      try{if(visible(frame))values.push(cleanText(frame.contentDocument?.body?.innerText));}catch{}
    }
    return uniq(values).filter(text=>text.length>15).join(' | ');
  }
  function reportSignature(root){
    return JSON.stringify((root.query?.('htmleditor')||[]).map(e=>cleanText(e.getValue?.()||'')));
  }
  async function supplementFromPatientScreen(patient,details){
    const source=extListSource();
    const record=source?.store.getRange().find(r=>norm(r.data.id)===norm(patient.ameliyatId));
    if(!record)throw new Error('Ameliyat kimliği açık listede doğrulanamadı');
    const node=source.grid.getView().getNode(record);
    if(!node)throw new Error('Ameliyat satırı ekranda yüklenemedi');
    click(node,true);await uiIdle();
    const loaded=await waitFor(()=>upper(currentPatientName())===upper(patient.name)&&norm(fieldValue('İşlem No :'))===norm(patient.operationNo),10000);
    if(!loaded)throw new Error('Hasta adı ve işlem numarası eşleşmedi');
    const tc=fieldValue('Kimlik No:');
    if(patient.tc&&norm(patient.tc)!==norm(tc))throw new Error('TC eşleşmedi');
    details.fields.tc=tc;
    const ageSex=fieldValue('Yaş / Cinsiyet / D.Tar:');
    details.fields.ageSex=ageSex;
    details.fields.gender=/\(Kadın\)/i.test(ageSex)?'Kadın':/\(Erkek\)/i.test(ageSex)?'Erkek':'';
    if(!details.fields.gender)throw new Error('Aktif hastanın cinsiyet alanı okunamadı; kayıt doğrulanamadı');
    const serviceButton=exactButton('Hizmet Listesi');if(serviceButton){click(serviceButton);await uiIdle();}
    const serviceGrid=uiComponents('gridpanel').find(g=>uiShown(g)&&/Makro/.test(g.el.dom.innerText)&&/Miktar/.test(g.el.dom.innerText)&&g!==source.grid);
    if(serviceGrid)details.history=uniq([...details.history,...serviceGrid.getStore().getRange().map(r=>objectText(r.data))]);
    const button=exactButton('Radyoloji');if(!button)throw new Error('Radyoloji sekmesi bulunamadı');
    click(button);await uiIdle();
    const ctx=radiologyRoot();if(!ctx)throw new Error('Radyoloji sonuç tablosu bulunamadı');
    const combo=(ctx.root.query('combobox')||[]).find(c=>/Bu Geliş|Tüm Gelişler|Bu Birim/.test(c.getRawValue?.()||''));
    if(!combo)throw new Error('Tüm Gelişler seçicisi bulunamadı');
    const option=combo.getStore().getRange().find(r=>Object.values(r.data).some(v=>norm(v)==='Tüm Gelişler'));
    if(!option)throw new Error('Tüm Gelişler seçeneği yüklenemedi');
    combo.setValue(option.get(combo.valueField));combo.fireEvent('select',combo,[option]);await uiIdle();
    if(norm(combo.getRawValue())!=='Tüm Gelişler')throw new Error('Tüm Gelişler seçilemedi');
    const store=ctx.grid.getStore();
    const total=store.getTotalCount?.()||store.getCount();
    if(total>store.getCount()){
      await new Promise((resolve,reject)=>{
        const timer=setTimeout(()=>reject(new Error('Radyoloji sayfaları zaman aşımı')),20000);
        store.load({params:{start:0,limit:total,page:1},callback:(r,o,ok)=>{clearTimeout(timer);ok?resolve():reject(new Error('Radyoloji sayfaları yüklenemedi'));}});
      });
    }
    if(store.getCount()<total)throw new Error(`Radyoloji listesi eksik: ${store.getCount()}/${total}`);
    const records=store.getRange(),reports=[],failures=[];
    for(let index=0;index<records.length;index++){
      if(state.stopped)throw new Error('Tarama durduruldu');
      while(state.paused&&!state.stopped)await sleep(250);
      log(`${patient.name}: radyoloji ${index+1}/${records.length}`);
      try{
        const before=reportSignature(ctx.root);
        const alreadySelected=ctx.grid.getSelectionModel().getSelection().includes(records[index]);
        activateGridRecord(ctx.grid,records[index]);await uiIdle();
        const tabs=ctx.root.query('tab')||[];
        const texts=[];
        for(const label of ['Sonuç Öneriler','Bulgular']){
          const t=tabs.find(t=>norm(t.text)===label);
          if(t?.el?.dom){click(t.el.dom);await uiIdle();const value=reportText(ctx.root);if(value)texts.push(value);}
        }
        const finalText=reportText(ctx.root);
        const text=uniq(texts.length?texts:[finalText]).join(' | ');
        if(!text)throw new Error('Rapor metni boş');
        if(!alreadySelected&&reportSignature(ctx.root)===before)throw new Error('Rapor değişimi doğrulanamadı; önceki rapor kopyalanmadı');
        reports.push({source:objectText(records[index].data),text});
      }catch(error){failures.push(`${index+1}: ${error.message}`);}
    }
    details.imaging=reports.map(r=>`${r.source} | ${r.text}`);
    details.radiologyAudit={total:records.length,read:reports.length,failures};
    details.fields.radiologyEvidence=reports;
  }
  async function scanPatient(patient){
    if(patient.mode==='fonet-list'){
      return await scanPatientBackground(patient);
    }
    let found,selected;
    if(patient.mode==='fonet-list'){
      found=openListRows();
      selected=found.find(r=>norm(r.innerText).includes(patient.operationNo))||chooseOperation(found.filter(r=>upper(r.innerText).includes(upper(patient.name))),patient.surgeryDate);
      let activated=false;
      for(let attempt=0;attempt<3;attempt++){
        if(selected){click(selected,true);activated=true;}
        activated=selectExtOperation(patient.operationNo,patient.listIndex)||activated;
        if(!activated)break;
        const ok=await waitFor(()=>{
          const shownName=upper(currentPatientName()),shownOperation=norm(fieldValue('İşlem No :'));
          return shownName===upper(patient.name)&&(!patient.operationNo||shownOperation===norm(patient.operationNo));
        },2800);
        if(ok){activated='loaded';break;}
        await sleep(120);
      }
      if(activated!=='loaded')throw new Error(activated?'Hasta bilgileri 3 denemede yüklenmedi':'Açık listedeki ameliyat satırı yeniden bulunamadı');
    }else{
      const controls=searchControls(); if(!controls)throw new Error('Ameliyat arama ekranı bulunamadı');
      click(controls.clear); await sleep(180); setValue(controls.tcInput,patient.tc); click(controls.query);
      found=await waitFor(()=>operationRows(controls.doc,patient),8000);
      if(!found||!found.length)throw new Error('TC için ameliyat kaydı bulunamadı');
      selected=chooseOperation(found,patient.surgeryDate); click(selected);
    }
    if(patient.mode!=='fonet-list'){
      await sleep(350);
      const loaded=await waitFor(()=>upper(currentPatientName())===upper(patient.name),10000);
      if(!loaded)throw new Error('Hasta bilgileri yüklenmedi');
    }
    patient.name=currentPatientName()||patient.name;patient.tc=currentTc()||patient.tc;
    if(patient.mode==='fonet-list'){
      const row=state.rows[patient.rowIndex];setIfFound(row,'name',patient.name);setIfFound(row,'tc',patient.tc);setIfFound(row,'operationNo',patient.operationNo);setIfFound(row,'surgeryDate',patient.surgeryDate);
    }
    const fields={operationNo:fieldValue('İşlem No :')||patient.operationNo,phone:fieldValue('Telefon')||fieldValue('Cep Telefonu'),ageSex:fieldValue('Yaş / Cinsiyet / D.Tar:'),asa:fieldValue('Asa/Euro Score:'),surgeryTimes:fieldValues('Ameliyat Saati').join(' ')};
    const surgeries=found.map(x=>norm(x.innerText));
    const selectedOperation=selected?norm(selected.innerText):`${patient.surgeryDate} ${patient.name} ${patient.operationNo}`;
    const note=await readTab('Ameliyat Notları',250);
    await readTab('Ameliyat Bilgileri',80);
    const materials=await readTab('İlaç Sarf',300);
    await readTab('Hizmet Listesi',80);
    const imaging=await readRadiology();
    await readTab('Hizmet Listesi',60);
    const stay=await readTab('Yatış Özet',220);
    const dischargeFields=uniq([...fieldValues('Taburcu Tarihi'),...fieldValues('Çıkış Tarihi'),...fieldValues('Bitiş Tarihi')]);
    await readTab('Hizmet Listesi',60);
    const history=patient.tc?await readHistory(patient.tc):[];
    return{fields,surgeries,selectedOperation,note,materials,history,stay,imaging,dischargeFields};
  }
  async function processPatient(p){
    try{const details=await scanPatient(p);if(state.stopped)return;const row=state.rows[p.rowIndex];p.tc=details.fields.tc||p.tc;p.name=p.name||details.fields.name;setIfFound(row,'name',p.name);setIfFound(row,'tc',p.tc);const derived=applyResult(p,details);const incomplete=details.radiologyAudit?.failures.length||0;const status=incomplete?'Eksik radyoloji':'Tamamlandı';if(incomplete)state.errors++;state.results[`${p.operationNo}|${p.surgeryDate}|${p.rowIndex}`]={status,details,derived:{prolenCount:derived.prolenCount,readmissions:derived.laterAdmissions.length}};log(`${p.name}: ${status}${details.radiologyAudit?`, radyoloji ${details.radiologyAudit.read}/${details.radiologyAudit.total}`:''}`,!!incomplete);}
    catch(error){if(state.stopped)return;state.errors++;state.results[`${p.operationNo}|${p.surgeryDate}|${p.rowIndex}`]={status:'Hata',error:String(error.message||error)};const r=state.rows[p.rowIndex];r[state.headerMap.get('FONET TARAMA DURUMU')]=`Hata: ${error.message||error}`;log(`${p.name||p.operationNo}: ${error.message||error}`,true);}
    state.current++;persist();updateStatus();
  }
  function consolidateDuplicatePatients(){
    if(state.mode!=='fonet-list')return 0;
    const groups=new Map();
    for(const patient of state.patients){
      const key=/^\d{11}$/.test(patient.tc)?patient.tc:'';
      if(!key)continue;
      if(!groups.has(key))groups.set(key,[]);groups.get(key).push(patient);
    }
    const removeRows=new Set();let merged=0;
    for(const group of groups.values()){
      if(group.length<2)continue;
      const ordered=[...group].sort((a,b)=>(parseDateTime(a.surgeryDate)||new Date(8640000000000000))-(parseDateTime(b.surgeryDate)||new Date(8640000000000000)));
      const primary=ordered[0],base=state.rows[primary.rowIndex];
      const distinctOperations=uniq(ordered.map(p=>`${p.surgeryDate} | ${p.operationNo}`));
      for(const patient of ordered.slice(1)){
        const other=state.rows[patient.rowIndex];
        for(let i=0;i<state.headers.length;i++)if((base[i]===''||base[i]==null)&&other[i]!==''&&other[i]!=null)base[i]=other[i];
        removeRows.add(patient.rowIndex);merged++;
      }
      setIfFound(base,'operationNo',uniq(ordered.map(p=>p.operationNo)).join('; '));
      setIfFound(base,'surgeryDate',uniq(ordered.map(p=>p.surgeryDate)).join('; '));
      const existingCounts=ordered.map(p=>Number(getCell(state.rows[p.rowIndex],'opCount'))||0);
      for(const key of ['defect','location']){
        const values=uniq(ordered.map(p=>norm(getCell(state.rows[p.rowIndex],key))).filter(Boolean));
        if(values.length)setIfFound(base,key,key==='location'?uniq(values.flatMap(v=>v.match(/M[1-5]/g)||[])).sort().join(', '):values.join('; '));
      }
      for(const header of ['RADYOLOJİ KAYNAKLARI','RADYOLOJİ OKUMA','MALİGNİTE KAYNAĞI']){
        const column=state.headerMap.get(header);
        base[column]=uniq(ordered.map(p=>state.rows[p.rowIndex][column]).filter(Boolean)).join('\n').slice(0,32000);
      }
      const opCount=Math.max(distinctOperations.length,...existingCounts,1);
      setIfFound(base,'opCount',opCount);
      if(opCount>1){
        setIfFound(base,'recurrence',1);
        const later=distinctOperations.slice(1).join('; ');
        if(later)setIfFound(base,'followRecurrence',`Evet – ${later}`);
      }
      for(const key of ['malign','benign','ht','dm','copd','hf','goiter','smoking','necrosis','vac','seroma','revision','mortality']){
        const values=ordered.map(p=>getCell(state.rows[p.rowIndex],key));
        if(values.some(v=>String(v).startsWith('1'))){
          const detailed=values.find(v=>String(v).startsWith('1 –'));
          setIfFound(base,key,detailed||1);
        }
      }
      if(String(getCell(base,'malign')).startsWith('1'))setIfFound(base,'benign',0);
    }
    if(removeRows.size)state.rows=state.rows.filter((row,index)=>index===0||!removeRows.has(index));
    return merged;
  }
  async function run(){
    if(!state.patients.length||state.running)return;
    state.running=true;state.stopped=false;state.errors=0;$('#fx-start').disabled=true;$('#fx-pause').disabled=false;$('#fx-stop').disabled=false;
    const queue=state.patients.filter(p=>!state.results[`${p.operationNo}|${p.surgeryDate}|${p.rowIndex}`]);
    state.current=state.patients.length-queue.length;
    const worker=async()=>{while(queue.length&&!state.stopped){while(state.paused&&!state.stopped)await sleep(250);const p=queue.shift();if(!p||state.stopped)return;await processPatient(p);}};
    if(state.mode==='fonet-list')await Promise.all(Array.from({length:3},()=>worker()));
    else for(const p of queue){if(state.stopped)break;while(state.paused&&!state.stopped)await sleep(250);if(!state.stopped)await processPatient(p);}
    state.running=false;
    const merged=!state.stopped?consolidateDuplicatePatients():0;
    if(state.destroyRequested){panel.remove();delete window.__fonetExcelHastaTarayici;return;}
    $('#fx-start').disabled=!state.stopped;$('#fx-pause').disabled=true;$('#fx-stop').disabled=true;$('#fx-export').disabled=false;
    updateStatus(state.stopped?'Tarama durduruldu. Sonuçlar kaydedildi.':`Tarama sona erdi. Hatalı veya eksik kayıt: ${state.errors}.${merged?` ${merged} yinelenen ameliyat satırı aynı TC altında birleştirildi.`:''} Radyoloji okuma sayılarını Excel'den kontrol edin.`);
  }
  function ensureOutputColumns(){
    for(const h of ['PROLEN MESH ADEDİ','MALZEME KAYDI','FONET TARAMA DURUMU','MALİGNİTE KAYNAĞI','RADYOLOJİ OKUMA','RADYOLOJİ KAYNAKLARI']){
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
  async function loadFonetList(){
    try{
      const source=extListSource();
      if(!source)throw new Error('Açık ameliyat veri tablosu bulunamadı.');
      updateStatus(`FONET listesinin tamamı alınıyor: ${source.total||source.count} kayıt...`);
      let records=source.store.getRange?.()||[];
      if(source.total>records.length){
        const proxy=source.store.getProxy?.(),rawUrl=norm(proxy?.url||proxy?.api?.read);
        if(rawUrl){
          const url=new URL(rawUrl,location.origin);
          url.searchParams.set('start','0');url.searchParams.set('page','1');url.searchParams.set('limit',String(source.total));
          for(const [key,value] of Object.entries(proxy?.extraParams||{}))url.searchParams.set(key,String(value));
          url.searchParams.set('_dc',String(Date.now()));
          const response=await fetch(url.href,{credentials:'include',headers:{Accept:'application/json, text/plain, */*'}});
          if(response.ok){const payload=await response.json();const all=payloadRows(payload);if(all.length>records.length)records=all.map(data=>({data}));}
        }
      }
      const parsed=records.length>openListRows().length
        ? records.map(extRecordData).filter(x=>x.operationNo)
        : openListRows().map(listRowData).filter(x=>x.operationNo);
      const unique=[...new Map(parsed.map(x=>[`${x.operationNo}|${x.surgeryDate}`,x])).values()];
      if(!unique.length)throw new Error('Açık ameliyat listesinde hasta satırı bulunamadı. Liste ekranını açık ve yüklenmiş bırakın.');
      state.mode='fonet-list';state.workbook=XLSX.utils.book_new();state.sheetName='Hastalar';state.headers=[...FONET_TEMPLATE_HEADERS];
      state.rows=[state.headers,...unique.map(x=>{const r=Array(state.headers.length).fill('');r[0]=x.name;r[2]=x.operationNo;r[7]=x.surgeryDate;return r;})];
      ensureOutputColumns();XLSX.utils.book_append_sheet(state.workbook,XLSX.utils.aoa_to_sheet(state.rows),state.sheetName);
      state.patients=unique.map((x,i)=>({rowIndex:i+1,name:x.name,tc:'',operationNo:x.operationNo,surgeryDate:x.surgeryDate,listIndex:x.index,mode:'fonet-list',ameliyatId:x.ameliyatId,gelisId:x.gelisId,birimSevkId:x.birimSevkId,isteyenBirimSevkId:x.isteyenBirimSevkId,kimlikId:x.kimlikId,raw:x.raw}));
      state.results={};state.current=0;state.errors=0;persist();$('#fx-start').disabled=false;$('#fx-export').disabled=false;
      updateStatus(`${state.patients.length} ameliyat FONET açık listesinden alındı. Taramayı Başlat'a basın.`);log(`FONET listesinden ${state.patients.length} kayıt hazırlandı`);
    }catch(error){updateStatus(error.message||String(error));log(error.message||String(error),true);}
  }
  function exportExcel(){
    if(!state.rows.length)return;
    const ws=XLSX.utils.aoa_to_sheet(state.rows);ws['!cols']=state.headers.map((h,i)=>({wch:Math.min(60,Math.max(12,Math.max(...state.rows.slice(0,46).map(r=>norm(r[i]).length))+2))}));
    state.workbook.Sheets[state.sheetName]=ws;
    const auditRows=[['TC','Hasta','Durum','Prolen mesh adedi','Yeniden yatış','Hata','Radyoloji toplam','Radyoloji okunan','Okunamayan raporlar']];
    for(const p of state.patients){const rr=state.results[`${p.operationNo}|${p.surgeryDate}|${p.rowIndex}`]||state.results[p.tc||`${p.operationNo}|${p.surgeryDate}`]||{};const a=rr.details?.radiologyAudit;auditRows.push([p.tc,p.name,rr.status||'Taranmadı',rr.derived?.prolenCount??'',rr.derived?.readmissions??'',rr.error||'',a?.total??'',a?.read??'',a?.failures.join('; ')||'']);}
    state.workbook.Sheets['FONET Tarama Kaydı']=XLSX.utils.aoa_to_sheet(auditRows);if(!state.workbook.SheetNames.includes('FONET Tarama Kaydı'))state.workbook.SheetNames.push('FONET Tarama Kaydı');
    XLSX.writeFile(state.workbook,`FONET_TARANMIS_${new Date().toISOString().slice(0,10)}.xlsx`,{compression:true,cellStyles:true});
  }
  $('#fx-load').onclick=loadExcel;$('#fx-fonet-list').onclick=loadFonetList;$('#fx-start').onclick=run;$('#fx-pause').onclick=()=>{state.paused=!state.paused;$('#fx-pause').textContent=state.paused?'Devam Et':'Duraklat';updateStatus(state.paused?'Tarama duraklatıldı.':'Tarama sürüyor...');};
  $('#fx-stop').onclick=()=>{state.stopped=true;state.paused=false;abortRequests();};$('#fx-export').onclick=exportExcel;
  $('#fx-close').onclick=()=>{
    state.stopped=true;state.paused=false;
    abortRequests();
    if(state.running){state.destroyRequested=true;panel.style.display='none';}
    else{panel.remove();delete window.__fonetExcelHastaTarayici;}
  };
})();
