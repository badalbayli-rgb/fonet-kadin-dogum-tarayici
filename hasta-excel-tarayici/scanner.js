(() => {
  'use strict';
  try {
    const grids=(window.Ext?.ComponentQuery?.query('gridpanel,grid')||[]).map(g=>{let s,p;try{s=g.getStore?.();p=s?.getProxy?.();}catch{}return{id:g.id,title:g.title||'',itemId:g.itemId||'',count:s?.getCount?.()||0,url:p?.url||p?.api?.read||'',extra:p?.extraParams||{},keys:s?.getAt?.(0)?Object.keys(s.getAt(0).data||{}).slice(0,40):[]};});
    console.log('FONET_GRID_DIAG',JSON.stringify(grids));
  } catch {}
  try {
    const forms=(window.Ext?.ComponentQuery?.query('form')||[]).map(form=>{
      let basic,record,values={};
      try{basic=form.getForm?.();record=basic?.getRecord?.()?.data||null;values=basic?.getValues?.()||{};}catch{}
      const text=JSON.stringify({record,values});
      return{id:form.id,itemId:form.itemId||'',record,values,interesting:/ameliyat|islemNo|saatBilgisi|protokolNo/i.test(text)};
    }).filter(x=>x.interesting);
    const fields=(window.Ext?.ComponentQuery?.query('field')||[]).map(field=>{
      let value='';try{value=field.getValue?.()??field.getRawValue?.()??field.value??'';}catch{}
      return{id:field.id,itemId:field.itemId||'',name:field.name||'',label:field.fieldLabel||'',value};
    }).filter(x=>/işlem no|islem no|ameliyat saati|an\.baş|an\.bas|post-op|prot\.no/i.test(`${x.label} ${x.name}`));
    console.log('FONET_FORM_DIAG',JSON.stringify({forms,fields}));
  } catch {}
  try {
    if(window.Ext?.Ajax&&!window.__fonetExcelAjaxDiag){
      window.__fonetExcelAjaxDiag=true;
      Ext.Ajax.on('beforerequest',(conn,options)=>{
        const url=String(options?.url||'');
        if(/Ameliyat|HastaBirimSevk|HastaHizmet/i.test(url))console.log('FONET_AJAX_DIAG',JSON.stringify({phase:'request',url,method:options?.method||'',params:options?.params||null,jsonData:options?.jsonData||null}));
      });
      Ext.Ajax.on('requestcomplete',(conn,response,options)=>{
        const url=String(options?.url||'');
        if(/Ameliyat|HastaBirimSevk|HastaHizmet/i.test(url))console.log('FONET_AJAX_DIAG',JSON.stringify({phase:'response',url,status:response?.status||0,body:String(response?.responseText||'').slice(0,1200)}));
      });
    }
  } catch {}
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
    <div class="head"><div class="title">FONET Hasta ve Excel Tarayıcı v1.3.4</div><button id="fx-close" class="close">Kapat</button></div>
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
    const cells=norm(text).split(/\s+/); const nums=cells.map(x=>Number(String(x).replace(',','.'))).filter(x=>Number.isFinite(x)&&x>0&&x<=10&&Number.isInteger(x));
    return nums.length?nums[nums.length-1]:1;
  }
  function classifyEhsLocation(imagingText,noteText){
    const source=upper(`${(imagingText||[]).join(' | ')} | ${noteText||''}`);
    const distance=(rx)=>{const m=source.match(rx);return m?Number(m[1].replace(',','.')):null;};
    const aboveUmb=distance(/G[ÖO]BE(?:K|Ğ[İI])(?:N)?\s+(\d+(?:[.,]\d+)?)\s*CM\s+(?:ÜST|YUKARI)/);
    const belowUmb=distance(/G[ÖO]BE(?:K|Ğ[İI])(?:N)?\s+(\d+(?:[.,]\d+)?)\s*CM\s+(?:ALT|AŞAĞI)/);
    const belowXiphoid=distance(/(?:KSİFOİD|KSIFOID|XİFOİD|XIPHOID)(?:İN)?\s+(\d+(?:[.,]\d+)?)\s*CM\s+ALT/);
    const abovePubis=distance(/PUBİ[SK](?:İN)?\s+(\d+(?:[.,]\d+)?)\s*CM\s+(?:ÜST|YUKARI)/);
    if(/SUBKSİFOİD|SUBKSIFOID|SUBXİPHOID|SUBXIPHOID/.test(source)||(belowXiphoid!==null&&belowXiphoid<=3))return'M1 (Subksifoid)';
    if(/SUPRAPUBİK|SUPRAPUBIK/.test(source)||(abovePubis!==null&&abovePubis<=3))return'M5 (Suprapubik)';
    if(/İNFRAUMBİLİKAL|INFRAUMBILIKAL|ALT ORTA HAT/.test(source)||(belowUmb!==null&&belowUmb>3)||(abovePubis!==null&&abovePubis>3))return'M4 (İnfraumbilikal)';
    if(/PERİUMBİLİKAL|PERIUMBILIKAL|UMBİLİKAL|UMBILIKAL|UMBLİKAL|G[ÖO]BEK ÇEVRES/.test(source)||(aboveUmb!==null&&aboveUmb<=3)||(belowUmb!==null&&belowUmb<=3))return'M3 (Umbilikal)';
    if(/EPİGASTRİK|EPIGASTRIK|ÜST ORTA HAT|SUPRAUMBİLİKAL|SUPRAUMBILIKAL/.test(source)||(aboveUmb!==null&&aboveUmb>3)||(belowXiphoid!==null&&belowXiphoid>3))return'M2 (Epigastrik)';
    return'';
  }
  function derive(patient, details){
    const surgeryDate=parseDateTime(patient.surgeryDate)||parseDateTime(details.selectedOperation);
    const note=details.note.join(' | '), allHistory=details.history.join(' | '), all=upper(`${note} ${allHistory}`);
    const opRows=uniq([...(details.surgeries||[]),...(details.history||[])]).map(text=>({text,date:parseDateTime(text)})).filter(x=>x.date);
    const previousHernia=details.history.filter(x=>/\bK43(?:\.|-)|İNSİZYONEL HERNİ|VENTRAL HERNİ/i.test(x)&&parseDate(x)&&surgeryDate&&parseDate(x)<surgeryDate);
    const laterHerniaOps=opRows.filter(x=>surgeryDate&&x.date>surgeryDate&&/İNSİZYONEL HERNİ|VENTRAL HERNİ/i.test(x.text));
    const laterDebridement=opRows.filter(x=>surgeryDate&&x.date>surgeryDate&&/YARA[^|]{0,40}DEBRİDMAN|YARA[^|]{0,40}DEBRIDMAN|DEBRİDMAN|DEBRIDMAN/i.test(x.text));
    const laterAdmissions=details.history.filter(x=>{const d=parseDate(x);return d&&surgeryDate&&dateDiffDays(surgeryDate,d)>2&&/\bYATIŞ\b/i.test(x);});
    const previousAbdominal=opRows.filter(x=>surgeryDate&&x.date<surgeryDate&&/ABDOM|LAPAROT|LAPAROSK|APPEN|KOLESİST|KOLEKT|REZEKS|GASTREK|HERNİ|FITIK|SEZARYEN|HİSTEREKT|OOFOREKT|SALPEN|KOLON|REKT|İLEOST|KOLOST|PANKREAT|SPLENEKT|BARSAK|BAĞIRSAK|UMBİLİK|MİDE|BYPASS|SLEEVE/i.test(x.text));
    const cancers=uniq(details.history.flatMap(x=>x.match(/\bC\d{2}(?:\.\d+)?-[^|]+/gi)||[]));
    const materialRows=details.materials.filter(x=>/MESH|MEŞ|CERRAHİ YAMA|HERNİ YAMASI|YAMA KOMPOZİT|PROLEN(?:E)?\s+(?:MESH|MEŞ|YAMA)/i.test(x)&&!/SÜTÜR|SUTUR|İĞNE|IGNE|YUVARLAK|ÖRGÜLÜ|EMİLEBİLEN|EMILEBILEN/i.test(x));
    const prolenRows=materialRows.filter(x=>/PROLEN|PROLENE/i.test(x));
    const prolenCount=prolenRows.reduce((sum,x)=>sum+quantityFromMaterialRow(x),0);
    const ageSex=details.fields.ageSex||''; const sex=/\((Kadın|Erkek)\)/i.exec(ageSex)?.[1]||''; const age=/^(\d+)/.exec(ageSex)?.[1]||'';
    const times=(details.fields.surgeryTimes||'').match(/\d{2}:\d{2}/g)||[]; let duration='';
    if(times.length>=2){const [h1,m1]=times[0].split(':').map(Number),[h2,m2]=times[times.length-1].split(':').map(Number);duration=(h2*60+m2)-(h1*60+m1);if(duration<0)duration+=1440;}
    const diagnoses=upper(allHistory);
    const complications=[]; if(/NEKROZ/.test(all))complications.push('Nekroz');if(/SEROMA/.test(all))complications.push('Seroma');if(/YARA ENFEKSİY|CERRAHİ ALAN ENFEKSİY|ENFEKSİYON/.test(all))complications.push('Enfeksiyon');if(/DEHİS|EVİSSER/.test(all))complications.push('Dehisens');
    const defect=(note.match(/(\d+(?:[.,]\d+)?)\s*(?:x|×|\*)\s*(\d+(?:[.,]\d+)?)\s*cm/i)||[]).slice(1).join(' x ');
    const location=classifyEhsLocation(details.imaging,note);
    const dischargeDates=(details.dischargeFields||[]).map(parseDateTime).filter(x=>x&&surgeryDate&&x>=surgeryDate).sort((a,b)=>a-b);
    const discharge=dischargeDates[0]||null;
    const stayDays=discharge&&surgeryDate?Math.max(0,Math.round(((discharge-surgeryDate)/86400000)*10)/10):'';
    return {surgeryDate,previousHernia,laterHerniaOps,laterDebridement,laterAdmissions,previousAbdominal,cancers,materialRows,prolenCount,sex,age,duration,diagnoses,all,note,complications,defect,location,stayDays};
  }
  function applyResult(patient, details){
    const row=state.rows[patient.rowIndex], d=derive(patient,details);
    setIfFound(row,'operationNo',details.fields.operationNo);
    setIfFound(row,'phone',details.fields.phone);
    setIfFound(row,'sex',d.sex); setIfFound(row,'age',d.age?Number(d.age):''); setIfFound(row,'asa',details.fields.asa);
    setIfFound(row,'duration',d.duration||''); setIfFound(row,'followup',d.surgeryDate?dateDiffHuman(d.surgeryDate):'');
    setIfFound(row,'defect',d.defect);setIfFound(row,'location',d.location);setIfFound(row,'stay',d.stayDays);
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
    if(d.laterDebridement.length){setIfFound(row,'necrosis',1);setIfFound(row,'vac',1);}
    if(/REVİZYON|REVIZYON|REAKSPLORASYON/.test(d.all))setIfFound(row,'revision',1);
    if(d.laterHerniaOps.length)setIfFound(row,'revision',1);
    if(/\bI10(?:\.|-)/.test(d.diagnoses))setIfFound(row,'ht',1); if(/\bE1[01](?:\.|-)/.test(d.diagnoses))setIfFound(row,'dm',1);
    if(/\bJ4[34](?:\.|-)/.test(d.diagnoses))setIfFound(row,'copd',1); if(/\bI50(?:\.|-)/.test(d.diagnoses))setIfFound(row,'hf',1);
    if(/\bE0[0-7](?:\.|-)|GUATR/.test(d.diagnoses))setIfFound(row,'goiter',1);
    if(/SİGARA.*(İÇİYOR|KULLANIYOR|AKTİF)|AKTİF SİGARA/.test(d.all))setIfFound(row,'smoking',1);
    if(d.materialRows.length)setIfFound(row,'graft',uniq(d.materialRows).join('; '));
    if(/DREN/.test(d.note)){const m=d.note.match(/DREN[^|]{0,60}?(\d+)\s*GÜN/i);if(m)setIfFound(row,'drain',`${m[1]} GÜN`);}
    row[state.headerMap.get('PROLEN MESH ADEDİ')] = d.prolenCount||0;
    row[state.headerMap.get('MALZEME KAYDI')] = d.materialRows.length?uniq(d.materialRows).join('; '):'FONET sarf kaydında mesh/yama bulunmadı';
    row[state.headerMap.get('FONET TARAMA DURUMU')] = 'Tamamlandı';
    return d;
  }

  function apiBase(){return `${location.origin}/hbys-rs/hbys`;}
  async function apiJson(path){
    const sep=path.includes('?')?'&':'?';
    const response=await fetch(`${apiBase()}${path}${sep}_dc=${Date.now()}`,{credentials:'include',headers:{Accept:'application/json, text/plain, */*'}});
    const body=await response.text();
    if(!response.ok)throw new Error(`HTTP ${response.status}`);
    try{return JSON.parse(body);}catch{throw new Error('FONET yanıtı okunamadı');}
  }
  function payloadRows(payload){
    if(Array.isArray(payload))return payload;
    for(const key of ['data','list','rows','result','content'])if(Array.isArray(payload?.[key]))return payload[key];
    return payload?.data&&typeof payload.data==='object'?[payload.data]:(payload&&typeof payload==='object'?[payload]:[]);
  }
  function objectText(value,depth=0,seen=new Set()){
    if(value==null||depth>7)return'';
    if(typeof value!=='object')return norm(value);
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
  async function settled(path){try{return await apiJson(path);}catch(error){return{__error:String(error.message||error)};}}
  async function operationHistory(patient){
    if(!patient.kimlikId)return[];
    const filter=encodeURIComponent(JSON.stringify([{index:1,property:'kimlikId',value:Number(patient.kimlikId)||patient.kimlikId,filterType:'kriterPanel',type:'Long',operator:'='}]));
    const payload=await settled(`/Ameliyat/Ameliyat/getKayitList?start=0&limit=1000&page=1&filter=${filter}`);
    return payload.__error?[]:payloadRows(payload).map(x=>objectText(x));
  }
  async function scanPatientBackground(patient){
    if(!patient.ameliyatId||!patient.gelisId||!patient.birimSevkId)throw new Error('Kayıt servis kimlikleri bulunamadı; listeyi yeniden alın');
    const [notePayload,servicePayload,patientPayload,consultPayload,history]=await Promise.all([
      settled(`/Ameliyat/Ameliyat/getAmeliyatPersonelList/${encodeURIComponent(patient.ameliyatId)}/-1`),
      settled(`/Tibbi/HastaHizmet/getHizmetList/${encodeURIComponent(patient.birimSevkId)}/${encodeURIComponent(patient.gelisId)}`),
      settled(`/Tibbi/HastaBirimSevk/getSevkUyariInfo/${encodeURIComponent(patient.birimSevkId)}`),
      settled(`/Poliklinik/Poliklinik/getHastaGelisKonsultasyonList/${encodeURIComponent(patient.gelisId)}/1`),
      operationHistory(patient)
    ]);
    const failed=[notePayload,servicePayload,patientPayload].filter(x=>x?.__error).length;
    if(failed===3)throw new Error('FONET arka plan servisleri yanıt vermedi');
    const noteRows=notePayload.__error?[]:payloadRows(notePayload);
    const serviceRows=servicePayload.__error?[]:payloadRows(servicePayload);
    const consultRows=consultPayload.__error?[]:payloadRows(consultPayload);
    const patientRoot=patientPayload.__error?{}:patientPayload;
    const note=noteRows.map(x=>norm(x?.notu||x?.hizmet||objectText(x))).filter(Boolean);
    const services=serviceRows.map(x=>objectText(x));
    const consultations=consultRows.map(x=>objectText(x));
    const patientText=objectText(patientRoot);
    const gender=deepValue(patientRoot,['cinsiyetAdi','cinsiyet']);
    const birthDate=deepValue(patientRoot,['dogumTarihi','doğumTarihi']);
    const birth=parseDateTime(birthDate),calculatedAge=birth?Math.max(0,new Date().getFullYear()-birth.getFullYear()-(new Date()<new Date(new Date().getFullYear(),birth.getMonth(),birth.getDate())?1:0)):'';
    const ageSex=deepValue(patientRoot,['yasCinsiyetDogumTarihi','yasCinsiyet'])||(gender?`${calculatedAge||''}Yıl (${gender}) / ${birthDate}`:'')||patientText.match(/\d+\s*(?:Yıl|Yaş)[^|]{0,40}\((?:Kadın|Erkek)\)/i)?.[0]||'';
    const fields={
      operationNo:patient.operationNo,
      tc:deepValue(patientRoot,['kimlikNo','tcKimlikNo','tckn']),
      name:deepValue(patientRoot,['adiSoyadi','adSoyad','hastaAdiSoyadi'])||patient.name,
      phone:deepValue(patientRoot,['cepTelefonu','cepTelefon','telefonNo','telefon','gsm','mobilTelefon']),
      ageSex,
      asa:deepValue(patient.raw||{},['asa','asaSkoru','asaEuroScore']),
      surgeryTimes:objectText(patient.raw||{})
    };
    const dischargeFields=Object.entries(flatObject(patientRoot)).filter(([k])=>/taburcu|cikis|bitiş|bitis/i.test(k)).map(([,v])=>norm(v));
    return{
      fields,selectedOperation:objectText(patient.raw||{})||`${patient.surgeryDate} ${patient.name}`,
      surgeries:history.length?history:[objectText(patient.raw||{})],note,
      materials:services.filter(x=>/MESH|MEŞ|YAMA|PROLEN|PROLENE/i.test(x)),
      history:uniq([...history,...consultations]),stay:[patientText],imaging:services.filter(x=>/RADYOLOJ|TOMOGRAF|\bBT\b|USG|MRG|ABDOM|BATIN/i.test(x)),dischargeFields
    };
  }
  async function scanPatient(patient){
    if(patient.mode==='fonet-list')return scanPatientBackground(patient);
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
    try{const details=await scanPatient(p);const row=state.rows[p.rowIndex];p.tc=details.fields.tc||p.tc;p.name=details.fields.name||p.name;setIfFound(row,'name',p.name);setIfFound(row,'tc',p.tc);const derived=applyResult(p,details);state.results[p.tc||`${p.operationNo}|${p.surgeryDate}`]={status:'Tamamlandı',details,derived:{prolenCount:derived.prolenCount,readmissions:derived.laterAdmissions.length}};log(`${p.name}: tamamlandı, Prolen mesh ${derived.prolenCount}, yeniden yatış ${derived.laterAdmissions.length}`);}
    catch(error){state.errors++;state.results[p.tc||`${p.operationNo}|${p.surgeryDate}`]={status:'Hata',error:String(error.message||error)};const r=state.rows[p.rowIndex];r[state.headerMap.get('FONET TARAMA DURUMU')]=`Hata: ${error.message||error}`;log(`${p.name||p.operationNo}: ${error.message||error}`,true);}
    state.current++;persist();updateStatus();
  }
  async function run(){
    if(!state.patients.length||state.running)return;
    state.running=true;state.stopped=false;state.errors=0;$('#fx-start').disabled=true;$('#fx-pause').disabled=false;$('#fx-stop').disabled=false;
    const queue=state.patients.slice(state.current);
    const worker=async()=>{while(queue.length&&!state.stopped){while(state.paused&&!state.stopped)await sleep(250);const p=queue.shift();if(!p||state.stopped)return;await processPatient(p);}};
    if(state.mode==='fonet-list')await Promise.all(Array.from({length:Math.min(6,queue.length)},worker));
    else for(const p of queue){if(state.stopped)break;while(state.paused&&!state.stopped)await sleep(250);if(!state.stopped)await processPatient(p);}
    state.running=false;
    if(state.destroyRequested){panel.remove();delete window.__fonetExcelHastaTarayici;return;}
    $('#fx-start').disabled=false;$('#fx-pause').disabled=true;$('#fx-stop').disabled=true;$('#fx-export').disabled=false;
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
    const auditRows=[['TC','Hasta','Durum','Prolen mesh adedi','Yeniden yatış','Hata']];
    for(const p of state.patients){const rr=state.results[p.tc||`${p.operationNo}|${p.surgeryDate}`]||{};auditRows.push([p.tc,p.name,rr.status||'Taranmadı',rr.derived?.prolenCount??'',rr.derived?.readmissions??'',rr.error||'']);}
    state.workbook.Sheets['FONET Tarama Kaydı']=XLSX.utils.aoa_to_sheet(auditRows);if(!state.workbook.SheetNames.includes('FONET Tarama Kaydı'))state.workbook.SheetNames.push('FONET Tarama Kaydı');
    XLSX.writeFile(state.workbook,`FONET_TARANMIS_${new Date().toISOString().slice(0,10)}.xlsx`,{compression:true,cellStyles:true});
  }
  $('#fx-load').onclick=loadExcel;$('#fx-fonet-list').onclick=loadFonetList;$('#fx-start').onclick=run;$('#fx-pause').onclick=()=>{state.paused=!state.paused;$('#fx-pause').textContent=state.paused?'Devam Et':'Duraklat';updateStatus(state.paused?'Tarama duraklatıldı.':'Tarama sürüyor...');};
  $('#fx-stop').onclick=()=>{state.stopped=true;state.paused=false;};$('#fx-export').onclick=exportExcel;
  $('#fx-close').onclick=()=>{
    state.stopped=true;state.paused=false;
    if(state.running){state.destroyRequested=true;panel.style.display='none';}
    else{panel.remove();delete window.__fonetExcelHastaTarayici;}
  };
})();
