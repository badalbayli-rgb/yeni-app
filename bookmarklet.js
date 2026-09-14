/*
 Yeni App — FONET Gebelikte Apandisit Hasta Tarayıcı v1.0.1
 Yalnızca okur: FONET'e kayıt eklemez/değiştirmez.
*/
(() => {
  'use strict';
  const APP='__YENI_APP_FONET_GEBELIK_APANDISIT_TARAYICI__';
  if (window.top !== window || window[APP]) return;
  window[APP]=true;

  const NA='Ulaşılamadı', sleep=ms=>new Promise(r=>setTimeout(r,ms));
  const norm=v=>String(v??'').replace(/\s+/g,' ').trim();
  const clean=v=>norm(String(v??'').replace(/<br\s*\/?>/gi,' ').replace(/<[^>]*>/g,' ').replace(/&nbsp;|&#160;/gi,' '));
  const up=v=>norm(v).toLocaleUpperCase('tr-TR');
  const uniq=a=>[...new Set(a.map(norm).filter(Boolean))];
  const now=()=>new Date().toLocaleTimeString('tr-TR');
  const active=new Set();
  const headers=['Adı','Soyadı','TC','İşlem No','Telefon','Yaş','ASA','Ameliyat tarihi','Ameliyat süresi (dk)','Yatış süresi','Kaçıncı trimester','Preoperatif görüntüleme','Patoloji','Negatif apendektomi mi?','Preoperatif görüntüleme / patoloji uyumu','Cerrahi yöntem','Apandisit tipi','Maternal komplikasyon','Fetal komplikasyon','Postoperatif SSO','Postoperatif analjezi ihtiyacı (doz)','DVT gelişimi','Tokoliz ihtiyacı','Tokolizde kullanılan ilaç','Postoperatif kontraksiyon / NST pozitifliği','İlerleyen süreçte canlı doğum','Preoperatif WBC','Preoperatif CRP','Preoperatif NLR','Preoperatif PLR','Preoperatif SII','KAYNAK / NOT','Tarama durumu'];
  const state={patients:[],results:[],running:false,paused:false,stop:false,done:0,errors:0};

  const panel=document.createElement('section'); panel.id='fga-panel';
  panel.innerHTML=`<style>#fga-panel{position:fixed;inset:14px;z-index:2147483647;background:#f8fafc;color:#102a43;border:1px solid #829ab1;border-radius:13px;box-shadow:0 18px 60px #0008;padding:14px;display:flex;flex-direction:column;gap:9px;font:13px Arial,sans-serif}#fga-panel *{box-sizing:border-box}#fga-panel header{display:flex;justify-content:space-between;align-items:center;background:#0f4c81;color:#fff;margin:-14px -14px 0;padding:13px 15px;border-radius:12px 12px 0 0}#fga-panel button{border:0;border-radius:6px;padding:8px 11px;background:#0878bd;color:#fff;font-weight:700;cursor:pointer;margin-right:5px}#fga-panel button:disabled{opacity:.45;cursor:not-allowed}#fga-stop{background:#b42318!important}#fga-pause,#fga-csv,#fga-close{background:#52606d!important}#fga-status{white-space:pre-wrap;min-height:36px}#fga-bar-wrap{height:9px;background:#d9e2ec;border-radius:8px;overflow:hidden}#fga-bar{height:100%;width:0;background:#16a34a}#fga-table{overflow:auto;flex:1;background:#fff;border:1px solid #bcccdc;border-radius:7px}#fga-table table{border-collapse:collapse;width:max-content;min-width:100%}#fga-table th,#fga-table td{padding:7px;vertical-align:top;text-align:left;border-bottom:1px solid #e5e7eb;max-width:260px;white-space:pre-wrap}#fga-table th{position:sticky;top:0;background:#d9eaf7}</style><header><b>Yeni App — Gebelikte Apandisit Hasta Tarayıcı v1.0.1</b><button id="fga-close">Kapat</button></header><div><input id="fga-file" type="file" accept=".xlsx,.xls,.csv"><button id="fga-excel">Excel'i Oku</button><button id="fga-open">Açık Listeyi Oku</button><button id="fga-start" disabled>Taramayı Başlat</button><button id="fga-pause" disabled>Duraklat</button><button id="fga-stop" disabled>Durdur</button><button id="fga-csv" disabled>CSV indir</button></div><div id="fga-status">Excel yükleyin veya FONET'teki ameliyat sorgu listesini açık bırakıp “Açık Listeyi Oku” seçin. Araç kayıt oluşturmaz/değiştirmez.</div><div id="fga-bar-wrap"><div id="fga-bar"></div></div><div id="fga-table"><table><thead><tr>${headers.map(h=>`<th>${h}</th>`).join('')}</tr></thead><tbody id="fga-body"></tbody></table></div>`;
  document.documentElement.append(panel);
  const $=s=>panel.querySelector(s), status=t=>{$('#fga-status').textContent=t??`İşlenen: ${state.done}/${state.patients.length} | Hata: ${state.errors}`;$('#fga-bar').style.width=state.patients.length?`${100*state.done/state.patients.length}%`:'0%';};
  const log=t=>status(`${now()}  ${t}`);
  const render=()=>{$('#fga-body').innerHTML=state.results.map(row=>`<tr>${headers.map(h=>`<td>${escapeHtml(row[h]??'')}</td>`).join('')}</tr>`).join('');};
  const escapeHtml=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const unavailable=(v,why)=>norm(v)||NA;

  function apiBase(){return `${location.origin}/hbys-rs/hbys`;}
  async function api(path){
    if(state.stop)throw Error('Tarama durduruldu'); const c=new AbortController();active.add(c);const timer=setTimeout(()=>c.abort(),16000);
    try {const sep=path.includes('?')?'&':'?';const r=await fetch(`${apiBase()}${path}${sep}_dc=${Date.now()}`,{credentials:'include',signal:c.signal,headers:{Accept:'application/json, text/plain, */*'}});const t=await r.text();if(!r.ok)throw Error(`HTTP ${r.status}`);return JSON.parse(t);}
    catch(e){throw Error(e.name==='AbortError'?'İstek zaman aşımı':e.message);}
    finally{clearTimeout(timer);active.delete(c);}
  }
  async function attempt(paths){for(const p of paths){try{const x=await api(p);if(x?.success!==false)return {payload:x,path:p};}catch{}}return null;}
  const rows=p=>Array.isArray(p)?p:Array.isArray(p?.data)?p.data:Array.isArray(p?.rows)?p.rows:Array.isArray(p?.list)?p.list:p?.data&&typeof p.data==='object'?[p.data]:p&&typeof p==='object'?[p]:[];
  function flat(o,p='',d=0,out={}){if(!o||typeof o!=='object'||d>8)return out;for(const[k,v]of Object.entries(o)){const q=p?`${p}.${k}`:k;if(v&&typeof v==='object')flat(v,q,d+1,out);else if(v!=null&&norm(v))out[q]=clean(v);}return out;}
  function val(o,names){const wanted=names.map(keyName);for(const[p,v]of Object.entries(flat(o))){const k=keyName(p.split('.').pop());if(wanted.includes(k)&&norm(v))return clean(v);}return '';}
  const text=o=>uniq(Object.values(flat(o)).map(clean)).join(' | ');
  const number=v=>{const n=Number(String(v??'').replace(',','.').replace(/[^0-9.-]/g,''));return Number.isFinite(n)?n:null;};
  const dateOf=v=>{const m=norm(v).match(/(\d{2})\.(\d{2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);return m?new Date(+m[3],+m[2]-1,+m[1],+(m[4]||0),+(m[5]||0)):null;};
  function keyName(v){return String(v??'').replace(/[^a-zA-Z0-9çğıöşüÇĞİÖŞÜ]/g,'').toLocaleLowerCase('tr-TR');}
  function docs(){const a=[];const visit=d=>{if(!d||a.includes(d))return;a.push(d);for(const f of d.querySelectorAll('iframe,frame'))try{visit(f.contentDocument)}catch{}};visit(document);return a;}
  const visible=el=>{if(!el)return false;const r=el.getBoundingClientRect(),s=el.ownerDocument.defaultView.getComputedStyle(el);return r.width>0&&r.height>0&&s.display!=='none'&&s.visibility!=='hidden';};
  const fireClick=el=>{if(!el)return;for(const type of ['mousedown','mouseup','click'])el.dispatchEvent(new MouseEvent(type,{bubbles:true,cancelable:true,view:el.ownerDocument.defaultView}));};
  const setInput=(el,value)=>{el.focus();const setter=Object.getOwnPropertyDescriptor(Object.getPrototypeOf(el),'value')?.set;if(setter)setter.call(el,value);else el.value=value;el.dispatchEvent(new Event('input',{bubbles:true}));el.dispatchEvent(new Event('change',{bubbles:true}));el.dispatchEvent(new KeyboardEvent('keyup',{bubbles:true,key:'0'}));};
  const waitFor=async(fn,timeout=12000)=>{const end=Date.now()+timeout;while(Date.now()<end){if(state.stop)throw Error('Tarama durduruldu');try{const x=fn();if(x)return x;}catch{}await sleep(180);}return null;};
  function searchControls(){for(const d of docs()){const buttons=[...d.querySelectorAll('button,a,[role="button"],.x-btn')].filter(visible);const query=buttons.find(x=>up(x.innerText)==='SORGULA'),clear=buttons.find(x=>up(x.innerText)==='TEMİZLE');if(!query||!clear)continue;let inputs=[...d.querySelectorAll('input')].filter(x=>visible(x)&&/Kimlik No/i.test(norm(x.closest('.x-field,.x-form-item,table')?.innerText)));if(!inputs.length)inputs=[...d.querySelectorAll('input')].filter(x=>visible(x)&&!x.disabled&&!x.readOnly);const tcInput=inputs.find(x=>!x.disabled&&!x.readOnly&&(!x.value||/^\d{0,11}$/.test(x.value)));if(tcInput)return {query,clear,tcInput};}return null;}
  function openOperations(){let best=[];for(const d of docs()){const Ext=d.defaultView.Ext;if(!Ext?.ComponentQuery)continue;for(const g of Ext.ComponentQuery.query('gridpanel')){const s=g.getStore?.(),sample=s?.getAt?.(0)?.data||{},keys=Object.keys(sample).map(keyName);if(s?.getCount?.()&&keys.some(k=>/islemno|ameliyatno/.test(k))&&keys.some(k=>/gelisid|birimsevkid/.test(k))){const a=s.getRange().map((r,i)=>({raw:r.data,index:i}));if(a.length>best.length)best=a;}}}return best;}
  function operation(raw,index=0){const all=text(raw);const name=val(raw,['adiSoyadi','adSoyad','hastaAdiSoyadi'])||'';const [first,...rest]=name.split(' ');return {index,raw,first,last:rest.join(' '),tc:val(raw,['tcKimlikNo','kimlikNo','tckn']),operationNo:val(raw,['islemNo','ameliyatNo','protokolNo'])||(all.match(/\b\d{6,}\b/g)||[]).at(-1)||'',date:val(raw,['ameliyatTarihi','islemTarihi','tarih'])||(all.match(/\d{2}\.\d{2}\.\d{4}(?:\s+\d{1,2}:\d{2})?/)||[])[0]||'',gelisId:val(raw,['gelisId','hastaGelisId']),birimSevkId:val(raw,['birimSevkId','hastaBirimSevkId']),ameliyatId:val(raw,['id','ameliyatId'])};}
  function findByInput(input,list){const tc=norm(input.tc), op=norm(input.operationNo);return list.find(x=>(tc&&x.tc===tc)||(op&&x.operationNo===op))||null;}
  function nearestOperation(list,targetDate){if(!list.length)return null;const target=dateOf(targetDate);if(!target)return list[0];return [...list].sort((a,b)=>Math.abs((dateOf(a.date)||new Date(0))-target)-Math.abs((dateOf(b.date)||new Date(0))-target))[0];}
  async function resolveFromFonet(p){
    if(p.gelisId&&p.birimSevkId)return p;
    if(!/^\d{11}$/.test(norm(p.tc)))throw Error('Geçerli 11 haneli TC bulunamadı');
    const controls=searchControls();if(!controls)throw Error('FONET Ameliyat arama ekranındaki Kimlik No/Sorgula alanları bulunamadı');
    const before=openOperations().map(x=>operation(x.raw)).map(x=>`${x.operationNo}|${x.gelisId}`).join(';');
    fireClick(controls.clear);await sleep(180);setInput(controls.tcInput,p.tc);fireClick(controls.query);
    const found=await waitFor(()=>{const list=openOperations().map((x,i)=>operation(x.raw,i));const signature=list.map(x=>`${x.operationNo}|${x.gelisId}`).join(';');return list.length&&(signature!==before||list.some(x=>x.operationNo===p.operationNo))?list:null;},15000);
    if(!found?.length)throw Error('TC için FONET ameliyat kaydı bulunamadı veya sorgu 15 saniyede tamamlanmadı');
    const selected=(p.operationNo&&found.find(x=>x.operationNo===p.operationNo))||nearestOperation(found,p.date);
    if(!selected?.gelisId||!selected?.birimSevkId)throw Error('Bulunan ameliyat satırında hasta geliş/birim sevk kimliği yok');
    Object.assign(p,selected,{tc:p.tc||selected.tc,first:p.first||selected.first,last:p.last||selected.last,date:p.date||selected.date,operationNo:p.operationNo||selected.operationNo});
    return p;
  }

  async function loadSheet(){
    const f=$('#fga-file').files[0];if(!f)throw Error('Önce Excel/CSV dosyasını seçin.');
    if(!window.XLSX){await new Promise((ok,bad)=>{const s=document.createElement('script');s.src='https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';s.onload=ok;s.onerror=()=>bad(Error('Excel kütüphanesi yüklenemedi'));document.head.append(s);});}
    const b=await f.arrayBuffer(), wb=XLSX.read(b,{type:'array'}), data=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
    const source=openOperations(); state.patients=data.map((r,i)=>{const x={tc:val(r,['tc','tcKimlikNo','tckn']),operationNo:val(r,['islemNo','islemno','ameliyatNo']),first:val(r,['adi','ad']),last:val(r,['soyadi','soyad']),date:val(r,['ameliyatTarihi','tarih'])};return {...(findByInput(x,source)||{}),...x,inputIndex:i};}).filter(x=>x.tc||x.operationNo);
    if(!state.patients.length)throw Error('Dosyada TC veya İşlem No sütunu bulunamadı.'); prepared(`${f.name}: ${state.patients.length} hasta hazır.`);
  }
  function prepared(msg){state.results=[];state.done=0;state.errors=0;$('#fga-start').disabled=false;$('#fga-csv').disabled=true;log(msg);render();}

  function classify(d){
    const all=[...d.imaging,...d.pathology,...d.orders,...d.consults,...d.services,...d.history].join('\n');
    const imaging=d.imaging.filter(x=>/\bUSG\b|ULTRASON|\bMR\b|MANYETİK|\bBT\b|TOMOGRAF|GÖRÜNTÜLEME/i.test(x));
    const path=d.pathology.filter(x=>/APEND|APPEND/i.test(x));
    const lab=extractLabs(d.labs);
    const trimester=(all.match(/(?:gebelik\s*)?(?:trimester|trimesterı|trimesteri)\s*[:=-]?\s*([123])/i)||[])[1]||(/\b([123])\.?\s*trimester/i.exec(all)||[])[1]||'';
    const surgical=/laparoskop/i.test(all)?'Laparoskopik':/açık|acik/i.test(all)?'Açık':'';
    const ptxt=path.join(' | '), itxt=imaging.join(' | ');
    const negative=/appendix\s+vermiformis.{0,100}(?:normal|doğal|dogal)|serozal|kataral/i.test(ptxt)?'Evet':/apandisit|appendisit/i.test(ptxt)?'Hayır':'';
    const type=/perfor/i.test(ptxt)?'Perfore apandisit':/flegmon/i.test(ptxt)?'Flegmon':/basit|akut apandisit/i.test(ptxt)?'Basit apandisit':'';
    const prePositive=/apandisit|appendisit/i.test(itxt), pathPositive=/apandisit|appendisit/i.test(ptxt);
    const concordance=(imaging.length&&path.length)?(prePositive===pathPositive?'1':'0'):'';
    const find=rx=>uniq(all.split(/[\n|]/).filter(x=>rx.test(x))).join(' | ');
    return {imaging:imaging.join('\n'),path:ptxt,negative,concordance,surgical,type,trimester,
      maternal:find(/kanama|enfeksiyon|sepsis|ileus|pnömoni|maternal.*komplikasyon/i), fetal:find(/fetal.*kayıp|abort|erken doğum|prematür|fetal.*komplikasyon/i),
      sso:find(/cerrahi alan|\bSSO\b|yara yeri|seroma|hematom|dehisans|enfeksiyon/i), dvt:find(/\bDVT\b|derin ven tromboz|pulmoner embol/i),
      tocolysis:find(/tokoliz|tokolitik/i),tocolyticDrug:find(/nifedipin|atosiban|indometazin|magnezyum|terbutalin|ritodrin/i),nst:find(/kontraksiyon|\bNST\b.*pozitif|NST.*reaktif/i),birth:find(/canlı doğum|canli dogum|doğum gerçekleş/i),analgesia:countAnalgesia(d.services.concat(d.orders)),...lab};
  }
  function extractLabs(items){const raw=items.join(' | '), hit=(names)=>{for(const n of names){const m=new RegExp(`(?:${n})\\s*[:=]?\\s*(\\d+(?:[.,]\\d+)?)`,'i').exec(raw);if(m)return number(m[1]);}return null;};const w=hit(['WBC','LÖKOSİT','LOKOSIT']),c=hit(['CRP']),neu=hit(['NÖTROFİL','NOTROFIL','NEU']),lym=hit(['LENFOSİT','LENFOSIT','LYM']),plt=hit(['PLT','TROMBOSİT','TROMBOSIT']);return {wbc:w??'',crp:c??'',nlr:neu!=null&&lym?round(neu/lym):'',plr:plt!=null&&lym?round(plt/lym):'',sii:plt!=null&&neu!=null&&lym?round(plt*neu/lym):''};}
  const round=n=>Math.round(n*100)/100;
  function countAnalgesia(items){const a=items.filter(x=>/parasetamol|paracetamol|tramadol|diklofenak|ibuprofen|morfin|analjezik/i.test(x));return a.length?String(a.length):'';}

  async function patientData(p){
    await resolveFromFonet(p);
    const detail=await attempt([`/Ameliyat/Ameliyat/getKayit/${encodeURIComponent(p.birimSevkId)}`]);if(!detail)throw Error('Ameliyat ayrıntısı okunamadı');
    const root=detail.payload.data||detail.payload, visit=root?.birimSevk?.hastaGelis||root?.hastaGelis||{}, patient=visit?.hasta||root?.hasta||{}, identity=patient?.kimlik||{};
    const verifiedGelis=String(visit.id||p.gelisId);if(String(p.gelisId)!==verifiedGelis)throw Error('Hasta geliş kimliği doğrulanamadı');
    const verifiedTc=norm(identity.tcKimlikNo||identity.kimlikNo||identity.tckn);if(p.tc&&verifiedTc&&norm(p.tc)!==verifiedTc)throw Error('FONET servisindeki TC, Excel satırıyla eşleşmedi; yanlış hastaya veri yazılmadı');
    const patientId=patient.id||p.hastaId;
    const filter=encodeURIComponent(JSON.stringify([{property:'hastaId',value:Number(patientId)||patientId,type:'Long',operator:'='}]));
    const [panel,consult,services,stay,rad,lab,pathology,orders,history]=await Promise.all([
      attempt([`/HastaKabul/HastaGelis/getHastaGelisPanelInfo/${p.gelisId}`]),attempt([`/Poliklinik/Poliklinik/getHastaGelisKonsultasyonList/${p.gelisId}/1`]),attempt([`/Tibbi/HastaHizmet/getHizmetList/${p.birimSevkId}/${p.gelisId}`]),attempt([`/Klinik/Klinik/getAmeliyathaneSevkList/${p.gelisId}?start=0&limit=2000&page=1`]),
      patientId?attempt([`/Ris/RisHizmetSonuc/getRisHizmetSonucInfoList?start=0&limit=1000&page=1&filter=${filter}`]):null,
      attempt([`/Laboratuvar/LaboratuvarSonuc/getHastaSonucList/${p.gelisId}`,`/Laboratuvar/LaboratuvarSonuc/getSonucList/${p.gelisId}`]),
      attempt([`/Patoloji/PatolojiSonuc/getHastaGelisList/${p.gelisId}`,`/Patoloji/PatolojiSonuc/getHastaPatolojiList/${p.gelisId}`]),
      attempt([`/Stok/EOrder/getKayitList?start=0&limit=2000&page=1&filter=${encodeURIComponent(JSON.stringify([{property:'hastaGelis.id',value:Number(p.gelisId)||p.gelisId,type:'Long',operator:'='}]))}`,`/Order/Order/getHastaGelisOrderList/${p.gelisId}`,`/Klinik/Order/getHastaGelisOrderList/${p.gelisId}`]),
      patientId?attempt([`/Tibbi/HastaBirimSevk/getKayitList?start=0&limit=1000&page=1&filter=${encodeURIComponent(JSON.stringify([{property:'hastaGelis.hasta.id',value:Number(patientId)||patientId,type:'Long',operator:'='}]))}`]):null
    ]);
    const imaging=[]; if(rad){for(const r of rows(rad.payload)){if(!r.raporId)continue;try{const report=await api(`/Ris/RisHizmetSonuc/getRisRaporSonucByRaporId/${encodeURIComponent(r.raporId)}`);const x=clean(report?.data?.raporTextByRapor||report?.data?.bulgular||'');if(x)imaging.push(`${text(r)} | ${x}`);}catch{}}}
    const arr=x=>x?rows(x.payload).map(text):[];
    const start=val(root,['ameliyatBaslangicTarihi','ameliyatBaslamaTarihi','ameliyatBaslangicSaati']);const end=val(root,['ameliyatBitisTarihi','ameliyatBitisSaati','ameliyatSonSaati']);let minutes='';if(dateOf(start)&&dateOf(end))minutes=Math.round((dateOf(end)-dateOf(start))/60000);
    const discharge=val(panel?.payload?.data||panel?.payload||{},['taburcuTarihi','cikisTarihi']);const stayDays=dateOf(p.date)&&dateOf(discharge)?round((dateOf(discharge)-dateOf(p.date))/86400000):'';
    return {root,identity,patient,imaging,pathology:arr(pathology),orders:arr(orders),consults:arr(consult),services:arr(services),history:arr(history),labs:arr(lab),stayDays,minutes,discharge};
  }
  function makeRow(p,d){const c=classify(d), first=val(d.identity,['adi','ad'])||p.first,last=val(d.identity,['soyadi','soyad'])||p.last,tc=val(d.identity,['tcKimlikNo','kimlikNo','tckn'])||p.tc;const birth=val(d.identity,['dogumTarihi']);const age=dateOf(birth)?new Date().getFullYear()-dateOf(birth).getFullYear():val(d.root,['yas','yaş']);const asa=val(d.root,['asaSkoruAdi','asaSinifi','asa']);const evidence=uniq([c.imaging&&`Görüntüleme: ${c.imaging}`,c.path&&`Patoloji: ${c.path}`,d.orders.length&&`Order: ${d.orders.join(' | ')}`,d.labs.length&&`Lab: ${d.labs.join(' | ')}`]).join('\n').slice(0,30000);const f=v=>unavailable(v);return {'Adı':f(first),'Soyadı':f(last),'TC':f(tc),'İşlem No':f(p.operationNo),'Telefon':f(val(d.identity,['telefon','cepTelefonu','gsm'])||val(d.patient,['telefon','cepTelefonu'])),'Yaş':f(age),'ASA':f(asa),'Ameliyat tarihi':f(p.date),'Ameliyat süresi (dk)':f(d.minutes),'Yatış süresi':f(d.stayDays),'Kaçıncı trimester':f(c.trimester),'Preoperatif görüntüleme':f(c.imaging),'Patoloji':f(c.path),'Negatif apendektomi mi?':f(c.negative),'Preoperatif görüntüleme / patoloji uyumu':f(c.concordance),'Cerrahi yöntem':f(c.surgical),'Apandisit tipi':f(c.type),'Maternal komplikasyon':f(c.maternal),'Fetal komplikasyon':f(c.fetal),'Postoperatif SSO':f(c.sso),'Postoperatif analjezi ihtiyacı (doz)':f(c.analgesia),'DVT gelişimi':f(c.dvt),'Tokoliz ihtiyacı':f(c.tocolysis),'Tokolizde kullanılan ilaç':f(c.tocolyticDrug),'Postoperatif kontraksiyon / NST pozitifliği':f(c.nst),'İlerleyen süreçte canlı doğum':f(c.birth),'Preoperatif WBC':f(c.wbc),'Preoperatif CRP':f(c.crp),'Preoperatif NLR':f(c.nlr),'Preoperatif PLR':f(c.plr),'Preoperatif SII':f(c.sii),'KAYNAK / NOT':f(evidence),'Tarama durumu':'Tamamlandı'};}
  async function run(){state.running=true;state.stop=false;state.paused=false;$('#fga-start').disabled=true;$('#fga-pause').disabled=false;$('#fga-stop').disabled=false;for(const p of state.patients){if(state.stop)break;while(state.paused&&!state.stop)await sleep(250);try{log(`${p.operationNo||p.tc} — FONET'te hasta ve tüm kayıtlar aranıyor...`);state.results.push(makeRow(p,await patientData(p)));}catch(e){state.errors++;const r=Object.fromEntries(headers.map(h=>[h,NA]));r['Adı']=p.first||NA;r['Soyadı']=p.last||NA;r['İşlem No']=p.operationNo||NA;r.TC=p.tc||NA;r['KAYNAK / NOT']=`Tarama hatası: ${e.message}`;r['Tarama durumu']=`Hata: ${e.message}`;state.results.push(r);}state.done++;render();status();}state.running=false;$('#fga-pause').disabled=true;$('#fga-stop').disabled=true;$('#fga-csv').disabled=!state.results.length;status(state.stop?'Tarama durduruldu. Mevcut sonuçlar indirilebilir.':`Tarama tamamlandı. ${state.errors} hata var. Hata nedeni KAYNAK / NOT ve Tarama durumu sütunlarında yazıyor.`);}
  function csv(){const esc=v=>'"'+String(v??'').replace(/"/g,'""')+'"';const out=[headers,...state.results.map(r=>headers.map(h=>r[h]??''))].map(r=>r.map(esc).join(';')).join('\r\n');const a=Object.assign(document.createElement('a'),{href:URL.createObjectURL(new Blob(['\uFEFF'+out],{type:'text/csv;charset=utf-8'})),download:`fonet-gebelik-apandisit-${new Date().toISOString().slice(0,10)}.csv`});a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1500);}
  $('#fga-close').onclick=()=>{for(const c of active)c.abort();panel.remove();window[APP]=false;};$('#fga-excel').onclick=()=>loadSheet().catch(e=>status(e.message));$('#fga-open').onclick=()=>{const x=openOperations().map((r,i)=>operation(r.raw,i));if(!x.length)return status('Açık ameliyat listesi bulunamadı. İşlem no, hasta geliş ID ve birim sevk ID içeren FONET sorgusunu açın.');state.patients=x;prepared(`${x.length} açık ameliyat kaydı hazır.`);};$('#fga-start').onclick=run;$('#fga-pause').onclick=()=>{state.paused=!state.paused;$('#fga-pause').textContent=state.paused?'Devam Et':'Duraklat';status(state.paused?'Tarama duraklatıldı.':'Tarama sürüyor.');};$('#fga-stop').onclick=()=>{state.stop=true;state.paused=false;for(const c of active)c.abort();};$('#fga-csv').onclick=csv;
})();
