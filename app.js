'use strict';
/* 棚卸アプリ: GitHub Pages(静的サイト)で動作。データは端末のIndexedDBに保存。 */
const $ = id => document.getElementById(id);
const COLS = {code:'資産コード', name:'名称', place:'置場・担当者', type:'型式', gen:'現物', dept:'所属', loc:'所在地', kiban:'原始管理番号', note:'備考'};
const state = {rows:[], header:[], idx:{}, byCode:new Map(), byKiban:new Map(), rec:new Map(), cur:null, meta:null,
  settings:{autoOcr:true, sound:true, photo:true}, sort:{k:'code',d:1}, limit:200, tab:'scan'};

/* ---------------- IndexedDB ---------------- */
const DB = {
  db:null,
  open(){ return new Promise((res,rej)=>{ const r=indexedDB.open('tanaoroshi',1);
    r.onupgradeneeded=()=>{ const d=r.result; d.createObjectStore('meta'); d.createObjectStore('records',{keyPath:'code'}); const p=d.createObjectStore('photos',{keyPath:'name'}); p.createIndex('code','code'); };
    r.onsuccess=()=>{this.db=r.result;res()}; r.onerror=()=>rej(r.error); }); },
  _s(store,mode){ return this.db.transaction(store,mode).objectStore(store); },
  _p(req){ return new Promise((res,rej)=>{req.onsuccess=()=>res(req.result);req.onerror=()=>rej(req.error)}); },
  get(s,k){ return this._p(this._s(s,'readonly').get(k)); },
  put(s,v,k){ return this._p(this._s(s,'readwrite').put(v,k)); },
  del(s,k){ return this._p(this._s(s,'readwrite').delete(k)); },
  all(s){ return this._p(this._s(s,'readonly').getAll()); },
  byCode(code){ return this._p(this._s('photos','readonly').index('code').getAll(code)); },
  clear(s){ return this._p(this._s(s,'readwrite').clear()); }
};

/* ---------------- 共通 ---------------- */
const norm = s => String(s==null?'':s).normalize('NFKC').trim().toUpperCase();
const pad = (n,l=2)=>String(n).padStart(l,'0');
const fmtTime = t => { if(!t) return ''; const d=new Date(t); return `${d.getMonth()+1}/${d.getDate()} ${pad(d.getHours())}:${pad(d.getMinutes())}`; };
const fmtFull = t => { const d=new Date(t); return `${d.getFullYear()}/${pad(d.getMonth()+1)}/${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`; };
function pop(msg){ const p=$('pop'); p.textContent=msg; p.hidden=false; clearTimeout(pop.t); pop.t=setTimeout(()=>p.hidden=true,2200); }
function beep(ok=true){
  if(!state.settings.sound) return;
  try{ navigator.vibrate&&navigator.vibrate(ok?60:[40,40,40]); }catch(e){}
  try{ const a=beep.ctx||(beep.ctx=new (window.AudioContext||window.webkitAudioContext)()); const o=a.createOscillator(),g=a.createGain();
    o.frequency.value=ok?1000:300; g.gain.value=.15; o.connect(g); g.connect(a.destination); o.start(); o.stop(a.currentTime+(ok?.1:.25)); }catch(e){}
}
function cellVal(v){ if(v==null) return ''; if(typeof v==='object'){ if(v.richText) return v.richText.map(t=>t.text).join(''); if(v.result!=null) return v.result; if(v.text!=null) return v.text; if(v instanceof Date) return v.toISOString().slice(0,10).replace(/-/g,'/'); } return v; }

/* ---------------- リスト読込 ---------------- */
async function importList(file){
  const wb=new ExcelJS.Workbook(); await wb.xlsx.load(await file.arrayBuffer());
  let ws=wb.getWorksheet('最新データ');
  if(!ws) ws=wb.worksheets.find(w=>{const v=w.getRow(1).getCell(1).value; return cellVal(v)==='資産コード';});
  if(!ws) throw new Error('「資産コード」列を持つシートが見つかりません');
  const header=[]; const hr=ws.getRow(1); const ncol=ws.columnCount;
  for(let c=1;c<=ncol;c++) header.push(String(cellVal(hr.getCell(c).value)||''));
  const rows=[];
  ws.eachRow({includeEmpty:false},(row,n)=>{ if(n===1) return; const arr=[]; for(let c=1;c<=ncol;c++) arr.push(cellVal(row.getCell(c).value)); if(arr[0]) rows.push(arr); });
  await DB.put('meta',{header,rows,filename:file.name,sheet:ws.name,at:Date.now()},'list');
  await loadList(); renderAll(); await distribute();
}
async function loadList(){
  const m=await DB.get('meta','list'); state.meta=m||null;
  state.rows=[];state.byCode.clear();state.byKiban.clear();
  if(!m){ return; }
  state.header=m.header; state.idx={}; m.header.forEach((h,i)=>{ if(!(h in state.idx)) state.idx[h]=i; });
  const g=(r,k)=>{const i=state.idx[COLS[k]]; return i==null?'':(r[i]==null?'':r[i]);};
  m.rows.forEach((raw,i)=>{ const o={i,raw,code:norm(g(raw,'code')),name:String(g(raw,'name')),place:String(g(raw,'place')),type:String(g(raw,'type')),
    gen0:String(g(raw,'gen')),dept:String(g(raw,'dept')),loc:String(g(raw,'loc')),kiban:norm(g(raw,'kiban')),note:String(g(raw,'note'))};
    o.search=(o.code+' '+o.name+' '+o.kiban+' '+o.place+' '+o.type).toUpperCase().normalize('NFKC');
    state.rows.push(o); state.byCode.set(o.code,o); if(o.kiban) state.byKiban.set(o.kiban,o); });
  const recs=await DB.all('records'); state.rec=new Map(recs.map(r=>[r.code,r]));
}
const genOf = r => { const rc=state.rec.get(r.code); return rc&&rc.gen!=null?rc.gen:r.gen0; };
const kibanFile = r => '機番'; /* 写真名は 資産コード_機番.jpg(「機番」は固定の文字) */

/* ---------------- 位置情報 ---------------- */
let lastPos=null;
function startGeo(){
  if(!navigator.geolocation) return;
  try{ navigator.geolocation.watchPosition(p=>{lastPos={lat:p.coords.latitude,lng:p.coords.longitude,acc:p.coords.accuracy,ts:Date.now()}},()=>{},{enableHighAccuracy:true,maximumAge:5000,timeout:20000}); }catch(e){}
}
function getPos(){
  if(lastPos&&Date.now()-lastPos.ts<30000) return Promise.resolve(lastPos);
  return new Promise(res=>{ if(!navigator.geolocation) return res(lastPos);
    navigator.geolocation.getCurrentPosition(p=>{lastPos={lat:p.coords.latitude,lng:p.coords.longitude,acc:p.coords.accuracy,ts:Date.now()};res(lastPos)},()=>res(lastPos),{enableHighAccuracy:true,timeout:3000,maximumAge:30000}); });
}

/* ---------------- カメラ ---------------- */
const video=$('video'); let stream=null, camOn=false, torchOn=false, detLoop=null, lastDetect=Date.now();
const zxCanvas=document.createElement('canvas'); let detector=null, zxReader=null, zxBusy=false;
async function startCam(){
  if(camOn) return;
  try{
    stream=await navigator.mediaDevices.getUserMedia({video:{facingMode:{ideal:'environment'},width:{ideal:1920},height:{ideal:1080}},audio:false});
    video.srcObject=stream; await video.play(); camOn=true; lastDetect=Date.now();
    $('btnStartCam').hidden=true; $('camState').textContent='読取中'; $('hud').textContent='QR・バーコードを枠に合わせてください'; $('btnCamToggle').textContent='停止';
    const tr=stream.getVideoTracks()[0]; const cap=tr.getCapabilities?tr.getCapabilities():{};
    $('btnTorch').hidden=!cap.torch;
    try{ if(cap.focusMode&&cap.focusMode.includes('continuous')) await tr.applyConstraints({advanced:[{focusMode:'continuous'}]}); }catch(e){}
    try{ if('wakeLock' in navigator) navigator.wakeLock.request('screen').catch(()=>{}); }catch(e){}
    initDetector(); scheduleDetect();
  }catch(e){
    $('camState').textContent='カメラを使えません'; $('btnStartCam').hidden=false;
    $('hud').textContent='カメラの許可が必要です。手入力は使えます。';
  }
}
function stopCam(){
  camOn=false; clearTimeout(detLoop);
  if(stream){ stream.getTracks().forEach(t=>t.stop()); stream=null; }
  video.srcObject=null; $('camState').textContent='停止中'; $('btnCamToggle').textContent='開始'; $('btnTorch').hidden=true; torchOn=false;
}
async function initDetector(){
  if(detector||zxReader) return;
  if('BarcodeDetector' in window){
    try{ const fm=await BarcodeDetector.getSupportedFormats(); const want=['qr_code','code_128','code_39','code_93','ean_13','ean_8','itf','codabar','data_matrix','upc_a','upc_e'].filter(f=>fm.includes(f));
      if(want.length){ detector=new BarcodeDetector({formats:want}); return; } }catch(e){}
  }
  const Z=ZXing, hints=new Map(); hints.set(Z.DecodeHintType.POSSIBLE_FORMATS,[Z.BarcodeFormat.QR_CODE,Z.BarcodeFormat.CODE_128,Z.BarcodeFormat.CODE_39,Z.BarcodeFormat.CODE_93,Z.BarcodeFormat.EAN_13,Z.BarcodeFormat.EAN_8,Z.BarcodeFormat.ITF,Z.BarcodeFormat.CODABAR,Z.BarcodeFormat.DATA_MATRIX]);
  hints.set(Z.DecodeHintType.TRY_HARDER,true); zxReader=new Z.MultiFormatReader(); zxReader.setHints(hints);
}
function scheduleDetect(){ clearTimeout(detLoop); if(camOn) detLoop=setTimeout(detectOnce,detector?110:160); }
async function detectOnce(){
  if(!camOn||document.hidden||state.tab!=='scan'||sheetOpen()){ return scheduleDetect(); }
  try{
    let text=null;
    if(video.readyState>=2){
      if(detector){ const r=await detector.detect(video); if(r.length) text=r[0].rawValue; }
      else if(zxReader){
        const vw=video.videoWidth, vh=video.videoHeight, sc=Math.min(1,1000/Math.max(vw,vh));
        zxCanvas.width=Math.round(vw*sc); zxCanvas.height=Math.round(vh*sc);
        const ctx=zxCanvas.getContext('2d',{willReadFrequently:true}); ctx.drawImage(video,0,0,zxCanvas.width,zxCanvas.height);
        const Z=ZXing; try{ const bb=new Z.BinaryBitmap(new Z.HybridBinarizer(new Z.HTMLCanvasElementLuminanceSource(zxCanvas))); text=zxReader.decodeWithState(bb).getText(); }catch(e){}
      }
    }
    if(text) { lastDetect=Date.now(); await onText(text,'QR/バーコード'); }
    else if(state.settings.autoOcr&&Date.now()-lastDetect>5000&&!ocrBusy&&!sheetOpen()){ lastDetect=Date.now(); runOcr(true); }
  }catch(e){}
  scheduleDetect();
}
const sheetOpen=()=>!$('sheet').hidden||!$('ocrSheet').hidden||!$('detail').hidden;
function grabFrame(maxSide=1920,q=0.85){
  return new Promise(res=>{ if(!camOn||video.readyState<2) return res(null);
    const vw=video.videoWidth, vh=video.videoHeight, sc=Math.min(1,maxSide/Math.max(vw,vh)); const c=document.createElement('canvas'); c.width=Math.round(vw*sc); c.height=Math.round(vh*sc);
    c.getContext('2d').drawImage(video,0,0,c.width,c.height); c.toBlob(b=>res(b),'image/jpeg',q); });
}
function flash(){ const f=$('flash'); f.classList.remove('go'); void f.offsetWidth; f.classList.add('go'); }

/* ---------------- コード解決 ---------------- */
const CODE_RE=/[A-Z]{3,4}\d{4}/g;
function resolveText(text){
  const t=norm(text), found=new Map();
  const add=(r)=>{ if(r) found.set(r.code,r); };
  add(state.byCode.get(t)); add(state.byKiban.get(t));
  (t.match(CODE_RE)||[]).forEach(c=>add(state.byCode.get(c)));
  t.split(/[^A-Z0-9]+/).forEach(tok=>{ if(tok.length>=8) add(state.byKiban.get(tok)); });
  return {rows:[...found.values()], raw:(t.match(CODE_RE)||[])};
}
const recentSeen=new Map();
async function onText(text,method){
  if(String(text).startsWith('TNPW1:')){ const pw=String(text).slice(6).trim(); const now=Date.now(); if(recentSeen.get('pw'+pw)>now-8000) return; recentSeen.set('pw'+pw,now); return receiveList(pw); }
  const {rows,raw}=resolveText(text);
  if(rows.length){ const r=rows[0]; const now=Date.now(); if(recentSeen.get(r.code)>now-4000) return; recentSeen.set(r.code,now); return recordScan(r,method); }
  if(raw.length){ const c=raw[0], now=Date.now(); if(recentSeen.get(c)>now-6000) return; recentSeen.set(c,now); beep(false); showToast(`リストにありません: ${c}`,'bad',{label:'リスト外として登録',fn:()=>recordUnlisted(c,method)}); }
  else { const now=Date.now(); if(recentSeen.get('x'+norm(text))>now-6000) return; recentSeen.set('x'+norm(text),now); $('hud').textContent='対象外のコード: '+String(text).slice(0,30); }
}

/* ---------------- 記録 ---------------- */
async function savePhoto(code,name,blob,kind){ await DB.put('photos',{name,code,blob,kind,time:Date.now()}); }
function nextSeq(rec){ let m=0; (rec.photos||[]).forEach(n=>{const x=n.match(/_(\d{3})\.jpg$/); if(x) m=Math.max(m,+x[1]);}); return pad(m+1,3); }
async function recordScan(r,method,{noPhoto=false}={}){
  let rec=state.rec.get(r.code);
  const isNew=!rec||!rec.time;
  if(!isNew){ state.cur=r; showCur(); beep(true); showToast(`記録済み: ${r.code}`,'warn'); return; }
  const blob=(state.settings.photo&&!noPhoto)?await grabFrame():null;
  const pos=await getPos();
  rec=Object.assign(rec||{code:r.code,photos:[]},{gen:'有',kiban:r.kiban,method,time:Date.now(),lat:pos?pos.lat:null,lng:pos?pos.lng:null,acc:pos?Math.round(pos.acc):null,listed:true});
  if(blob){ const name=`${r.code}_${kibanFile(r)}.jpg`; await savePhoto(r.code,name,blob,'main'); rec.photos=[name,...(rec.photos||[]).filter(n=>n!==name)]; flash(); }
  await DB.put('records',rec); state.rec.set(r.code,rec); state.cur=r; beep(true);
  showCur(); showToast(`${r.code} を記録しました`,'ok'); renderProgress(); invalidateList();
}
async function recordUnlisted(code,method){
  const rec={code,gen:'有',kiban:'',method,time:Date.now(),photos:[],listed:false};
  const pos=await getPos(); if(pos){rec.lat=pos.lat;rec.lng=pos.lng;rec.acc=Math.round(pos.acc);}
  const blob=state.settings.photo?await grabFrame():null;
  if(blob){ const name=`${code}_機番.jpg`; await savePhoto(code,name,blob,'main'); rec.photos=[name]; flash(); }
  await DB.put('records',rec); state.rec.set(code,rec);
  state.cur={i:-1,code,name:'(リスト外)',place:'',kiban:'',gen0:'',dept:'',loc:'',unlisted:true}; showCur(); showToast(`${code} をリスト外として記録`,'warn'); invalidateList();
}
async function addPhotoForCurrent(){
  const r=state.cur; if(!r) return; const rec=state.rec.get(r.code); if(!rec) return;
  const blob=await grabFrame(); if(!blob){ pop('カメラが動いていません'); return; }
  const name=`${r.code}_${nextSeq(rec)}.jpg`; await savePhoto(r.code,name,blob,'extra'); rec.photos.push(name);
  await DB.put('records',rec); flash(); beep(true); showCur(); invalidateList(); pop(`保存: ${name}`);
}
async function setGen(r,v){
  let rec=state.rec.get(r.code);
  if(v===''){ if(rec){ if(rec.time||rec.photos.length){ rec.gen=null; await DB.put('records',rec);} else { await DB.del('records',r.code); state.rec.delete(r.code);} } }
  else {
    if(!rec){ rec={code:r.code,photos:[],kiban:r.kiban,method:'手動',listed:!r.unlisted}; const pos=await getPos(); if(pos){rec.lat=pos.lat;rec.lng=pos.lng;rec.acc=Math.round(pos.acc);} rec.time=Date.now(); }
    rec.gen=v; await DB.put('records',rec); state.rec.set(r.code,rec);
  }
  renderProgress(); invalidateList(); if(state.cur&&state.cur.code===r.code) showCur();
}
async function undoCurrent(){
  const r=state.cur; if(!r) return; const rec=state.rec.get(r.code); if(!rec) return;
  const ps=await DB.byCode(r.code); for(const p of ps) await DB.del('photos',p.name);
  await DB.del('records',r.code); state.rec.delete(r.code); recentSeen.delete(r.code);
  state.cur=null; $('cur').hidden=true; showToast(`${r.code} の記録を取り消しました`,'warn'); renderProgress(); invalidateList();
}

/* ---------------- スキャン画面UI ---------------- */
function showToast(msg,kind='ok',action){
  const t=$('toast'); t.className='toast '+kind; t.hidden=false; t.innerHTML=''; const s=document.createElement('span'); s.textContent=msg; t.appendChild(s);
  if(action){ const b=document.createElement('button'); b.className='btn'; b.textContent=action.label; b.onclick=()=>{t.hidden=true;action.fn()}; t.appendChild(b); }
  clearTimeout(showToast.t); showToast.t=setTimeout(()=>{t.hidden=true},action?7000:3500);
}
function showCur(){
  const r=state.cur; if(!r){ $('cur').hidden=true; return; }
  const rec=state.rec.get(r.code)||{photos:[]};
  $('cur').hidden=false; $('curCode').textContent=r.code; $('curName').textContent=r.name; $('curSub').textContent=[r.place,r.dept].filter(Boolean).join(' / ');
  $('curMethod').textContent=rec.method||''; const g=genOf(r.unlisted?{code:r.code,gen0:''}:r);
  [...$('curGen').children].forEach(b=>b.classList.toggle('on',b.dataset.v===g));
  const np=(rec.photos||[]).length; $('curPhotos').textContent=np?`写真 ${np}枚 (${rec.photos[rec.photos.length-1]})`:'写真なし';
  $('curLoc').textContent=rec.lat!=null?`位置 ${rec.lat.toFixed(5)}, ${rec.lng.toFixed(5)} (±${rec.acc}m) ${fmtTime(rec.time)}`:(rec.time?'位置情報なし':'');
}
function renderProgress(){
  const n=state.rows.length; let done=0; state.rows.forEach(r=>{ const rc=state.rec.get(r.code); if(rc&&rc.time&&rc.listed!==false) done++; });
  $('pDone').textContent=done; $('pAll').textContent=n; $('pBar').style.width=n?(done/n*100)+'%':'0';
}

/* ---------------- 手入力 ---------------- */
function openManual(){ $('sheet').hidden=false; $('manualIn').value=''; renderSug(''); setTimeout(()=>$('manualIn').focus(),50); }
function renderSug(q){
  const ul=$('manualSug'); ul.innerHTML=''; const t=norm(q); if(!t) return;
  let hits=[]; const ex=state.byCode.get(t)||state.byKiban.get(t); if(ex) hits.push(ex);
  for(const r of state.rows){ if(hits.length>=8) break; if(r!==ex&&r.search.includes(t)) hits.push(r); }
  hits.forEach(r=>ul.appendChild(sugItem(r,()=>{ $('sheet').hidden=true; recordScan(r,'手入力'); })));
  if(!hits.length&&/^[A-Z]{3,4}\d{4}$/.test(t)){ const li=document.createElement('li'); const b=document.createElement('button'); b.innerHTML=`<b>${t}</b><span>リストにありません。リスト外として登録</span>`; b.onclick=()=>{$('sheet').hidden=true;recordUnlisted(t,'手入力')}; li.appendChild(b); ul.appendChild(li); }
}
function sugItem(r,fn){ const li=document.createElement('li'),b=document.createElement('button'); const rc=state.rec.get(r.code);
  b.innerHTML=''; const a=document.createElement('b'); a.textContent=r.code+(rc&&rc.time?'  (記録済)':''); const s=document.createElement('span'); s.textContent=[r.name,r.place].filter(Boolean).join(' / ');
  b.append(a,s); b.onclick=fn; li.appendChild(b); return li; }

/* ---------------- OCR ---------------- */
let ocrWorker=null, ocrBusy=false;
async function getWorker(){
  if(ocrWorker) return ocrWorker;
  if(typeof Tesseract==='undefined') await new Promise((ok,ng)=>{ const s=document.createElement('script'); s.src='vendor/tess/tesseract.min.js'; s.onload=ok; s.onerror=ng; document.head.appendChild(s); });
  const base=new URL('vendor/tess/',location.href).href;
  ocrWorker=await Tesseract.createWorker('eng',1,{workerPath:base+'worker.min.js',corePath:base,langPath:base,gzip:true,workerBlobURL:false});
  await ocrWorker.setParameters({tessedit_char_whitelist:'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789',tessedit_pageseg_mode:'11'});
  return ocrWorker;
}
function fixCode(w){
  const L=w.length, a=w.slice(0,L-4).replace(/0/g,'O').replace(/1/g,'I').replace(/5/g,'S').replace(/8/g,'B').replace(/2/g,'Z'),
        d=w.slice(-4).replace(/[OQD]/g,'0').replace(/[IL]/g,'1').replace(/S/g,'5').replace(/B/g,'8').replace(/Z/g,'2');
  return a+d;
}
function ocrCandidates(text){
  const out=new Map(); const t=norm(text);
  const lines=t.split(/\n/).map(l=>l.replace(/[^A-Z0-9]/g,'')).filter(l=>l.length>=7);
  const toks=t.split(/[^A-Z0-9]+/).filter(x=>x.length>=7);
  for(const s of [...lines,...toks]){
    for(const L of [8,7]) for(let i=0;i+L<=s.length;i++){ const w=s.substr(i,L); for(const c of new Set([w,fixCode(w)])){ const r=state.byCode.get(c); if(r) out.set(r.code,r); } }
    for(let i=0;i+11<=s.length;i++){ const r=state.byKiban.get(s.substr(i,11)); if(r) out.set(r.code,r); }
  }
  return [...out.values()];
}
async function runOcr(auto=false){
  if(ocrBusy) return; ocrBusy=true;
  if(!auto) { $('btnOcr').disabled=true; $('btnOcr').textContent='読み取り中…'; } else $('camState').textContent='OCR中';
  try{
    let blob=await grabFrame(1400,0.9); if(!blob){ if(!auto) pop('カメラが動いていません'); return; }
    const w=await getWorker(); const res=await w.recognize(blob); const text=res.data.text||'';
    const cands=ocrCandidates(text);
    if(cands.length===1){ const r=cands[0]; lastDetect=Date.now(); const now=Date.now(); if(!(recentSeen.get(r.code)>now-4000)){ recentSeen.set(r.code,now); await recordScan(r,'OCR'); } }
    else if(cands.length>1){ showOcrSheet(cands,text); }
    else if(!auto){ showOcrSheet([],text); }
  }catch(e){ if(!auto) pop('OCRを実行できませんでした'); }
  finally{ ocrBusy=false; $('btnOcr').disabled=false; $('btnOcr').textContent='文字を読む(OCR)'; if(camOn) $('camState').textContent='読取中'; }
}
function showOcrSheet(cands,text){
  $('ocrSheet').hidden=false; $('ocrTitle').textContent=cands.length?'候補を選んでください':'コードを読み取れませんでした';
  $('ocrText').textContent='認識した文字: '+(text.replace(/\s+/g,' ').trim().slice(0,80)||'(なし)');
  const ul=$('ocrSug'); ul.innerHTML=''; cands.slice(0,8).forEach(r=>ul.appendChild(sugItem(r,()=>{ $('ocrSheet').hidden=true; recordScan(r,'OCR'); })));
}

/* ---------------- 一覧 ---------------- */
let listDirty=true, filtered=[];
function invalidateList(){ listDirty=true; if(state.tab==='list') renderList(); }
function fillSelects(){
  const uniq=k=>[...new Set(state.rows.map(r=>r[k]).filter(Boolean))].sort();
  const set=(id,label,arr)=>{ const el=$(id), cur=el.value; el.innerHTML=`<option value="">${label}: すべて</option>`+arr.map(v=>`<option>${v.replace(/</g,'&lt;')}</option>`).join(''); el.value=cur; };
  set('fdept','所属',uniq('dept')); set('floc','所在地',uniq('loc'));
}
function applyFilter(){
  const q=norm($('fq').value), g=$('fgen').value, d=$('fdept').value, l=$('floc').value;
  filtered=state.rows.filter(r=>{
    if(d&&r.dept!==d) return false; if(l&&r.loc!==l) return false;
    if(q&&!r.search.includes(q)) return false;
    const rc=state.rec.get(r.code), gv=genOf(r);
    if(g==='__blank'&&gv) return false; if(g==='__mine'&&!(rc&&rc.time)) return false;
    if(g==='__nophoto'&&rc&&rc.photos&&rc.photos.length) return false;
    if(['有','無','確認中'].includes(g)&&gv!==g) return false;
    return true; });
  const {k,d:dir}=state.sort; const key=r=>{ const rc=state.rec.get(r.code); return k==='gen'?genOf(r):k==='time'?(rc&&rc.time||0):r[k]; };
  filtered.sort((a,b)=>{ const x=key(a),y=key(b); return (x>y?1:x<y?-1:0)*dir; });
}
function renderList(more=false){
  if(!more){ applyFilter(); state.limit=200; }
  const tb=$('tbody'); const frag=document.createDocumentFragment(); const upto=Math.min(state.limit,filtered.length);
  for(let i=0;i<upto;i++){ const r=filtered[i], rc=state.rec.get(r.code), gv=genOf(r); const tr=document.createElement('tr'); if(rc&&rc.time) tr.className='mine'; tr.dataset.code=r.code;
    const td=(txt,cls)=>{const c=document.createElement('td'); if(cls) c.className=cls; c.textContent=txt; tr.appendChild(c); return c;};
    td(r.code,'code'); td(r.name,'ellip'); td(r.place,'ellip'); td(r.type,'ellip'); td(r.dept);
    const g=td(''); const p=document.createElement('span'); p.className='pill '+(gv==='有'?'y':gv==='無'?'n':gv==='確認中'?'c':''); p.textContent=gv||'—'; g.appendChild(p);
    td(r.code,'code');
    const loc=td(''); if(rc&&rc.lat!=null){ const a=document.createElement('a'); a.href=`https://www.google.com/maps?q=${rc.lat},${rc.lng}`; a.target='_blank'; a.rel='noopener'; a.textContent='地図'; a.onclick=e=>e.stopPropagation(); loc.appendChild(a); }
    td(rc&&rc.photos&&rc.photos.length?rc.photos.length+'枚':''); td(rc&&rc.time?fmtTime(rc.time):'');
    frag.appendChild(tr); }
  tb.replaceChildren(frag);
  $('listCount').textContent=`${filtered.length.toLocaleString()}件 / 全${state.rows.length.toLocaleString()}件`+(upto<filtered.length?`(先頭${upto}件を表示)`:'');
  $('btnMore').hidden=upto>=filtered.length; listDirty=false;
}

/* ---------------- 詳細 ---------------- */
let detailRow=null;
async function openDetail(code){
  const r=state.byCode.get(code)||(state.rec.get(code)&&{code,name:'(リスト外)',place:'',type:'',dept:'',loc:'',kiban:'',gen0:'',unlisted:true}); if(!r) return;
  detailRow=r; $('detail').hidden=false; await renderDetail();
}
async function renderDetail(){
  const r=detailRow, rc=state.rec.get(r.code);
  $('dCode').textContent=r.code; const dl=document.createElement('dl'); dl.className='dinfo';
  const add=(k,v)=>{ if(!v) return; const a=document.createElement('dt'),b=document.createElement('dd'); a.textContent=k; b.textContent=v; dl.append(a,b); };
  add('名称',r.name); add('置場・担当者',r.place); add('型式',r.type); add('所属',r.dept); add('所在地',r.loc); add('備考',r.note);
  if(rc&&rc.time){ add('記録',`${fmtFull(rc.time)} (${rc.method||''})`); if(rc.lat!=null) add('位置',`${rc.lat.toFixed(6)}, ${rc.lng.toFixed(6)} ±${rc.acc}m`); }
  $('dInfo').replaceChildren(dl);
  const gv=genOf(r.unlisted?{code:r.code,gen0:''}:r); [...$('dGen').children].forEach(b=>b.classList.toggle('on',b.dataset.v===gv&&gv!==''));
  const gal=$('dPhotos'); gal.innerHTML=''; const ps=(await DB.byCode(r.code)).sort((a,b)=>a.name<b.name?-1:1);
  ps.forEach(p=>{ const f=document.createElement('figure'); const im=document.createElement('img'); im.src=URL.createObjectURL(p.blob); im.alt=p.name; im.onload=()=>URL.revokeObjectURL(im.src);
    const c=document.createElement('figcaption'); const s=document.createElement('span'); s.textContent=p.name; const d=document.createElement('button'); d.textContent='削除';
    d.onclick=async()=>{ await DB.del('photos',p.name); const rec=state.rec.get(r.code); if(rec){ rec.photos=rec.photos.filter(n=>n!==p.name); await DB.put('records',rec);} renderDetail(); invalidateList(); };
    c.append(s,d); f.append(im,c); gal.appendChild(f); });
}
async function addFilesToDetail(files){
  const r=detailRow; let rec=state.rec.get(r.code);
  if(!rec){ await setGen(r,genOf(r)||'有'); rec=state.rec.get(r.code); }
  for(const f of files){ const blob=await resizeImage(f,1920);
    let name; if(!rec.photos.length) name=`${r.code}_${kibanFile(r)}.jpg`; else name=`${r.code}_${nextSeq(rec)}.jpg`;
    await savePhoto(r.code,name,blob,rec.photos.length?'extra':'main'); rec.photos.push(name); }
  await DB.put('records',rec); await renderDetail(); invalidateList();
}
function resizeImage(file,max){ return new Promise((res,rej)=>{ const im=new Image(); const u=URL.createObjectURL(file); im.onload=()=>{ const sc=Math.min(1,max/Math.max(im.width,im.height)); const c=document.createElement('canvas'); c.width=Math.round(im.width*sc); c.height=Math.round(im.height*sc); c.getContext('2d').drawImage(im,0,0,c.width,c.height); URL.revokeObjectURL(u); c.toBlob(b=>res(b),'image/jpeg',.85); }; im.onerror=rej; im.src=u; }); }

/* ---------------- 出力 ---------------- */
function download(blob,name){ const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=name; document.body.appendChild(a); a.click(); setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},4000); }
const stamp=()=>{const d=new Date();return `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`};
function recordRows(){ return [...state.rec.values()].filter(r=>r.time).sort((a,b)=>a.time-b.time); }
async function exportExcel(){
  if(!state.meta) return pop('リストが未読み込みです');
  const wb=new ExcelJS.Workbook(); const ws=wb.addWorksheet('最新データ'); ws.addRow(state.header); ws.getRow(1).font={bold:true};
  const gi=state.idx[COLS.gen]; let changed=0;
  state.rows.forEach(r=>{ const row=state.meta.rows[r.i].slice(); const nv=genOf(r); const rc=state.rec.get(r.code); const edit=rc&&rc.gen!=null&&rc.gen!==r.gen0; if(gi!=null&&edit) row[gi]=nv;
    const xr=ws.addRow(row.map(v=>v===''?null:v)); if(edit){ xr.getCell(gi+1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFFF00'}}; changed++; } });
  ws.views=[{state:'frozen',ySplit:1}];
  const w2=wb.addWorksheet('記録'); w2.addRow(['資産コード','名称','機番','現物','方法','記録日時','緯度','経度','精度(m)','写真名','リスト']); w2.getRow(1).font={bold:true};
  recordRows().forEach(rc=>{ const r=state.byCode.get(rc.code); w2.addRow([rc.code,r?r.name:'',rc.kiban||'',rc.gen||'',rc.method||'',fmtFull(rc.time),rc.lat,rc.lng,rc.acc,(rc.photos||[]).join(' '),rc.listed===false?'リスト外':'']); });
  w2.columns.forEach(c=>c.width=16); ws.columns.forEach(c=>c.width=14);
  const buf=await wb.xlsx.writeBuffer(); download(new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),`棚卸回答_${stamp()}.xlsx`);
  $('expInfo').textContent=`Excelを出力しました(変更 ${changed}件を黄色表示)。`;
}
function csvEsc(v){ v=v==null?'':String(v); return /[",\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v; }
function buildCsv(){ const h=['資産コード','名称','機番','現物','方法','記録日時','緯度','経度','精度(m)','写真名']; const lines=[h.join(',')];
  recordRows().forEach(rc=>{ const r=state.byCode.get(rc.code); lines.push([rc.code,r?r.name:'',rc.kiban||'',rc.gen||'',rc.method||'',fmtFull(rc.time),rc.lat,rc.lng,rc.acc,(rc.photos||[]).join(' ')].map(csvEsc).join(',')); });
  return '﻿'+lines.join('\r\n'); }
function exportCsv(){ download(new Blob([buildCsv()],{type:'text/csv;charset=utf-8'}),`棚卸記録_${stamp()}.csv`); }
async function exportZip(){
  const photos=await DB.all('photos'); if(!photos.length) return pop('写真がありません');
  const per=150, parts=Math.ceil(photos.length/per); $('btnZip').disabled=true;
  try{ for(let p=0;p<parts;p++){ const z=new JSZip(); if(p===0) z.file('記録.csv',buildCsv());
      photos.slice(p*per,(p+1)*per).forEach(ph=>z.file(ph.name,ph.blob));
      $('expInfo').textContent=`ZIP作成中 ${p+1}/${parts}`; const blob=await z.generateAsync({type:'blob',compression:'STORE'});
      download(blob,`棚卸写真_${stamp()}${parts>1?`_${p+1}of${parts}`:''}.zip`); await new Promise(r=>setTimeout(r,700)); }
    $('expInfo').textContent=`写真${photos.length}枚をZIP ${parts}個で出力しました。`; }
  finally{ $('btnZip').disabled=false; }
}


/* ---------------- リスト配布(暗号化) ---------------- */
const KEEP=['資産コード','名称','置場・担当者','型式','現物','所属','所在地','備考'];
const PW_CHARS='ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789';
const PBKDF2_ITER=300000;
const b64=u8=>{let s='';for(let i=0;i<u8.length;i+=0x8000)s+=String.fromCharCode.apply(null,u8.subarray(i,i+0x8000));return btoa(s)};
const unb64=s=>Uint8Array.from(atob(s),c=>c.charCodeAt(0));
function genPassword(){ const r=new Uint32Array(8); crypto.getRandomValues(r); return [...r].map(x=>PW_CHARS[x%PW_CHARS.length]).join(''); }
async function deriveKey(pw,salt){ const k=await crypto.subtle.importKey('raw',new TextEncoder().encode(pw),'PBKDF2',false,['deriveKey']);
  return crypto.subtle.deriveKey({name:'PBKDF2',salt,iterations:PBKDF2_ITER,hash:'SHA-256'},k,{name:'AES-GCM',length:256},false,['encrypt','decrypt']); }
async function gz(u8,dir){ if(typeof CompressionStream==='undefined') return null; const s=new Blob([u8]).stream().pipeThrough(new (dir==='c'?CompressionStream:DecompressionStream)('gzip')); return new Uint8Array(await new Response(s).arrayBuffer()); }
async function encryptList(meta,pw){
  const idx=KEEP.map(k=>meta.header.indexOf(k));
  const payload={header:KEEP,rows:meta.rows.map(r=>idx.map(i=>i<0?'':r[i])),filename:meta.filename,sheet:meta.sheet,at:Date.now()};
  let raw=new TextEncoder().encode(JSON.stringify(payload)); let z=false; const c=await gz(raw,'c'); if(c){raw=c;z=true;}
  const salt=crypto.getRandomValues(new Uint8Array(16)), iv=crypto.getRandomValues(new Uint8Array(12)), key=await deriveKey(pw,salt);
  const ct=new Uint8Array(await crypto.subtle.encrypt({name:'AES-GCM',iv},key,raw));
  return JSON.stringify({v:1,z,iter:PBKDF2_ITER,salt:b64(salt),iv:b64(iv),data:b64(ct),at:payload.at});
}
async function decryptList(text,pw){
  const o=JSON.parse(text); const key=await deriveKey(pw,unb64(o.salt));
  let raw=new Uint8Array(await crypto.subtle.decrypt({name:'AES-GCM',iv:unb64(o.iv)},key,unb64(o.data)));
  if(o.z){ raw=await gz(raw,'d'); if(!raw) throw new Error('この端末は圧縮データを展開できません'); }
  return JSON.parse(new TextDecoder().decode(raw));
}
async function receiveList(pw){
  if(receiveList.busy) return; receiveList.busy=true;
  try{
    showToast('リストを受け取り中…','ok');
    const res=await fetch('list.enc?ts='+Date.now(),{cache:'no-store'}); if(!res.ok) throw new Error('list.enc が見つかりません');
    let p; try{ p=await decryptList(await res.text(),pw); }catch(e){ if(e.message&&e.message.includes('圧縮')) throw e; throw new Error('パスワードが合いません。新しいQRを読み取ってください'); }
    await DB.put('meta',{header:p.header,rows:p.rows,filename:p.filename,sheet:p.sheet,at:p.at,received:true},'list');
    await loadList(); renderAll(); beep(true); showToast(`リストを受け取りました(${state.rows.length.toLocaleString()}件)`,'ok');
  }catch(e){ beep(false); showToast(e.message||'受け取れませんでした','bad'); }
  finally{ receiveList.busy=false; }
}
function qrDataUrl(text){ const q=qrcode(0,'M'); q.addData(text); q.make(); return q.createDataURL(8,4); }
let lastEnc=null;
async function distribute(){
  const meta=state.meta; if(!meta){ $('distStatus').textContent='先にリストを読み込んでください。'; return; }
  $('distStatus').textContent='暗号化しています…';
  const pw=genPassword(); lastEnc=await encryptList(meta,pw);
  $('distOut').hidden=false; $('pwQr').src=qrDataUrl('TNPW1:'+pw); $('pwText').textContent=pw; $('pwInfo').textContent=fmtFull(Date.now())+' 発行';
  const tok=$('ghToken').value.trim(), repo=$('ghRepo').value.trim();
  if(!tok){ $('distStatus').textContent='パスワードを作りました。list.encを保存して、GitHubのリポジトリに上書きアップロードしてください(自動公開を使うには、下の設定にトークンを入れます)。'; return; }
  try{
    const api=`https://api.github.com/repos/${repo}/contents/list.enc`, hd={Authorization:'Bearer '+tok,Accept:'application/vnd.github+json'};
    let sha; const g=await fetch(api,{headers:hd,cache:'no-store'}); if(g.ok) sha=(await g.json()).sha; else if(g.status!==404) throw new Error('GitHubに接続できません('+g.status+')');
    const put=await fetch(api,{method:'PUT',headers:hd,body:JSON.stringify({message:'リスト更新 '+fmtFull(Date.now()),content:b64(new TextEncoder().encode(lastEnc)),sha})});
    if(!put.ok) throw new Error('公開に失敗しました('+put.status+')。トークンの権限を確認してください');
    $('distStatus').textContent='公開しました。反映まで1〜2分かかります。反映後に、このQRをスマホで読み取ってください。';
  }catch(e){ $('distStatus').textContent=e.message+'。list.encを保存して、手動でアップロードしてください。'; }
}

/* ---------------- 設定 ---------------- */
function loadSettings(){ try{ $('ghToken').value=localStorage.getItem('tn-gh-token')||''; const rp=localStorage.getItem('tn-gh-repo'); $('ghRepo').value=rp||(location.hostname.endsWith('github.io')?location.hostname.split('.')[0]+'/'+(location.pathname.split('/')[1]||''):''); }catch(e){}
  try{ Object.assign(state.settings,JSON.parse(localStorage.getItem('tn-settings')||'{}')); }catch(e){}
  $('sAutoOcr').checked=state.settings.autoOcr; $('sSound').checked=state.settings.sound; $('sPhoto').checked=state.settings.photo; }
function saveSettings(){ state.settings={autoOcr:$('sAutoOcr').checked,sound:$('sSound').checked,photo:$('sPhoto').checked}; try{ localStorage.setItem('tn-settings',JSON.stringify(state.settings)); }catch(e){} }
async function renderStore(){
  const m=state.meta; $('listInfo').textContent=m?`${m.filename}(シート「${m.sheet}」${m.rows.length.toLocaleString()}件・${fmtFull(m.at)}に読込)`:'未読み込み';
  let t=''; try{ const e=await navigator.storage.estimate(); t=`端末内の使用量: ${(e.usage/1048576).toFixed(1)}MB`; }catch(e){}
  const np=(await DB.all('photos')).length; $('storeInfo').textContent=`記録 ${recordRows().length}件 / 写真 ${np}枚 ${t}`;
}
function renderAll(){ fillSelects(); renderProgress(); const has=state.rows.length>0; $('nolist').hidden=has; $('btnManual').disabled=!has; renderStore(); invalidateList(); if(state.tab==='scan') showCur(); }

/* ---------------- タブ ---------------- */
function goTab(t){
  state.tab=t; document.querySelectorAll('.tabs button').forEach(b=>b.classList.toggle('on',b.dataset.tab===t));
  ['scan','list','settings'].forEach(v=>$('v-'+v).hidden=v!==t);
  if(t==='scan') { startCam(); } else stopCam();
  if(t==='list'&&listDirty) renderList(); if(t==='settings') renderStore();
  window.scrollTo(0,0);
}

/* ---------------- 起動 ---------------- */
async function init(){
  await DB.open(); loadSettings(); await loadList(); startGeo();
  try{ navigator.storage&&navigator.storage.persist&&navigator.storage.persist(); }catch(e){}
  document.querySelectorAll('.tabs button').forEach(b=>b.onclick=()=>goTab(b.dataset.tab));
  document.querySelectorAll('[data-goto]').forEach(b=>b.onclick=()=>goTab(b.dataset.goto));
  $('btnManual').onclick=openManual; $('sheetClose').onclick=()=>$('sheet').hidden=true;
  $('manualIn').oninput=e=>renderSug(e.target.value);
  $('manualIn').onkeydown=e=>{ if(e.key==='Enter'){ const f=$('manualSug').querySelector('button'); if(f) f.click(); } };
  $('btnOcr').onclick=()=>runOcr(false); $('ocrClose').onclick=()=>$('ocrSheet').hidden=true; $('ocrManual').onclick=()=>{$('ocrSheet').hidden=true;openManual()};
  $('btnCamToggle').onclick=()=>camOn?stopCam():startCam(); $('btnStartCam').onclick=startCam;
  $('btnTorch').onclick=async()=>{ try{ torchOn=!torchOn; await stream.getVideoTracks()[0].applyConstraints({advanced:[{torch:torchOn}]}); $('btnTorch').classList.toggle('primary',torchOn);}catch(e){} };
  $('btnShot').onclick=addPhotoForCurrent; $('btnUndo').onclick=undoCurrent;
  $('curGen').onclick=e=>{ const v=e.target.dataset&&e.target.dataset.v; if(v!=null&&state.cur) setGen(state.cur.unlisted?{code:state.cur.code,unlisted:true}:state.cur,v); };
  $('dGen').onclick=e=>{ const v=e.target.dataset&&e.target.dataset.v; if(v!=null) setGen(detailRow,v).then(renderDetail); };
  $('dClose').onclick=()=>$('detail').hidden=true; $('dFile').onchange=async e=>{ await addFilesToDetail([...e.target.files]); e.target.value=''; };
  ['fq','fgen','fdept','floc'].forEach(id=>$(id).addEventListener(id==='fq'?'input':'change',()=>renderList()));
  $('btnMore').onclick=()=>{ state.limit+=300; renderList(true); };
  $('tbody').onclick=e=>{ const tr=e.target.closest('tr'); if(tr) openDetail(tr.dataset.code); };
  document.querySelectorAll('th[data-sort]').forEach(th=>th.onclick=()=>{ const k=th.dataset.sort; state.sort=state.sort.k===k?{k,d:-state.sort.d}:{k,d:1}; renderList(); });
  $('fileList').onchange=async e=>{ const f=e.target.files[0]; if(!f) return; $('listInfo').textContent='読み込み中…'; try{ await importList(f); pop('リストを読み込みました'); }catch(err){ $('listInfo').textContent='読み込めませんでした: '+err.message; } e.target.value=''; };
  $('btnReissue').onclick=distribute; $('btnEncSave').onclick=()=>{ if(!lastEnc) return pop('先にパスワードを発行してください'); download(new Blob([lastEnc],{type:'application/json'}),'list.enc'); };
  $('ghToken').onchange=()=>{ try{localStorage.setItem('tn-gh-token',$('ghToken').value.trim())}catch(e){} }; $('ghRepo').onchange=()=>{ try{localStorage.setItem('tn-gh-repo',$('ghRepo').value.trim())}catch(e){} };
  $('btnPwManual').onclick=()=>{ openPwSheet(); };
  $('btnExcel').onclick=exportExcel; $('btnZip').onclick=exportZip; $('btnCsv').onclick=exportCsv;
  ['sAutoOcr','sSound','sPhoto'].forEach(id=>$(id).onchange=saveSettings);
  $('btnReset').onclick=()=>$('resetConfirm').hidden=false; $('btnResetNo').onclick=()=>$('resetConfirm').hidden=true;
  $('btnResetYes').onclick=async()=>{ await DB.clear('records'); await DB.clear('photos'); state.rec.clear(); state.cur=null; $('cur').hidden=true; $('resetConfirm').hidden=true; recentSeen.clear(); renderAll(); pop('回答データを削除しました'); };
  document.addEventListener('visibilitychange',()=>{ if(document.hidden) stopCam(); else if(state.tab==='scan') startCam(); });
  renderAll(); goTab('scan');
  if('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(()=>{});
  window.__app={state,onText,recordScan,resolveText,ocrCandidates,runOcr};
}
init();

function restoreManual(){ $('manualIn').oninput=e=>renderSug(e.target.value); $('manualIn').placeholder='資産コード・機番・名称の一部'; $('sheetClose').onclick=()=>$('sheet').hidden=true; }
function openPwSheet(){ $('sheet').hidden=false; $('manualIn').value=''; $('manualIn').placeholder='8桁のパスワード'; $('manualSug').innerHTML=''; const b=document.createElement('li'); const bt=document.createElement('button'); bt.textContent='このパスワードでリストを受け取る'; bt.onclick=()=>{ const v=$('manualIn').value.trim(); if(v){ $('sheet').hidden=true; restoreManual(); receiveList(v);} }; b.appendChild(bt); $('manualSug').appendChild(b); $('manualIn').oninput=null; setTimeout(()=>$('manualIn').focus(),50); $('sheetClose').onclick=()=>{$('sheet').hidden=true; restoreManual();}; }
