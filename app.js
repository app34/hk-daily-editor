const STATUSES=[['OCC','OCCUPIED'],['ARR','ARRIVAL'],['DEP','DEPATURE'],['B2B','BACK TO BACK'],['VAC','VACANT'],['OOO','OUT OF ORDER'],['SV','SHOW VILLA'],['STB','STAND BY'],['DU','DAY USE'],['OS','OUT OF SERVICE'],['DC','DEEP CLEANING'],['VM-ARR','VILLA MOVE ARRIVAL']];
const CATS=[['BHV','Beach Villa'],['DVP','Deluxe Pool'],['BFS','Beach Family'],['WAV','Water Villa'],['WVP','Water Pool'],['HWS','Horizon Water']];
const WEATHER=[['💧','Water / humid'],['☀️','Sunny'],['🌤️','Mostly sunny'],['⛅','Partly cloudy'],['☁️','Cloudy'],['🌦️','Showers'],['🌧️','Rain'],['⛈️','Thunderstorm'],['☔','Umbrella'],['🌫️','Fog'],['💨','Windy'],['🌊','Rough sea'],['🌙','Night'],['❄️','Cool']];
const COLORS=['#5b9bd5','#c45911','#70ad47','#ed7d31','#9b6bdf','#2e75b6','#548235','#203864','#00b0f0'];
const G1=[1,5,9,13,17,21,25,29];
const G2=[1,5,9,13,17,21,25,29,33];
const SHEET='JULY- 2026';
let originalBuf=null,fileName='',model=null,pick=null,savedAt='';
let viewMode=localStorage.getItem('hk-daily-view')||'full';
let masterQ='';
let selectedSrc='';
let undoStack=[], redoStack=[];
const VIEWS=[
  ['full','Full daily sheet'],
  ['cards','Attendant cards'],
  ['sheet','Allocation board'],
  ['master','Master list'],
  ['status','Status board'],
  ['compact','Compact grid']
];
const DB_NAME='hk-daily-oblu';
const DB_STORE='files';

function openDb(){
  return new Promise((resolve,reject)=>{
    const req=indexedDB.open(DB_NAME,1);
    req.onupgradeneeded=()=>req.result.createObjectStore(DB_STORE);
    req.onsuccess=()=>resolve(req.result);
    req.onerror=()=>reject(req.error);
  });
}
async function idbSet(key,value){
  const db=await openDb();
  return new Promise((resolve,reject)=>{
    const tx=db.transaction(DB_STORE,'readwrite');
    tx.objectStore(DB_STORE).put(value,key);
    tx.oncomplete=()=>resolve();
    tx.onerror=()=>reject(tx.error);
  });
}
async function idbGet(key){
  const db=await openDb();
  return new Promise((resolve,reject)=>{
    const tx=db.transaction(DB_STORE,'readonly');
    const q=tx.objectStore(DB_STORE).get(key);
    q.onsuccess=()=>resolve(q.result);
    q.onerror=()=>reject(q.error);
  });
}
async function persist(msg){
  if(!originalBuf||!model) return;
  savedAt=new Date().toISOString();
  try{
    await idbSet('pack',{
      fileName,
      savedAt,
      buf:originalBuf,
      model,
      undoStack,
      redoStack
    });
    localStorage.setItem('hk-daily-meta', JSON.stringify({fileName,savedAt}));
    updateFileLabel();
    if(msg!==false) toast(msg||'Saved on this phone');
  }catch(err){
    toast('Save failed: '+(err.message||err));
  }
}
function scheduleSave(){
  clearTimeout(window.__sv);
  window.__sv=setTimeout(()=>persist(false),400);
}
function updateFileLabel(){
  const t=savedAt? new Date(savedAt).toLocaleString():'';
  document.getElementById('fileName').textContent = fileName + (t?(' · saved '+t):'');
}
function showEditor(){
  document.getElementById('gate').hidden=true;
  document.getElementById('page').hidden=false;
  document.getElementById('btnOut').disabled=false;
  document.getElementById('btnSave').disabled=false;
  document.getElementById('btnClear').hidden=false;
  document.getElementById('btnText').disabled=false;
  document.getElementById('btnRoll').disabled=false;
  document.getElementById('btnUndo').disabled=false;
  document.getElementById('btnRedo').disabled=false;
  document.getElementById('viewBar').hidden=false;
  updateFileLabel();
}

function toast(t){const el=document.getElementById('toast');el.textContent=t;el.classList.add('show');clearTimeout(window.__t);window.__t=setTimeout(()=>el.classList.remove('show'),1600);}
function esc(s){return String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));}
function colL(n){let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;}
function addr(r,c){return colL(c)+r;}
function val(ws,a){const c=ws[a];if(!c||c.v==null)return '';if(c.t==='d'||c.v instanceof Date){const d=c.v instanceof Date?c.v:new Date(c.v);if(!isNaN(d))return d.toISOString().slice(0,10);}if(typeof c.v==='number'&&c.z&&/d|m|y/i.test(String(c.z))){const d=XLSX.SSF.parse_date_code(c.v);if(d)return d.y+'-'+String(d.m).padStart(2,'0')+'-'+String(d.d).padStart(2,'0');}return String(c.v).trim();}

function parseBlocks(ws, headerRow, dataRows, cols){
  return cols.map((c,i)=>({
    name:val(ws,addr(headerRow,c)),
    nameCell:addr(headerRow,c),
    col:c,
    headerRow,
    color:COLORS[i%COLORS.length],
    extraStart: dataRows[dataRows.length-1]+1,
    rows:dataRows.map(r=>({
      villa:val(ws,addr(r,c)),
      cat:val(ws,addr(r,c+1)),
      status:val(ws,addr(r,c+2)),
      cells:[addr(r,c),addr(r,c+1),addr(r,c+2)],
      extra:false
    }))
  }));
}
function parseWorkbook(wb){
  const ws=wb.Sheets[SHEET]||wb.Sheets[wb.SheetNames[0]];
  const r1=[]; for(let r=36;r<=47;r++) r1.push(r);
  const r2=[]; for(let r=52;r<=64;r++) r2.push(r);
  const moves=[];
  [[1,2],[3,4],[6,7]].forEach(([a,b])=>{
    for(let r=20;r<=29;r++){
      const from=val(ws,addr(r,a)),to=val(ws,addr(r,b));
      if(from||to) moves.push({from,to,fromCell:addr(r,a),toCell:addr(r,b)});
    }
  });
  const pack=(rs,cs)=>{const a=[];rs.forEach(r=>cs.forEach(c=>{const v=val(ws,addr(r,c));if(v)a.push({v,cell:addr(r,c)});}));return a;};
  const leave=[];
  for(let r=67;r<=78;r++){const name=val(ws,addr(r,1));if(name)leave.push({name,type:val(ws,addr(r,4)),start:val(ws,addr(r,7)).slice(0,10),end:val(ws,addr(r,10)).slice(0,10),cells:{name:addr(r,1),type:addr(r,4),start:addr(r,7),end:addr(r,10)}});}
  const laundry=[];
  for(let r=67;r<=81;r++){const name=val(ws,addr(r,26));if(name)laundry.push({no:val(ws,addr(r,25)),name,timing:val(ws,addr(r,30)),desig:val(ws,addr(r,36)),contact:val(ws,addr(r,38)),cells:{no:addr(r,25),name:addr(r,26),timing:addr(r,30),desig:addr(r,36),contact:addr(r,38)}});}
  const supervisors=[];
  for(let r=87;r<=94;r++){const name=val(ws,addr(r,1));if(name)supervisors.push({name,role:val(ws,addr(r,4)),status:val(ws,addr(r,8)),section:val(ws,addr(r,11)),contact:val(ws,addr(r,16)),cells:{name:addr(r,1),role:addr(r,4),status:addr(r,8),section:addr(r,11),contact:addr(r,16)}});}
  const minibar=[];
  for(let r=87;r<=90;r++){const name=val(ws,addr(r,19));if(name&&name!=='TEAM MEMBER')minibar.push({no:val(ws,addr(r,18)),name,role:val(ws,addr(r,23)),status:val(ws,addr(r,28)),section:val(ws,addr(r,31)),cells:{no:addr(r,18),name:addr(r,19),role:addr(r,23),status:addr(r,28),section:addr(r,31)}});}
  const office=[];
  for(let r=94;r<=97;r++){const name=val(ws,addr(r,19));if(name&&name!=='TEAM MEMBER')office.push({no:val(ws,addr(r,18)),name,role:val(ws,addr(r,23)),status:val(ws,addr(r,28)),contact:val(ws,addr(r,31)),cells:{no:addr(r,18),name:addr(r,19),role:addr(r,23),status:addr(r,28),contact:addr(r,31)}});}
  const publicArea=[];
  for(let r=101;r<=110;r++){const name=val(ws,addr(r,2));if(name)publicArea.push({no:val(ws,addr(r,1)),name,role:val(ws,addr(r,6)),section:val(ws,addr(r,8)),cells:{no:addr(r,1),name:addr(r,2),role:addr(r,6),section:addr(r,8)}});}
  const tasks=[];
  for(let r=101;r<=108;r++){tasks.push({detail:val(ws,addr(r,19)),area:val(ws,addr(r,24)),time:val(ws,addr(r,34)),cells:{detail:addr(r,19),area:addr(r,24),time:addr(r,34)}});}
  return {
    date:val(ws,'O4').slice(0,10),
    duty:{mod:val(ws,'AG1'),hkDay:val(ws,'AG2'),hkNight:val(ws,'AG3'),security:val(ws,'AG4'),eng:val(ws,'AG5')},
    core:val(ws,'I7'), adults:val(ws,'F16'), children:val(ws,'F17'),
    forecast:{
      weather: (val(ws,'F8') && !String(val(ws,'F8')).startsWith('=')) ? val(ws,'F8') : '💧',
      occupancyPct: pctFromCell(val(ws,'F10')),
      arrival: numOrBlank(val(ws,'F11')),
      departure: numOrBlank(val(ws,'F12')),
      villaMove: numOrBlank(val(ws,'F13')),
      occupied: numOrBlank(val(ws,'F14')),
      vacant: numOrBlank(val(ws,'F15')),
      auto:true
    },
    beach:parseBlocks(ws,32,r1,G1),
    water:parseBlocks(ws,48,r2,G2),
    moves, arrivals:pack(range(19,29),range(9,15)), departures:pack(range(19,29),range(17,23)),
    honeymoon:pack(range(19,29),[25]), birthday:pack(range(19,29),[29]), anniversary:pack(range(19,29),[33]), upon:pack(range(19,29),[37]),
    leave, laundry, supervisors, minibar, office, publicArea, tasks
  };
}
function padGroup(list, group){
  const n=group==='beach'?12:13;
  const start=group==='beach'?36:52;
  (list||[]).forEach(p=>{
    while(p.rows.length<n){
      const r=start+p.rows.length;
      p.rows.push({villa:'',cat:'',status:'',extra:false,cells:[addr(r,p.col),addr(r,p.col+1),addr(r,p.col+2)]});
    }
  });
}
function ensureModel(m){
  if(!m) return m;
  ['moves','arrivals','departures','honeymoon','birthday','anniversary','upon','leave','laundry','supervisors','minibar','office','publicArea','tasks','beach','water'].forEach(k=>{ if(!m[k]) m[k]=[]; });
  if(!m.duty) m.duty={};
  padGroup(m.beach,'beach');
  padGroup(m.water,'water');
  if(!m.forecast) m.forecast={weather:'☀️',occupancyPct:'',arrival:'',departure:'',villaMove:'',occupied:'',vacant:'',auto:true};
  if(m.forecast.auto!==false) m.forecast.auto=true;
  return m;
}
function numOrBlank(v){
  if(v===''||v==null) return '';
  const s=String(v).trim();
  if(s.startsWith('=')) return '';
  const n=Number(s.replace('%',''));
  return isNaN(n) ? '' : n;
}
function pctFromCell(v){
  if(v===''||v==null) return '';
  const s=String(v).trim();
  if(s.startsWith('=')) return '';
  const n=Number(s.replace('%',''));
  if(isNaN(n)) return '';
  return n<=1 ? (n*100).toFixed(2) : String(n);
}
function allocCounts(){
  const vs=people().flatMap(p=>filled(p));
  const c={}; vs.forEach(v=>c[v.status]=(c[v.status]||0)+1);
  return {
    arrival:c.ARR||0,
    departure:c.DEP||0,
    occupied:(c.OCC||0)+(c.ARR||0)+(c.B2B||0)+(c.DU||0)+(c['VM-ARR']||0),
    vacant:c.VAC||0
  };
}
function syncForecast(){
  if(!model||!model.forecast||model.forecast.auto===false) return;
  const a=allocCounts();
  model.forecast.arrival=a.arrival;
  model.forecast.departure=a.departure;
  model.forecast.occupied=a.occupied;
  model.forecast.vacant=a.vacant;
}
function allocSnap(){
  return JSON.parse(JSON.stringify({beach:model.beach,water:model.water,forecast:model.forecast,moves:model.moves}));
}
function applySnap(s){
  if(!s) return;
  model.beach=s.beach; model.water=s.water; model.forecast=s.forecast; model.moves=s.moves;
}
function pushUndo(){
  undoStack.push(allocSnap());
  if(undoStack.length>30) undoStack.shift();
  redoStack=[];
}
function doUndo(){
  if(!undoStack.length){ toast('Nothing to undo'); return; }
  redoStack.push(allocSnap());
  applySnap(undoStack.pop());
  render(); persist(false); toast('Undo — back to previous statuses');
}
function doRedo(){
  if(!redoStack.length){ toast('Nothing to redo'); return; }
  undoStack.push(allocSnap());
  applySnap(redoStack.pop());
  render(); persist(false); toast('Redo — statuses applied again');
}
function parseRooms(text){
  return [...new Set(String(text||'').split(/[^0-9]+/).filter(Boolean))];
}
function parseMoveLines(text){
  const out=[];
  String(text||'').split(/\n+/).forEach(line=>{
    const nums=String(line).split(/[^0-9]+/).filter(Boolean);
    if(nums.length>=2) out.push({from:nums[0],to:nums[1]});
  });
  return out;
}
function findVilla(num){
  const hits=[];
  ['beach','water'].forEach(g=>model[g].forEach((p,pi)=>p.rows.forEach((r,ri)=>{
    if(String(r.villa)===String(num)) hits.push(r);
  })));
  return hits;
}
function rollYesterdayToToday(){
  if(!model){ toast('Import XLSM first'); return; }
  pushUndo();
  let n=0;
  ['beach','water'].forEach(g=>model[g].forEach(p=>p.rows.forEach(r=>{
    if(!String(r.villa||'').trim()) return;
    const y=r.status;
    if(y==='DEP'){ r.status='VAC'; n++; }
    else if(y==='ARR' || y==='B2B' || y==='VM-ARR'){ r.status='OCC'; n++; }
  })));
  syncForecast();
  render();
  persist('Rolled yesterday to today');
  toast(n+' rooms changed to today. Undo if it is wrong');
}
function applyTextImport(){
  const A=new Set(parseRooms(document.getElementById('txArr').value));
  const D=new Set(parseRooms(document.getElementById('txDep').value));
  const O=new Set(parseRooms(document.getElementById('txOcc').value));
  const V=new Set(parseRooms(document.getElementById('txVac').value));
  const mv=parseMoveLines(document.getElementById('txMove').value);
  if(!A.size && !D.size && !O.size && !V.size && !mv.length){ toast('Paste at least one villa number'); return; }
  pushUndo();
  let missing=[];
  const listed=[...A,...D,...O,...V,...mv.map(x=>x.from),...mv.map(x=>x.to)];
  listed.forEach(num=>{ if(!findVilla(num).length) missing.push(num); });
  ['beach','water'].forEach(g=>model[g].forEach(p=>p.rows.forEach(r=>{
    const v=String(r.villa||'').trim();
    if(!v) return;
    if(A.has(v) && D.has(v)) r.status='B2B';
    else if(A.has(v)) r.status='ARR';
    else if(D.has(v)) r.status='DEP';
    if(O.has(v)) r.status='OCC';
    if(V.has(v)) r.status='VAC';
  })));
  mv.forEach(m=>{
    findVilla(m.from).forEach(r=>{ if(!A.has(String(r.villa))) r.status='DEP'; });
    findVilla(m.to).forEach(r=>{ r.status = (A.has(String(r.villa))&&D.has(String(r.villa))) ? 'B2B' : 'VM-ARR'; });
    if(!model.moves.some(x=>String(x.from)===String(m.from)&&String(x.to)===String(m.to)))
      model.moves.push({from:m.from,to:m.to,fromCell:'',toCell:''});
  });
  syncForecast();
  closeTextModal();
  render();
  persist('Statuses updated from text');
  toast((missing.length? (missing.length+' numbers not on board. ') : '')+'Updated. Use Undo if needed');
}
function openTextModal(){
  document.getElementById('overlay').classList.add('show');
  document.getElementById('textModal').classList.add('show');
}
function closeTextModal(){
  document.getElementById('textModal').classList.remove('show');
  if(!document.getElementById('picker').classList.contains('show'))
    document.getElementById('overlay').classList.remove('show');
}
function range(a,b){const x=[];for(let i=a;i<=b;i++)x.push(i);return x;}
function people(){return [...(model.beach||[]),...(model.water||[])];}
function filled(p){return p.rows.filter(r=>String(r.villa).trim());}
function metrics(){
  const vs=people().flatMap(p=>filled(p));
  const c={}; vs.forEach(v=>c[v.status]=(c[v.status]||0)+1);
  return {rooms:vs.length, occ:(c.OCC||0)+(c.ARR||0)+(c.B2B||0)+(c.DU||0), vac:c.VAC||0, arr:c.ARR||0, dep:c.DEP||0, b2b:c.B2B||0};
}

function setPath(path,value){
  const parts=path.split('.');
  let cur=model;
  for(let i=0;i<parts.length-1;i++) cur=cur[parts[i]];
  cur[parts[parts.length-1]]=value;
}
function dd(path,val,kind){
  return `<button type="button" class="dd st-${esc(val||'')}" data-dd="${kind}" data-path="${path}">${esc(val||'Select')}</button>`;
}

function attCard(group, gi, p, pi){
  const rooms=p.rows;
  return `<article class="att dropzone" data-drop="${group}.${pi}">
    <div class="hd" style="background:${p.color}">
      <input data-path="${group}.${pi}.name" value="${esc(p.name)}" placeholder="New section / attendant">
      <span class="cnt">${filled(p).length} rooms</span>
    </div>
    <table>
      <tr><th>V#</th><th>CAT</th><th>STAT</th><th></th></tr>
      ${rooms.map((r,ri)=>{
        const id=group+'.'+pi+'.'+ri;
        const empty=!String(r.villa||'').trim();
        return `<tr class="slot ${empty?'empty':''} ${selectedSrc===id?'on':''}" data-slot="${id}">
        <td><input data-path="${group}.${pi}.rows.${ri}.villa" value="${esc(r.villa)}" placeholder="${empty?'empty slot':''}"></td>
        <td>${dd(group+'.'+pi+'.rows.'+ri+'.cat', r.cat,'cat')}</td>
        <td>${dd(group+'.'+pi+'.rows.'+ri+'.status', r.status,'status')}</td>
        <td><button class="xbtn" data-del="${group}.${pi}.${ri}">×</button></td>
      </tr>`;}).join('')}
    </table>
    <button class="add" data-add="${group}.${pi}">+ Add extra room</button>
  </article>`;
}

function viewBar(){
  document.getElementById('viewBar').innerHTML =
    VIEWS.map(([id,lab])=>`<button data-view="${id}" class="${viewMode===id?'on':''}">${lab}</button>`).join('')+
    `<button data-textimport="1">Import from text</button>
     <button data-undo="1" ${undoStack.length?'':'disabled'} style="opacity:${undoStack.length?1:.45}">Undo</button>
     <button data-redo="1" ${redoStack.length?'':'disabled'} style="opacity:${redoStack.length?1:.45}">Redo</button>`;
}
function moveHint(){
  const sel=selectedSrc?(()=>{const [g,i,r]=selectedSrc.split('.'); const row=model[g][+i].rows[+r]; return (row&&row.villa||'')+' · '+(row&&row.cat||'')+' · '+(row&&row.status||'');})():'';
  return `<div class="card" style="margin:0 0 10px;${selectedSrc?'background:#9a3412;color:#fff;border-color:#9a3412':''}">
    <b>${selectedSrc?('SELECTED '+sel):'No room selected'}</b>
    <div class="hint" style="margin:4px 0 0;${selectedSrc?'color:#ffedd5':''}">Double-tap a room to select. Single tap edits the number. Then tap an empty slot to place V# / CAT / STAT.</div>
  </div>`;
}
function pathOf(group,pi,ri,field){return group+'.'+pi+'.rows.'+ri+'.'+field;}
function sheetBlock(title, group, list){
  const max=Math.max(...list.map(p=>p.rows.length),1);
  return `<div class="secbar"><h3 style="color:#1f4e79;margin:0">${title}</h3>
      <button class="add" style="width:auto;padding:8px 10px" data-addsec="${group}">+ New section</button></div>
    <div class="sheetwrap"><table class="sheetgrid">
      <tr>${list.map((p,i)=>`<th class="nm dropzone" data-drop="${group}.${i}" colspan="3" style="background:${p.color}"><input data-path="${group}.${i}.name" value="${esc(p.name)}" style="color:#222;font-weight:800"></th>`).join('')}</tr>
      <tr>${list.map(()=>'<th>V#</th><th>CAT</th><th>STAT</th>').join('')}</tr>
      ${Array.from({length:max},(_,r)=>'<tr>'+list.map((p,i)=>{
        const row=p.rows[r];
        if(!row) return '<td></td><td></td><td></td>';
        const id=group+'.'+i+'.'+r;
        const empty=!String(row.villa||'').trim();
        return `<td class="slot ${empty?'empty':''} ${selectedSrc===id?'on':''}" data-slot="${id}"><input data-path="${pathOf(group,i,r,'villa')}" value="${esc(row.villa)}" placeholder="${empty?'+':''}"></td>
                <td class="slot ${empty?'empty':''} ${selectedSrc===id?'on':''}" data-slot="${id}">${dd(pathOf(group,i,r,'cat'),row.cat,'cat')}</td>
                <td class="slot ${empty?'empty':''} ${selectedSrc===id?'on':''}" data-slot="${id}">${dd(pathOf(group,i,r,'status'),row.status,'status')}</td>`;
      }).join('')+'</tr>').join('')}
    </table></div>
    <div style="display:flex;gap:6px;flex-wrap:wrap;margin:6px 0 0">${list.map((p,i)=>`<button class="add" style="width:auto;padding:8px 10px" data-add="${group}.${i}">+ Room ${esc(p.name||'section')}</button>`).join('')}</div>`;
}
function masterRows(){
  const q=(masterQ||'').toLowerCase();
  const rows=[];
  ['beach','water'].forEach(g=>model[g].forEach((p,pi)=>p.rows.forEach((r,ri)=>{
    if(!q || String(r.villa).includes(q) || (p.name||'').toLowerCase().includes(q) || (r.status||'').toLowerCase().includes(q) || (r.cat||'').toLowerCase().includes(q))
      rows.push({g,pi,ri,p,r});
  })));
  rows.sort((a,b)=>{
    const ae=!String(a.r.villa||'').trim(), be=!String(b.r.villa||'').trim();
    if(ae!==be) return ae-be;
    const an=(a.p.name||'').localeCompare(b.p.name||'');
    if(an) return an;
    return String(a.r.villa).localeCompare(String(b.r.villa),'en',{numeric:true});
  });
  return rows;
}
function chips(list, path){
  if(!list||!list.length) return '<span class="tiny">—</span>';
  return list.map((x,i)=>`<span style="display:inline-block;min-width:42px">${path?`<input data-path="${path}.${i}.v" value="${esc(x.v)}">`:esc(x.v)}</span>`).join(' ');
}
function fullSheetView(){
  const m=metrics();
  const grid=(group,list)=>{
    const max=Math.max(...list.map(p=>p.rows.length),1);
    return `<table>
      <tr>${list.map((p,i)=>`<td class="dropzone" data-drop="${group}.${i}" colspan="3" style="background:${p.color};padding:5px 3px"><input class="allocname" data-path="${group}.${i}.name" value="${esc(p.name)}"></td>`).join('')}</tr>
      <tr>${list.map(()=>'<td class="sec">V#</td><td class="sec">CAT.</td><td class="sec">STAT.</td>').join('')}</tr>
      ${Array.from({length:max},(_,r)=>'<tr>'+list.map((p,i)=>{
        const row=p.rows[r]; if(!row) return '<td></td><td></td><td></td>';
        const id=group+'.'+i+'.'+r;
        const empty=!String(row.villa||'').trim();
        return `<td class="slot ${empty?'empty':''} ${selectedSrc===id?'on':''}" data-slot="${id}"><div class="vcell"><input data-path="${pathOf(group,i,r,'villa')}" value="${esc(row.villa)}" placeholder="${empty?'+':''}"></div></td><td class="slot ${empty?'empty':''} ${selectedSrc===id?'on':''} st-${esc(row.cat||'')}" data-slot="${id}">${dd(pathOf(group,i,r,'cat'),row.cat,'cat')}</td><td class="slot ${empty?'empty':''} ${selectedSrc===id?'on':''} st-${esc(row.status||'')}" data-slot="${id}">${dd(pathOf(group,i,r,'status'),row.status,'status')}</td>`;
      }).join('')+'</tr>').join('')}
    </table>`;
  };
  return `<div class="fullwrap ${selectedSrc?'ready':''}"><div class="full">
    <table>
      <tr>
        <td class="brand" colspan="6">OBLU SELECT<br><span class="tiny">Sangeli · Housekeeping</span></td>
        <td class="htitle" colspan="12">Housekeeping Daily</td>
        <td colspan="9">
          <table>
            <tr><td>Manager on duty</td><td><input data-path="duty.mod" value="${esc(model.duty.mod)}"></td></tr>
            <tr><td>HK Day Duty</td><td><input data-path="duty.hkDay" value="${esc(model.duty.hkDay)}"></td></tr>
            <tr><td>HK NIGHT Duty</td><td><input data-path="duty.hkNight" value="${esc(model.duty.hkNight)}"></td></tr>
            <tr><td>Security on duty</td><td><input data-path="duty.security" value="${esc(model.duty.security)}"></td></tr>
            <tr><td>Eng Hot line</td><td><input data-path="duty.eng" value="${esc(model.duty.eng)}"></td></tr>
          </table>
        </td>
      </tr>
      <tr><td colspan="27" style="text-align:center;font-weight:800">Date: <input data-path="date" type="date" value="${esc(model.date)}" style="width:auto;font-weight:800;font-size:14px"></td></tr>
      <tr><td class="core" colspan="6">CORE VALUE OF THE DAY</td><td class="core" colspan="21"><input data-path="core" value="${esc(model.core)}"></td></tr>
    </table>
    <table style="margin-top:4px">
      <tr><td class="sec">FORECAST <button type="button" class="dd" data-dd="weather" data-path="forecast.weather" style="width:auto;background:#fff;font-size:18px;padding:2px 8px">${esc(model.forecast.weather||'💧')}</button></td><td class="sec">VIP / NOTES</td><td class="sec2">Next day arrival guest preferences</td></tr>
      <tr>
        <td>
          <table>
            <tr><td>WEATHER</td><td><button type="button" class="dd" data-dd="weather" data-path="forecast.weather" style="font-size:20px">${esc(model.forecast.weather||'☀️')}</button></td></tr>
            <tr><td>OCCUPANCY %</td><td><input data-path="forecast.occupancyPct" value="${esc(model.forecast.occupancyPct)}" placeholder="72.99"> <button type="button" class="mv" id="calcOcc">From rooms</button></td></tr>
            <tr><td class="kpi-arr">ARRIVAL</td><td class="kpi-arr"><input data-path="forecast.arrival" value="${esc(model.forecast.arrival)}"></td></tr>
            <tr><td class="kpi-dep">DEPATURE</td><td class="kpi-dep"><input data-path="forecast.departure" value="${esc(model.forecast.departure)}"></td></tr>
            <tr><td class="kpi-move">VILLA MOVE</td><td class="kpi-move"><input data-path="forecast.villaMove" value="${esc(model.forecast.villaMove)}"></td></tr>
            <tr><td class="kpi-occ">OCCUPIED VILLA</td><td class="kpi-occ"><input data-path="forecast.occupied" value="${esc(model.forecast.occupied)}"></td></tr>
            <tr><td class="kpi-vac">VACANT VILLA</td><td class="kpi-vac"><input data-path="forecast.vacant" value="${esc(model.forecast.vacant)}"></td></tr>
            <tr><td class="kpi-ad">GUEST INHOUSE ADULT</td><td class="kpi-ad"><input data-path="adults" value="${esc(model.adults)}"></td></tr>
            <tr><td>GUEST INHOUSE CHILD</td><td><input data-path="children" value="${esc(model.children)}"></td></tr>
          </table>
          <label class="tiny" style="display:flex;gap:6px;align-items:center;margin-top:4px">
            <input id="autoCount" type="checkbox" ${model.forecast.auto!==false?'checked':''}> Auto ARR / DEP / OCC / VAC from allocation
          </label>
        </td>
        <td class="tiny">VIP list stays on the orange sheet.</td>
        <td class="tiny">Preferences stay in the original sheet cells when present.</td>
      </tr>
    </table>
    <table style="margin-top:4px">
      <tr><td class="sec">VILLA MOVE</td><td class="sec">ARRIVAL</td><td class="sec2">DEPATURE</td><td class="sec">HONEYMOON</td><td class="sec">BIRTHDAY</td><td class="sec">ANNIVERSARY</td></tr>
      <tr>
        <td>${(model.moves||[]).map((x,i)=>`<div>${`<input data-path="moves.${i}.from" value="${esc(x.from)}" style="width:46%">`} → ${`<input data-path="moves.${i}.to" value="${esc(x.to)}" style="width:46%">`}</div>`).join('')}<button class="add" data-addmove="1">+ move</button></td>
        <td>${chips(model.arrivals,'arrivals')}</td>
        <td>${chips(model.departures,'departures')}</td>
        <td>${chips(model.honeymoon,'honeymoon')}</td>
        <td>${chips(model.birthday,'birthday')}</td>
        <td>${chips(model.anniversary,'anniversary')}</td>
      </tr>
    </table>
    <div class="sec" style="margin-top:6px;padding:5px">VILLA ATTENDANT ALLOCATION</div>
    <div class="tools">
      <b>Edit</b>
      <span>Double-tap a room to select it. Then tap an empty slot to place V# / CAT / STAT. Single tap only edits the number.</span>
      <span id="selLabel" style="${selectedSrc?'background:#9a3412;color:#fff;padding:4px 8px;border-radius:4px;font-weight:800':''}">${selectedSrc?('SELECTED '+(()=>{const [g,i,r]=selectedSrc.split('.'); const row=model[g][+i].rows[+r]; return (row&&row.villa||'')+' · '+(row&&row.cat||'')+' · '+(row&&row.status||'');})()):'No room selected'}</span>
      <select id="moveWho" style="border:0;border-radius:4px;padding:5px 6px;font-size:11px">
        <option value="">Move selected to…</option>
        ${['beach','water'].flatMap(g=>model[g].map((p,i)=>`<option value="${g}.${i}">${esc(p.name||g)}</option>`)).join('')}
      </select>
      <button id="moveGo">Move</button>
      <button data-addsec="beach">+ Section</button>
      <button data-addsec="water">+ Water section</button>
      <select id="addRoomWho" style="border:0;border-radius:4px;padding:5px 6px;font-size:11px">
        ${['beach','water'].flatMap(g=>model[g].map((p,i)=>`<option value="${g}.${i}">+ Room · ${esc(p.name||g)}</option>`)).join('')}
      </select>
      <button id="addRoomGo">Add room</button>
      <button id="delSel">Remove selected</button>
    </div>
    ${grid('beach',model.beach)}
    ${grid('water',model.water)}
    <table style="margin-top:6px">
      <tr>
        <td>
          <div class="sec">VILLA ATTENDANT ON LEAVE</div>
          <table><tr><td class="tiny">NAME</td><td class="tiny">TYPE</td><td class="tiny">START</td><td class="tiny">END</td></tr>
          ${(model.leave||[]).map((l,i)=>`<tr>
            <td><input data-path="leave.${i}.name" value="${esc(l.name)}"></td>
            <td><input data-path="leave.${i}.type" value="${esc(l.type)}"></td>
            <td><input data-path="leave.${i}.start" type="date" value="${esc(l.start)}"></td>
            <td><input data-path="leave.${i}.end" type="date" value="${esc(l.end)}"></td>
          </tr>`).join('')}
          </table>
        </td>
        <td>
          <div class="sec">COLOUR LEGEND & SHORT CODE</div>
          <table>${STATUSES.map(([c,n])=>`<tr><td class="legend-${c}">${n}</td><td class="legend-${c}"><b>${c}</b></td></tr>`).join('')}</table>
        </td>
        <td>
          <div class="sec">LAUNDRY</div>
          <table>${(model.laundry||[]).map((x,i)=>`<tr>
            <td>${esc(x.no)}</td>
            <td><input data-path="laundry.${i}.name" value="${esc(x.name)}"></td>
            <td><input data-path="laundry.${i}.timing" value="${esc(x.timing)}"></td>
            <td><input data-path="laundry.${i}.desig" value="${esc(x.desig)}"></td>
          </tr>`).join('')}</table>
        </td>
      </tr>
    </table>
    <table style="margin-top:6px">
      <tr>
        <td>
          <div class="sec">HOUSEKEEPING SUPERVISOR</div>
          <table>${(model.supervisors||[]).map((x,i)=>`<tr>
            <td><input data-path="supervisors.${i}.name" value="${esc(x.name)}"></td>
            <td><input data-path="supervisors.${i}.role" value="${esc(x.role)}"></td>
            <td><input data-path="supervisors.${i}.status" value="${esc(x.status)}"></td>
            <td><input data-path="supervisors.${i}.section" value="${esc(x.section)}"></td>
          </tr>`).join('')}</table>
        </td>
        <td>
          <div class="sec">MINIBAR DUTY</div>
          <table>${(model.minibar||[]).map((x,i)=>`<tr>
            <td><input data-path="minibar.${i}.name" value="${esc(x.name)}"></td>
            <td><input data-path="minibar.${i}.status" value="${esc(x.status)}"></td>
            <td><input data-path="minibar.${i}.section" value="${esc(x.section)}"></td>
          </tr>`).join('')}</table>
          <div class="sec">OFFICE DUTY</div>
          <table>${(model.office||[]).map((x,i)=>`<tr>
            <td><input data-path="office.${i}.name" value="${esc(x.name)}"></td>
            <td><input data-path="office.${i}.role" value="${esc(x.role)}"></td>
            <td><input data-path="office.${i}.status" value="${esc(x.status)}"></td>
          </tr>`).join('')}</table>
        </td>
      </tr>
    </table>
    <table style="margin-top:6px">
      <tr>
        <td>
          <div class="sec">PUBLIC AREA</div>
          <table>${(model.publicArea||[]).map((x,i)=>`<tr>
            <td>${esc(x.no)}</td>
            <td><input data-path="publicArea.${i}.name" value="${esc(x.name)}"></td>
            <td><input data-path="publicArea.${i}.role" value="${esc(x.role)}"></td>
            <td><input data-path="publicArea.${i}.section" value="${esc(x.section)}"></td>
          </tr>`).join('')}</table>
        </td>
        <td>
          <div class="sec">ADDITIONAL TASK</div>
          <table>${(model.tasks||[]).map((x,i)=>`<tr>
            <td><input data-path="tasks.${i}.detail" value="${esc(x.detail)}"></td>
            <td><input data-path="tasks.${i}.area" value="${esc(x.area)}"></td>
            <td><input data-path="tasks.${i}.time" value="${esc(x.time)}"></td>
          </tr>`).join('')}</table>
        </td>
      </tr>
    </table>
  </div></div>`;
}
function viewBody(){
  if(viewMode==='full') return fullSheetView();
  if(viewMode==='cards'){
    return `${moveHint()}
      <div class="secbar"><h3 style="color:#1f4e79;margin:0">Beach / island · 12 slots each</h3><button class="add" style="width:auto;padding:8px 10px" data-addsec="beach">+ New section</button></div>
      <div class="attgrid ${selectedSrc?'ready':''}">${model.beach.map((p,i)=>attCard('beach',0,p,i)).join('')}</div>
      <div class="secbar"><h3 style="color:#1f4e79;margin:0">Water villas · 13 slots each</h3><button class="add" style="width:auto;padding:8px 10px" data-addsec="water">+ New section</button></div>
      <div class="attgrid ${selectedSrc?'ready':''}">${model.water.map((p,i)=>attCard('water',0,p,i)).join('')}</div>`;
  }
  if(viewMode==='sheet'){
    return `${moveHint()}
      <div class="${selectedSrc?'ready':''}">
      ${sheetBlock('Beach / island attendants · 12 slots','beach',model.beach)}
      ${sheetBlock('Water villa attendants · 13 slots','water',model.water)}
      </div>`;
  }
  if(viewMode==='master'){
    const rows=masterRows();
    return `${moveHint()}
      <input class="filter" id="masterQ" value="${esc(masterQ)}" placeholder="Search villa / attendant / status / empty">
      <div class="sheetwrap ${selectedSrc?'ready':''}"><table class="master">
        <tr><th>Villa</th><th>Cat</th><th>Status</th><th>Attendant</th></tr>
        ${rows.map(x=>{
          const id=x.g+'.'+x.pi+'.'+x.ri;
          const empty=!String(x.r.villa||'').trim();
          return `<tr class="slot ${empty?'empty':''} ${selectedSrc===id?'on':''}" data-slot="${id}">
          <td><input data-path="${pathOf(x.g,x.pi,x.ri,'villa')}" value="${esc(x.r.villa)}" placeholder="${empty?'empty slot':''}"></td>
          <td>${dd(pathOf(x.g,x.pi,x.ri,'cat'),x.r.cat,'cat')}</td>
          <td>${dd(pathOf(x.g,x.pi,x.ri,'status'),x.r.status,'status')}</td>
          <td>${esc(x.p.name)} ${empty?'· empty':''}</td>
        </tr>`;}).join('')}
      </table></div>`;
  }
  if(viewMode==='status'){
    const cols=STATUSES.map(([code,name])=>{
      const items=[];
      ['beach','water'].forEach(g=>model[g].forEach((p,pi)=>p.rows.forEach((r,ri)=>{
        if(r.status===code && r.villa) items.push({g,pi,ri,p,r});
      })));
      return `<div class="col"><h4 class="st-${code}">${code} · ${name} · ${items.length}</h4>
        ${items.map(x=>{
          const id=x.g+'.'+x.pi+'.'+x.ri;
          return `<div class="vchip slot st-${code} ${selectedSrc===id?'on':''}" data-slot="${id}">
          <b>${esc(x.r.villa)}</b>
          <div class="muted">${esc(x.p.name)} · ${esc(x.r.cat)}</div>
          ${dd(pathOf(x.g,x.pi,x.ri,'status'),x.r.status,'status')}
        </div>`;}).join('')||'<div class="hint" style="padding:8px">None</div>'}
      </div>`;
    }).join('');
    const empties=['beach','water'].flatMap(g=>model[g].flatMap((p,pi)=>p.rows.map((r,ri)=>({g,pi,ri,p,r})).filter(x=>!String(x.r.villa||'').trim())));
    return `${moveHint()}
      <div class="kanban">${cols}</div>
      <h3 style="color:#1f4e79;margin:14px 0 6px">Empty slots — tap to place selected room</h3>
      <div class="attgrid ${selectedSrc?'ready':''}">${empties.map(x=>{
        const id=x.g+'.'+x.pi+'.'+x.ri;
        return `<div class="slot empty ${selectedSrc===id?'on':''}" data-slot="${id}" style="padding:10px;border:1px dashed #c2410c;border-radius:10px;background:#fff">
          <b>${esc(x.p.name)}</b><div class="hint" style="margin:0">empty slot</div></div>`;
      }).join('')}</div>`;
  }
  return `${moveHint()}
    <div class="${selectedSrc?'ready':''}">
    ${sheetBlock('Beach · 12 slots','beach',model.beach).replace('sheetgrid','sheetgrid dense')}
    ${sheetBlock('Water · 13 slots','water',model.water).replace('sheetgrid','sheetgrid dense')}
    </div>`;
}
function extrasBlock(){
  return `<div class="card list" style="margin-top:14px">
      <h3>Villa moves</h3>
      <table><tr><th>From</th><th>To</th></tr>
      ${model.moves.map((x,i)=>`<tr><td><input data-path="moves.${i}.from" value="${esc(x.from)}"></td><td><input data-path="moves.${i}.to" value="${esc(x.to)}"></td></tr>`).join('')}
      </table>
      <button class="add" data-addmove="1">+ Add move</button>
    </div>
    <div class="card list" style="margin-top:10px">
      <h3>Leave</h3>
      <table><tr><th>Name</th><th>Type</th><th>Start</th><th>End</th></tr>
      ${model.leave.map((l,i)=>`<tr>
        <td><input data-path="leave.${i}.name" value="${esc(l.name)}"></td>
        <td><input data-path="leave.${i}.type" value="${esc(l.type)}"></td>
        <td><input data-path="leave.${i}.start" type="date" value="${esc(l.start)}"></td>
        <td><input data-path="leave.${i}.end" type="date" value="${esc(l.end)}"></td>
      </tr>`).join('')}
      </table>
    </div>`;
}
function render(){
  if(!model) return;
  if(!model.forecast) model.forecast={weather:'☀️',auto:true};
  syncForecast();
  const m=metrics();
  viewBar();
  document.getElementById('page').innerHTML=`
    <div class="bar">
      <div class="card"><h3>Duty</h3>
        <div class="field"><label>Date</label><input data-path="date" type="date" value="${esc(model.date)}"></div>
        <div class="field"><label>Manager on duty</label><input data-path="duty.mod" value="${esc(model.duty.mod)}"></div>
        <div class="field"><label>HK Day</label><input data-path="duty.hkDay" value="${esc(model.duty.hkDay)}"></div>
        <div class="field"><label>HK Night</label><input data-path="duty.hkNight" value="${esc(model.duty.hkNight)}"></div>
      </div>
      <div class="card"><h3>Contacts / forecast entry</h3>
        <div class="field"><label>Security</label><input data-path="duty.security" value="${esc(model.duty.security)}"></div>
        <div class="field"><label>Eng hotline</label><input data-path="duty.eng" value="${esc(model.duty.eng)}"></div>
        <div class="field"><label>In-house adults</label><input data-path="adults" value="${esc(model.adults)}"></div>
        <div class="field"><label>In-house children</label><input data-path="children" value="${esc(model.children)}"></div>
      </div>
      <div class="card" style="grid-column:1/-1"><h3>Forecast</h3>
        <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:8px">
          <div class="field"><label>Weather</label><button type="button" class="dd" data-dd="weather" data-path="forecast.weather" style="font-size:22px">${esc(model.forecast.weather||'☀️')}</button></div>
          <div class="field"><label>Occupancy %</label><input data-path="forecast.occupancyPct" value="${esc(model.forecast.occupancyPct)}"></div>
          <div class="field"><label>Arrival</label><input data-path="forecast.arrival" value="${esc(model.forecast.arrival)}"></div>
          <div class="field"><label>Departure</label><input data-path="forecast.departure" value="${esc(model.forecast.departure)}"></div>
          <div class="field"><label>Villa move</label><input data-path="forecast.villaMove" value="${esc(model.forecast.villaMove)}"></div>
          <div class="field"><label>Occupied villa</label><input data-path="forecast.occupied" value="${esc(model.forecast.occupied)}"></div>
          <div class="field"><label>Vacant villa</label><input data-path="forecast.vacant" value="${esc(model.forecast.vacant)}"></div>
        </div>
        <label class="tiny"><input id="autoCount" type="checkbox" ${model.forecast.auto!==false?'checked':''}> Auto ARR / DEP / OCC / VAC from allocation statuses</label>
      </div>
      <div class="card" style="grid-column:1/-1"><h3>Core value</h3>
        <div class="field"><input data-path="core" value="${esc(model.core)}"></div>
      </div>
    </div>
    <div class="kpis">
      <div class="kpi"><b>${m.rooms}</b><span>Assigned rooms</span></div>
      <div class="kpi"><b>${m.occ}</b><span>Occupied + ARR/B2B</span></div>
      <div class="kpi"><b>${m.arr}</b><span>ARR</span></div>
      <div class="kpi"><b>${m.dep}</b><span>DEP</span></div>
      <div class="kpi"><b>${m.b2b}</b><span>B2B</span></div>
      <div class="kpi"><b>${m.vac}</b><span>VAC</span></div>
    </div>
    ${viewBody()}
    ${viewMode==='full'?'':extrasBlock()}
  `;
}

function addRoom(group, pi, data){
  const p=model[group][pi];
  const payload=data||{villa:'',cat:'',status:''};
  const empty=p.rows.find(r=>!String(r.villa||'').trim());
  if(empty){
    empty.villa=payload.villa; empty.cat=payload.cat; empty.status=payload.status;
    return empty;
  }
  const rowIndex=p.extraStart + (p.rows.filter(r=>r.extra).length);
  const row={villa:payload.villa,cat:payload.cat,status:payload.status,extra:true,cells:[addr(rowIndex,p.col),addr(rowIndex,p.col+1),addr(rowIndex,p.col+2)]};
  p.rows.push(row);
  return row;
}
function addSection(group){
  const used=model[group].map(p=>p.col);
  let col=group==='beach'?33:37;
  while(used.includes(col)) col+=4;
  const headerRow=group==='beach'?32:48;
  const start=group==='beach'?36:52;
  const count=group==='beach'?12:13;
  const rows=[];
  for(let i=0;i<count;i++){
    const r=start+i;
    rows.push({villa:'',cat:'',status:'',extra:false,cells:[addr(r,col),addr(r,col+1),addr(r,col+2)]});
  }
  model[group].push({
    name:'NEW SECTION',
    nameCell:addr(headerRow,col),
    col, headerRow,
    color:COLORS[model[group].length%COLORS.length],
    extraStart:start+count,
    rows, added:true
  });
  render();
  persist('New allocation section added');
}
function takeRoom(g,pi,ri){
  const p=model[g][pi];
  const row=p.rows[ri];
  if(!row) return null;
  const data={villa:row.villa,cat:row.cat,status:row.status};
  if(row.extra) p.rows.splice(ri,1);
  else {row.villa='';row.cat='';row.status='';}
  return data;
}
function moveRoom(src, dest){
  const parts=dest.split('.');
  if(parts.length>=3) return moveToSlot(src, dest);
  const [sg,spi,sri]=src.split('.');
  const [dg,dpi]=parts;
  if(sg===dg && +spi===+dpi) return;
  const data=takeRoom(sg,+spi,+sri);
  if(!data) return;
  addRoom(dg,+dpi,data);
  render();
  persist((data.villa||'Room')+' moved with CAT '+(data.cat||'-')+' STAT '+(data.status||'-'));
}
function moveToSlot(src, dest){
  if(!src||!dest||src===dest) return;
  const [sg,spi,sri]=src.split('.');
  const [dg,dpi,dri]=dest.split('.');
  const a=model[sg]&&model[sg][+spi]&&model[sg][+spi].rows[+sri];
  const b=model[dg]&&model[dg][+dpi]&&model[dg][+dpi].rows[+dri];
  if(!a||!b) return;
  const tmp={villa:a.villa,cat:a.cat,status:a.status};
  a.villa=b.villa; a.cat=b.cat; a.status=b.status;
  b.villa=tmp.villa; b.cat=tmp.cat; b.status=tmp.status;
  selectedSrc='';
  syncForecast();
  render();
  persist((tmp.villa||'Room')+' placed with CAT '+(tmp.cat||'-')+' STAT '+(tmp.status||'-'));
}
function openMovePicker(src){
  const [g,pi,ri]=src.split('.');
  const row=model[g][+pi].rows[+ri];
  pick={kind:'move',src};
  document.getElementById('overlay').classList.add('show');
  const p=document.getElementById('picker');
  p.classList.add('show');
  const peopleHtml=['beach','water'].map(gr=>`<div class="hint">${gr==='beach'?'Beach':'Water'}</div>`+model[gr].map((x,i)=>`<div class="opt" data-dest="${gr}.${i}"><b>${esc(x.name||'Section '+(i+1))}</b><small>${filled(x).length} rooms</small></div>`).join('')).join('');
  p.innerHTML='<h3>Move villa '+(esc(row&&row.villa)||'')+' · '+(esc(row&&row.cat)||'')+' · '+(esc(row&&row.status)||'')+'</h3>'+peopleHtml;
}

function openPicker(kind,path,cur){
  pick={kind,path};
  const list=kind==='weather'?WEATHER:kind==='status'?STATUSES:CATS;
  const title=kind==='weather'?'Weather':kind==='status'?'Status':'Category';
  document.getElementById('overlay').classList.add('show');
  const p=document.getElementById('picker');
  p.classList.add('show');
  p.innerHTML='<h3>'+title+'</h3>'+
    list.map(([c,n])=>`<div class="opt ${c===cur?'on':''}" data-val="${c}"><b style="font-size:${kind==='weather'?'22px':'13px'}">${c}</b><small>${n}</small></div>`).join('')+
    (kind==='weather'?'':`<div class="opt" data-val=""><b>Clear</b></div>`);
}
function closePicker(){document.getElementById('overlay').classList.remove('show');document.getElementById('picker').classList.remove('show');pick=null;}

document.body.addEventListener('input',e=>{
  if(e.target.id==='masterQ'){ masterQ=e.target.value; render(); const n=document.getElementById('masterQ'); if(n){n.focus(); n.setSelectionRange(n.value.length,n.value.length);} return; }
  if(e.target.id==='autoCount'){
    model.forecast.auto=!!e.target.checked;
    if(model.forecast.auto){ syncForecast(); render(); }
    persist(false);
    return;
  }
  const el=e.target.closest('[data-path]');
  if(!el) return;
  setPath(el.dataset.path, el.value);
  if(/^forecast\.(arrival|departure|occupied|vacant)$/.test(el.dataset.path)) model.forecast.auto=false;
  scheduleSave();
});
document.body.addEventListener('click',e=>{
  const vw=e.target.closest('[data-view]');
  if(vw){
    viewMode=vw.dataset.view;
    localStorage.setItem('hk-daily-view', viewMode);
    render();
    return;
  }
  if(e.target.closest('[data-textimport]')){ openTextModal(); return; }
  if(e.target.closest('[data-undo]')){ doUndo(); return; }
  if(e.target.closest('[data-redo]')){ doRedo(); return; }
  if(e.target.id==='txCancel'){ closeTextModal(); return; }
  if(e.target.id==='txApply'){ applyTextImport(); return; }
  if(!e.target.closest('[data-dd],.opt,button,.filebtn')){
    const slot=e.target.closest('[data-slot]');
    if(slot){
      const id=slot.dataset.slot;
      if(selectedSrc && selectedSrc!==id){ moveToSlot(selectedSrc,id); return; }
      const now=Date.now();
      if(window.__tap && window.__tap.id===id && now-window.__tap.t<420){
        const [g,i,r]=id.split('.');
        const row=model[g]&&model[g][+i]&&model[g][+i].rows[+r];
        if(row && String(row.villa||'').trim()){
          selectedSrc = selectedSrc===id?'':id;
          window.__ignoreDbl=true;
          render();
          toast(selectedSrc?'Room selected — tap empty slot to place':'Selection cleared');
        }
        window.__tap=null;
        return;
      }
      window.__tap={id,t:now};
    }
  }
  if(e.target.id==='moveGo'){
    const dest=document.getElementById('moveWho')?.value;
    if(!selectedSrc){ toast('Double-tap a room first'); return; }
    if(!dest){ toast('Choose who receives the room'); return; }
    moveRoom(selectedSrc, dest); selectedSrc=''; return;
  }
  if(e.target.id==='delSel'){
    if(!selectedSrc){ toast('Double-tap a room first'); return; }
    const [g,pi,ri]=selectedSrc.split('.');
    const row=model[g][+pi].rows[+ri];
    if(!row) return;
    if(row.extra) model[g][+pi].rows.splice(+ri,1);
    else {row.villa='';row.cat='';row.status='';}
    selectedSrc=''; render(); persist(false); return;
  }
  if(e.target.id==='calcOcc'){
    const occ=Number(model.forecast.occupied||allocCounts().occupied)||0;
    model.forecast.occupancyPct=((occ/137)*100).toFixed(2);
    render(); persist(false); toast('Occupancy set from occupied villas / 137');
    return;
  }
  if(e.target.id==='addRoomGo'){
    const v=document.getElementById('addRoomWho')?.value;
    if(v){ const [g,i]=v.split('.'); addRoom(g,+i); render(); persist('Extra room added'); }
    return;
  }
  const add=e.target.closest('[data-add]');
  if(add){const [g,i]=add.dataset.add.split('.'); addRoom(g,+i); syncForecast(); render(); persist('Extra room added'); return;}
  const addsec=e.target.closest('[data-addsec]');
  if(addsec){ addSection(addsec.dataset.addsec); return; }
  const mv=e.target.closest('[data-move]');
  if(mv){ openMovePicker(mv.dataset.move); return; }
  if(e.target.closest('[data-addmove]')){model.moves.push({from:'',to:'',fromCell:'',toCell:''});render();persist(false);return;}
  const del=e.target.closest('[data-del]');
  if(del){
    const [g,pi,ri]=del.dataset.del.split('.');
    const row=model[g][+pi].rows[+ri];
    if(row.extra) model[g][+pi].rows.splice(+ri,1);
    else {row.villa='';row.cat='';row.status='';}
    syncForecast(); render(); persist(false); return;
  }
  const btn=e.target.closest('[data-dd]');
  if(btn){openPicker(btn.dataset.dd, btn.dataset.path, btn.textContent.trim()==='Select'?'':btn.textContent.trim());return;}
  const opt=e.target.closest('.opt');
  if(opt&&pick){
    if(pick.kind==='move' && opt.dataset.dest){ const src=pick.src; closePicker(); moveRoom(src, opt.dataset.dest); return; }
    setPath(pick.path,opt.dataset.val);
    if(String(pick.path).includes('.status')) syncForecast();
    closePicker();render();persist(false);return;
  }
  if(e.target.id==='overlay'){ closePicker(); closeTextModal(); }
});
document.body.addEventListener('dblclick',e=>{
  if(window.__ignoreDbl){ window.__ignoreDbl=false; return; }
  const slot=e.target.closest('[data-slot]');
  if(!slot) return;
  const id=slot.dataset.slot;
  const [g,i,r]=id.split('.');
  const row=model[g]&&model[g][+i]&&model[g][+i].rows[+r];
  if(!row || !String(row.villa||'').trim()) return;
  e.preventDefault();
  selectedSrc = selectedSrc===id?'':id;
  render();
  toast(selectedSrc?'Room selected — tap an empty slot to place':'Selection cleared');
});
document.body.addEventListener('dragstart',e=>{
  if(e.target.closest('input,button.dd,.dd')){ e.preventDefault(); return; }
  const grip=e.target.closest('.grip,[data-src]');
  if(!grip || !grip.dataset.src) return;
  e.dataTransfer.setData('text/plain', grip.dataset.src);
  e.dataTransfer.effectAllowed='move';
});
document.body.addEventListener('dragover',e=>{
  const z=e.target.closest('[data-drop]');
  if(!z) return;
  e.preventDefault();
  z.classList.add('dragover');
});
document.body.addEventListener('dragleave',e=>{
  const z=e.target.closest('[data-drop]');
  if(z) z.classList.remove('dragover');
});
document.body.addEventListener('drop',e=>{
  const z=e.target.closest('[data-drop]');
  if(!z) return;
  e.preventDefault();
  z.classList.remove('dragover');
  const src=e.dataTransfer.getData('text/plain');
  if(src) moveRoom(src, z.dataset.drop);
});

let ptr=null;
document.body.addEventListener('pointerdown',e=>{
  const grip=e.target.closest('.grip');
  if(!grip) return;
  selectedSrc=grip.dataset.src;
  ptr={src:grip.dataset.src,x:e.clientX,y:e.clientY,started:false,id:e.pointerId};
  try{ grip.setPointerCapture(e.pointerId); }catch(err){}
});
document.body.addEventListener('pointermove',e=>{
  if(!ptr || e.pointerId!==ptr.id) return;
  const dx=e.clientX-ptr.x, dy=e.clientY-ptr.y;
  if(!ptr.started && (dx*dx+dy*dy)>80){
    ptr.started=true;
    const [g,i,r]=ptr.src.split('.');
    const row=model[g][+i].rows[+r]||{};
    const gEl=document.createElement('div');
    gEl.className='drag-ghost';
    gEl.id='dragGhost';
    gEl.textContent=(row.villa||'')+' · '+(row.cat||'')+' · '+(row.status||'');
    document.body.appendChild(gEl);
  }
  const ghost=document.getElementById('dragGhost');
  if(ghost){ ghost.style.left=(e.clientX+12)+'px'; ghost.style.top=(e.clientY+12)+'px'; }
  document.querySelectorAll('.dragover').forEach(n=>n.classList.remove('dragover'));
  const z=document.elementFromPoint(e.clientX,e.clientY)?.closest('[data-drop]');
  if(z) z.classList.add('dragover');
});
function endPtr(e){
  if(!ptr || (e && e.pointerId!==ptr.id)) return;
  const ghost=document.getElementById('dragGhost');
  if(ghost) ghost.remove();
  document.querySelectorAll('.dragover').forEach(n=>n.classList.remove('dragover'));
  if(ptr.started){
    const z=document.elementFromPoint(e.clientX,e.clientY)?.closest('[data-drop]');
    if(z) moveRoom(ptr.src, z.dataset.drop);
  } else {
    render();
    toast('Room selected — type the number, or Move to…');
  }
  ptr=null;
}
document.body.addEventListener('pointerup',endPtr);
document.body.addEventListener('pointercancel',endPtr);

function excelDate(iso){if(!iso||!/^\d{4}-\d{2}-\d{2}/.test(iso))return '';const [y,m,d]=iso.split('-').map(Number);return Math.floor(Date.UTC(y,m-1,d)/86400000)+25569;}
function xmlEsc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function patchCell(xml,cell,value,type){
  if(!cell) return xml;
  const re=new RegExp('<c r="'+cell+'"([^>]*)(?:/>|>[\\s\\S]*?</c>)');
  const style=(xml.match(re)?(xml.match(re)[1].match(/\bs="(\d+)"/)||[])[1]:null);
  const sAttr=style?(' s="'+style+'"'):'';
  let neu;
  if(value===''||value==null) neu='<c r="'+cell+'"'+sAttr+'/>';
  else if(type==='n') neu='<c r="'+cell+'"'+sAttr+'><v>'+value+'</v></c>';
  else neu='<c r="'+cell+'"'+sAttr+' t="inlineStr"><is><t xml:space="preserve">'+xmlEsc(value)+'</t></is></c>';
  if(re.test(xml)) return xml.replace(re,neu);
  const rowNum=parseInt(cell.replace(/^[A-Z]+/,''),10);
  const rowRe=new RegExp('(<row r="'+rowNum+'"[^>]*>)([\\s\\S]*?)(</row>)');
  if(rowRe.test(xml)) return xml.replace(rowRe,(_,a,mid,b)=>a+mid+neu+b);
  const sheetData=xml.indexOf('<sheetData>');
  if(sheetData>=0){
    return xml.replace('<sheetData>','<sheetData><row r="'+rowNum+'">'+neu+'</row>');
  }
  return xml;
}
function collectWrites(){
  const out=[];
  const add=(cell,value,type='s')=>{if(cell) out.push({cell,value,type});};
  add('AG1',model.duty.mod);add('AG2',model.duty.hkDay);add('AG3',model.duty.hkNight);
  add('AG4',model.duty.security);add('AG5',model.duty.eng);add('I7',model.core);
  add('F8', model.forecast&&model.forecast.weather);
  if(model.forecast){
    const pct=Number(String(model.forecast.occupancyPct).replace('%',''));
    add('F10', isNaN(pct)?'':(pct>1?pct/100:pct),'n');
    add('F11', model.forecast.arrival,'n');
    add('F12', model.forecast.departure,'n');
    add('F13', model.forecast.villaMove,'n');
    add('F14', model.forecast.occupied,'n');
    add('F15', model.forecast.vacant,'n');
  }
  add('F16',model.adults,'n');add('F17',model.children,'n');
  if(model.date) add('O4',excelDate(model.date),'n');
  people().forEach(p=>{
    add(p.nameCell,p.name);
    p.rows.forEach(r=>{
      add(r.cells[0], r.villa, /^\d+$/.test(String(r.villa))?'n':'s');
      add(r.cells[1], r.cat);
      add(r.cells[2], r.status);
    });
  });
  model.moves.forEach(m=>{add(m.fromCell,m.from,/^\d+$/.test(m.from)?'n':'s');add(m.toCell,m.to,/^\d+$/.test(m.to)?'n':'s');});
  model.leave.forEach(l=>{add(l.cells.name,l.name);add(l.cells.type,l.type);add(l.cells.start,l.start?excelDate(l.start):'','n');add(l.cells.end,l.end?excelDate(l.end):'','n');});
  (model.laundry||[]).forEach(x=>{if(!x.cells)return;add(x.cells.name,x.name);add(x.cells.timing,x.timing);add(x.cells.desig,x.desig);add(x.cells.contact,x.contact);});
  (model.supervisors||[]).forEach(x=>{if(!x.cells)return;add(x.cells.name,x.name);add(x.cells.role,x.role);add(x.cells.status,x.status);add(x.cells.section,x.section);add(x.cells.contact,x.contact);});
  (model.minibar||[]).forEach(x=>{if(!x.cells)return;add(x.cells.name,x.name);add(x.cells.status,x.status);add(x.cells.section,x.section);});
  (model.office||[]).forEach(x=>{if(!x.cells)return;add(x.cells.name,x.name);add(x.cells.role,x.role);add(x.cells.status,x.status);add(x.cells.contact,x.contact);});
  (model.publicArea||[]).forEach(x=>{if(!x.cells)return;add(x.cells.name,x.name);add(x.cells.role,x.role);add(x.cells.section,x.section);});
  (model.tasks||[]).forEach(x=>{if(!x.cells)return;add(x.cells.detail,x.detail);add(x.cells.area,x.area);add(x.cells.time,x.time);});
  ['honeymoon','birthday','anniversary','upon'].forEach(k=> (model[k]||[]).forEach(x=>add(x.cell,x.v,/^\d+$/.test(String(x.v))?'n':'s')));
  return out;
}
async function downloadXlsm(){
  if(!originalBuf||!model){toast('Import XLSM first');return;}
  const zip=await JSZip.loadAsync(originalBuf);
  let xml=await zip.file('xl/worksheets/sheet1.xml').async('string');
  collectWrites().forEach(w=>{xml=patchCell(xml,w.cell,w.value,w.type);});
  zip.file('xl/worksheets/sheet1.xml',xml);
  const blob=await zip.generateAsync({type:'blob',mimeType:'application/vnd.ms-excel.sheet.macroEnabled.12'});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download=(fileName||'HK-Daily').replace(/\.xlsm$/i,'')+'-updated.xlsm';
  a.click();
  await persist(false);
  toast('XLSM downloaded and saved');
}
function applyImported(){
  showEditor();
  render();
}
async function importFile(file){
  if(!file) return;
  try{
    if(typeof XLSX==='undefined'){
      toast('Excel library not loaded. Open this file with internet once');
      return;
    }
    fileName=file.name;
    originalBuf=await file.arrayBuffer();
    const wb=XLSX.read(new Uint8Array(originalBuf),{type:'array',cellDates:true, raw:false});
    if(!wb.SheetNames||!wb.SheetNames.length) throw new Error('No sheets in file');
    model=ensureModel(parseWorkbook(wb));
    syncForecast();
    const rooms=people().reduce((n,p)=>n+filled(p).length,0);
    applyImported();
    await persist('Imported '+fileName+' · '+rooms+' rooms');
  }catch(err){
    console.error(err);
    toast('Import failed: '+(err.message||err));
  }
}
async function restoreSaved(){
  try{
    const pack=await idbGet('pack');
    if(!pack||!pack.buf||!pack.model) return false;
    fileName=pack.fileName||'HK Daily.xlsm';
    savedAt=pack.savedAt||'';
    originalBuf=pack.buf;
    model=ensureModel(pack.model);
    undoStack=Array.isArray(pack.undoStack)?pack.undoStack:[];
    redoStack=Array.isArray(pack.redoStack)?pack.redoStack:[];
    if(model.forecast&&model.forecast.auto!==false) syncForecast();
    applyImported();
    toast('Opened last saved file');
    return true;
  }catch(e){ return false; }
}
document.getElementById('fileIn').onchange=e=>importFile(e.target.files[0]);
document.getElementById('fileIn2').onchange=e=>importFile(e.target.files[0]);
document.getElementById('btnOut').onclick=downloadXlsm;
document.getElementById('btnSave').onclick=()=>persist('Saved on this phone');
document.getElementById('btnRoll').onclick=()=>rollYesterdayToToday();
document.getElementById('btnText').onclick=()=>openTextModal();
document.getElementById('btnUndo').onclick=()=>doUndo();
document.getElementById('btnRedo').onclick=()=>doRedo();
document.getElementById('btnOpenSaved').onclick=()=>restoreSaved();
document.getElementById('btnClear').onclick=async()=>{
  if(!confirm('Clear saved file and start again? Edits on this phone will be removed. The original Excel on your phone is not deleted.')) return;
  try{ await idbSet('pack', null); }catch(e){}
  localStorage.removeItem('hk-daily-meta');
  originalBuf=null; model=null; fileName=''; savedAt='';
  document.getElementById('gate').hidden=false;
  document.getElementById('page').hidden=true;
  document.getElementById('page').innerHTML='';
  document.getElementById('btnOut').disabled=true;
  document.getElementById('btnSave').disabled=true;
  document.getElementById('btnClear').hidden=true;
  document.getElementById('viewBar').hidden=true;
  document.getElementById('fileName').textContent='Import orange HK Daily first';
  toast('Saved copy cleared');
};
(async()=>{
  const meta=(()=>{try{return JSON.parse(localStorage.getItem('hk-daily-meta')||'null');}catch(e){return null;}})();
  if(meta&&meta.fileName){
    const b=document.getElementById('btnOpenSaved');
    b.style.display='block';
    b.textContent='Open last saved: '+meta.fileName;
  }
  await restoreSaved();
})();

/* PWA install + online badge */
(function(){
  let deferred;
  window.addEventListener('beforeinstallprompt', function(e){
    e.preventDefault();
    deferred = e;
    const b = document.getElementById('btnInstall');
    if(b) b.hidden = false;
  });
  window.addEventListener('appinstalled', function(){
    const b = document.getElementById('btnInstall');
    if(b) b.hidden = true;
    deferred = null;
    if(typeof toast==='function') toast('HK Daily installed on this phone');
  });
  document.addEventListener('click', function(e){
    if(e.target && e.target.id==='btnInstall'){
      if(deferred){ deferred.prompt(); deferred.userChoice.finally(function(){ deferred=null; const b=document.getElementById('btnInstall'); if(b) b.hidden=true; }); }
      else if(typeof toast==='function') toast('Use browser menu → Add to Home Screen');
    }
  });
})();
