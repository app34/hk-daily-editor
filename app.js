const STATUSES=[['OCC','OCCUPIED'],['ARR','ARRIVAL'],['DEP','DEPATURE'],['B2B','BACK TO BACK'],['VAC','VACANT'],['OOO','OUT OF ORDER'],['SV','SHOW VILLA'],['STB','STAND BY'],['DU','DAY USE'],['OS','OUT OF SERVICE'],['DC','DEEP CLEANING'],['VM-ARR','VILLA MOVE ARRIVAL']];
const CATS=[['BHV','Beach Villa'],['DVP','Deluxe Pool'],['BFS','Beach Family'],['WAV','Water Villa'],['WVP','Water Pool'],['HWS','Horizon Water']];
const WEATHER=[['💧','Water / humid'],['☀️','Sunny'],['🌤️','Mostly sunny'],['⛅','Partly cloudy'],['☁️','Cloudy'],['🌦️','Showers'],['🌧️','Rain'],['⛈️','Thunderstorm'],['☔','Umbrella'],['🌫️','Fog'],['💨','Windy'],['🌊','Rough sea'],['🌙','Night'],['❄️','Cool']];
const COLORS=['#5b9bd5','#c45911','#70ad47','#ed7d31','#9b6bdf','#2e75b6','#548235','#203864','#00b0f0'];
const ALLOC_COLS=[1,5,9,13,17,21,25,29,33,37];
const G1=ALLOC_COLS.slice();
const G2=ALLOC_COLS.slice();
const SHEET='JULY- 2026';
let originalBuf=null,sourceBuf=null,fileName='',model=null,pick=null,savedAt='';
let viewMode=localStorage.getItem('hk-daily-view')||'full';
let sheetZoom=Number(localStorage.getItem('hk-zoom')||100);
let sheetScroll={x:0,y:0,wx:0,wy:0};
let masterQ='';
let selectedSrc='';
let selectedSec='';
let undoStack=[], redoStack=[];
const STATUS_RGB={
  ARR:'FFE181',DEP:'F8A968',VAC:'F9B9F4',B2B:'F56B96',
  OCC:'A9D098','VM-ARR':'F3FDC3',SV:'B6FCC3',DC:'FCFEE8',
  DU:'F4B183',STB:'D4F5FC',OS:'6DD9FF',OOO:'BFBFBF'
};
function inHouse(st){return ['OCC','ARR','B2B','DU','VM-ARR'].includes(st);}
function isVacantLike(st){return !st || st==='VAC';}
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
  if(!(sourceBuf||originalBuf)||!model) return;
  savedAt=new Date().toISOString();
  try{
    await idbSet('pack',{
      fileName,
      savedAt,
      buf:sourceBuf||originalBuf,
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
  let extra=fileName + (t?(' · saved '+t):'');
  if(saveMode==='folder' && saveDirHandle) extra+=' · folder '+saveDirHandle.name;
  document.getElementById('fileName').textContent = extra;
}
function showEditor(){
  document.getElementById('gate').hidden=true;
  document.getElementById('page').hidden=false;
  document.getElementById('btnOut').disabled=false;
  document.getElementById('btnSave').disabled=false;
  document.getElementById('btnClear').hidden=false;
  document.getElementById('btnText').disabled=false;
  const pickBtn=document.getElementById('btnPick');
  if(pickBtn) pickBtn.disabled=false;
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
  for(let r=67;r<=78;r++){
    leave.push({
      name:val(ws,addr(r,1)),
      type:val(ws,addr(r,4)),
      start:isoDate(val(ws,addr(r,7))),
      end:isoDate(val(ws,addr(r,10))),
      days:val(ws,addr(r,12)),
      cells:{name:addr(r,1),type:addr(r,4),start:addr(r,7),end:addr(r,10),days:addr(r,12)}
    });
  }
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
    vip:pack(range(10,17),range(7,11)),
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
  ['moves','arrivals','departures','honeymoon','birthday','anniversary','upon','vip','leave','laundry','supervisors','minibar','office','publicArea','tasks','beach','water'].forEach(k=>{ if(!m[k]) m[k]=[]; });
  if(!m.duty) m.duty={};
  padGroup(m.beach,'beach');
  padGroup(m.water,'water');
  padEmptySections(m,'beach');
  padEmptySections(m,'water');
  padTrafficLists(m);
  ensureLeave(m);
  seedNames(m);
  seedLeaveTypes(m);
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
  return JSON.parse(JSON.stringify({beach:model.beach,water:model.water,forecast:model.forecast,moves:model.moves,arrivals:model.arrivals,departures:model.departures,honeymoon:model.honeymoon,birthday:model.birthday,anniversary:model.anniversary,upon:model.upon,vip:model.vip}));
}
function applySnap(s){
  if(!s) return;
  model.beach=s.beach; model.water=s.water; model.forecast=s.forecast; model.moves=s.moves;
  if(s.arrivals) model.arrivals=s.arrivals;
  if(s.departures) model.departures=s.departures;
  if(s.honeymoon) model.honeymoon=s.honeymoon;
  if(s.birthday) model.birthday=s.birthday;
  if(s.anniversary) model.anniversary=s.anniversary;
  if(s.upon) model.upon=s.upon;
  if(s.vip) model.vip=s.vip;
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
function isoDate(v){
  const s=String(v||'').trim();
  if(!s) return '';
  const m=s.match(/(\d{4}-\d{2}-\d{2})/);
  if(m) return m[1];
  return s.slice(0,10);
}
function leaveDays(start,end){
  const a=isoDate(start), b=isoDate(end);
  if(!a||!b) return '';
  const x=new Date(a+'T00:00:00'), y=new Date(b+'T00:00:00');
  if(isNaN(x)||isNaN(y)) return '';
  const n=Math.round((y-x)/86400000)+1;
  return n>0?String(n):'';
}
function leaveTypeCode(t){
  const s=String(t||'').trim().toUpperCase();
  if(s==='AL'||s==='ANNUAL'||s==='ANNUAL LEAVE') return 'AL';
  if(s==='OFF'||s==='OFF DAY'||s==='OFF DAYS') return 'OFF';
  if(s==='ML'||s==='MEDICAL'||s==='MEDICAL LEAVE') return 'ML';
  if(s==='EL'||s==='EMERGENCY'||s==='EMERGENCY LEAVE') return 'EL';
  return s;
}
function leaveSummary(){
  const rows=(model.leave||[]).filter(l=>String(l.name||'').trim());
  const c={AL:0,OFF:0,ML:0,EL:0,other:0};
  rows.forEach(l=>{
    const k=leaveTypeCode(l.type);
    if(c[k]==null) c.other++; else c[k]++;
  });
  return {total:rows.length, AL:c.AL, OFF:c.OFF, ML:c.ML, EL:c.EL};
}
function ensureLeave(m){
  if(!m) return m;
  const slots=[];
  for(let r=67;r<=78;r++){
    slots.push({name:'',type:'',start:'',end:'',days:'',cells:{name:addr(r,1),type:addr(r,4),start:addr(r,7),end:addr(r,10),days:addr(r,12)}});
  }
  (m.leave||[]).forEach((item,i)=>{
    let hit=item.cells&&item.cells.name?slots.find(s=>s.cells.name===item.cells.name):null;
    if(!hit) hit=slots.find(s=>!s.name);
    if(!hit) hit=slots[Math.min(i,slots.length-1)];
    if(!hit) return;
    hit.name=item.name||'';
    hit.type=item.type||'';
    hit.start=isoDate(item.start);
    hit.end=isoDate(item.end);
    hit.days=leaveDays(hit.start,hit.end);
  });
  m.leave=slots;
  return m;
}
function paintLeaveDays(){
  (model.leave||[]).forEach((l,i)=>{
    l.days=leaveDays(l.start,l.end);
    const n=document.querySelector('[data-leavedays="'+i+'"]');
    if(n) n.textContent=l.days||'';
  });
  const s=leaveSummary();
  const box=document.getElementById('leaveSummary');
  if(box) box.innerHTML=`<b>On leave ${s.total}</b> · Annual ${s.AL} · Off day ${s.OFF} · Medical ${s.ML} · Emergency ${s.EL}`;
}
function parseRooms(text){
  const out=[];
  String(text||'').split(/[^0-9]+/).filter(Boolean).forEach(chunk=>{
    if(chunk.length<=3){ out.push(chunk); return; }
    for(let i=0;i<chunk.length;i+=3){
      const part=chunk.slice(i,i+3);
      if(part.length===3) out.push(part);
    }
  });
  return [...new Set(out)];
}
function formatRoomText(text){ return parseRooms(text).join(' '); }
function parseMoveLines(text){
  const out=[];
  String(text||'').split(/\n+/).forEach(line=>{
    const nums=parseRooms(line);
    for(let i=0;i+1<nums.length;i+=2) out.push({from:nums[i],to:nums[i+1]});
  });
  return out;
}
function formatMoveText(text){
  return parseMoveLines(text).map(m=>m.from+'-'+m.to).join('\n');
}
const TRAFFIC={
  arrivals:{rows:[19,29],cols:[9,15]},
  departures:{rows:[19,29],cols:[17,23]},
  honeymoon:{rows:[19,29],cols:[25,27]},
  birthday:{rows:[19,29],cols:[29,31]},
  anniversary:{rows:[19,29],cols:[33,35]},
  upon:{rows:[19,29],cols:[37,39]},
  vip:{rows:[10,17],cols:[7,11],fill:'col'}
};
function padChipSlots(list, rows, cols){
  const slots=[];
  for(let r=rows[0];r<=rows[1];r++){
    for(let c=cols[0];c<=cols[1];c++) slots.push({v:'',cell:addr(r,c)});
  }
  (list||[]).forEach(item=>{
    if(!item) return;
    const hit=slots.find(s=>s.cell===item.cell);
    if(hit) hit.v=item.v||'';
    else {
      const empty=slots.find(s=>!String(s.v||'').trim());
      if(empty) empty.v=item.v||'';
    }
  });
  return slots;
}
function padMoveSlots(list){
  const slots=[];
  [[1,2],[3,4],[6,7]].forEach(([a,b])=>{
    for(let r=20;r<=29;r++) slots.push({from:'',to:'',fromCell:addr(r,a),toCell:addr(r,b)});
  });
  (list||[]).forEach(item=>{
    const hit=item.fromCell?slots.find(s=>s.fromCell===item.fromCell):null;
    if(hit){ hit.from=item.from||''; hit.to=item.to||''; }
    else {
      const empty=slots.find(s=>!s.from && !s.to);
      if(empty){ empty.from=item.from||''; empty.to=item.to||''; }
    }
  });
  return slots;
}
function padTrafficLists(m){
  if(!m) return m;
  Object.keys(TRAFFIC).forEach(k=>{
    const spec=TRAFFIC[k];
    m[k]=padChipSlots(m[k], spec.rows, spec.cols);
  });
  m.moves=padMoveSlots(m.moves);
  return m;
}
function fillChipList(key, nums){
  const list=model[key]||[];
  list.forEach(x=>x.v='');
  const spec=TRAFFIC[key];
  if(spec && spec.fill==='col'){
    const cols=spec.cols[1]-spec.cols[0]+1;
    const rows=spec.rows[1]-spec.rows[0]+1;
    nums.forEach((n,i)=>{
      const c=Math.floor(i/rows);
      const r=i%rows;
      const idx=r*cols+c;
      if(list[idx]) list[idx].v=n;
    });
    return;
  }
  nums.forEach((n,i)=>{ if(list[i]) list[i].v=n; });
}
function fillMoveList(pairs){
  (model.moves||[]).forEach(x=>{ x.from=''; x.to=''; });
  pairs.forEach((m,i)=>{
    if(model.moves[i]){ model.moves[i].from=m.from; model.moves[i].to=m.to; }
  });
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
function normalizeTextFields(){
  const roomIds=['txArr','txDep','txOcc','txVac','txHoney','txBirth','txAnn','txUpon','txVip'];
  roomIds.forEach(id=>{
    const el=document.getElementById(id);
    if(el && el.value.trim()) el.value=formatRoomText(el.value);
  });
  const mv=document.getElementById('txMove');
  if(mv && mv.value.trim()) mv.value=formatMoveText(mv.value);
}
function applyImportBags(bags, source){
  const A=new Set(bags.arrivals||[]);
  const D=new Set(bags.departures||[]);
  const O=new Set(bags.occupied||[]);
  const V=new Set(bags.vacant||[]);
  const mv=bags.moves||[];
  const Hn=bags.honeymoon||[];
  const Bd=bags.birthday||[];
  const An=bags.anniversary||[];
  const Up=bags.upon||[];
  const Vp=bags.vip||[];
  if(!A.size && !D.size && !O.size && !V.size && !mv.length && !Hn.length && !Bd.length && !An.length && !Up.length && !Vp.length){ toast('Select or paste at least one villa'); return; }
  pushUndo();
  const fromSet=new Set(mv.map(x=>String(x.from)));
  const toSet=new Set(mv.map(x=>String(x.to)));
  let missing=[], flagged=0, changed=0;
  const listed=[...A,...D,...O,...V,...fromSet,...toSet];
  listed.forEach(num=>{ if(!findVilla(num).length) missing.push(num); });
  ['beach','water'].forEach(g=>model[g].forEach(p=>p.rows.forEach(r=>{
    const v=String(r.villa||'').trim();
    if(!v){ r.flag=''; return; }
    r.flag='';
    const incoming=A.has(v) || toSet.has(v);
    const leaving=D.has(v) || fromSet.has(v);
    const cur=r.status;
    if(inHouse(cur) && A.has(v) && !leaving){
      r.flag='occ-arr'; flagged++; return;
    }
    if(isVacantLike(cur) && D.has(v) && !incoming){
      r.flag='vac-dep'; flagged++; return;
    }
    if(isVacantLike(cur) && fromSet.has(v) && !incoming){
      r.flag='vac-move'; flagged++; return;
    }
    if(inHouse(cur) && toSet.has(v) && !leaving){
      r.flag='occ-movein'; flagged++; return;
    }
    let next=cur;
    if(leaving && incoming) next='B2B';
    else if(A.has(v)) next='ARR';
    else if(toSet.has(v)) next='VM-ARR';
    else if(D.has(v) || fromSet.has(v)) next='DEP';
    else if(O.has(v)) next='OCC';
    else if(V.has(v)) next='VAC';
    if(next!==cur){ r.status=next; changed++; }
  })));
  if(A.size) fillChipList('arrivals',[...A]);
  if(D.size) fillChipList('departures',[...D]);
  if(mv.length) fillMoveList(mv);
  if(Hn.length) fillChipList('honeymoon',Hn);
  if(Bd.length) fillChipList('birthday',Bd);
  if(An.length) fillChipList('anniversary',An);
  if(Up.length) fillChipList('upon',Up);
  if(Vp.length) fillChipList('vip',Vp);
  if(model.forecast && model.forecast.auto!==false){
    model.forecast.villaMove=mv.length||model.forecast.villaMove;
  }
  syncForecast();
  closeTextModal();
  closePickModal();
  render();
  persist(source==='pick'?'Statuses updated from room list':'Statuses updated from text');
  const bits=[];
  if(changed) bits.push(changed+' rooms updated');
  if(flagged) bits.push(flagged+' need a manual check (red highlight)');
  if(missing.length) bits.push(missing.length+' numbers not on board');
  toast((bits.join('. ')||'No status change')+'. Undo if needed');
}
function applyTextImport(){
  normalizeTextFields();
  applyImportBags({
    arrivals:parseRooms(document.getElementById('txArr').value),
    departures:parseRooms(document.getElementById('txDep').value),
    occupied:parseRooms(document.getElementById('txOcc').value),
    vacant:parseRooms(document.getElementById('txVac').value),
    moves:parseMoveLines(document.getElementById('txMove').value),
    honeymoon:parseRooms(document.getElementById('txHoney')?document.getElementById('txHoney').value:''),
    birthday:parseRooms(document.getElementById('txBirth')?document.getElementById('txBirth').value:''),
    anniversary:parseRooms(document.getElementById('txAnn')?document.getElementById('txAnn').value:''),
    upon:parseRooms(document.getElementById('txUpon')?document.getElementById('txUpon').value:''),
    vip:parseRooms(document.getElementById('txVip')?document.getElementById('txVip').value:'')
  }, 'text');
}
const PICK_MODES=[
  ['arrivals','Arrival'],
  ['departures','Departure'],
  ['moves','Room move'],
  ['vip','VIP Arrival'],
  ['honeymoon','Honeymoon'],
  ['birthday','Birthday'],
  ['anniversary','Anniversary'],
  ['upon','Upon Arrival'],
  ['occupied','Occupied'],
  ['vacant','Vacant']
];
let pickBags=emptyPickBags();
let pickMode='arrivals';
let pickFrom='';
function emptyPickBags(){
  return {arrivals:[],departures:[],occupied:[],vacant:[],vip:[],honeymoon:[],birthday:[],anniversary:[],upon:[],moves:[]};
}
function allResortRooms(){
  const out=[];
  [[101,139],[201,214],[301,307],[401,445],[501,526],[601,606]].forEach(([a,b])=>{
    for(let i=a;i<=b;i++) out.push(String(i));
  });
  return out;
}
function dropPickRoom(num){
  num=String(num);
  if(pickMode==='moves'){
    pickBags.moves=pickBags.moves.filter(m=>m.from!==num && m.to!==num);
    if(pickFrom===num) pickFrom='';
    return;
  }
  const list=pickBags[pickMode]||[];
  const i=list.indexOf(num);
  if(i>=0) list.splice(i,1);
}
function togglePickRoom(num){
  num=String(num);
  if(pickMode==='moves'){
    if(pickBags.moves.some(m=>m.from===num||m.to===num)){
      dropPickRoom(num);
      renderPickModal();
      return;
    }
    if(!pickFrom){ pickFrom=num; renderPickModal(); return; }
    if(pickFrom===num){ pickFrom=''; renderPickModal(); return; }
    pickBags.moves.push({from:pickFrom,to:num});
    pickFrom='';
    renderPickModal();
    return;
  }
  const list=pickBags[pickMode]||(pickBags[pickMode]=[]);
  const i=list.indexOf(num);
  if(i>=0) list.splice(i,1); else list.push(num);
  renderPickModal();
}
function roomSelected(num){
  num=String(num);
  if(pickMode==='moves'){
    if(pickFrom===num) return 'from';
    if(pickBags.moves.some(m=>m.from===num||m.to===num)) return 'on';
    return '';
  }
  return (pickBags[pickMode]||[]).includes(num)?'on':'';
}
function renderPickModal(){
  const box=document.getElementById('pickBody');
  if(!box) return;
  const rooms=allResortRooms();
  const lab=(PICK_MODES.find(x=>x[0]===pickMode)||[])[1]||'';
  const bits=pickMode==='moves'?pickBags.moves.map(m=>m.from+'→'+m.to):(pickBags[pickMode]||[]);
  const empty=pickMode==='moves'?(pickFrom?'FROM '+pickFrom+' — tap TO':'Tap FROM then TO'):'None yet';
  box.innerHTML=`
    <div class="pmode">${PICK_MODES.map(([k,l])=>`<button type="button" class="${pickMode===k?'on':''}" data-pmode="${k}">${l}${k===pickMode&&bits.length?' · '+bits.length:''}</button>`).join('')}</div>
    <div class="pnow"><b>${lab}</b><span>${bits.length?(pickMode==='moves'?pickBags.moves.map((m,i)=>`<button type="button" class="pch" data-prm="${i}">${m.from}→${m.to} ×</button>`).join(''):bits.map(n=>`<button type="button" class="pch" data-punpick="${n}">${n} ×</button>`).join('')):empty}</span></div>
    <div class="prooms">${rooms.map(n=>`<button type="button" class="proom ${roomSelected(n)}" data-proom="${n}">${n}</button>`).join('')}</div>`;
}
function openPickModal(){
  if(!model){ toast('Import XLSM first'); return; }
  pickBags=emptyPickBags();
  pickMode='arrivals';
  pickFrom='';
  renderPickModal();
  document.getElementById('overlay').classList.add('show');
  document.getElementById('pickModal').classList.add('show');
}
function closePickModal(){
  const m=document.getElementById('pickModal');
  if(m) m.classList.remove('show');
  if(!document.getElementById('picker').classList.contains('show') && !document.getElementById('textModal').classList.contains('show') && !document.getElementById('peopleModal').classList.contains('show'))
    document.getElementById('overlay').classList.remove('show');
}
function applyPickImport(){
  applyImportBags({
    arrivals:pickBags.arrivals.slice(),
    departures:pickBags.departures.slice(),
    occupied:pickBags.occupied.slice(),
    vacant:pickBags.vacant.slice(),
    moves:pickBags.moves.slice(),
    honeymoon:pickBags.honeymoon.slice(),
    birthday:pickBags.birthday.slice(),
    anniversary:pickBags.anniversary.slice(),
    upon:pickBags.upon.slice(),
    vip:pickBags.vip.slice()
  }, 'pick');
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
  if(/^leave\.\d+\.(start|end|type|name)$/.test(path)){
    if(/\.type$/.test(path) && value){
      seedLeaveTypes(model);
      const t=String(value).trim();
      if(t && !model.leaveTypes.some(x=>x.toLowerCase()===t.toLowerCase())) model.leaveTypes.push(t);
    }
    paintLeaveDays();
  }
}
const NAME_GROUPS=[
  ['villa','Villa attendants'],
  ['laundry','Laundry'],
  ['public','Public area'],
  ['supervisor','Supervisors'],
  ['minibar','Minibar'],
  ['office','Office']
];
const DEFAULT_LEAVE_TYPES=['TASK FORCE','ON DUTY','AL','OFF','ML','NP','EL'];
function seedLeaveTypes(m){
  if(!m) return m;
  if(!Array.isArray(m.leaveTypes) || !m.leaveTypes.length) m.leaveTypes=DEFAULT_LEAVE_TYPES.slice();
  return m;
}
function seedNames(m){
  if(!m) return m;
  if(!m.names) m.names={villa:[],laundry:[],public:[],supervisor:[],minibar:[],office:[]};
  const add=(key,val)=>{
    const n=String(val||'').trim();
    if(!n) return;
    if(!m.names[key]) m.names[key]=[];
    if(!m.names[key].some(x=>x.toLowerCase()===n.toLowerCase())) m.names[key].push(n);
  };
  (m.beach||[]).forEach(p=>add('villa',p.name));
  (m.water||[]).forEach(p=>add('villa',p.name));
  (m.laundry||[]).forEach(p=>add('laundry',p.name));
  (m.publicArea||[]).forEach(p=>add('public',p.name));
  (m.supervisors||[]).forEach(p=>add('supervisor',p.name));
  (m.minibar||[]).forEach(p=>add('minibar',p.name));
  (m.office||[]).forEach(p=>add('office',p.name));
  (m.leave||[]).forEach(p=>add('villa',p.name));
  return m;
}
function nameInput(path,value,list){
  const key=String(list||'villa').replace(/^dl-/,'');
  return `<span class="namepick"><input data-path="${path}" value="${esc(value)}" autocomplete="off" placeholder="Type or pick"><button type="button" class="namebtn" data-dd="name" data-nlist="${key}" data-path="${path}" aria-label="Pick name">▾</button></span>`;
}
function nameLists(){ return ''; }
function packSection(p){
  if(!p||!p.rows) return;
  const data=p.rows.map(r=>({villa:r.villa,cat:r.cat,status:r.status,flag:r.flag})).filter(r=>String(r.villa||'').trim());
  p.rows.forEach((r,i)=>{
    const d=data[i];
    if(d){ r.villa=d.villa; r.cat=d.cat; r.status=d.status; r.flag=d.flag; }
    else { r.villa=''; r.cat=''; r.status=''; r.flag=''; }
  });
}
function dd(path,val,kind){
  return `<button type="button" class="dd st-${esc(val||'')}" data-dd="${kind}" data-path="${path}">${esc(val||'Select')}</button>`;
}

function attCard(group, gi, p, pi){
  const rooms=p.rows;
  return `<article class="att dropzone" data-drop="${group}.${pi}">
    <div class="hd" style="background:${p.color}">
      ${nameInput(group+'.'+pi+'.name', p.name, 'villa')}
      <span class="cnt">${filled(p).length} rooms</span>
    </div>
    <table>
      <tr><th>V#</th><th>CAT</th><th>STAT</th><th></th></tr>
      ${rooms.map((r,ri)=>{
        const id=group+'.'+pi+'.'+ri;
        const empty=!String(r.villa||'').trim();
        return `<tr class="${slotClass(group,pi,ri,r)}" data-slot="${id}" title="${esc(flagTitle(r))}">
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
  const secLab=selectedSec?(()=>{const [g,i]=selectedSec.split('.'); const p=model[g][+i]; return (p&&p.name)||'Empty section';})():'';
  return `<div class="card" style="margin:0 0 10px;${selectedSrc||selectedSec?'background:#9a3412;color:#fff;border-color:#9a3412':''}">
    <b>${selectedSec?('SELECTED SECTION '+secLab):selectedSrc?('SELECTED '+sel):'No room selected'}</b>
    <div class="hint" style="margin:4px 0 0;${selectedSrc||selectedSec?'color:#ffedd5':''}">${selectedSec?'Section selected — tap an empty section name to place it.':'Double-tap a room or a section name. Empty sections stay visible so a full column can move.'}</div>
  </div>`;
}
function pathOf(group,pi,ri,field){return group+'.'+pi+'.rows.'+ri+'.'+field;}
function slotClass(group,pi,ri,row){
  const id=group+'.'+pi+'.'+ri;
  const empty=!String(row&&row.villa||'').trim();
  const st=esc(row&&row.status||'');
  return `slot ${empty?'empty':''} ${selectedSrc===id?'on':''} ${row&&row.flag?'flag':''} st-${st}`;
}
function flagTitle(row){
  const m={
    'occ-arr':'Occupied + arrival, no departure — not auto-updated',
    'vac-dep':'Vacant + departure — not auto-updated',
    'vac-move':'Vacant + room move out — not auto-updated',
    'occ-movein':'Occupied + room move in, no departure — not auto-updated'
  };
  return row&&row.flag? m[row.flag]||'Check this room':'';
}
function isSectionEmpty(p){
  return !String(p.name||'').trim() && !filled(p).length;
}
function padEmptySections(m, group){
  if(!m[group]) m[group]=[];
  const have=new Set(m[group].map(p=>p.col));
  ALLOC_COLS.forEach(c=>{
    if(!have.has(c)) addSection(group, {silent:true, empty:true, modelRef:m, col:c});
  });
  m[group].sort((a,b)=>(a.col||0)-(b.col||0));
}
function flagBanner(){
  const items=[];
  ['beach','water'].forEach(g=>model[g].forEach((p,pi)=>p.rows.forEach((r,ri)=>{
    if(r.flag && String(r.villa||'').trim()) items.push({g,pi,ri,p,r});
  })));
  if(!items.length) return '';
  return `<div class="flagbar"><b>${items.length} rooms need a manual check</b>
    <div class="hint" style="margin:4px 0 0;color:#7f1d1d">Import did not change these. Occupied + arrival with no departure, vacant + departure, or a room move that does not fit. Tap the red V# / CAT / STAT and set the status yourself.</div>
    <div class="flaglist">${items.map(x=>`<button type="button" class="flagchip" data-slot="${x.g}.${x.pi}.${x.ri}">${esc(x.r.villa)} · ${esc(x.r.status||'-')} · ${esc(flagTitle(x.r))}</button>`).join('')}</div>
  </div>`;
}

function sheetBlock(title, group, list){
  const max=Math.max(...list.map(p=>p.rows.length),1);
  return `<div class="secbar"><h3 style="color:#1f4e79;margin:0">${title}</h3>
      <button class="add" style="width:auto;padding:8px 10px" data-addsec="${group}">+ New section</button></div>
    <div class="sheetwrap"><table class="sheetgrid">
      <tr>${list.map((p,i)=>{
        const sid=group+'.'+i;
        const emptySec=isSectionEmpty(p);
        return `<th class="nm dropzone ${selectedSec===sid?'sec-on':''} ${emptySec?'sec-empty':''}" data-drop="${sid}" data-sec="${sid}" colspan="3" style="background:${emptySec?'#e8e0d6':p.color}"><span class="namepick"><input class="allocname" data-path="${group}.${i}.name" value="${esc(p.name)}" autocomplete="off" placeholder="${emptySec?'Empty section':''}" style="color:${emptySec?'#6b5a4a':'#222'};font-weight:800"><button type="button" class="namebtn light" data-dd="name" data-nlist="villa" data-path="${group}.${i}.name">▾</button></span></th>`;
      }).join('')}</tr>
      <tr>${list.map(()=>'<th>V#</th><th>CAT</th><th>STAT</th>').join('')}</tr>
      ${Array.from({length:max},(_,r)=>'<tr>'+list.map((p,i)=>{
        const row=p.rows[r];
        if(!row) return '<td></td><td></td><td></td>';
        const id=group+'.'+i+'.'+r;
        const empty=!String(row.villa||'').trim();
        const cls=slotClass(group,i,r,row);
        return `<td class="${cls}" data-slot="${id}" title="${esc(flagTitle(row))}"><input data-path="${pathOf(group,i,r,'villa')}" value="${esc(row.villa)}" placeholder="${empty?'+':''}"></td>
                <td class="${cls}" data-slot="${id}" title="${esc(flagTitle(row))}">${dd(pathOf(group,i,r,'cat'),row.cat,'cat')}</td>
                <td class="${cls}" data-slot="${id}" title="${esc(flagTitle(row))}">${dd(pathOf(group,i,r,'status'),row.status,'status')}</td>`;
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
function trafficVisible(list, hasVal){
  const filled=[], empty=[];
  (list||[]).forEach((x,i)=> (hasVal(x)?filled:empty).push(i));
  return [...filled, ...empty.slice(0,3)];
}
function trafficCols(key){
  const spec=TRAFFIC[key];
  return spec? spec.cols[1]-spec.cols[0]+1 : 1;
}
function trafficGrid(key, cls){
  const list=model[key]||[];
  const cols=trafficCols(key);
  const rows=Math.max(1, Math.ceil(list.length/cols));
  let body='';
  for(let r=0;r<rows;r++){
    body+='<tr>';
    for(let c=0;c<cols;c++){
      const i=r*cols+c;
      const x=list[i];
      body += x? `<td><input data-path="${key}.${i}.v" value="${esc(x.v)}" inputmode="numeric" enterkeyhint="next"></td>` : '<td></td>';
    }
    body+='</tr>';
  }
  const n=(list||[]).filter(x=>String(x.v||'').trim()).length;
  const titles={arrivals:'ARRIVAL',departures:'DEPATURE',honeymoon:'HONEYMOON',birthday:'BIRTHDAY',anniversary:'ANNIVERSARY',upon:'UPON ARRIVAL',vip:'VIP ARRIVAL'};
  return `<td class="tblock ${cls||''}">
    <div class="sec ttitle">${titles[key]||key.toUpperCase()} <small>${n}</small></div>
    <table class="tgrid">${body}</table>
  </td>`;
}
function moveGrid(){
  const moves=model.moves||[];
  const pairs=3, rows=10;
  let body='<tr>'+Array.from({length:pairs},()=>'<th>FROM</th><th>TO</th>').join('')+'</tr>';
  for(let r=0;r<rows;r++){
    body+='<tr>';
    for(let p=0;p<pairs;p++){
      const i=p*rows+r;
      const x=moves[i]||{from:'',to:''};
      body += `<td><input data-path="moves.${i}.from" value="${esc(x.from)}" inputmode="numeric" enterkeyhint="next"></td><td><input data-path="moves.${i}.to" value="${esc(x.to)}" inputmode="numeric" enterkeyhint="next"></td>`;
    }
    body+='</tr>';
  }
  const mn=(model.moves||[]).filter(x=>String(x.from||'').trim()||String(x.to||'').trim()).length;
  return `<td class="tblock">
    <div class="sec ttitle">VILLA MOVE <small>${mn}</small></div>
    <table class="tgrid tmoves">${body}</table>
  </td>`;
}
function trafficBoard(){
  const n=function(key){ return (model[key]||[]).filter(x=>String(x.v||'').trim()).length; };
  const mn=(model.moves||[]).filter(x=>String(x.from||'').trim()||String(x.to||'').trim()).length;
  return `<table class="traffic">
      <tr>
        <td class="sec" colspan="1">VILLA MOVE <small>${mn}</small></td>
        <td class="sec">ARRIVAL <small>${n('arrivals')}</small></td>
        <td class="sec2">DEPATURE <small>${n('departures')}</small></td>
        <td class="sec">HONEYMOON <small>${n('honeymoon')}</small></td>
        <td class="sec">BIRTHDAY <small>${n('birthday')}</small></td>
        <td class="sec">ANNIVERSARY <small>${n('anniversary')}</small></td>
        <td class="sec">UPON ARRIVAL <small>${n('upon')}</small></td>
      </tr>
      <tr>
        ${moveGrid()}
        ${trafficGrid('arrivals','kpi-arr')}
        ${trafficGrid('departures','kpi-dep')}
        ${trafficGrid('honeymoon','')}
        ${trafficGrid('birthday','')}
        ${trafficGrid('anniversary','')}
        ${trafficGrid('upon','')}
      </tr>
    </table>
    <div class="tiny" style="margin:4px 0 6px">Same layout as the orange sheet: moves are 3 FROM/TO blocks, arrival and departure are 7×11, honeymoon / birthday / anniversary / upon arrival are 3×11. Type here or use From text.</div>`;
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
      <tr>${list.map((p,i)=>{
        const sid=group+'.'+i;
        const emptySec=isSectionEmpty(p);
        return `<td class="dropzone sechead ${selectedSec===sid?'sec-on':''} ${emptySec?'sec-empty':''}" data-drop="${sid}" data-sec="${sid}" colspan="3" style="background:${emptySec?'#e8e0d6':p.color};padding:5px 3px"><span class="namepick"><input class="allocname" data-path="${group}.${i}.name" value="${esc(p.name)}" autocomplete="off" placeholder="${emptySec?'Empty section — drop here':''}"><button type="button" class="namebtn light" data-dd="name" data-nlist="villa" data-path="${group}.${i}.name">▾</button></span></td>`;
      }).join('')}</tr>
      <tr>${list.map(()=>'<td class="sec">V#</td><td class="sec">CAT.</td><td class="sec">STAT.</td>').join('')}</tr>
      ${Array.from({length:max},(_,r)=>'<tr class="'+(selectedSec?(selectedSec.startsWith(group+'.')?'':''):'')+'">'+list.map((p,i)=>{
        const row=p.rows[r]; if(!row) return '<td></td><td></td><td></td>';
        const id=group+'.'+i+'.'+r;
        const empty=!String(row.villa||'').trim();
        const cls=slotClass(group,i,r,row)+(selectedSec===group+'.'+i?' sec-col':'');
        return `<td class="${cls}" data-slot="${id}" title="${esc(flagTitle(row))}"><div class="vcell"><input data-path="${pathOf(group,i,r,'villa')}" value="${esc(row.villa)}" placeholder="${empty?'+':''}" inputmode="numeric" enterkeyhint="next"></div></td><td class="${cls}" data-slot="${id}" title="${esc(flagTitle(row))}">${dd(pathOf(group,i,r,'cat'),row.cat,'cat')}</td><td class="${cls}" data-slot="${id}" title="${esc(flagTitle(row))}">${dd(pathOf(group,i,r,'status'),row.status,'status')}</td>`;
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
      <tr><td class="sec">FORECAST <button type="button" class="dd" data-dd="weather" data-path="forecast.weather" style="width:auto;background:#fff;font-size:18px;padding:2px 8px">${esc(model.forecast.weather||'💧')}</button></td><td class="sec">VIP ARRIVAL <small>${(model.vip||[]).filter(x=>String(x.v||'').trim()).length}</small></td><td class="sec2">Next day arrival guest preferences</td></tr>
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
        ${trafficGrid('vip','vipbox')}
        <td class="tiny">Preferences stay in the original sheet cells when present.</td>
      </tr>
    </table>
    ${trafficBoard()}
    <div class="sec" style="margin-top:6px;padding:5px">VILLA ATTENDANT ALLOCATION</div>
    <div class="tools">
      <b>Edit</b>
      <span>Double-tap a room to move it. Double-tap a section name to move the whole column onto an empty section.</span>
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
          <table><tr><td class="tiny">NAME</td><td class="tiny">TYPE</td><td class="tiny">START</td><td class="tiny">END</td><td class="tiny">DAYS</td></tr>
          ${(model.leave||[]).map((l,i)=>`<tr>
            <td>${nameInput('leave.'+i+'.name',l.name,'dl-villa')}</td>
            <td>${nameInput('leave.'+i+'.type',l.type,'leaveType')}</td>
            <td><input data-path="leave.${i}.start" type="date" value="${esc(isoDate(l.start))}"></td>
            <td><input data-path="leave.${i}.end" type="date" value="${esc(isoDate(l.end))}"></td>
            <td><b data-leavedays="${i}">${esc(leaveDays(l.start,l.end)||l.days||'')}</b></td>
          </tr>`).join('')}
          </table>
          <div id="leaveSummary" class="tiny" style="margin-top:6px;font-weight:800;color:#1f4e79">${(()=>{const s=leaveSummary();return 'On leave '+s.total+' · Annual '+s.AL+' · Off day '+s.OFF+' · Medical '+s.ML+' · Emergency '+s.EL;})()}</div>
        </td>
        <td>
          <div class="sec">COLOUR LEGEND & SHORT CODE</div>
          <table>${STATUSES.map(([c,n])=>`<tr><td class="legend-${c}">${n}</td><td class="legend-${c}"><b>${c}</b></td></tr>`).join('')}</table>
        </td>
        <td>
          <div class="sec">LAUNDRY</div>
          <table>${(model.laundry||[]).map((x,i)=>`<tr>
            <td>${esc(x.no)}</td>
            <td>${nameInput('laundry.'+i+'.name',x.name,'dl-laundry')}</td>
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
            <td>${nameInput('supervisors.'+i+'.name',x.name,'dl-supervisor')}</td>
            <td><input data-path="supervisors.${i}.role" value="${esc(x.role)}"></td>
            <td><input data-path="supervisors.${i}.status" value="${esc(x.status)}"></td>
            <td><input data-path="supervisors.${i}.section" value="${esc(x.section)}"></td>
          </tr>`).join('')}</table>
        </td>
        <td>
          <div class="sec">MINIBAR DUTY</div>
          <table>${(model.minibar||[]).map((x,i)=>`<tr>
            <td>${nameInput('minibar.'+i+'.name',x.name,'dl-minibar')}</td>
            <td><input data-path="minibar.${i}.status" value="${esc(x.status)}"></td>
            <td><input data-path="minibar.${i}.section" value="${esc(x.section)}"></td>
          </tr>`).join('')}</table>
          <div class="sec">OFFICE DUTY</div>
          <table>${(model.office||[]).map((x,i)=>`<tr>
            <td>${nameInput('office.'+i+'.name',x.name,'dl-office')}</td>
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
            <td>${nameInput('publicArea.'+i+'.name',x.name,'dl-public')}</td>
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
      <table><tr><th>Name</th><th>Type</th><th>Start</th><th>End</th><th>Days</th></tr>
      ${model.leave.map((l,i)=>`<tr>
        <td>${nameInput('leave.'+i+'.name',l.name,'dl-villa')}</td>
        <td>${nameInput('leave.'+i+'.type',l.type,'leaveType')}</td>
        <td><input data-path="leave.${i}.start" type="date" value="${esc(isoDate(l.start))}"></td>
        <td><input data-path="leave.${i}.end" type="date" value="${esc(isoDate(l.end))}"></td>
        <td data-leavedays="${i}">${esc(leaveDays(l.start,l.end)||l.days||'')}</td>
      </tr>`).join('')}
      </table>
    </div>`;
}
function grabSheetScroll(){
  const wrap=document.querySelector('.fullwrap')||document.querySelector('.sheetwrap');
  if(wrap) sheetScroll={x:wrap.scrollLeft,y:wrap.scrollTop,wx:window.scrollX||0,wy:window.scrollY||window.pageYOffset||0};
  return sheetScroll;
}
function putSheetScroll(){
  const s=sheetScroll||{x:0,y:0,wx:0,wy:0};
  const wrap=document.querySelector('.fullwrap')||document.querySelector('.sheetwrap');
  if(wrap){ wrap.scrollLeft=s.x||0; wrap.scrollTop=s.y||0; }
  if(s.wx||s.wy) window.scrollTo(s.wx||0,s.wy||0);
}
function render(opt){
  opt=opt||{};
  if(!model) return;
  if(!opt.resetScroll) grabSheetScroll();
  else sheetScroll={x:0,y:0,wx:0,wy:0};
  if(!model.forecast) model.forecast={weather:'☀️',auto:true};
  syncForecast();
  const m=metrics();
  viewBar();
  document.getElementById('page').innerHTML=`
    ${nameLists()}
    <datalist id="dl-leave-type"><option value="AL"></option><option value="OFF"></option><option value="EL"></option><option value="ML"></option><option value="TASK FORCE"></option></datalist>
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
    ${flagBanner()}
    ${viewBody()}
    ${viewMode==='full'?'':extrasBlock()}
  `;
  applyZoom();
  putSheetScroll();
  requestAnimationFrame(()=>{ putSheetScroll(); });
}
function applyZoom(){
  const z=Math.max(70, Math.min(160, sheetZoom||100));
  sheetZoom=z;
  localStorage.setItem('hk-zoom', String(z));
  const el=document.querySelector('.fullwrap')||document.getElementById('page');
  if(el && el.style.zoom!==String(z/100)){
    el.style.zoom=String(z/100);
    el.style.transform='';
  }
  const lab=document.getElementById('zoomLabel');
  if(lab) lab.textContent=z+'%';
}
function bumpZoom(delta){
  sheetZoom=(sheetZoom||100)+delta;
  applyZoom();
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
function addSection(group, opt){
  opt=opt||{};
  const target=opt.modelRef||model;
  if(!target[group]) target[group]=[];
  const used=target[group].map(p=>p.col);
  let col=opt.col;
  if(col==null){
    col=ALLOC_COLS.find(c=>!used.includes(c));
    if(col==null){
      if(!opt.silent) toast('The orange sheet has 10 section columns only');
      return;
    }
  }
  const headerRow=group==='beach'?32:48;
  const start=group==='beach'?36:52;
  const count=group==='beach'?12:13;
  const rows=[];
  for(let i=0;i<count;i++){
    const r=start+i;
    rows.push({villa:'',cat:'',status:'',flag:'',extra:false,cells:[addr(r,col),addr(r,col+1),addr(r,col+2)]});
  }
  target[group].push({
    name: opt.empty?'':(opt.name||'NEW SECTION'),
    nameCell:addr(headerRow,col),
    col, headerRow,
    color:COLORS[target[group].length%COLORS.length],
    extraStart:start+count,
    rows, added:true, landing:!!opt.empty
  });
  if(!opt.silent){
    render();
    persist(opt.empty?'Empty section added':'New allocation section added');
  }
}
function moveSection(src, dest){
  if(!src||!dest||src===dest) return;
  const [sg,si]=src.split('.');
  const [dg,di]=dest.split('.');
  const a=model[sg]&&model[sg][+si];
  const b=model[dg]&&model[dg][+di];
  if(!a||!b) return;
  const n=Math.max(a.rows.length,b.rows.length);
  while(a.rows.length<n) a.rows.push({villa:'',cat:'',status:'',flag:'',extra:true,cells:[]});
  while(b.rows.length<n) b.rows.push({villa:'',cat:'',status:'',flag:'',extra:true,cells:[]});
  const tmpName=a.name; a.name=b.name; b.name=tmpName;
  for(let i=0;i<n;i++){
    const ta={villa:a.rows[i].villa,cat:a.rows[i].cat,status:a.rows[i].status,flag:a.rows[i].flag};
    a.rows[i].villa=b.rows[i].villa; a.rows[i].cat=b.rows[i].cat; a.rows[i].status=b.rows[i].status; a.rows[i].flag=b.rows[i].flag;
    b.rows[i].villa=ta.villa; b.rows[i].cat=ta.cat; b.rows[i].status=ta.status; b.rows[i].flag=ta.flag;
  }
  packSection(a); packSection(b);
  selectedSec='';
  syncForecast();
  render();
  persist((b.name||'Section')+' moved');
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
  packSection(model[sg][+spi]);
  packSection(model[dg][+dpi]);
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
  packSection(model[sg][+spi]);
  if(!(sg===dg && +spi===+dpi)) packSection(model[dg][+dpi]);
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

function openPicker(kind,path,cur,nlist){
  pick={kind,path,nlist};
  document.getElementById('overlay').classList.add('show');
  const p=document.getElementById('picker');
  p.classList.add('show');
  if(kind==='name'){
    seedNames(model);
    seedLeaveTypes(model);
    const key=nlist||'villa';
    const names=key==='leaveType'?(model.leaveTypes||[]):((model.names&&model.names[key])||[]);
    const title=({villa:'Villa attendant',laundry:'Laundry',public:'Public area',supervisor:'Supervisor',minibar:'Minibar',office:'Office',leaveType:'Leave type'}[key]||'Name');
    p.innerHTML='<h3>'+title+'</h3>'+
      (names.length?names.map(n=>`<div class="opt ${n===cur?'on':''}" data-val="${esc(n)}"><b>${esc(n)}</b></div>`).join(''):'<div class="hint" style="padding:8px">No saved items yet. Add them in the menu, or type in the box.</div>')+
      `<div class="opt" data-val=""><b>Clear</b></div>
       <div class="hint" style="padding:8px 4px 0">Or type in the field — it stays editable.</div>`;
    return;
  }
  const list=kind==='weather'?WEATHER:kind==='status'?STATUSES:CATS;
  const title=kind==='weather'?'Weather':kind==='status'?'Status':'Category';
  p.innerHTML='<h3>'+title+'</h3>'+
    list.map(([c,n])=>`<div class="opt ${c===cur?'on':''}" data-val="${c}"><b style="font-size:${kind==='weather'?'22px':'13px'}">${c}</b><small>${n}</small></div>`).join('')+
    (kind==='weather'?'':`<div class="opt" data-val=""><b>Clear</b></div>`);
}
function closePicker(){document.getElementById('overlay').classList.remove('show');document.getElementById('picker').classList.remove('show');pick=null;}

function applyField(el){
  if(!el) return false;
  if(el.id==='masterQ'){ masterQ=el.value; render(); const n=document.getElementById('masterQ'); if(n){n.focus(); try{n.setSelectionRange(n.value.length,n.value.length);}catch(err){} } return true; }
  if(el.id==='autoCount'){
    model.forecast.auto=!!el.checked;
    if(model.forecast.auto){ syncForecast(); render(); }
    persist(false);
    return true;
  }
  const pathEl=el.closest('[data-path]');
  if(!pathEl) return false;
  setPath(pathEl.dataset.path, pathEl.type==='checkbox'?pathEl.checked:pathEl.value);
  if(/^forecast\.(arrival|departure|occupied|vacant)$/.test(pathEl.dataset.path)) model.forecast.auto=false;
  if(/\.rows\.\d+\.(villa|status|cat)$/.test(pathEl.dataset.path)){
    const parts=pathEl.dataset.path.split('.');
    const row=model[parts[0]][+parts[1]].rows[+parts[3]];
    if(row && row.flag && parts[4]!=='villa'){ row.flag=''; }
  }
  scheduleSave();
  return true;
}
document.addEventListener('scroll',e=>{
  const el=e.target;
  if(el&&el.classList&&(el.classList.contains('fullwrap')||el.classList.contains('sheetwrap'))){
    sheetScroll={x:el.scrollLeft,y:el.scrollTop,wx:window.scrollX||0,wy:window.scrollY||window.pageYOffset||0};
  }
}, true);
document.body.addEventListener('input',e=>{
  applyField(e.target);
});
document.body.addEventListener('change',e=>{
  if(e.target && (e.target.tagName==='SELECT' || e.target.type==='checkbox' || e.target.type==='date')) applyField(e.target);
});
document.body.addEventListener('click',e=>{
  const vw=e.target.closest('[data-view]');
  if(vw){
    viewMode=vw.dataset.view;
    localStorage.setItem('hk-daily-view', viewMode);
    render({resetScroll:true});
    return;
  }
  if(e.target.closest('[data-textimport]')){ openTextModal(); return; }
  if(e.target.closest('[data-undo]')){ doUndo(); return; }
  if(e.target.closest('[data-redo]')){ doRedo(); return; }
  if(e.target.id==='txCancel'){ closeTextModal(); return; }
  if(e.target.id==='txApply'){ applyTextImport(); return; }
  const pmode=e.target.closest('[data-pmode]');
  if(pmode){ pickMode=pmode.dataset.pmode; pickFrom=''; renderPickModal(); return; }
  const pun=e.target.closest('[data-punpick]');
  if(pun){ dropPickRoom(pun.dataset.punpick); renderPickModal(); return; }
  const prm=e.target.closest('[data-prm]');
  if(prm){ pickBags.moves.splice(+prm.dataset.prm,1); renderPickModal(); return; }
  const proom=e.target.closest('[data-proom]');
  if(proom){ togglePickRoom(proom.dataset.proom); return; }
  if(e.target.id==='pkCancel' || e.target.id==='pkClose'){ closePickModal(); return; }
  if(e.target.id==='pkApply'){ applyPickImport(); return; }
  if(e.target.id==='pkClear'){ pickBags=emptyPickBags(); pickFrom=''; renderPickModal(); return; }
  const secEl=e.target.closest('[data-sec]');
  if(secEl && !e.target.closest('input,button,select')){
    const sid=secEl.dataset.sec;
    if(selectedSec && selectedSec!==sid){ pushUndo(); moveSection(selectedSec, sid); return; }
    const now=Date.now();
    if(window.__secTap && window.__secTap.id===sid && now-window.__secTap.t<450){
      selectedSrc='';
      selectedSec = selectedSec===sid?'':sid;
      window.__ignoreDbl=true;
      render();
      toast(selectedSec?'Section selected — tap an empty section to place it':'Section cleared');
      window.__secTap=null;
      return;
    }
    window.__secTap={id:sid,t:now};
    return;
  }
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
  if(addsec){ addSection(addsec.dataset.addsec, {empty:true}); return; }
  const mv=e.target.closest('[data-move]');
  if(mv){ openMovePicker(mv.dataset.move); return; }
  if(e.target.closest('[data-addmove]')){
    const empty=(model.moves||[]).find(x=>!String(x.from||'').trim()&&!String(x.to||'').trim());
    if(empty){ empty.from=' '; empty.from=empty.from.trim(); }
    else toast('No empty villa-move cell left on the sheet');
    render(); persist(false); return;
  }
  const addchip=e.target.closest('[data-addchip]');
  if(addchip){
    const key=addchip.dataset.addchip;
    const empty=(model[key]||[]).find(x=>!String(x.v||'').trim());
    if(empty) empty.v=' ';
    else toast('No empty cell left on the orange sheet');
    if(empty) empty.v='';
    render(); persist(false); return;
  }
  const del=e.target.closest('[data-del]');
  if(del){
    const [g,pi,ri]=del.dataset.del.split('.');
    const row=model[g][+pi].rows[+ri];
    if(row.extra) model[g][+pi].rows.splice(+ri,1);
    else {row.villa='';row.cat='';row.status='';}
    syncForecast(); render(); persist(false); return;
  }
  const btn=e.target.closest('[data-dd]');
  if(btn){
    const wrap=btn.closest('.namepick');
    const typed=wrap&&wrap.querySelector('input')?wrap.querySelector('input').value:'';
    const cur=btn.dataset.dd==='name'?typed:(btn.textContent.trim()==='Select'?'':btn.textContent.trim());
    openPicker(btn.dataset.dd, btn.dataset.path, cur, btn.dataset.nlist);
    return;
  }
  const opt=e.target.closest('.opt');
  if(opt&&pick){
    if(pick.kind==='move' && opt.dataset.dest){ const src=pick.src; closePicker(); moveRoom(src, opt.dataset.dest); return; }
    setPath(pick.path,opt.dataset.val);
    if(pick.kind==='name') seedNames(model);
    if(String(pick.path).includes('.status')){
      const parts=pick.path.split('.');
      if(parts[2]==='rows'){
        const row=model[parts[0]][+parts[1]].rows[+parts[3]];
        if(row) row.flag='';
      }
      syncForecast();
    }
    closePicker();render();persist(false);return;
  }
  if(e.target.id==='overlay'){ closePicker(); closeTextModal(); if(typeof closePickModal==='function') closePickModal(); if(typeof closePeopleModal==='function') closePeopleModal(); if(typeof closeLeaveTypeModal==='function') closeLeaveTypeModal(); }
});
document.body.addEventListener('dblclick',e=>{
  if(window.__ignoreDbl){ window.__ignoreDbl=false; return; }
  const sec=e.target.closest('[data-sec]');
  if(sec){
    e.preventDefault();
    selectedSrc='';
    selectedSec = selectedSec===sec.dataset.sec?'':sec.dataset.sec;
    render();
    toast(selectedSec?'Section selected — tap an empty section to place it':'Section cleared');
    return;
  }
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
function patchCell(xml,cell,value,type,styleId){
  if(!cell) return xml;
  const re=new RegExp('<c\\b(?=[^>]*\\br="'+cell+'")[^>]*?(?:/>|>[\\s\\S]*?</c>)');
  const hit=xml.match(re);
  const style=styleId!=null&&styleId!==''?String(styleId):(hit&&((hit[0].match(/\bs="(\d+)"/)||[])[1])||null);
  const sAttr=(style!=null&&style!=='')?(' s="'+style+'"'):'';
  let neu;
  if(value===''||value==null) neu='<c r="'+cell+'"'+sAttr+'/>';
  else if(type==='n') neu='<c r="'+cell+'"'+sAttr+' t="n"><v>'+value+'</v></c>';
  else neu='<c r="'+cell+'"'+sAttr+' t="inlineStr"><is><t xml:space="preserve">'+xmlEsc(value)+'</t></is></c>';
  if(hit) return xml.replace(re,neu);
  const rowNum=parseInt(cell.replace(/^[A-Z]+/,''),10);
  const rowRe=new RegExp('(<row r="'+rowNum+'"[^>]*>)([\\s\\S]*?)(</row>)');
  if(rowRe.test(xml)) return xml.replace(rowRe,(_,a,mid,b)=>a+mid+neu+b);
  const sheetData=xml.indexOf('<sheetData>');
  if(sheetData>=0){
    return xml.replace('<sheetData>','<sheetData><row r="'+rowNum+'">'+neu+'</row>');
  }
  return xml;
}
function collectWrites(styleMap){
  const out=[];
  const add=(cell,value,type='s',style)=>{if(cell) out.push({cell,value,type,style});};
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
      const sid=styleMap&&r.status?styleMap[r.status]:undefined;
      add(r.cells[0], r.villa, /^\d+$/.test(String(r.villa))?'n':'s', sid);
      add(r.cells[1], r.cat, 's', sid);
      add(r.cells[2], r.status, 's', sid);
    });
  });
  model.moves.forEach(m=>{add(m.fromCell,m.from,/^\d+$/.test(m.from)?'n':'s');add(m.toCell,m.to,/^\d+$/.test(m.to)?'n':'s');});
  (model.leave||[]).forEach(l=>{
    if(!l.cells) return;
    add(l.cells.name,l.name);
    add(l.cells.type,l.type);
    add(l.cells.start,l.start?excelDate(isoDate(l.start)):'','n');
    add(l.cells.end,l.end?excelDate(isoDate(l.end)):'','n');
  });
  (model.laundry||[]).forEach(x=>{if(!x.cells)return;add(x.cells.name,x.name);add(x.cells.timing,x.timing);add(x.cells.desig,x.desig);add(x.cells.contact,x.contact);});
  (model.supervisors||[]).forEach(x=>{if(!x.cells)return;add(x.cells.name,x.name);add(x.cells.role,x.role);add(x.cells.status,x.status);add(x.cells.section,x.section);add(x.cells.contact,x.contact);});
  (model.minibar||[]).forEach(x=>{if(!x.cells)return;add(x.cells.name,x.name);add(x.cells.status,x.status);add(x.cells.section,x.section);});
  (model.office||[]).forEach(x=>{if(!x.cells)return;add(x.cells.name,x.name);add(x.cells.role,x.role);add(x.cells.status,x.status);add(x.cells.contact,x.contact);});
  (model.publicArea||[]).forEach(x=>{if(!x.cells)return;add(x.cells.name,x.name);add(x.cells.role,x.role);add(x.cells.section,x.section);});
  (model.tasks||[]).forEach(x=>{if(!x.cells)return;add(x.cells.detail,x.detail);add(x.cells.area,x.area);add(x.cells.time,x.time);});
  ['arrivals','departures','honeymoon','birthday','anniversary','upon','vip'].forEach(k=> (model[k]||[]).forEach(x=>add(x.cell,x.v,/^\d+$/.test(String(x.v))?'n':'s')));
  return out;
}
function ensureStatusStyles(stylesXml){
  const map={};
  if(!stylesXml) return {xml:stylesXml, map};
  const fillBlock=stylesXml.match(/<fills([^>]*)>([\s\S]*?)<\/fills>/);
  if(!fillBlock) return {xml:stylesXml, map};
  let fillsInner=fillBlock[2];
  const fillEls=[...fillsInner.matchAll(/<fill\b[\s\S]*?<\/fill>|<fill\/>/g)].map(m=>m[0]);
  const rgbToId={};
  fillEls.forEach((f,i)=>{
    const rgb=(f.match(/rgb="([A-Fa-f0-9]+)"/)||[])[1];
    if(rgb) rgbToId[rgb.toUpperCase().replace(/^FF/,'')]=i;
  });
  let fillCount=fillEls.length;
  Object.entries(STATUS_RGB).forEach(([code,rgb])=>{
    const key=rgb.toUpperCase();
    if(rgbToId[key]==null){
      fillsInner += '<fill><patternFill patternType="solid"><fgColor rgb="FF'+key+'"/><bgColor rgb="FF'+key+'"/></patternFill></fill>';
      rgbToId[key]=fillCount++;
    }
  });
  stylesXml=stylesXml.replace(/<fills([^>]*)>([\s\S]*?)<\/fills>/, '<fills count="'+fillCount+'">'+fillsInner+'</fills>');
  const xfBlock=stylesXml.match(/<cellXfs([^>]*)>([\s\S]*?)<\/cellXfs>/);
  if(!xfBlock) return {xml:stylesXml, map};
  let xfsInner=xfBlock[2];
  const xfs=[...xfsInner.matchAll(/<xf\b[^>]*\/>|<xf\b[\s\S]*?<\/xf>/g)].map(m=>m[0]);
  const fillToXf={};
  xfs.forEach((xf,i)=>{
    const fid=(xf.match(/fillId="(\d+)"/)||[])[1];
    if(fid!=null) fillToXf[fid]=i;
  });
  let xfCount=xfs.length;
  Object.entries(STATUS_RGB).forEach(([code,rgb])=>{
    const fid=rgbToId[rgb.toUpperCase()];
    if(fid==null) return;
    if(fillToXf[String(fid)]==null){
      xfsInner += '<xf numFmtId="0" fontId="0" fillId="'+fid+'" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>';
      fillToXf[String(fid)]=xfCount++;
    }
    map[code]=fillToXf[String(fid)];
  });
  stylesXml=stylesXml.replace(/<cellXfs([^>]*)>([\s\S]*?)<\/cellXfs>/, '<cellXfs count="'+xfCount+'">'+xfsInner+'</cellXfs>');
  return {xml:stylesXml, map};
}
let saveDirHandle=null;
let saveMode=localStorage.getItem('hk-save-mode')||'download';
async function loadSaveFolder(){
  try{
    const pack=await idbGet('saveFolder');
    if(pack && pack.handle) saveDirHandle=pack.handle;
    if(pack && pack.mode) saveMode=pack.mode;
    else saveMode=localStorage.getItem('hk-save-mode')||saveMode||'download';
  }catch(err){}
}
async function rememberSaveFolder(handle, mode){
  if(handle) saveDirHandle=handle;
  if(mode) saveMode=mode;
  localStorage.setItem('hk-save-mode', saveMode);
  try{ await idbSet('saveFolder', {handle:saveDirHandle||null, mode:saveMode, name:saveDirHandle?saveDirHandle.name:''}); }catch(err){}
  updateFileLabel();
}
async function ensureFolderPerm(){
  if(!saveDirHandle || typeof saveDirHandle.queryPermission!=='function') return false;
  const opts={mode:'readwrite'};
  try{
    let q=await saveDirHandle.queryPermission(opts);
    if(q!=='granted') q=await saveDirHandle.requestPermission(opts);
    return q==='granted';
  }catch(err){ return false; }
}
async function chooseSaveFolder(){
  if(!window.showDirectoryPicker){
    toast('This browser cannot keep a folder. Use Download for the normal Downloads folder.');
    return;
  }
  try{
    const handle=await window.showDirectoryPicker({id:'hk-daily-xlsm', mode:'readwrite', startIn:'documents'});
    await rememberSaveFolder(handle, 'folder');
    toast('Folder remembered: '+handle.name+'. Download will save there.');
  }catch(err){
    if(err && err.name==='AbortError') return;
    toast('Folder not selected');
  }
}
async function useNormalDownload(){
  await rememberSaveFolder(saveDirHandle, 'download');
  toast('Normal download — files go to the Downloads folder');
}
function xlsmStamp(){
  const d=new Date();
  const p=n=>String(n).padStart(2,'0');
  return p(d.getDate())+'-'+p(d.getMonth()+1)+'-'+d.getFullYear()+'_'+p(d.getHours())+p(d.getMinutes())+p(d.getSeconds());
}
async function writeToFolder(blob, name){
  const ok=await ensureFolderPerm();
  if(!ok) throw new Error('Folder permission needed');
  async function put(n){
    const fh=await saveDirHandle.getFileHandle(n, {create:true});
    const w=await fh.createWritable({keepExistingData:false});
    await w.write(blob);
    await w.close();
    return n;
  }
  try{
    return await put(name);
  }catch(err){
    const alt=String(name||'HK Daily').replace(/\.xlsm$/i,'')+'-'+xlsmStamp()+'.xlsm';
    return await put(alt);
  }
}
async function buildUpdatedXlsm(){
  const base=sourceBuf||originalBuf;
  if(!base||!model||typeof JSZip==='undefined') return null;
  const zip=await JSZip.loadAsync(base);
  const sheet=zip.file('xl/worksheets/sheet1.xml');
  if(!sheet) return null;
  let xml=await sheet.async('string');
  let styleMap={};
  const stylesFile=zip.file('xl/styles.xml');
  if(stylesFile){
    const packed=ensureStatusStyles(await stylesFile.async('string'));
    zip.file('xl/styles.xml', packed.xml);
    styleMap=packed.map;
  }
  collectWrites(styleMap).forEach(w=>{xml=patchCell(xml,w.cell,w.value,w.type,w.style);});
  zip.file('xl/worksheets/sheet1.xml',xml);
  const wbFile=zip.file('xl/workbook.xml');
  if(wbFile){
    let wbxml=await wbFile.async('string');
    if(/<calcPr[\s>]/.test(wbxml)){
      if(!/fullCalcOnLoad=/.test(wbxml)) wbxml=wbxml.replace('<calcPr','<calcPr fullCalcOnLoad="1"');
    } else {
      wbxml=wbxml.replace('</workbook>','<calcPr fullCalcOnLoad="1" calcMode="auto"/></workbook>');
    }
    zip.file('xl/workbook.xml',wbxml);
  }
  const buf=await zip.generateAsync({type:'arraybuffer'});
  const blob=new Blob([buf],{type:'application/vnd.ms-excel.sheet.macroEnabled.12'});
  const name=(fileName&&/\.xlsm$/i.test(fileName))?fileName:((fileName||'HK Daily')+'.xlsm');
  return {buf,blob,name};
}
async function downloadXlsm(){
  if(!originalBuf||!model){toast('Import XLSM first');return;}
  try{
    const packed=await buildUpdatedXlsm();
    if(!packed){ toast('This workbook sheet could not be written'); return; }
    const blob=packed.blob;
    const name=packed.name;
    if(saveMode==='folder' && saveDirHandle){
      try{
        const savedAs=await writeToFolder(blob, name);
        await persist(false);
        toast('Saved '+savedAs+' in '+saveDirHandle.name);
        return;
      }catch(err){
        console.error(err);
        toast('Folder save failed — downloading instead');
      }
    }
    const file=new File([blob], name, {type:'application/vnd.ms-excel.sheet.macroEnabled.12'});
    if(navigator.share && navigator.canShare && navigator.canShare({files:[file]})){
      try{ await navigator.share({files:[file], title:name}); await persist(false); toast('XLSM shared and saved'); return; }
      catch(err){ if(err && err.name==='AbortError') return; }
    }
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url;
    a.download=name;
    a.rel='noopener';
    a.style.display='none';
    document.body.appendChild(a);
    a.click();
    setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 2500);
    await persist(false);
    toast('XLSM downloaded to Downloads');
  }catch(err){
    console.error(err);
    toast('Download failed: '+(err.message||err));
  }
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
    sourceBuf=originalBuf;
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
    sourceBuf=pack.buf;
    model=ensureModel(pack.model);
    undoStack=Array.isArray(pack.undoStack)?pack.undoStack:[];
    redoStack=Array.isArray(pack.redoStack)?pack.redoStack:[];
    if(model.forecast&&model.forecast.auto!==false) syncForecast();
    applyImported();
    toast('Opened last saved file');
    return true;
  }catch(e){ return false; }
}
function takeFile(el){
  const f=el&&el.files&&el.files[0];
  if(el) el.value='';
  return importFile(f);
}
document.getElementById('fileIn').onchange=e=>takeFile(e.target);
document.getElementById('fileIn2').onchange=e=>takeFile(e.target);
document.getElementById('btnOut').onclick=downloadXlsm;
document.getElementById('btnSave').onclick=()=>persist('Saved on this phone');
document.getElementById('btnRoll').onclick=()=>rollYesterdayToToday();
document.getElementById('btnText').onclick=()=>openTextModal();
if(document.getElementById('btnPick')) document.getElementById('btnPick').onclick=()=>openPickModal();
function liveFormatArea(el){
  if(!el || el.tagName!=='TEXTAREA') return;
  const atEnd=el.selectionStart===el.value.length;
  if(el.id==='txMove'){
    const next=formatMoveText(el.value);
    if(next && atEnd) el.value=next;
    return;
  }
  if(/^tx(Arr|Dep|Occ|Vac|Honey|Birth|Ann|Upon|Vip)$/.test(el.id)){
    const next=formatRoomText(el.value);
    if(next && atEnd) el.value=next;
  }
}
document.getElementById('textModal').addEventListener('blur', e=>liveFormatArea(e.target), true);
document.getElementById('textModal').addEventListener('input', e=>liveFormatArea(e.target));
document.getElementById('btnUndo').onclick=()=>doUndo();
document.getElementById('btnUndo').onclick=()=>doUndo();
document.getElementById('btnRedo').onclick=()=>doRedo();
document.getElementById('btnOpenSaved').onclick=()=>restoreSaved();
document.getElementById('btnClear').onclick=async()=>{
  if(!confirm('Clear saved file and start again? Edits on this phone will be removed. The original Excel on your phone is not deleted.')) return;
  try{ await idbSet('pack', null); }catch(e){}
  localStorage.removeItem('hk-daily-meta');
  originalBuf=null; sourceBuf=null; model=null; fileName=''; savedAt='';
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
  await loadSaveFolder();
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
  const standalone = window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone;
  const isiOS = /iphone|ipad|ipod/i.test(navigator.userAgent);
  if(isiOS && !standalone){
    const b=document.getElementById('btnInstall');
    if(b) b.hidden=false;
  }
  document.addEventListener('click', function(e){
    if(e.target && e.target.id==='btnInstall'){
      if(deferred){ deferred.prompt(); deferred.userChoice.finally(function(){ deferred=null; const b=document.getElementById('btnInstall'); if(b) b.hidden=true; }); }
      else if(typeof toast==='function') toast(isiOS ? 'Share → Add to Home Screen' : 'Use the browser menu → Install / Add to Home Screen');
    }
  });
})();

function openLeaveTypeModal(){
  if(!model){ toast('Import XLSM first'); return; }
  seedLeaveTypes(model);
  renderLeaveTypes();
  document.getElementById('overlay').classList.add('show');
  document.getElementById('leaveTypeModal').classList.add('show');
}
function closeLeaveTypeModal(){
  const m=document.getElementById('leaveTypeModal');
  if(m) m.classList.remove('show');
  if(!document.getElementById('picker').classList.contains('show') && !document.getElementById('textModal').classList.contains('show') && !document.getElementById('peopleModal').classList.contains('show') && !document.getElementById('pickModal').classList.contains('show'))
    document.getElementById('overlay').classList.remove('show');
  render();
}
function renderLeaveTypes(){
  seedLeaveTypes(model);
  const box=document.getElementById('leaveTypeBody');
  if(!box) return;
  const list=model.leaveTypes||[];
  box.innerHTML=`<div class="plist"><h4>Leave types</h4>
    ${list.map((n,i)=>`<div class="prow"><input data-ltype="${i}" value="${esc(n)}"><button data-ltdel="${i}">Remove</button></div>`).join('')}
    <div class="prow"><input id="ltadd" placeholder="Add type e.g. NP"><button data-ltadd="1">Add</button></div>
    <p class="hint">Used in the leave table dropdown. You can still type a new type in the sheet.</p>
  </div>`;
}
function toggleMenu(on){
  const m=document.getElementById('appMenu');
  if(!m) return;
  if(on==null) m.hidden=!m.hidden;
  else m.hidden=!on;
}
function openPeopleModal(){
  if(!model){ toast('Import XLSM first'); return; }
  seedNames(model);
  renderPeople();
  document.getElementById('overlay').classList.add('show');
  document.getElementById('peopleModal').classList.add('show');
}
function closePeopleModal(){
  document.getElementById('peopleModal').classList.remove('show');
  if(!document.getElementById('picker').classList.contains('show') && !document.getElementById('textModal').classList.contains('show'))
    document.getElementById('overlay').classList.remove('show');
  render();
}
function renderPeople(){
  seedNames(model);
  const box=document.getElementById('peopleBody');
  if(!box) return;
  box.innerHTML=NAME_GROUPS.map(([key,title])=>{
    const list=model.names[key]||[];
    return `<div class="plist"><h4>${title}</h4>
      ${list.map((n,i)=>`<div class="prow"><input data-pname="${key}.${i}" value="${esc(n)}"><button data-pdel="${key}.${i}">Remove</button></div>`).join('')}
      <div class="prow"><input id="padd-${key}" placeholder="Add name"><button data-padd="${key}">Add</button></div>
    </div>`;
  }).join('');
}
document.getElementById('btnMenu').onclick=()=>toggleMenu();
document.getElementById('btnZoomIn').onclick=()=>bumpZoom(10);
document.getElementById('btnZoomOut').onclick=()=>bumpZoom(-10);
document.getElementById('peopleClose').onclick=()=>closePeopleModal();
if(document.getElementById('leaveTypeClose')) document.getElementById('leaveTypeClose').onclick=()=>closeLeaveTypeModal();
document.getElementById('leaveTypeModal').addEventListener('click',e=>{
  if(e.target.closest('[data-ltadd]')){
    const inp=document.getElementById('ltadd');
    const n=(inp&&inp.value||'').trim();
    if(!n) return;
    seedLeaveTypes(model);
    if(!model.leaveTypes.some(x=>x.toLowerCase()===n.toLowerCase())) model.leaveTypes.push(n);
    persist(false); renderLeaveTypes(); return;
  }
  const del=e.target.closest('[data-ltdel]');
  if(del){
    model.leaveTypes.splice(+del.dataset.ltdel,1);
    persist(false); renderLeaveTypes();
  }
});
document.getElementById('leaveTypeModal').addEventListener('change',e=>{
  const el=e.target.closest('[data-ltype]');
  if(!el) return;
  const i=+el.dataset.ltype;
  if(model.leaveTypes && model.leaveTypes[i]!=null) model.leaveTypes[i]=el.value.trim();
  persist(false);
});
document.getElementById('appMenu').addEventListener('click',e=>{
  const b=e.target.closest('[data-menu]');
  if(!b) return;
  toggleMenu(false);
  const act=b.dataset.menu;
  if(act==='people') openPeopleModal();
  else if(act==='leavetypes') openLeaveTypeModal();
  else if(act==='text') openTextModal();
  else if(act==='pick') openPickModal();
  else if(act==='roll') rollYesterdayToToday();
  else if(act==='save') persist('Saved on this phone');
  else if(act==='out') downloadXlsm();
  else if(act==='folder') chooseSaveFolder();
  else if(act==='downloads') useNormalDownload();
  else if(act==='zoomin') bumpZoom(10);
  else if(act==='zoomout') bumpZoom(-10);
  else if(act==='install') document.getElementById('btnInstall').click();
});
document.getElementById('peopleModal').addEventListener('click',e=>{
  const add=e.target.closest('[data-padd]');
  if(add){
    const key=add.dataset.padd;
    const inp=document.getElementById('padd-'+key);
    const n=(inp&&inp.value||'').trim();
    if(!n) return;
    if(!model.names[key]) model.names[key]=[];
    if(!model.names[key].some(x=>x.toLowerCase()===n.toLowerCase())) model.names[key].push(n);
    persist(false); renderPeople(); return;
  }
  const del=e.target.closest('[data-pdel]');
  if(del){
    const [key,i]=del.dataset.pdel.split('.');
    (model.names[key]||[]).splice(+i,1);
    persist(false); renderPeople();
  }
});
document.getElementById('peopleModal').addEventListener('change',e=>{
  const el=e.target.closest('[data-pname]');
  if(!el) return;
  const [key,i]=el.dataset.pname.split('.');
  if(model.names[key] && model.names[key][+i]!=null) model.names[key][+i]=el.value.trim();
  persist(false);
});
document.addEventListener("visibilitychange",()=>{ if(document.visibilityState==="hidden") persist(false); });
window.addEventListener("pagehide",()=>persist(false));
