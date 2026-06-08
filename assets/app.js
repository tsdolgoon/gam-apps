/* ============================================================
   GAM Back Office — application logic
   Golomt Asset Management · subscription / redemption / NAV
   Data layer: MSAL + Microsoft Graph + SheetJS, backend = one
   shared .xlsx on the Golomt Asset Management SharePoint site.
   ============================================================ */
'use strict';

/* ---------- SharePoint configuration ---------- */
const SHAREPOINT_SITE = "https://golomtasset.sharepoint.com";
const SHARED_FILE_PATH = "/Shared Documents/GAM Back Office/GAM_BackOffice.xlsx";

/* ---------- Schema: sheet -> ordered columns ---------- */
const SCHEMA = {
  Funds: ['FundID','NameMN','NameEN','ShortName','Type','Currency','RegisteredDate','TermYears',
          'TargetRaise','AuthorizedUnits','NominalPrice','MgmtFeePct','PerfFeePct',
          'Custodian','ManagementCo','Status','Color','RoutesTo'],
  Investors: ['InvestorID','RegDate','InvestorType','NameMN','NameEN','RegNo','ContactPerson',
              'Phone','Email','Address','BankName','BankAccount','AccountName','Status','Notes'],
  Transactions: ['TxnID','TradeDate','Type','FundID','InvestorID','Units','UnitPrice','GrossAmount',
                 'FeePct','FeeAmount','NetAmount','SettlementDate','Status','PaymentRef','Notes','CreatedAt'],
  NAVHistory: ['NavID','Date','FundID','TotalNAV','UnitsOutstanding','UnitPrice','Source','Notes'],
  Fees: ['FeeID','CalcDate','FundID','Period','OpeningNAV','ClosingNAV','MgmtFeePct','MgmtFee',
         'TargetReturnPct','TargetGrowth','NavGrowth','ExcessGrowth','PerfFeePct','PerfFee','TotalFee','Status','Notes'],
  ECIFFirms: ['FirmID','JoinDate','NameMN','NameEN','RegNo','EmployeeCount','ContactPerson','Phone','Email',
              'Address','BankName','BankAccount','InvestorID','PortfolioVariant','Status','Notes'],
  ECIFEmployees: ['EmployeeID','FirmID','RegNo','NameMN','Position','JoinDate','MonthlySalary','Status','Notes'],
  ECIFContributions: ['ContribID','Date','Period','FirmID','EmployeeID','EmployeeAmount','EmployerAmount',
                      'TotalAmount','UnitPrice','Units','Type','TxnID','Notes'],
  Meta: ['Key','Value'],
};

const SCHEMA_VERSION = '1.0';

/* ---------- Seed data: the two funds from GAM - Fund info.xlsx ---------- */
const SEED_FUNDS = [
  {
    FundID:'SGF', NameMN:'«Тогтмол Орлого» Хувийн ХОС', NameEN:'Stable Growth Private Fund',
    ShortName:'Тогтмол Орлого', Type:'Хувийн хөрөнгө оруулалтын сан', Currency:'MNT',
    RegisteredDate:'2023-12-28', TermYears:10, TargetRaise:100000000000, AuthorizedUnits:1000000000,
    NominalPrice:100, MgmtFeePct:1.5, PerfFeePct:20,
    Custodian:'Голомт банк — Кастодианы үйлчилгээний газар', ManagementCo:'«Голомт Ассет Менежмент ҮЦК» ХХК',
    Status:'Идэвхтэй', Color:'#1366c4'
  },
  {
    FundID:'USF', NameMN:'«Урбан Скай» Хувийн ХОС', NameEN:'Urban Sky Private Fund',
    ShortName:'Урбан Скай', Type:'Хувийн (хаалттай) хөрөнгө оруулалтын сан', Currency:'MNT',
    RegisteredDate:'2024-07-25', TermYears:10, TargetRaise:150000000000, AuthorizedUnits:150000,
    NominalPrice:1000000, MgmtFeePct:2.5, PerfFeePct:20,
    Custodian:'Голомт банк — Кастодианы үйлчилгээний газар', ManagementCo:'«Голомт Ассет Менежмент ҮЦК» ХХК',
    Status:'Идэвхтэй', Color:'#c8a14b'
  },
  {
    FundID:'ECIF', NameMN:'«Ажилтны хуримтлалын сан» Хувийн ХОС', NameEN:'Employee Contribution Investment Fund',
    ShortName:'ECIF', Type:'Ажилчдын хуримтлалын хөтөлбөр — Тогтмол Орлого сангаар дамжина', Currency:'MNT',
    RegisteredDate:'2023-12-28', TermYears:10, TargetRaise:100000000000, AuthorizedUnits:1000000000,
    NominalPrice:100, MgmtFeePct:1.5, PerfFeePct:20,
    Custodian:'Голомт банк — Кастодианы үйлчилгээний газар', ManagementCo:'«Голомт Ассет Менежмент ҮЦК» ХХК',
    Status:'Идэвхтэй', Color:'#1c8a4a', RoutesTo:'SGF'
  },
];

const TXN_STATUS = ['Хүлээгдэж буй','Баталгаажсан','Цуцлагдсан']; // Pending / Confirmed / Cancelled
const INV_TYPES  = ['Хувь хүн','Хуулийн этгээд'];                 // Individual / Company
const INV_STATUS = ['Идэвхтэй','Идэвхгүй'];

/* ============================================================
   State
   ============================================================ */
const state = {
  account: null,
  data: { Funds:[], Investors:[], Transactions:[], NAVHistory:[], Fees:[], ECIFFirms:[], ECIFEmployees:[], ECIFContributions:[], Meta:[] },
  dirty: false,
  view: 'dashboard',
  savedAt: null,
};

/* ============================================================
   Tiny helpers
   ============================================================ */
const $  = (s,r=document)=>r.querySelector(s);
const $$ = (s,r=document)=>[...r.querySelectorAll(s)];
const el = (tag,attrs={},...kids)=>{
  const n=document.createElement(tag);
  for(const[k,v] of Object.entries(attrs)){
    if(k==='class') n.className=v;
    else if(k==='html') n.innerHTML=v;
    else if(k.startsWith('on')&&typeof v==='function') n.addEventListener(k.slice(2),v);
    else if(v!=null) n.setAttribute(k,v);
  }
  for(const kid of kids.flat()){ if(kid==null||kid===false)continue; n.append(kid.nodeType?kid:document.createTextNode(kid)); }
  return n;
};
const esc = s => String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const todayISO = ()=> new Date().toISOString().slice(0,10);
const num = v => { const n=parseFloat(String(v??'').replace(/,/g,'')); return isFinite(n)?n:0; };
const fmtMoney = (v,dec=2)=> num(v).toLocaleString('en-US',{minimumFractionDigits:dec,maximumFractionDigits:dec});
const fmtUnits = v => { const n=num(v); return n.toLocaleString('en-US',{maximumFractionDigits:4}); };
const fmtInt   = v => num(v).toLocaleString('en-US',{maximumFractionDigits:0});
const fmtDate  = v => v? String(v).slice(0,10) : '';

function nextId(prefix, list, key){
  let max=0;
  for(const r of list){ const m=String(r[key]||'').match(/(\d+)\s*$/); if(m) max=Math.max(max,+m[1]); }
  return `${prefix}-${String(max+1).padStart(4,'0')}`;
}
function fundById(id){ return state.data.Funds.find(f=>f.FundID===id); }
function investorById(id){ return state.data.Investors.find(i=>i.InvestorID===id); }
function isProgramFund(f){ return !!(f&&f.RoutesTo); }
/** funds that take direct investment (excludes ECIF program which routes to SGF) */
function investableFunds(){ return state.data.Funds.filter(f=>!isProgramFund(f)); }
function firmById(id){ return state.data.ECIFFirms.find(f=>f.FirmID===id); }
function employeeById(id){ return state.data.ECIFEmployees.find(e=>e.EmployeeID===id); }

/* ---------- toast ---------- */
function toast(msg, kind='ok', title){
  const t=el('div',{class:`toast ${kind}`},
    title?el('strong',{},title):null,
    el('span',{},msg));
  $('#toastHost').append(t);
  setTimeout(()=>{t.style.opacity='0';t.style.transition='.3s';setTimeout(()=>t.remove(),320);},3200);
}

/* ============================================================
   Microsoft Authentication — MSAL.js 2.x
   ============================================================ */
const MSAL_CONFIG = {
  auth: {
    clientId: '0873a639-ed1f-4351-8d83-290e42775899',
    authority: 'https://login.microsoftonline.com/3aeede26-447f-4658-9da0-918778fa4255',
    redirectUri: window.location.origin + window.location.pathname,
  },
  cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: true },
};
const GRAPH_SCOPES = ['User.Read', 'Files.ReadWrite'];
const FILE_NAME = SHARED_FILE_PATH.split('/').pop();
const _SP_HOST = new URL(SHAREPOINT_SITE).hostname;
// Drive-relative path: strip the library name (first path segment) from SHARED_FILE_PATH
const _SP_DRIVE_PATH = '/' + SHARED_FILE_PATH.split('/').filter(Boolean).slice(1).join('/');
const msalApp = new msal.PublicClientApplication(MSAL_CONFIG);

async function getToken(){
  const acct = msalApp.getActiveAccount();
  if(!acct) throw new Error('Нэвтрээгүй байна.');
  try{
    const r = await msalApp.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: acct });
    return r.accessToken;
  }catch{
    await msalApp.acquireTokenRedirect({ scopes: GRAPH_SCOPES });
    throw new Error('Redirecting for token…');
  }
}

async function gFetch(path, opts={}){
  const token = await getToken();
  const resp = await fetch('https://graph.microsoft.com/v1.0'+path, {
    ...opts,
    headers: { Authorization: 'Bearer '+token, ...(opts.headers||{}) },
  });
  if(resp.status===404){ const e=new Error('Graph404'); e.status=404; throw e; }
  if(!resp.ok) throw new Error('Graph '+resp.status+': '+(await resp.text().catch(()=>'')));
  return resp;
}

/* ============================================================
   SharePoint workbook operations
   ============================================================ */
async function signIn(){
  try{
    await msalApp.loginRedirect({ scopes: GRAPH_SCOPES });
    // browser navigates away; init() handles the return via handleRedirectPromise()
  }catch(e){
    gateError('Нэвтрэхэд алдаа: '+(e.message||String(e)));
  }
}

async function signOut(){
  try{ await msalApp.logoutRedirect({ account: msalApp.getActiveAccount() }); }catch{}
  state.account = null;
  state.data = { Funds:[], Investors:[], Transactions:[], NAVHistory:[], Fees:[],
    ECIFFirms:[], ECIFEmployees:[], ECIFContributions:[], Meta:[] };
  state.dirty = false;
  state.savedAt = null;
  $('#app').classList.add('hidden');
  $('#gate').classList.remove('hidden');
  $('#gateHint').textContent = '';
  updateTopBarUser();
  updateFileStatus();
}

async function findOrCreateWorkbook(){
  try{
    const resp = await gFetch('/sites/'+_SP_HOST+'/drive/root:'+_SP_DRIVE_PATH+':/content');
    loadWorkbookBuffer(await resp.arrayBuffer());
    enterApp();
    toast('SharePoint-аас файл ачааллаа.','ok','Холбогдлоо');
  }catch(e){
    if(e.status===404){
      await createWorkbookOnSharePoint();
    }else{
      gateError('SharePoint алдаа: '+e.message);
      console.error(e);
    }
  }
}

async function createWorkbookOnSharePoint(){
  gateInfo(FILE_NAME+' файлыг SharePoint дээр үүсгэж байна…');
  state.data = {
    Funds: SEED_FUNDS.map(f=>({...f})),
    Investors:[], Transactions:[], NAVHistory:[], Fees:[],
    ECIFFirms:[], ECIFEmployees:[], ECIFContributions:[],
    Meta:[{Key:'SchemaVersion',Value:SCHEMA_VERSION},
          {Key:'CreatedAt',Value:new Date().toISOString()},
          {Key:'App',Value:'GAM Back Office'}],
  };
  for(const f of state.data.Funds){
    state.data.NAVHistory.push({
      NavID: nextId('NAV',state.data.NAVHistory,'NavID'),
      Date: f.RegisteredDate, FundID: f.FundID,
      TotalNAV: 0, UnitsOutstanding: 0, UnitPrice: f.NominalPrice,
      Source: 'Нэрлэсэн үнэ', Notes: 'Анхны нэгж эрхийн үнэ',
    });
  }
  await saveToSharePoint(true);
  toast('Шинэ '+FILE_NAME+' файл үүсгэж SharePoint-д хадгаллаа.','ok','Бэлэн боллоо');
  enterApp();
}

function loadWorkbookBuffer(buf){
  const wb = XLSX.read(buf,{type:'array'});
  const data = {};
  for(const sheet of Object.keys(SCHEMA)){
    const ws = wb.Sheets[sheet];
    data[sheet] = ws ? XLSX.utils.sheet_to_json(ws,{defval:''}) : [];
  }
  if(!data.Funds||!data.Funds.length) data.Funds = SEED_FUNDS.map(f=>({...f}));
  for(const sf of SEED_FUNDS){ if(!data.Funds.some(f=>f.FundID===sf.FundID)) data.Funds.push({...sf}); }
  if(!data.Meta||!data.Meta.length) data.Meta = [{Key:'SchemaVersion',Value:SCHEMA_VERSION}];
  state.data = data;
  state.dirty = false;
  state.savedAt = new Date();
}

async function ensureSpFolder(){
  const folderPath = _SP_DRIVE_PATH.split('/').slice(0, -1).join('/');
  if(!folderPath) return;
  try{
    await gFetch('/sites/'+_SP_HOST+'/drive/root:'+folderPath);
  }catch(e){
    if(e.status===404){
      const name = folderPath.split('/').pop();
      await gFetch('/sites/'+_SP_HOST+'/drive/root/children',{
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body: JSON.stringify({name, folder:{}, '@microsoft.graph.conflictBehavior':'replace'}),
      });
    }else{ throw e; }
  }
}

async function saveToSharePoint(silent){
  if(!state.account) return;
  const wb = XLSX.utils.book_new();
  for(const[sheet,cols] of Object.entries(SCHEMA)){
    const rows = (state.data[sheet]||[]).map(r=>{ const o={}; for(const c of cols) o[c]=r[c]??''; return o; });
    const ws = XLSX.utils.json_to_sheet(rows,{header:cols});
    XLSX.utils.book_append_sheet(wb,ws,sheet);
  }
  const out = XLSX.write(wb,{bookType:'xlsx',type:'array'});
  await ensureSpFolder();
  await gFetch('/sites/'+_SP_HOST+'/drive/root:'+_SP_DRIVE_PATH+':/content',{
    method: 'PUT',
    headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
    body: new Blob([out],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),
  });
  state.dirty = false;
  state.savedAt = new Date();
  updateFileStatus();
  if(!silent) toast('SharePoint рүү хадгалагдлаа.','ok','Хадгалсан');
}

async function reloadFromSharePoint(){
  if(!state.account) return;
  try{
    const resp = await gFetch('/sites/'+_SP_HOST+'/drive/root:'+_SP_DRIVE_PATH+':/content');
    loadWorkbookBuffer(await resp.arrayBuffer());
    render();
    toast('SharePoint-аас дахин ачаалагдлаа.','ok');
  }catch(e){
    toast('Дахин ачааллахад алдаа: '+e.message,'err');
  }
}

/** mutate + autosave helper */
async function commit(fn, okMsg){
  try{ fn(); state.dirty=true; updateFileStatus(); await saveToSharePoint(true);
    if(okMsg) toast(okMsg,'ok','Амжилттай'); render();
  }catch(e){ toast(e.message||'Алдаа гарлаа','err','Алдаа'); console.error(e); }
}

function gateError(msg){ const h=$('#gateHint'); h.textContent=msg; h.style.color='var(--red)'; }
function gateInfo(msg){ const h=$('#gateHint'); h.textContent=msg; h.style.color='var(--gam-blue)'; }

/* ============================================================
   Computed: balances & fund stats
   ============================================================ */
function confirmedTxns(){ return state.data.Transactions.filter(t=>t.Status==='Баталгаажсан'); }

/** units held by an investor in a fund (confirmed only) */
function investorUnits(investorId, fundId){
  let u=0;
  for(const t of confirmedTxns()){
    if(t.InvestorID!==investorId||t.FundID!==fundId) continue;
    u += t.Type==='Худалдан авалт' ? num(t.Units) : -num(t.Units);
  }
  return u;
}
/** total invested (net cash in) by investor in fund */
function investorNetCash(investorId, fundId){
  let c=0;
  for(const t of confirmedTxns()){
    if(t.InvestorID!==investorId||t.FundID!==fundId) continue;
    c += t.Type==='Худалдан авалт' ? num(t.NetAmount) : -num(t.NetAmount);
  }
  return c;
}
/** outstanding units of a fund */
function fundUnitsOutstanding(fundId){
  let u=0;
  for(const t of confirmedTxns()){
    if(t.FundID!==fundId) continue;
    u += t.Type==='Худалдан авалт' ? num(t.Units) : -num(t.Units);
  }
  return u;
}
/** latest unit price for a fund (from NAV history; fallback nominal) */
function latestUnitPrice(fundId){
  const f=fundById(fundId);
  const rows=state.data.NAVHistory.filter(n=>n.FundID===fundId&&num(n.UnitPrice)>0)
    .sort((a,b)=>String(b.Date).localeCompare(String(a.Date)));
  return rows.length ? num(rows[0].UnitPrice) : (f?num(f.NominalPrice):0);
}
/** investors holding >0 units in a fund */
function fundHolders(fundId){
  return state.data.Investors.filter(i=>investorUnits(i.InvestorID,fundId)>0.00001).length;
}
function fundAUM(fundId){ return fundUnitsOutstanding(fundId)*latestUnitPrice(fundId); }
/** most recent recorded Total NAV for a fund (0 if none) */
function latestTotalNAV(fundId){
  const rows=state.data.NAVHistory.filter(n=>n.FundID===fundId&&num(n.TotalNAV)>0)
    .sort((a,b)=>String(b.Date).localeCompare(String(a.Date)));
  return rows.length? num(rows[0].TotalNAV) : 0;
}
/** Total NAV recorded on or before a given date (for prior-year-end opening) */
function navAsOf(fundId,dateStr){
  const rows=state.data.NAVHistory.filter(n=>n.FundID===fundId&&num(n.TotalNAV)>0&&String(n.Date)<=dateStr)
    .sort((a,b)=>String(b.Date).localeCompare(String(a.Date)));
  return rows.length? num(rows[0].TotalNAV) : 0;
}

/* ---------- ECIF (Employee Contribution Investment Fund) ---------- */
const ECIF_TARGET_FUND='SGF'; // ECIF firms are nominal holders of Stable Growth Fund
function ecifFirmEmployees(firmId){ return state.data.ECIFEmployees.filter(e=>e.FirmID===firmId); }
function ecifEmployeeUnits(empId){
  return state.data.ECIFContributions.filter(c=>c.EmployeeID===empId).reduce((s,c)=>s+num(c.Units),0);
}
function ecifEmployeeContributed(empId){
  return state.data.ECIFContributions.filter(c=>c.EmployeeID===empId&&num(c.TotalAmount)>0).reduce((s,c)=>s+num(c.TotalAmount),0);
}
function ecifFirmUnits(firmId){
  // firm's nominal SGF holding = sum of its employees' units (== firm investor units in SGF)
  return state.data.ECIFContributions.filter(c=>c.FirmID===firmId).reduce((s,c)=>s+num(c.Units),0);
}
function ecifFirmContributed(firmId){
  return state.data.ECIFContributions.filter(c=>c.FirmID===firmId&&num(c.TotalAmount)>0).reduce((s,c)=>s+num(c.TotalAmount),0);
}
function ecifValue(units){ return units*latestUnitPrice(ECIF_TARGET_FUND); }
function ecifTotalUnits(){ return state.data.ECIFContributions.reduce((s,c)=>s+num(c.Units),0); }
function ecifTotalAUM(){ return ecifValue(ecifTotalUnits()); }
function ecifActiveEmployees(){ return state.data.ECIFEmployees.filter(e=>e.Status!=='Гарсан').length; }

/* ============================================================
   App shell wiring
   ============================================================ */
function enterApp(){
  $('#gate').classList.add('hidden');
  $('#app').classList.remove('hidden');
  updateFileStatus();
  updateTopBarUser();
  render();
}
function updateTopBarUser(){
  const info=$('#userInfo');
  if(state.account){
    $('#userName').textContent = state.account.name||state.account.username||'';
    $('#userEmail').textContent = state.account.username||'';
    info.classList.remove('hidden');
  }else{
    info.classList.add('hidden');
  }
}
function updateFileStatus(){
  const dot=$('#connDot'), name=$('#fileName'), saved=$('#fileSaved');
  if(!state.account){ dot.className='dot dot-off'; name.textContent='Холбогдоогүй · Not connected'; saved.textContent=''; return; }
  name.textContent=FILE_NAME;
  if(state.dirty){ dot.className='dot dot-dirty'; saved.textContent='Хадгалаагүй өөрчлөлт байна'; }
  else{ dot.className='dot dot-on'; saved.textContent= state.savedAt? 'Хадгалсан: '+state.savedAt.toLocaleString('en-GB') : 'Synced'; }
}

/* ============================================================
   Router / render
   ============================================================ */
const VIEWS={};
function render(){
  $$('.nav-item').forEach(b=>b.classList.toggle('active',b.dataset.view===state.view));
  const fn=VIEWS[state.view]||VIEWS.dashboard;
  const root=$('#content'); root.innerHTML=''; root.append(fn());
  updateFileStatus();
}
function go(view){ state.view=view; render(); }

/* ---------- shared UI bits ---------- */
function viewHead(titleMN,titleEN,actions){
  return el('div',{class:'view-head'},
    el('div',{},el('h2',{},titleMN), el('div',{class:'sub'},titleEN)),
    actions?el('div',{class:'view-actions'},actions):null);
}
function fundChip(fundId){
  const f=fundById(fundId); if(!f) return el('span',{},fundId||'');
  return el('span',{class:'fund-chip'},
    el('span',{class:'fund-dot',style:`background:${f.Color||'#888'}`}), f.ShortName||f.FundID);
}
function statusBadge(s){
  const map={'Баталгаажсан':'green','Хүлээгдэж буй':'amber','Цуцлагдсан':'red','Идэвхтэй':'green','Идэвхгүй':'gray'};
  return el('span',{class:`badge badge-${map[s]||'gray'}`},s);
}
function emptyRow(cols,msg){ return el('tr',{},el('td',{class:'empty',colspan:cols},msg)); }

/* ============================================================
   VIEW: Dashboard
   ============================================================ */
VIEWS.dashboard=function(){
  const wrap=el('div',{});
  wrap.append(viewHead('Хяналтын самбар','Dashboard overview'));

  // KPIs
  const totalAUM=investableFunds().reduce((s,f)=>s+fundAUM(f.FundID),0);
  const cards=el('div',{class:'cards'});
  cards.append(kpi('Нийт хөрөнгө (AUM)', '₮'+fmtMoney(totalAUM,0), investableFunds().length+' сан'));
  cards.append(kpi('Хөрөнгө оруулагч', fmtInt(state.data.Investors.length), 'нийт бүртгэлтэй'));
  const pend=state.data.Transactions.filter(t=>t.Status==='Хүлээгдэж буй').length;
  cards.append(kpi('Гүйлгээ', fmtInt(state.data.Transactions.length), pend?(pend+' хүлээгдэж буй'):'бүгд баталгаажсан'));
  if(state.data.ECIFFirms.length)
    cards.append(kpi('ECIF хуримтлал', '₮'+fmtMoney(ecifTotalAUM(),0), state.data.ECIFFirms.length+' байгууллага · '+ecifActiveEmployees()+' ажилтан'));
  wrap.append(cards);

  // per-fund panel
  const panel=el('div',{class:'panel'});
  panel.append(el('div',{class:'panel-head'},el('h3',{},'Сан тус бүрийн төлөв · Per-fund summary')));
  const t=el('table',{class:'grid'});
  t.innerHTML=`<thead><tr><th>Сан</th><th class="num">Нэгж эрхийн үнэ</th><th class="num">Гүйлгээнд буй нэгж</th>
    <th class="num">Эзэмшигч</th><th class="num">Цэвэр хөрөнгө (AUM)</th><th class="num">Татан төвлөрүүлэлт</th></tr></thead>`;
  const tb=el('tbody');
  for(const f of investableFunds()){
    const aum=fundAUM(f.FundID), pct= num(f.TargetRaise)? aum/num(f.TargetRaise)*100:0;
    tb.append(el('tr',{},
      el('td',{},fundChip(f.FundID)),
      el('td',{class:'num'},'₮'+fmtMoney(latestUnitPrice(f.FundID),2)),
      el('td',{class:'num'},fmtUnits(fundUnitsOutstanding(f.FundID))),
      el('td',{class:'num'},fmtInt(fundHolders(f.FundID))),
      el('td',{class:'num'},'₮'+fmtMoney(aum,0)),
      el('td',{class:'num'},pct.toFixed(1)+'%'),
    ));
  }
  t.append(tb); panel.append(el('div',{class:'table-wrap'},t)); wrap.append(panel);

  // recent transactions
  const recent=[...state.data.Transactions].sort((a,b)=>String(b.CreatedAt||b.TradeDate).localeCompare(String(a.CreatedAt||a.TradeDate))).slice(0,8);
  const rp=el('div',{class:'panel'});
  rp.append(el('div',{class:'panel-head'},el('h3',{},'Сүүлийн гүйлгээ · Recent transactions'),
    el('button',{class:'btn btn-ghost btn-sm',onclick:()=>go('transactions')},'Бүгдийг харах →')));
  rp.append(el('div',{class:'table-wrap'},txnTable(recent,true)));
  wrap.append(rp);
  return wrap;
};
function kpi(label,value,sub){
  return el('div',{class:'card kpi'},el('div',{class:'k-label'},label),el('div',{class:'k-value'},value),sub?el('div',{class:'k-sub'},sub):null);
}

/* ============================================================
   VIEW: Investors
   ============================================================ */
VIEWS.investors=function(){
  const wrap=el('div',{});
  wrap.append(viewHead('Хөрөнгө оруулагч','Investor registry',
    el('button',{class:'btn btn-primary',onclick:()=>investorForm()},'＋ Шинэ хөрөнгө оруулагч')));

  const search=el('input',{type:'text',placeholder:'Нэр, регистр, утсаар хайх…',oninput:e=>filterRows(e.target.value)});
  wrap.append(el('div',{class:'toolbar'},el('div',{class:'search'},search)));

  const panel=el('div',{class:'panel'});
  const t=el('table',{class:'grid'});
  t.innerHTML=`<thead><tr><th>ID</th><th>Нэр / Name</th><th>Төрөл</th><th>Регистр</th><th>Утас</th>
    <th class="num">Эзэмшил</th><th>Төлөв</th><th></th></tr></thead>`;
  const tb=el('tbody',{id:'invBody'});
  const list=[...state.data.Investors].sort((a,b)=>String(a.InvestorID).localeCompare(String(b.InvestorID)));
  if(!list.length) tb.append(emptyRow(8,'Хөрөнгө оруулагч бүртгэгдээгүй байна. «Шинэ хөрөнгө оруулагч» дарж эхлүүлнэ үү.'));
  for(const i of list) tb.append(investorRow(i));
  t.append(tb); panel.append(el('div',{class:'table-wrap'},t)); wrap.append(panel);
  return wrap;
};
function investorRow(i){
  const holdings=state.data.Funds.map(f=>{const u=investorUnits(i.InvestorID,f.FundID); return u>0?`${f.ShortName}: ${fmtUnits(u)}`:null;}).filter(Boolean);
  const tr=el('tr',{'data-search':`${i.InvestorID} ${i.NameMN} ${i.NameEN} ${i.RegNo} ${i.Phone}`.toLowerCase()},
    el('td',{},i.InvestorID),
    el('td',{},el('div',{},el('strong',{},i.NameMN||''), i.NameEN?el('div',{class:'small muted'},i.NameEN):null)),
    el('td',{},i.InvestorType||''),
    el('td',{},i.RegNo||''),
    el('td',{},i.Phone||''),
    el('td',{class:'num'},holdings.length?el('span',{class:'small'},holdings.join(' · ')):el('span',{class:'muted'},'—')),
    el('td',{},statusBadge(i.Status||'Идэвхтэй')),
    el('td',{},el('div',{class:'row-actions'},
      el('button',{class:'btn btn-ghost btn-sm',onclick:()=>investorDetail(i.InvestorID)},'Дэлгэрэнгүй'),
      el('button',{class:'btn btn-ghost btn-sm',onclick:()=>investorForm(i)},'Засах'))),
  );
  return tr;
}
function filterRows(q){
  q=q.toLowerCase().trim();
  $$('#invBody tr').forEach(tr=>{ const s=tr.getAttribute('data-search')||''; tr.style.display=(!q||s.includes(q))?'':'none'; });
}

function investorForm(existing){
  const isEdit=!!existing;
  const i=existing||{InvestorID:nextId('INV',state.data.Investors,'InvestorID'),RegDate:todayISO(),InvestorType:'Хувь хүн',Status:'Идэвхтэй'};
  const f=(name,labelMN,labelEN,opts={})=>fieldHTML(name,labelMN,labelEN,i[name],opts);
  const body=el('div',{},
    el('div',{class:'form-grid'},
      f('InvestorID','Дугаар','ID',{readonly:true}),
      f('RegDate','Бүртгэсэн огноо','Reg. date',{type:'date'}),
      f('InvestorType','Төрөл','Type',{type:'select',options:INV_TYPES}),
      f('Status','Төлөв','Status',{type:'select',options:INV_STATUS}),
      f('NameMN','Нэр (Монгол)','Name (MN)',{required:true,full:true}),
      f('NameEN','Нэр (English)','Name (EN)',{full:true}),
      f('RegNo','Регистр / Улсын бүртгэл','Registration no.'),
      f('ContactPerson','Холбоо барих хүн','Contact person'),
      f('Phone','Утас','Phone'),
      f('Email','И-мэйл','Email',{type:'email'}),
      f('Address','Хаяг','Address',{full:true}),
      f('BankName','Банк','Bank'),
      f('BankAccount','Дансны дугаар','Account no.'),
      f('AccountName','Данс эзэмшигчийн нэр','Account name',{full:true}),
      f('Notes','Тэмдэглэл','Notes',{type:'textarea',full:true}),
    )
  );
  modal(isEdit?'Хөрөнгө оруулагч засах':'Шинэ хөрөнгө оруулагч', body,
    {okText:isEdit?'Хадгалах':'Бүртгэх', onOk:()=>{
      const rec=readForm(body,SCHEMA.Investors,i);
      if(!rec.NameMN){ toast('Нэр заавал бөглөнө.','err'); return false; }
      commit(()=>{ if(isEdit){ Object.assign(existing,rec);} else { state.data.Investors.push(rec);} },
        isEdit?'Хөрөнгө оруулагч шинэчлэгдлээ':'Хөрөнгө оруулагч бүртгэгдлээ');
    }});
}

function investorDetail(id){
  const i=investorById(id); if(!i)return;
  const body=el('div',{});
  const dl=el('dl',{class:'def-list'});
  const add=(k,v)=>{ dl.append(el('dt',{},k),el('dd',{},v||'—')); };
  add('Дугаар',i.InvestorID); add('Төрөл',i.InvestorType); add('Регистр',i.RegNo);
  add('Утас',i.Phone); add('И-мэйл',i.Email); add('Банк',`${i.BankName||''} ${i.BankAccount||''}`);
  body.append(el('h3',{style:'margin:0 0 4px;color:var(--gam-navy)'},i.NameMN||''),
    i.NameEN?el('div',{class:'muted',style:'margin-bottom:12px'},i.NameEN):el('div',{style:'margin-bottom:12px'}),
    dl);
  // holdings
  body.append(el('h4',{style:'margin:18px 0 8px;color:var(--gam-navy)'},'Эзэмшил · Holdings'));
  const t=el('table',{class:'grid'});
  t.innerHTML=`<thead><tr><th>Сан</th><th class="num">Нэгж эрх</th><th class="num">Цэвэр хөрөнгө оруулалт</th><th class="num">Одоогийн үнэлгээ</th></tr></thead>`;
  const tb=el('tbody'); let any=false;
  for(const f of state.data.Funds){
    const u=investorUnits(id,f.FundID); if(u<=0.00001)continue; any=true;
    const val=u*latestUnitPrice(f.FundID);
    tb.append(el('tr',{},el('td',{},fundChip(f.FundID)),el('td',{class:'num'},fmtUnits(u)),
      el('td',{class:'num'},'₮'+fmtMoney(investorNetCash(id,f.FundID),0)),el('td',{class:'num'},'₮'+fmtMoney(val,0))));
  }
  if(!any) tb.append(emptyRow(4,'Идэвхтэй эзэмшилгүй.'));
  t.append(tb); body.append(el('div',{class:'table-wrap'},t));
  // transactions
  const txns=state.data.Transactions.filter(x=>x.InvestorID===id)
    .sort((a,b)=>String(b.TradeDate).localeCompare(String(a.TradeDate)));
  body.append(el('h4',{style:'margin:18px 0 8px;color:var(--gam-navy)'},'Гүйлгээ · Transactions'));
  body.append(el('div',{class:'table-wrap'},txnTable(txns,true)));
  modal('Хөрөнгө оруулагчийн дэлгэрэнгүй', body, {okText:'Хаах', cancel:false, wide:true});
}

/* ============================================================
   VIEW: Subscription (Худалдан авалт)
   ============================================================ */
VIEWS.subscribe=function(){ return tradeView('Худалдан авалт'); };
VIEWS.redeem   =function(){ return tradeView('Буцаан худалдалт'); };

function tradeView(kind){
  const isSub = kind==='Худалдан авалт';
  const wrap=el('div',{});
  wrap.append(viewHead(
    isSub?'Нэгж эрх худалдан авалт':'Нэгж эрх буцаан худалдалт',
    isSub?'Subscription — investor buys fund units':'Redemption — investor sells units back'));

  if(!state.data.Investors.length){
    wrap.append(el('div',{class:'panel'},el('div',{class:'panel-body'},
      el('p',{},'Эхлээд хөрөнгө оруулагч бүртгэнэ үү.'),
      el('button',{class:'btn btn-primary',onclick:()=>go('investors')},'→ Хөрөнгө оруулагч руу'))));
    return wrap;
  }

  const panel=el('div',{class:'panel'});
  panel.append(el('div',{class:'panel-head'},el('h3',{},isSub?'Шинэ худалдан авалт':'Шинэ буцаан худалдалт')));
  const b=el('div',{class:'panel-body'});

  // form refs
  const fundSel=selectEl('FundID',investableFunds().map(f=>({v:f.FundID,t:`${f.ShortName} (${f.FundID})`})));
  const invSel =selectEl('InvestorID',state.data.Investors.map(i=>({v:i.InvestorID,t:`${i.NameMN} (${i.InvestorID})`})));
  const dateIn =el('input',{type:'date',value:todayISO()});
  const priceIn=el('input',{type:'number',step:'0.01',min:'0'});
  const amtIn  =el('input',{type:'number',step:'0.01',min:'0',placeholder:'Төгрөг'});
  const unitsIn=el('input',{type:'number',step:'0.0001',min:'0',placeholder:'Нэгж эрх'});
  const feePctIn=el('input',{type:'number',step:'0.01',min:'0',value:'0'});
  const settleIn=el('input',{type:'date',value:todayISO()});
  const refIn  =el('input',{type:'text',placeholder:'Гүйлгээний утга / баримт'});
  const noteIn =el('input',{type:'text'});
  const holdingNote=el('div',{class:'hint'});

  const calc=el('div',{class:'calc-box'});

  function syncPrice(){ const p=latestUnitPrice(fundSel.value); if(p && !priceIn.value) priceIn.value=p; recalc(); }
  function updateHolding(){
    if(isSub){ holdingNote.textContent=''; return; }
    const u=investorUnits(invSel.value,fundSel.value);
    holdingNote.textContent=`Эзэмшиж буй нэгж эрх: ${fmtUnits(u)}`;
  }
  function recalc(){
    const price=num(priceIn.value);
    let units, gross;
    if(isSub){ // amount-driven
      gross=num(amtIn.value); units= price? gross/price : 0; unitsIn.value= units? units.toFixed(4):'';
    }else{ // units-driven
      units=num(unitsIn.value); gross=units*price; amtIn.value= gross? gross.toFixed(2):'';
    }
    const feePct=num(feePctIn.value), fee=gross*feePct/100;
    const net= isSub ? gross /* investor pays gross, fee taken from units side optionally */ : gross-fee;
    const netLabel= isSub? gross-fee : gross-fee;
    calc.innerHTML='';
    calc.append(
      crow('Нэгж эрхийн үнэ', '₮'+fmtMoney(price,2)),
      crow('Нэгж эрх', fmtUnits(units)),
      crow('Дүн (gross)', '₮'+fmtMoney(gross,2)),
      crow(`Шимтгэл (${feePct||0}%)`, '−₮'+fmtMoney(fee,2)),
      crow(isSub?'Цэвэр оруулах дүн':'Хөрөнгө оруулагчид төлөх', '₮'+fmtMoney(gross-fee,2), true),
    );
  }
  fundSel.addEventListener('change',()=>{priceIn.value='';syncPrice();updateHolding();});
  invSel.addEventListener('change',updateHolding);
  [amtIn,unitsIn,priceIn,feePctIn].forEach(x=>x.addEventListener('input',recalc));

  const grid=el('div',{class:'form-grid'});
  grid.append(
    wrapField('Сан','Fund',fundSel,true),
    wrapField('Хөрөнгө оруулагч','Investor',invSel,true),
    wrapField('Арилжааны огноо','Trade date',dateIn),
    wrapField('Нэгж эрхийн үнэ','Unit price (₮)',priceIn,true),
    isSub? wrapField('Оруулах дүн','Amount (₮)',amtIn,true) : wrapField('Буцаах нэгж эрх','Units to redeem',unitsIn,true),
    isSub? wrapField('Нэгж эрх (тооцоолсон)','Units (calculated)',unitsIn) : wrapField('Дүн (тооцоолсон)','Amount (calculated)',amtIn),
    wrapField('Шимтгэл %','Fee %',feePctIn),
    wrapField('Төлбөр хийгдэх огноо','Settlement date',settleIn),
    wrapField('Гүйлгээний утга','Reference',refIn),
    wrapField('Тэмдэглэл','Notes',noteIn),
  );
  if(!isSub){ unitsIn.readOnly=false; amtIn.readOnly=true; } else { amtIn.readOnly=false; unitsIn.readOnly=true; }
  b.append(grid);
  if(!isSub) b.append(holdingNote);
  b.append(el('div',{class:'two-col',style:'margin-top:14px'},
    calc,
    el('div',{style:'display:flex;align-items:flex-end;justify-content:flex-end;gap:10px'},
      el('button',{class:'btn btn-ghost',onclick:()=>go('transactions')},'Цуцлах'),
      el('button',{class:'btn btn-primary',onclick:submit}, isSub?'Худалдан авалт бүртгэх':'Буцаан худалдалт бүртгэх'))));

  panel.append(b); wrap.append(panel);
  syncPrice(); updateHolding(); recalc();

  function submit(){
    const fundId=fundSel.value, invId=invSel.value, price=num(priceIn.value);
    const units=num(unitsIn.value), gross=num(amtIn.value), feePct=num(feePctIn.value), fee=gross*feePct/100;
    if(!fundId||!invId){ toast('Сан болон хөрөнгө оруулагч сонгоно уу.','err'); return; }
    if(price<=0){ toast('Нэгж эрхийн үнэ оруулна уу.','err'); return; }
    if(units<=0){ toast('Нэгж эрх / дүн оруулна уу.','err'); return; }
    if(!isSub){
      const have=investorUnits(invId,fundId);
      if(units>have+0.00001){ toast(`Эзэмшил хүрэлцэхгүй. Боломжит: ${fmtUnits(have)} нэгж эрх.`,'err'); return; }
    }
    const rec={
      TxnID:nextId('TXN',state.data.Transactions,'TxnID'), TradeDate:dateIn.value, Type:kind,
      FundID:fundId, InvestorID:invId, Units:units, UnitPrice:price, GrossAmount:gross,
      FeePct:feePct, FeeAmount:fee, NetAmount:gross-fee, SettlementDate:settleIn.value,
      Status:'Баталгаажсан', PaymentRef:refIn.value, Notes:noteIn.value, CreatedAt:new Date().toISOString(),
    };
    commit(()=>state.data.Transactions.push(rec),
      (isSub?'Худалдан авалт':'Буцаан худалдалт')+' бүртгэгдлээ ('+rec.TxnID+')');
    go('transactions');
  }
  return wrap;
}
function crow(l,v,total){ return el('div',{class:'calc-row'+(total?' total':'')},el('span',{},l),el('span',{},v)); }
function selectEl(name,opts){ const s=el('select',{name}); s.append(el('option',{value:''},'— сонгох —'));
  for(const o of opts) s.append(el('option',{value:o.v},o.t)); return s; }
function wrapField(mn,en,inputEl,req){
  return el('div',{class:'field'+(inputEl.classList.contains('full')?' full':'')},
    el('label',{},mn+' ',el('span',{class:'en'},'· '+en),req?el('span',{class:'req'},' *'):''),inputEl);
}

/* ============================================================
   VIEW: Transactions
   ============================================================ */
VIEWS.transactions=function(){
  const wrap=el('div',{});
  wrap.append(viewHead('Гүйлгээний түүх','Transaction history',
    el('div',{},
      el('button',{class:'btn btn-outline btn-sm',onclick:()=>go('subscribe')},'＋ Худалдан авалт'),
      ' ',
      el('button',{class:'btn btn-outline btn-sm',onclick:()=>go('redeem')},'－ Буцаан худалдалт'))));

  // filters
  let fFund='',fType='';
  const fundSel=el('select',{onchange:e=>{fFund=e.target.value;draw();}});
  fundSel.append(el('option',{value:''},'Бүх сан'));
  investableFunds().forEach(f=>fundSel.append(el('option',{value:f.FundID},f.ShortName)));
  const typeSel=el('select',{onchange:e=>{fType=e.target.value;draw();}});
  ['','Худалдан авалт','Буцаан худалдалт'].forEach(t=>typeSel.append(el('option',{value:t},t||'Бүх төрөл')));
  wrap.append(el('div',{class:'toolbar'},fundSel,typeSel,
    el('button',{class:'btn btn-ghost btn-sm',style:'margin-left:auto',onclick:exportCSV},'⤓ CSV татах')));

  const panel=el('div',{class:'panel'});
  const host=el('div',{class:'table-wrap'});
  panel.append(host); wrap.append(panel);
  function draw(){ let list=[...state.data.Transactions];
    if(fFund) list=list.filter(t=>t.FundID===fFund);
    if(fType) list=list.filter(t=>t.Type===fType);
    list.sort((a,b)=>String(b.TradeDate).localeCompare(String(a.TradeDate))||String(b.TxnID).localeCompare(String(a.TxnID)));
    host.innerHTML=''; host.append(txnTable(list,false));
  }
  draw();
  function exportCSV(){
    const cols=SCHEMA.Transactions;
    const lines=[cols.join(',')].concat(state.data.Transactions.map(r=>cols.map(c=>JSON.stringify(r[c]??'')).join(',')));
    const blob=new Blob([lines.join('\n')],{type:'text/csv'});
    const a=el('a',{href:URL.createObjectURL(blob),download:'GAM_transactions.csv'}); a.click();
  }
  return wrap;
};

function txnTable(list,compact){
  const t=el('table',{class:'grid'});
  t.innerHTML=`<thead><tr><th>Огноо</th><th>Дугаар</th><th>Төрөл</th><th>Сан</th>${compact?'':'<th>Хөрөнгө оруулагч</th>'}
    <th class="num">Нэгж эрх</th><th class="num">Үнэ</th><th class="num">Дүн</th><th>Төлөв</th>${compact?'':'<th></th>'}</tr></thead>`;
  const tb=el('tbody');
  if(!list.length){ tb.append(emptyRow(compact?8:10,'Гүйлгээ алга.')); }
  for(const x of list){
    const inv=investorById(x.InvestorID);
    const typeBadge=el('span',{class:'badge '+(x.Type==='Худалдан авалт'?'badge-blue':'badge-amber')},x.Type);
    const cells=[ el('td',{},fmtDate(x.TradeDate)), el('td',{},x.TxnID), el('td',{},typeBadge), el('td',{},fundChip(x.FundID)) ];
    if(!compact) cells.push(el('td',{},inv?inv.NameMN:x.InvestorID));
    cells.push(el('td',{class:'num'},fmtUnits(x.Units)), el('td',{class:'num'},'₮'+fmtMoney(x.UnitPrice,2)),
      el('td',{class:'num'},'₮'+fmtMoney(x.NetAmount,0)), el('td',{},statusBadge(x.Status)));
    if(!compact) cells.push(el('td',{},el('div',{class:'row-actions'},
      x.Status!=='Цуцлагдсан'? el('button',{class:'btn btn-ghost btn-sm',onclick:()=>cancelTxn(x.TxnID)},'Цуцлах'):null)));
    tb.append(el('tr',{},cells));
  }
  t.append(tb); return t;
}
function cancelTxn(id){
  const x=state.data.Transactions.find(t=>t.TxnID===id); if(!x)return;
  confirmModal('Гүйлгээ цуцлах','Энэ гүйлгээг цуцлах уу? Эзэмшлийн тооцоо шинэчлэгдэнэ.',()=>{
    commit(()=>{ x.Status='Цуцлагдсан'; }, 'Гүйлгээ цуцлагдлаа');
  });
}

/* ============================================================
   VIEW: Funds
   ============================================================ */
VIEWS.funds=function(){
  const wrap=el('div',{});
  wrap.append(viewHead('Сангууд','Funds master data'));
  for(const f of state.data.Funds){
    const panel=el('div',{class:'panel'});
    panel.append(el('div',{class:'panel-head'},
      el('h3',{},fundChip(f.FundID),' — ',f.NameMN),
      el('button',{class:'btn btn-ghost btn-sm',onclick:()=>fundForm(f)},'Засах')));
    const body=el('div',{class:'panel-body two-col'});
    const dl1=el('dl',{class:'def-list'});
    const dl2=el('dl',{class:'def-list'});
    const add=(dl,k,v)=>{dl.append(el('dt',{},k),el('dd',{},v));};
    add(dl1,'English name',f.NameEN); add(dl1,'Төрөл',f.Type); add(dl1,'Бүртгэсэн',fmtDate(f.RegisteredDate));
    add(dl1,'Хугацаа',f.TermYears+' жил'); add(dl1,'Кастодиан',f.Custodian);
    add(dl2,'Татан төвлөрүүлэлт','₮'+fmtMoney(f.TargetRaise,0)); add(dl2,'Нийт нэгж эрх',fmtInt(f.AuthorizedUnits));
    add(dl2,'Нэрлэсэн үнэ','₮'+fmtMoney(f.NominalPrice,2)); add(dl2,'Удирдлагын шимтгэл',f.MgmtFeePct+'%');
    add(dl2,'Гүйцэтгэлийн шимтгэл',f.PerfFeePct+'%');
    body.append(dl1,dl2);
    // live stats
    if(isProgramFund(f)){
      const tgt=fundById(f.RoutesTo);
      body.append(el('div',{class:'full',style:'grid-column:1/-1'},
        el('div',{class:'small muted',style:'margin:2px 0 4px'},
          '↪ Хөрөнгө оруулалт ',el('strong',{},tgt?tgt.ShortName:f.RoutesTo),
          ' сангаар дамжина. Оролцогч байгууллага бүр уг санд нэрлэсэн данс эзэмшигч болно.'),
        el('div',{class:'cards',style:'margin:6px 0 0'},
          kpi('Нэгж эрхийн үнэ','₮'+fmtMoney(latestUnitPrice(ECIF_TARGET_FUND),2),(tgt?tgt.ShortName:'')+' үнэ'),
          kpi('Оролцогч байгууллага',fmtInt(state.data.ECIFFirms.length),ecifActiveEmployees()+' ажилтан'),
          kpi('Хуримтлалын үнэлгээ','₮'+fmtMoney(ecifTotalAUM(),0),fmtUnits(ecifTotalUnits())+' нэгж эрх'))));
    } else {
      body.append(el('div',{class:'full',style:'grid-column:1/-1'},
        el('div',{class:'cards',style:'margin:6px 0 0'},
          kpi('Нэгж эрхийн үнэ','₮'+fmtMoney(latestUnitPrice(f.FundID),2),'хамгийн сүүлийн'),
          kpi('Гүйлгээнд буй нэгж',fmtUnits(fundUnitsOutstanding(f.FundID)),''),
          kpi('Цэвэр хөрөнгө','₮'+fmtMoney(fundAUM(f.FundID),0),fmtInt(fundHolders(f.FundID))+' эзэмшигч'))));
    }
    panel.append(body); wrap.append(panel);
  }
  return wrap;
};
function fundForm(f){
  const g=(name,mn,en,opts={})=>fieldHTML(name,mn,en,f[name],opts);
  const body=el('div',{},el('div',{class:'form-grid'},
    g('FundID','Дугаар','ID',{readonly:true}),
    g('ShortName','Богино нэр','Short name'),
    g('NameMN','Нэр (Монгол)','Name (MN)',{full:true}),
    g('NameEN','Нэр (English)','Name (EN)',{full:true}),
    g('Type','Төрөл','Type',{full:true}),
    g('RegisteredDate','Бүртгэсэн огноо','Reg. date',{type:'date'}),
    g('TermYears','Хугацаа (жил)','Term (yrs)',{type:'number'}),
    g('TargetRaise','Татан төвлөрүүлэлт','Target raise',{type:'number'}),
    g('AuthorizedUnits','Нийт нэгж эрх','Authorized units',{type:'number'}),
    g('NominalPrice','Нэрлэсэн үнэ','Nominal price',{type:'number'}),
    g('MgmtFeePct','Удирдлагын шимтгэл %','Mgmt fee %',{type:'number'}),
    g('PerfFeePct','Гүйцэтгэлийн шимтгэл %','Perf fee %',{type:'number'}),
    g('Custodian','Кастодиан','Custodian',{full:true}),
    g('Status','Төлөв','Status'),
  ));
  modal('Сан засах', body, {okText:'Хадгалах', onOk:()=>{
    const rec=readForm(body,SCHEMA.Funds,f); Object.assign(f,rec);
    commit(()=>{}, 'Сангийн мэдээлэл шинэчлэгдлээ');
  }});
}

/* ============================================================
   VIEW: NAV
   ============================================================ */
VIEWS.nav=function(){
  const wrap=el('div',{});
  wrap.append(viewHead('НЦХҮ / Нэгж эрхийн үнэ','Net Asset Value & unit price',
    el('button',{class:'btn btn-primary',onclick:()=>navForm()},'＋ НЦХҮ бүртгэх')));

  wrap.append(el('div',{class:'panel'},el('div',{class:'panel-body'},
    el('p',{class:'muted',style:'margin:0'},
      'Өдөр тутмын цэвэр хөрөнгийн үнэлгээ (НЦХҮ) болон нэгж эрхийн үнийг энд бүртгэнэ. ',
      'Нэгж эрхийн үнэ = Сангийн цэвэр хөрөнгө ÷ Гүйлгээнд буй нэгж эрх. ',
      el('strong',{},'Тэмдэглэл:'),' НЦХҮ-г Голомт банкны кастодиантай тулган баталгаажуулна.'))));

  for(const f of investableFunds()){
    const rows=state.data.NAVHistory.filter(n=>n.FundID===f.FundID)
      .sort((a,b)=>String(b.Date).localeCompare(String(a.Date)));
    const panel=el('div',{class:'panel'});
    panel.append(el('div',{class:'panel-head'},el('h3',{},fundChip(f.FundID),' — НЦХҮ түүх'),
      el('button',{class:'btn btn-ghost btn-sm',onclick:()=>navForm(f.FundID)},'＋ Бүртгэх')));
    const t=el('table',{class:'grid'});
    t.innerHTML=`<thead><tr><th>Огноо</th><th class="num">Цэвэр хөрөнгө</th><th class="num">Гүйлгээнд буй нэгж</th>
      <th class="num">Нэгж эрхийн үнэ</th><th>Эх сурвалж</th><th>Тэмдэглэл</th></tr></thead>`;
    const tb=el('tbody');
    if(!rows.length) tb.append(emptyRow(6,'Бүртгэл алга.'));
    for(const n of rows) tb.append(el('tr',{},el('td',{},fmtDate(n.Date)),
      el('td',{class:'num'},'₮'+fmtMoney(n.TotalNAV,0)),el('td',{class:'num'},fmtUnits(n.UnitsOutstanding)),
      el('td',{class:'num'},'₮'+fmtMoney(n.UnitPrice,2)),el('td',{},n.Source||''),el('td',{},n.Notes||'')));
    t.append(tb); panel.append(el('div',{class:'table-wrap'},t)); wrap.append(panel);
  }
  return wrap;
};
function navForm(presetFund){
  const fundSel=selectEl('FundID',investableFunds().map(f=>({v:f.FundID,t:f.ShortName})));
  if(presetFund) fundSel.value=presetFund;
  const dateIn=el('input',{type:'date',value:todayISO()});
  const navIn =el('input',{type:'number',step:'0.01',min:'0',placeholder:'Нийт цэвэр хөрөнгө ₮'});
  const unitsIn=el('input',{type:'number',step:'0.0001',min:'0'});
  const priceOut=el('input',{type:'text',readonly:true});
  const srcIn=el('input',{type:'text',value:'Дотоод тооцоо'});
  const noteIn=el('input',{type:'text'});
  function syncUnits(){ if(fundSel.value && !unitsIn.value) unitsIn.value=fundUnitsOutstanding(fundSel.value)||''; recalc(); }
  function recalc(){ const nav=num(navIn.value),u=num(unitsIn.value); priceOut.value= u? (nav/u).toFixed(4):''; }
  fundSel.addEventListener('change',()=>{unitsIn.value='';syncUnits();});
  [navIn,unitsIn].forEach(x=>x.addEventListener('input',recalc));
  const body=el('div',{class:'form-grid'},
    wrapField('Сан','Fund',fundSel,true),
    wrapField('Огноо','Date',dateIn,true),
    wrapField('Нийт цэвэр хөрөнгө','Total NAV (₮)',navIn,true),
    wrapField('Гүйлгээнд буй нэгж','Units outstanding',unitsIn),
    wrapField('Нэгж эрхийн үнэ','Unit price (auto)',priceOut),
    wrapField('Эх сурвалж','Source',srcIn),
    wrapField('Тэмдэглэл','Notes',noteIn),
  );
  syncUnits();
  modal('НЦХҮ бүртгэх', body, {okText:'Бүртгэх', onOk:()=>{
    if(!fundSel.value){toast('Сан сонгоно уу.','err');return false;}
    const nav=num(navIn.value),u=num(unitsIn.value);
    const rec={NavID:nextId('NAV',state.data.NAVHistory,'NavID'),Date:dateIn.value,FundID:fundSel.value,
      TotalNAV:nav,UnitsOutstanding:u,UnitPrice:u?nav/u:0,Source:srcIn.value,Notes:noteIn.value};
    commit(()=>state.data.NAVHistory.push(rec),'НЦХҮ бүртгэгдлээ');
  }});
}

/* ============================================================
   VIEW: Fees — management & performance fee calculation
   ============================================================ */
VIEWS.fees=function(){
  const wrap=el('div',{});
  wrap.append(viewHead('Шимтгэлийн тооцоо','Management & performance fees',
    el('button',{class:'btn btn-primary',onclick:()=>feeForm()},'＋ Шимтгэл тооцох')));

  wrap.append(el('div',{class:'panel'},el('div',{class:'panel-body'},
    el('div',{class:'def-list',style:'grid-template-columns:auto 1fr'},
      el('dt',{},'Удирдлагын төлбөр'),
      el('dd',{class:'small',style:'font-weight:400'},'= Өмнөх жилийн цэвэр хөрөнгө × Удирдлагын шимтгэл % (Тогтмол Орлого 1.5%, Урбан Скай 1.5–2.5%)'),
      el('dt',{},'Гүйцэтгэлийн төлбөр'),
      el('dd',{class:'small',style:'font-weight:400'},'Цэвэр хөрөнгийн өсөлт нь Зорилтот өгөөжөөс давсан тохиолдолд, давсан өсөлтийн 20%. ',
        el('strong',{},'Зорилтот өгөөж'),'-ийг гараар оруулна (ж: Монголбанкны бодлогын хүү + 4%).'))) ));

  const rows=[...state.data.Fees].sort((a,b)=>String(b.CalcDate).localeCompare(String(a.CalcDate)));
  const panel=el('div',{class:'panel'});
  panel.append(el('div',{class:'panel-head'},el('h3',{},'Тооцооны түүх · Fee calculation history')));
  const t=el('table',{class:'grid'});
  t.innerHTML=`<thead><tr><th>Огноо</th><th>Сан</th><th>Хугацаа</th><th class="num">Эхний ЦХ</th><th class="num">Эцсийн ЦХ</th>
    <th class="num">Удирдлага</th><th class="num">Гүйцэтгэл</th><th class="num">Нийт шимтгэл</th><th></th></tr></thead>`;
  const tb=el('tbody');
  if(!rows.length) tb.append(emptyRow(9,'Тооцоо хийгдээгүй байна. «Шимтгэл тооцох» дарна уу.'));
  for(const r of rows) tb.append(el('tr',{},
    el('td',{},fmtDate(r.CalcDate)), el('td',{},fundChip(r.FundID)), el('td',{},r.Period||''),
    el('td',{class:'num'},'₮'+fmtMoney(r.OpeningNAV,0)), el('td',{class:'num'},'₮'+fmtMoney(r.ClosingNAV,0)),
    el('td',{class:'num'},'₮'+fmtMoney(r.MgmtFee,0)), el('td',{class:'num'},'₮'+fmtMoney(r.PerfFee,0)),
    el('td',{class:'num'},el('strong',{},'₮'+fmtMoney(r.TotalFee,0))),
    el('td',{},el('button',{class:'btn btn-ghost btn-sm',onclick:()=>deleteFee(r.FeeID)},'Устгах'))));
  t.append(tb); panel.append(el('div',{class:'table-wrap'},t)); wrap.append(panel);
  return wrap;
};
function deleteFee(id){
  confirmModal('Тооцоо устгах','Энэ шимтгэлийн тооцоог устгах уу?',()=>
    commit(()=>{ state.data.Fees=state.data.Fees.filter(f=>f.FeeID!==id); },'Тооцоо устгагдлаа'));
}
function feeForm(){
  const fundSel=selectEl('FundID',investableFunds().map(f=>({v:f.FundID,t:f.ShortName})));
  const periodIn=el('input',{type:'text',value:String(new Date().getFullYear()-1),placeholder:'ж: 2025'});
  const openIn =el('input',{type:'number',step:'0.01',min:'0',placeholder:'Өмнөх жилийн эцсийн цэвэр хөрөнгө'});
  const closeIn=el('input',{type:'number',step:'0.01',min:'0',placeholder:'Тайлант жилийн цэвэр хөрөнгө'});
  const mgmtIn =el('input',{type:'number',step:'0.01',min:'0'});
  const tgtIn  =el('input',{type:'number',step:'0.01',min:'0',placeholder:'ж: 16 (= бодлогын хүү + 4%)'});
  const perfIn =el('input',{type:'number',step:'0.01',min:'0',value:'20'});
  const noteIn =el('input',{type:'text'});
  const calc=el('div',{class:'calc-box'});

  function onFund(){
    const f=fundById(fundSel.value); if(!f){calc.innerHTML='';return;}
    mgmtIn.value=num(f.MgmtFeePct)||'';
    // auto-pull NAVs: opening = NAV as of prior year-end, closing = latest
    const yr=parseInt(periodIn.value,10);
    if(yr) openIn.value=navAsOf(f.FundID, yr+'-12-31')||'';
    closeIn.value=latestTotalNAV(f.FundID)||'';
    recalc();
  }
  function recalc(){
    const opening=num(openIn.value), closing=num(closeIn.value);
    const mgmtPct=num(mgmtIn.value), tgtPct=num(tgtIn.value), perfPct=num(perfIn.value);
    const mgmtFee=opening*mgmtPct/100;
    const navGrowth=closing-opening;
    const targetGrowth=opening*tgtPct/100;
    const excess=navGrowth-targetGrowth;
    const perfFee= excess>0 ? excess*perfPct/100 : 0;
    const growthPct= opening? navGrowth/opening*100 : 0;
    calc.innerHTML='';
    calc.append(
      crow('Цэвэр хөрөнгийн өсөлт', '₮'+fmtMoney(navGrowth,0)+'  ('+growthPct.toFixed(2)+'%)'),
      crow('Зорилтот өсөлт ('+ (tgtPct||0) +'%)', '₮'+fmtMoney(targetGrowth,0)),
      crow('Давсан өсөлт', (excess>0?'₮':'−₮')+fmtMoney(Math.abs(excess),0)),
      el('div',{class:'calc-row',style:'border-top:1px dashed #cfe2fb;margin-top:6px;padding-top:8px'},
        el('span',{},'Удирдлагын төлбөр ('+(mgmtPct||0)+'%)'), el('span',{},'₮'+fmtMoney(mgmtFee,0))),
      crow('Гүйцэтгэлийн төлбөр ('+(perfPct||0)+'% × давсан)', '₮'+fmtMoney(perfFee,0)),
      crow('Нийт шимтгэл', '₮'+fmtMoney(mgmtFee+perfFee,0), true),
    );
  }
  fundSel.addEventListener('change',onFund);
  [periodIn].forEach(x=>x.addEventListener('input',onFund));
  [openIn,closeIn,mgmtIn,tgtIn,perfIn].forEach(x=>x.addEventListener('input',recalc));

  const body=el('div',{},
    el('div',{class:'form-grid'},
      wrapField('Сан','Fund',fundSel,true),
      wrapField('Тайлант хугацаа','Period',periodIn),
      wrapField('Өмнөх жилийн цэвэр хөрөнгө','Opening NAV (₮)',openIn,true),
      wrapField('Тайлант цэвэр хөрөнгө','Closing NAV (₮)',closeIn,true),
      wrapField('Удирдлагын шимтгэл %','Mgmt fee %',mgmtIn,true),
      wrapField('Зорилтот өгөөж %','Target return %',tgtIn),
      wrapField('Гүйцэтгэлийн шимтгэл %','Perf fee %',perfIn),
      wrapField('Тэмдэглэл','Notes',noteIn),
    ),
    el('div',{class:'small muted',style:'margin:4px 2px'},'НЦХҮ түүхээс цэвэр хөрөнгийн дүн автоматаар бөглөгдөнө — шаардвал гараар засаарай.'),
    calc);
  onFund();
  modal('Шимтгэл тооцох', body, {okText:'Хадгалах', wide:true, onOk:()=>{
    if(!fundSel.value){toast('Сан сонгоно уу.','err');return false;}
    const opening=num(openIn.value), closing=num(closeIn.value);
    const mgmtPct=num(mgmtIn.value), tgtPct=num(tgtIn.value), perfPct=num(perfIn.value);
    const mgmtFee=opening*mgmtPct/100, navGrowth=closing-opening, targetGrowth=opening*tgtPct/100;
    const excess=navGrowth-targetGrowth, perfFee= excess>0? excess*perfPct/100 : 0;
    const rec={FeeID:nextId('FEE',state.data.Fees,'FeeID'),CalcDate:todayISO(),FundID:fundSel.value,Period:periodIn.value,
      OpeningNAV:opening,ClosingNAV:closing,MgmtFeePct:mgmtPct,MgmtFee:mgmtFee,TargetReturnPct:tgtPct,
      TargetGrowth:targetGrowth,NavGrowth:navGrowth,ExcessGrowth:excess>0?excess:0,PerfFeePct:perfPct,PerfFee:perfFee,
      TotalFee:mgmtFee+perfFee,Status:'Тооцоолсон',Notes:noteIn.value};
    commit(()=>state.data.Fees.push(rec),'Шимтгэлийн тооцоо хадгалагдлаа');
  }});
}

/* ============================================================
   VIEW: ECIF — Employee Contribution Investment Fund
   Each firm = one nominal holder in Stable Growth Fund (SGF).
   Employees are sub-accounts; their contributions buy SGF units.
   ============================================================ */
VIEWS.ecif=function(){
  const wrap=el('div',{});
  wrap.append(viewHead('ECIF — Ажилчдын хуримтлалын сан','Employee Contribution Investment Fund',
    el('div',{},
      el('button',{class:'btn btn-outline btn-sm',onclick:()=>contributionForm()},'＋ Хуримтлал бүртгэх'),' ',
      el('button',{class:'btn btn-primary btn-sm',onclick:()=>firmForm()},'＋ Байгууллага нэмэх'))));

  const tgt=fundById(ECIF_TARGET_FUND);
  wrap.append(el('div',{class:'panel'},el('div',{class:'panel-body small muted'},
    'Оролцогч байгууллага бүр ', el('strong',{},tgt?tgt.ShortName:'SGF'),
    ' санд нэг ', el('strong',{},'нэрлэсэн данс эзэмшигч'), ' болно. Ажилтнуудын хуримтлал тухайн байгууллагын дансаар дамжин ',
    (tgt?tgt.ShortName:'SGF'),'-ийн нэгж эрх худалдан авна. Нэгж эрхийн одоогийн үнэ: ',
    el('strong',{},'₮'+fmtMoney(latestUnitPrice(ECIF_TARGET_FUND),2)),'.')));

  // KPIs
  wrap.append(el('div',{class:'cards'},
    kpi('Оролцогч байгууллага',fmtInt(state.data.ECIFFirms.length),''),
    kpi('Нийт ажилтан',fmtInt(ecifActiveEmployees()),'идэвхтэй'),
    kpi('Нийт хуримтлал','₮'+fmtMoney(state.data.ECIFContributions.filter(c=>num(c.TotalAmount)>0).reduce((s,c)=>s+num(c.TotalAmount),0),0),'оруулсан'),
    kpi('Одоогийн үнэлгээ','₮'+fmtMoney(ecifTotalAUM(),0),fmtUnits(ecifTotalUnits())+' нэгж эрх')));

  const panel=el('div',{class:'panel'});
  panel.append(el('div',{class:'panel-head'},el('h3',{},'Оролцогч байгууллагууд · Member firms')));
  const t=el('table',{class:'grid'});
  t.innerHTML=`<thead><tr><th>ID</th><th>Байгууллага</th><th>Регистр</th><th class="num">Ажилтан</th>
    <th class="num">Хуримтлал</th><th class="num">Нэгж эрх (SGF)</th><th class="num">Үнэлгээ</th><th>Төлөв</th><th></th></tr></thead>`;
  const tb=el('tbody');
  if(!state.data.ECIFFirms.length) tb.append(emptyRow(9,'Оролцогч байгууллага алга. «Байгууллага нэмэх» дарж эхэлнэ үү.'));
  for(const f of state.data.ECIFFirms){
    const u=ecifFirmUnits(f.FirmID);
    tb.append(el('tr',{},
      el('td',{},f.FirmID),
      el('td',{},el('strong',{},f.NameMN||''), f.NameEN?el('div',{class:'small muted'},f.NameEN):null),
      el('td',{},f.RegNo||''),
      el('td',{class:'num'},fmtInt(ecifFirmEmployees(f.FirmID).filter(e=>e.Status!=='Гарсан').length)),
      el('td',{class:'num'},'₮'+fmtMoney(ecifFirmContributed(f.FirmID),0)),
      el('td',{class:'num'},fmtUnits(u)),
      el('td',{class:'num'},'₮'+fmtMoney(ecifValue(u),0)),
      el('td',{},statusBadge(f.Status||'Идэвхтэй')),
      el('td',{},el('div',{class:'row-actions'},
        el('button',{class:'btn btn-ghost btn-sm',onclick:()=>firmDetail(f.FirmID)},'Дэлгэрэнгүй')))));
  }
  t.append(tb); panel.append(el('div',{class:'table-wrap'},t)); wrap.append(panel);
  return wrap;
};

function firmForm(existing){
  const isEdit=!!existing;
  const f=existing||{FirmID:nextId('FIRM',state.data.ECIFFirms,'FirmID'),JoinDate:todayISO(),PortfolioVariant:'Хувилбар 1 (Бонд 40 / Хувьцаа 45 / Альт 15)',Status:'Идэвхтэй'};
  const g=(name,mn,en,opts={})=>fieldHTML(name,mn,en,f[name],opts);
  const body=el('div',{},el('div',{class:'form-grid'},
    g('FirmID','Дугаар','ID',{readonly:true}),
    g('JoinDate','Элссэн огноо','Join date',{type:'date'}),
    g('NameMN','Байгууллагын нэр (Монгол)','Name (MN)',{required:true,full:true}),
    g('NameEN','Нэр (English)','Name (EN)',{full:true}),
    g('RegNo','Улсын бүртгэлийн дугаар','Registration no.'),
    g('EmployeeCount','Нийт ажилтан (тоо)','Headcount',{type:'number'}),
    g('PortfolioVariant','Багцын хувилбар','Portfolio variant',{type:'select',
      options:['Хувилбар 1 (Бонд 40 / Хувьцаа 45 / Альт 15)','Хувилбар 2 (Бонд 85 / Хувьцаа 15)'],full:true}),
    g('ContactPerson','Холбоо барих хүн','Contact'),
    g('Phone','Утас','Phone'),
    g('Email','И-мэйл','Email',{type:'email'}),
    g('BankName','Банк','Bank'),
    g('BankAccount','Дансны дугаар','Account no.'),
    g('Address','Хаяг','Address',{full:true}),
    g('Status','Төлөв','Status',{type:'select',options:INV_STATUS}),
    g('Notes','Тэмдэглэл','Notes',{type:'textarea',full:true}),
  ));
  modal(isEdit?'Байгууллага засах':'Шинэ оролцогч байгууллага', body,
    {okText:isEdit?'Хадгалах':'Бүртгэх', wide:true, onOk:()=>{
      const rec=readForm(body,SCHEMA.ECIFFirms,f);
      if(!rec.NameMN){ toast('Байгууллагын нэр заавал.','err'); return false; }
      commit(()=>{
        if(isEdit){
          Object.assign(existing,rec);
          const inv=investorById(existing.InvestorID);
          if(inv){ inv.NameMN=rec.NameMN+' (ECIF)'; inv.NameEN=rec.NameEN; inv.RegNo=rec.RegNo;
            inv.Phone=rec.Phone; inv.Email=rec.Email; inv.Address=rec.Address;
            inv.BankName=rec.BankName; inv.BankAccount=rec.BankAccount; }
        } else {
          // auto-create the nominal SGF investor for this firm
          const invId=nextId('INV',state.data.Investors,'InvestorID');
          state.data.Investors.push({InvestorID:invId,RegDate:rec.JoinDate,InvestorType:'ECIF нэрлэсэн данс',
            NameMN:rec.NameMN+' (ECIF)',NameEN:rec.NameEN,RegNo:rec.RegNo,ContactPerson:rec.ContactPerson,
            Phone:rec.Phone,Email:rec.Email,Address:rec.Address,BankName:rec.BankName,BankAccount:rec.BankAccount,
            AccountName:rec.NameMN,Status:'Идэвхтэй',Notes:'ECIF нэрлэсэн данс — '+rec.FirmID});
          rec.InvestorID=invId;
          state.data.ECIFFirms.push(rec);
        }
      }, isEdit?'Байгууллага шинэчлэгдлээ':'Байгууллага бүртгэгдэж, SGF-д нэрлэсэн данс үүслээ');
    }});
}

function employeeForm(firmId, existing){
  const isEdit=!!existing;
  const firm=firmById(firmId||(existing&&existing.FirmID));
  const e=existing||{EmployeeID:nextId('EMP',state.data.ECIFEmployees,'EmployeeID'),FirmID:firm.FirmID,JoinDate:todayISO(),Status:'Идэвхтэй'};
  const g=(name,mn,en,opts={})=>fieldHTML(name,mn,en,e[name],opts);
  const body=el('div',{},
    el('div',{class:'small muted',style:'margin-bottom:10px'},'Байгууллага: ',el('strong',{},firm.NameMN)),
    el('div',{class:'form-grid'},
      g('EmployeeID','Дугаар','ID',{readonly:true}),
      g('JoinDate','Элссэн огноо','Join date',{type:'date'}),
      g('NameMN','Ажилтны нэр','Employee name',{required:true,full:true}),
      g('RegNo','Регистрийн дугаар','Reg. no.'),
      g('Position','Албан тушаал','Position'),
      g('MonthlySalary','Сарын цалин','Monthly salary',{type:'number'}),
      g('Status','Төлөв','Status',{type:'select',options:['Идэвхтэй','Гарсан']}),
      g('Notes','Тэмдэглэл','Notes',{type:'textarea',full:true}),
    ));
  modal(isEdit?'Ажилтан засах':'Шинэ ажилтан', body, {okText:isEdit?'Хадгалах':'Бүртгэх', onOk:()=>{
    const rec=readForm(body,SCHEMA.ECIFEmployees,e);
    if(!rec.NameMN){ toast('Ажилтны нэр заавал.','err'); return false; }
    commit(()=>{ if(isEdit) Object.assign(existing,rec); else state.data.ECIFEmployees.push(rec); },
      isEdit?'Ажилтан шинэчлэгдлээ':'Ажилтан бүртгэгдлээ');
    firmDetail(firm.FirmID); return false;
  }});
}

function firmDetail(firmId){
  const firm=firmById(firmId); if(!firm)return;
  const body=el('div',{});
  body.append(el('h3',{style:'margin:0 0 2px;color:var(--gam-navy)'},firm.NameMN),
    el('div',{class:'muted small',style:'margin-bottom:10px'},
      (firm.NameEN||'')+' · '+(firm.FirmID)+' · '+(firm.PortfolioVariant||'')));
  const u=ecifFirmUnits(firmId);
  body.append(el('div',{class:'cards',style:'margin-bottom:6px'},
    kpi('Ажилтан',fmtInt(ecifFirmEmployees(firmId).length),''),
    kpi('Хуримтлал','₮'+fmtMoney(ecifFirmContributed(firmId),0),''),
    kpi('SGF нэгж эрх',fmtUnits(u),''),
    kpi('Үнэлгээ','₮'+fmtMoney(ecifValue(u),0),'')));

  body.append(el('div',{style:'display:flex;gap:8px;margin:10px 0'},
    el('button',{class:'btn btn-outline btn-sm',onclick:()=>employeeForm(firmId)},'＋ Ажилтан нэмэх'),
    el('button',{class:'btn btn-outline btn-sm',onclick:()=>contributionForm(firmId)},'＋ Хуримтлал бүртгэх'),
    el('button',{class:'btn btn-ghost btn-sm',onclick:()=>firmForm(firm)},'Байгууллага засах')));

  // employees
  body.append(el('h4',{style:'margin:8px 0 6px;color:var(--gam-navy)'},'Ажилтнууд · Employees'));
  const te=el('table',{class:'grid'});
  te.innerHTML=`<thead><tr><th>ID</th><th>Нэр</th><th>Албан тушаал</th><th class="num">Хуримтлал</th>
    <th class="num">Нэгж эрх</th><th class="num">Үнэлгээ</th><th>Төлөв</th><th></th></tr></thead>`;
  const tbe=el('tbody');
  const emps=ecifFirmEmployees(firmId);
  if(!emps.length) tbe.append(emptyRow(8,'Ажилтан бүртгэгдээгүй.'));
  for(const e of emps){
    const eu=ecifEmployeeUnits(e.EmployeeID);
    tbe.append(el('tr',{},
      el('td',{},e.EmployeeID), el('td',{},e.NameMN||''), el('td',{},e.Position||''),
      el('td',{class:'num'},'₮'+fmtMoney(ecifEmployeeContributed(e.EmployeeID),0)),
      el('td',{class:'num'},fmtUnits(eu)),
      el('td',{class:'num'},'₮'+fmtMoney(ecifValue(eu),0)),
      el('td',{},statusBadge(e.Status||'Идэвхтэй')),
      el('td',{},el('div',{class:'row-actions'},
        eu>0.00001?el('button',{class:'btn btn-ghost btn-sm',onclick:()=>withdrawalForm(e.EmployeeID)},'Татан авах'):null,
        el('button',{class:'btn btn-ghost btn-sm',onclick:()=>employeeForm(null,e)},'Засах')))));
  }
  te.append(tbe); body.append(el('div',{class:'table-wrap'},te));

  // contribution history
  const contribs=state.data.ECIFContributions.filter(c=>c.FirmID===firmId)
    .sort((a,b)=>String(b.Date).localeCompare(String(a.Date)));
  body.append(el('h4',{style:'margin:16px 0 6px;color:var(--gam-navy)'},'Хуримтлалын түүх · Contribution history'));
  const tc=el('table',{class:'grid'});
  tc.innerHTML=`<thead><tr><th>Огноо</th><th>Хугацаа</th><th>Ажилтан</th><th>Төрөл</th>
    <th class="num">Ажилтан</th><th class="num">Ажил олгогч</th><th class="num">Нийт</th><th class="num">Нэгж эрх</th></tr></thead>`;
  const tbc=el('tbody');
  if(!contribs.length) tbc.append(emptyRow(8,'Хуримтлал бүртгэгдээгүй.'));
  for(const c of contribs){
    const emp=employeeById(c.EmployeeID);
    tbc.append(el('tr',{},el('td',{},fmtDate(c.Date)),el('td',{},c.Period||''),el('td',{},emp?emp.NameMN:c.EmployeeID),
      el('td',{},el('span',{class:'badge '+(num(c.Units)<0?'badge-amber':'badge-blue')},c.Type||(num(c.Units)<0?'Татан авалт':'Хуримтлал'))),
      el('td',{class:'num'},'₮'+fmtMoney(c.EmployeeAmount,0)),el('td',{class:'num'},'₮'+fmtMoney(c.EmployerAmount,0)),
      el('td',{class:'num'},'₮'+fmtMoney(c.TotalAmount,0)),el('td',{class:'num'},fmtUnits(c.Units))));
  }
  tc.append(tbc); body.append(el('div',{class:'table-wrap'},tc));

  modal('Байгууллагын дэлгэрэнгүй', body, {okText:'Хаах', cancel:false, wide:true});
}

function contributionForm(presetFirm){
  if(!state.data.ECIFFirms.length){ toast('Эхлээд оролцогч байгууллага нэмнэ үү.','err'); return; }
  const firmSel=selectEl('FirmID',state.data.ECIFFirms.map(f=>({v:f.FirmID,t:f.NameMN})));
  if(presetFirm) firmSel.value=presetFirm;
  const dateIn=el('input',{type:'date',value:todayISO()});
  const periodIn=el('input',{type:'text',value:todayISO().slice(0,7),placeholder:'ж: 2026-06'});
  const price=latestUnitPrice(ECIF_TARGET_FUND);
  const linesHost=el('div',{});
  const totalsBox=el('div',{class:'calc-box'});

  function buildLines(){
    linesHost.innerHTML='';
    const emps=ecifFirmEmployees(firmSel.value).filter(e=>e.Status!=='Гарсан');
    if(!emps.length){ linesHost.append(el('div',{class:'small muted'},'Энэ байгууллагад идэвхтэй ажилтан алга. Эхлээд ажилтан нэмнэ үү.')); recalc(); return; }
    const tbl=el('table',{class:'grid'});
    tbl.innerHTML=`<thead><tr><th>Ажилтан</th><th class="num">Ажилтны хувь нэмэр (₮)</th><th class="num">Ажил олгогчийн хувь нэмэр (₮)</th></tr></thead>`;
    const tb=el('tbody');
    for(const e of emps){
      const empIn=el('input',{type:'number',step:'0.01',min:'0',value:'0','data-emp':e.EmployeeID,'data-role':'emp'});
      const erIn =el('input',{type:'number',step:'0.01',min:'0',value:'0','data-emp':e.EmployeeID,'data-role':'er'});
      [empIn,erIn].forEach(x=>x.addEventListener('input',recalc));
      tb.append(el('tr',{},el('td',{},e.NameMN),el('td',{class:'num'},empIn),el('td',{class:'num'},erIn)));
    }
    tbl.append(tb); linesHost.append(el('div',{class:'table-wrap'},tbl));
    recalc();
  }
  function recalc(){
    let total=0;
    linesHost.querySelectorAll('input[type=number]').forEach(i=>total+=num(i.value));
    totalsBox.innerHTML='';
    totalsBox.append(
      crow('Нэгж эрхийн үнэ (SGF)','₮'+fmtMoney(price,2)),
      crow('Нийт хувь нэмэр','₮'+fmtMoney(total,2)),
      crow('Худалдан авах нэгж эрх', price? fmtUnits(total/price):'0', true));
  }
  firmSel.addEventListener('change',buildLines);

  const body=el('div',{},
    el('div',{class:'form-grid'},
      wrapField('Байгууллага','Firm',firmSel,true),
      wrapField('Огноо','Date',dateIn),
      wrapField('Хугацаа (сар)','Period',periodIn),
    ),
    el('div',{class:'small muted',style:'margin:6px 0'},'Ажилтан бүрийн хувь нэмрийг оруулна. Нийт дүн SGF-ийн нэгж эрхийг ',
      el('strong',{},'₮'+fmtMoney(price,2)),' үнээр худалдан авч, ажилтан тус бүрд хуваарилагдана.'),
    linesHost,
    totalsBox);
  buildLines();

  modal('Хуримтлал бүртгэх', body, {okText:'Бүртгэх', wide:true, onOk:()=>{
    const firm=firmById(firmSel.value);
    if(!firm){ toast('Байгууллага сонгоно уу.','err'); return false; }
    const lines=[];
    const byEmp={};
    linesHost.querySelectorAll('input[type=number]').forEach(i=>{
      const id=i.getAttribute('data-emp'), role=i.getAttribute('data-role');
      byEmp[id]=byEmp[id]||{emp:0,er:0}; byEmp[id][role==='emp'?'emp':'er']=num(i.value);
    });
    let batchTotal=0;
    for(const[empId,v] of Object.entries(byEmp)){
      const tot=v.emp+v.er; if(tot<=0) continue;
      lines.push({empId,emp:v.emp,er:v.er,tot,units:tot/price}); batchTotal+=tot;
    }
    if(!lines.length){ toast('Дор хаяж нэг ажилтанд хувь нэмэр оруулна уу.','err'); return false; }

    commit(()=>{
      // one aggregated SGF subscription for the firm's nominal account
      const txnId=nextId('TXN',state.data.Transactions,'TxnID');
      state.data.Transactions.push({TxnID:txnId,TradeDate:dateIn.value,Type:'Худалдан авалт',FundID:ECIF_TARGET_FUND,
        InvestorID:firm.InvestorID,Units:batchTotal/price,UnitPrice:price,GrossAmount:batchTotal,FeePct:0,FeeAmount:0,
        NetAmount:batchTotal,SettlementDate:dateIn.value,Status:'Баталгаажсан',
        PaymentRef:'ECIF '+periodIn.value,Notes:'ECIF хуримтлал — '+firm.NameMN,CreatedAt:new Date().toISOString()});
      // per-employee ledger rows
      for(const ln of lines){
        state.data.ECIFContributions.push({ContribID:nextId('CTR',state.data.ECIFContributions,'ContribID'),
          Date:dateIn.value,Period:periodIn.value,FirmID:firm.FirmID,EmployeeID:ln.empId,
          EmployeeAmount:ln.emp,EmployerAmount:ln.er,TotalAmount:ln.tot,UnitPrice:price,Units:ln.units,
          Type:'Хуримтлал',TxnID:txnId,Notes:''});
      }
    }, 'Хуримтлал бүртгэгдэж, '+(batchTotal/price).toFixed(2)+' нэгж эрх худалдаж авлаа');
    if(presetFirm){ firmDetail(firm.FirmID); return false; }
  }});
}

function withdrawalForm(empId){
  const e=employeeById(empId); if(!e)return;
  const firm=firmById(e.FirmID);
  const have=ecifEmployeeUnits(empId);
  const price=latestUnitPrice(ECIF_TARGET_FUND);
  const dateIn=el('input',{type:'date',value:todayISO()});
  const unitsIn=el('input',{type:'number',step:'0.0001',min:'0',max:String(have),value:have.toFixed(4)});
  const amtOut=el('input',{type:'text',readonly:true});
  const noteIn=el('input',{type:'text'});
  function recalc(){ amtOut.value='₮'+fmtMoney(num(unitsIn.value)*price,2); }
  unitsIn.addEventListener('input',recalc); recalc();
  const body=el('div',{},
    el('div',{class:'small muted',style:'margin-bottom:8px'},'Ажилтан: ',el('strong',{},e.NameMN),
      ' · Эзэмшил: ',el('strong',{},fmtUnits(have)+' нэгж эрх')),
    el('div',{class:'form-grid'},
      wrapField('Огноо','Date',dateIn),
      wrapField('Татан авах нэгж эрх','Units to withdraw',unitsIn,true),
      wrapField('Нэгж эрхийн үнэ','Unit price',el('input',{type:'text',readonly:true,value:'₮'+fmtMoney(price,2)})),
      wrapField('Олгох дүн','Payout amount',amtOut),
      wrapField('Тэмдэглэл','Notes',noteIn),
    ));
  modal('Хуримтлал татан авах', body, {okText:'Татан авах', onOk:()=>{
    const units=num(unitsIn.value);
    if(units<=0){ toast('Нэгж эрх оруулна уу.','err'); return false; }
    if(units>have+0.00001){ toast('Эзэмшлээс хэтэрсэн. Боломжит: '+fmtUnits(have),'err'); return false; }
    const amt=units*price;
    commit(()=>{
      const txnId=nextId('TXN',state.data.Transactions,'TxnID');
      state.data.Transactions.push({TxnID:txnId,TradeDate:dateIn.value,Type:'Буцаан худалдалт',FundID:ECIF_TARGET_FUND,
        InvestorID:firm.InvestorID,Units:units,UnitPrice:price,GrossAmount:amt,FeePct:0,FeeAmount:0,NetAmount:amt,
        SettlementDate:dateIn.value,Status:'Баталгаажсан',PaymentRef:'ECIF татан авалт',
        Notes:'ECIF татан авалт — '+e.NameMN,CreatedAt:new Date().toISOString()});
      state.data.ECIFContributions.push({ContribID:nextId('CTR',state.data.ECIFContributions,'ContribID'),
        Date:dateIn.value,Period:dateIn.value.slice(0,7),FirmID:firm.FirmID,EmployeeID:empId,
        EmployeeAmount:-amt,EmployerAmount:0,TotalAmount:-amt,UnitPrice:price,Units:-units,
        Type:'Татан авалт',TxnID:txnId,Notes:noteIn.value});
    }, 'Татан авалт бүртгэгдлээ (₮'+fmtMoney(amt,0)+')');
    firmDetail(firm.FirmID); return false;
  }});
}

/* ============================================================
   Generic form helpers (modal forms)
   ============================================================ */
function fieldHTML(name,mn,en,value,opts={}){
  const id='f_'+name;
  let input;
  if(opts.type==='select'){
    input=el('select',{id,name});
    for(const o of (opts.options||[])) input.append(el('option',{value:o,selected:o===value?'':null},o));
    if(value && !opts.options.includes(value)) input.append(el('option',{value,selected:''},value));
  }else if(opts.type==='textarea'){
    input=el('textarea',{id,name,rows:'2'}); input.value=value??'';
  }else{
    input=el('input',{id,name,type:opts.type||'text',value:value??''});
    if(opts.readonly) input.readOnly=true;
  }
  const fieldCls='field'+(opts.full?' full':'');
  return el('div',{class:fieldCls},
    el('label',{for:id}, mn+' ', el('span',{class:'en'},'· '+en), opts.required?el('span',{class:'req'},' *'):''),
    input,
    opts.hint?el('span',{class:'hint'},opts.hint):null);
}
function readForm(root,cols,base){
  const rec={...base};
  for(const c of cols){
    const inp=root.querySelector(`[name="${c}"]`);
    if(inp) rec[c]=inp.value;
  }
  return rec;
}

/* ============================================================
   Modal
   ============================================================ */
function modal(title, bodyNode, opts={}){
  const backdrop=$('#modalBackdrop'), host=$('#modal');
  host.className='modal'+(opts.wide?' wide':'');
  host.innerHTML='';
  host.append(
    el('div',{class:'modal-head'},el('h3',{},title),el('button',{class:'modal-close',onclick:close},'×')),
    el('div',{class:'modal-body'},bodyNode),
    el('div',{class:'modal-foot'},
      opts.cancel===false?null:el('button',{class:'btn btn-ghost',onclick:close},'Цуцлах'),
      el('button',{class:'btn btn-primary',onclick:async()=>{ if(opts.onOk){ const r=await opts.onOk(); if(r===false)return; } close(); }},
        opts.okText||'OK')),
  );
  backdrop.classList.remove('hidden');
  function close(){ backdrop.classList.add('hidden'); }
  backdrop.onclick=e=>{ if(e.target===backdrop) close(); };
}
function confirmModal(title,msg,onYes){
  modal(title, el('p',{},msg), {okText:'Тийм', onOk:onYes});
}

/* ============================================================
   Init
   ============================================================ */
async function init(){
  $('#btnSignIn').onclick = signIn;
  $('#btnSave').onclick   = ()=>saveToSharePoint(false);
  $('#btnReload').onclick = reloadFromSharePoint;
  $('#btnSignOut').onclick= signOut;
  $$('.nav-item').forEach(b=>b.onclick=()=>go(b.dataset.view));

  window.addEventListener('beforeunload',e=>{ if(state.dirty){ e.preventDefault(); e.returnValue=''; } });

  // handle redirect return (login or token refresh) or restore a cached session
  try{
    const result = await msalApp.handleRedirectPromise();
    const account = result ? result.account : (msalApp.getAllAccounts()[0] || null);
    if(account){
      msalApp.setActiveAccount(account);
      state.account = account;
      updateTopBarUser();
      gateInfo('SharePoint-аас файлыг хайж байна…');
      await findOrCreateWorkbook();
    }
  }catch(e){ console.warn('MSAL restore:', e); }
}
document.addEventListener('DOMContentLoaded',init);
