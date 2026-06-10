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
          'TargetRaise','AuthorizedUnits','NominalPrice','MgmtFeePct','PerfFeePct','HurdleRatePct',
          'Custodian','ManagementCo','Status','Color','RoutesTo'],
  Investors: ['InvestorID','RegDate','InvestorType','Residency','NameMN','NameEN','RegNo','ContactPerson',
              'Phone','Email','Address','BankName','BankAccount','AccountName','Status','Notes'],
  Transactions: ['TxnID','TradeDate','Type','FundID','InvestorID','Units','UnitPrice','GrossAmount',
                 'FeePct','FeeAmount','NetAmount','SettlementDate','Status','PaymentRef','Notes','CreatedAt'],
  NAVHistory: ['NavID','Date','FundID','TotalNAV','UnitsOutstanding','UnitPrice','Source','Notes'],
  Bonds: ['BondID','IssuerMN','IssuerEN','FundID','BondType','FaceValue','CouponRate','CouponFreq',
          'IssueDate','MaturityDate','PurchasePrice','PurchaseDate','MarketPrice','MarketPriceDate',
          'Currency','YTM','AccruedInterest','MarketValue','UnrealizedPL','Status','Notes'],
  Instruments: ['InsID','InsType','NameMN','NameEN','Issuer','ISIN','Currency','FundID','Status',
          'PurchaseDate','PurchasePrice','Quantity','CurrentPrice','CurrentPriceDate',
          'FaceValue','CouponRate','CouponFreq','IssueDate','MaturityDate','YTM','AccruedInterest',
          'Exchange','Ticker','DividendYield','FundManager','FundUnitType',
          'MarketValue','UnrealizedPL','GainPct','Notes'],
  PriceHistory: ['PriceID','InsID','Date','Price','Source','Notes'],
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
    NominalPrice:100, MgmtFeePct:1.5, PerfFeePct:20, HurdleRatePct:8,
    Custodian:'Голомт банк — Кастодианы үйлчилгээний газар', ManagementCo:'«Голомт Ассет Менежмент ҮЦК» ХХК',
    Status:'Идэвхтэй', Color:'#1366c4'
  },
  {
    FundID:'USF', NameMN:'«Урбан Скай» Хувийн ХОС', NameEN:'Urban Sky Private Fund',
    ShortName:'Урбан Скай', Type:'Хувийн (хаалттай) хөрөнгө оруулалтын сан', Currency:'MNT',
    RegisteredDate:'2024-07-25', TermYears:10, TargetRaise:150000000000, AuthorizedUnits:150000,
    NominalPrice:1000000, MgmtFeePct:2.5, PerfFeePct:20, HurdleRatePct:8,
    Custodian:'Голомт банк — Кастодианы үйлчилгээний газар', ManagementCo:'«Голомт Ассет Менежмент ҮЦК» ХХК',
    Status:'Идэвхтэй', Color:'#c8a14b'
  },
  {
    FundID:'ECIF', NameMN:'«Ажилтны хуримтлалын сан» Хувийн ХОС', NameEN:'Employee Contribution Investment Fund',
    ShortName:'ECIF', Type:'Ажилчдын хуримтлалын хөтөлбөр — Тогтмол Орлого сангаар дамжина', Currency:'MNT',
    RegisteredDate:'2023-12-28', TermYears:10, TargetRaise:100000000000, AuthorizedUnits:1000000000,
    NominalPrice:100, MgmtFeePct:1.5, PerfFeePct:20, HurdleRatePct:8,
    Custodian:'Голомт банк — Кастодианы үйлчилгээний газар', ManagementCo:'«Голомт Ассет Менежмент ҮЦК» ХХК',
    Status:'Идэвхтэй', Color:'#1c8a4a', RoutesTo:'SGF'
  },
];

const TXN_STATUS = ['Хүлээгдэж буй','Баталгаажсан','Цуцлагдсан']; // Pending / Confirmed / Cancelled
const INV_TYPES  = ['Хувь хүн','Хуулийн этгээд'];                 // Individual / Company
const INV_STATUS = ['Идэвхтэй','Идэвхгүй'];
const INV_RESIDENCY = ['Дотоодын (Монгол)','Гадаадын'];          // Domestic (Mongolian) / Foreign

/* ---------- Transaction types ----------
   Trades move units (Subscription / Redemption). Fee & tax rows are
   booked as separate non-unit lines so the ledger shows every cash flow
   without disturbing unit / AUM math. */
const TXN_TYPE = {
  SUB:     'Худалдан авалт',        // Subscription
  RED:     'Буцаан худалдалт',      // Redemption
  MGMT:    'Удирдлагын шимтгэл',    // Management fee (taken upfront on subscription)
  SUCCESS: 'Гүйцэтгэлийн шимтгэл',  // Success / performance fee (after 1 year, above hurdle)
  TAX:     'Татвар',                // Withholding tax on returns
};
function isTradeType(t){ return t===TXN_TYPE.SUB || t===TXN_TYPE.RED; }
/** withholding tax rate: foreign investors 20%, domestic (Mongolian) 10% */
function investorTaxRate(inv){
  return /гадаад|foreign/i.test(String(inv&&inv.Residency||'')) ? 0.20 : 0.10;
}
/** true if `dateStr` is at least one year before `asOf` (default today) */
function isOneYearOld(dateStr, asOf){
  if(!dateStr) return false;
  const d=new Date(dateStr); if(isNaN(d)) return false;
  d.setFullYear(d.getFullYear()+1);
  return d <= new Date(asOf||todayISO());
}
/** sum NetAmount of a given fee/tax transaction type, optionally filtered by investor & fund */
function sumTxnType(type, investorId, fundId){
  return state.data.Transactions
    .filter(t=>t.Type===type && t.Status!=='Цуцлагдсан'
      && (investorId==null||t.InvestorID===investorId)
      && (fundId==null||t.FundID===fundId))
    .reduce((s,t)=>s+num(t.NetAmount),0);
}

/* ============================================================
   State
   ============================================================ */
const state = {
  account: null,
  data: { Funds:[], Investors:[], Transactions:[], NAVHistory:[], Bonds:[], Instruments:[], PriceHistory:[], Fees:[], ECIFFirms:[], ECIFEmployees:[], ECIFContributions:[], Meta:[] },
  dirty: false,
  view: 'dashboard',
  savedAt: null,
};

/* ---------- Bond reference data ---------- */
const BOND_TYPES   = ['Засгийн газрын (Government)','Корпорацийн (Corporate)','Банкны (Bank)'];
const COUPON_FREQ  = ['Сар бүр (Monthly)','Улирал бүр (Quarterly)','Жилд нэг (Annually)'];
const BOND_CCY     = ['MNT','USD','CNY','EUR','JPY'];
const BOND_STATUS  = ['Идэвхтэй','Дууссан','Зарсан']; // Active / Matured / Sold

/* ---------- Investment instrument registry reference data ---------- */
const INS_TYPES    = ['Бонд (Bond)','Хувьцаа (Equity)','Сангийн нэгж (Fund Unit)','Мөнгөний зах зээл (Money Market)','Бусад (Other)'];
const INS_STATUS   = ['Идэвхтэй (Active)','Дууссан (Matured)','Зарагдсан (Sold)'];
const INS_EXCHANGES= ['МХБ (MSE)','NYSE','NASDAQ','HKEX','Бусад (Other)'];

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
const _SP_DRIVE_PATH = SHARED_FILE_PATH;
/* Address the shared workbook through the signed-in user's own drive.
   The shared SharePoint folder is shared with the user's M365 account, so
   /me/drive resolves it — no /sites lookup (Sites.Read.All) required. */
const _DRIVE_CONTENT = '/me/drive/root:'+SHARED_FILE_PATH+':/content';
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
   SharePoint workbook operations (via /me/drive)
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
  state.data = { Funds:[], Investors:[], Transactions:[], NAVHistory:[], Bonds:[], Instruments:[], PriceHistory:[], Fees:[],
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
    const resp = await gFetch(_DRIVE_CONTENT);            // GET /me/drive/root:<path>:/content
    loadWorkbookBuffer(await resp.arrayBuffer());
    enterApp();
    toast('SharePoint-аас файл ачааллаа.','ok','Холбогдлоо');
  }catch(e){
    if(e.status===404){ await createWorkbookOnSharePoint(); return; }   // file not there yet → create it
    console.error('[GAM] findOrCreateWorkbook:', e);
    gateError('SharePoint алдаа: '+(e.message||String(e)));
  }
}

async function createWorkbookOnSharePoint(){
  gateInfo(FILE_NAME+' файлыг SharePoint дээр үүсгэж байна…');
  state.data = {
    Funds: SEED_FUNDS.map(f=>({...f})),
    Investors:[], Transactions:[], NAVHistory:[], Bonds:[], Instruments:[], PriceHistory:[], Fees:[],
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
  const parts = SHARED_FILE_PATH.split('/').filter(Boolean);
  const folderParts = parts.slice(0, -1); // e.g. ['Shared Documents', 'GAM Back Office']
  if(!folderParts.length) return;
  const folderDrivePath = '/' + folderParts.join('/');
  try{
    await gFetch('/me/drive/root:'+folderDrivePath);
  }catch(e){
    if(e.status===404){
      const name = folderParts[folderParts.length-1];
      const parentDrivePath = '/' + folderParts.slice(0,-1).join('/');
      const childrenPath = parentDrivePath==='/'
        ? '/me/drive/root/children'
        : '/me/drive/root:'+parentDrivePath+':/children';
      await gFetch(childrenPath,{
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body: JSON.stringify({name, folder:{}, '@microsoft.graph.conflictBehavior':'replace'}),
      });
    }else{ throw e; }
  }
}

async function saveToSharePoint(silent){
  if(!state.account) return;
  refreshBondComputed();       // snapshot auto-computed bond metrics into the workbook
  refreshInstrumentComputed(); // snapshot auto-computed instrument metrics into the workbook
  const wb = XLSX.utils.book_new();
  for(const[sheet,cols] of Object.entries(SCHEMA)){
    const rows = (state.data[sheet]||[]).map(r=>{ const o={}; for(const c of cols) o[c]=r[c]??''; return o; });
    const ws = XLSX.utils.json_to_sheet(rows,{header:cols});
    XLSX.utils.book_append_sheet(wb,ws,sheet);
  }
  const out = XLSX.write(wb,{bookType:'xlsx',type:'array'});
  await ensureSpFolder();
  await gFetch(_DRIVE_CONTENT,{                            // PUT /me/drive/root:<path>:/content
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
    const resp = await gFetch(_DRIVE_CONTENT);            // GET /me/drive/root:<path>:/content
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
    if(t.InvestorID!==investorId||t.FundID!==fundId||!isTradeType(t.Type)) continue;
    u += t.Type===TXN_TYPE.SUB ? num(t.Units) : -num(t.Units);
  }
  return u;
}
/** total invested (net cash in) by investor in fund */
function investorNetCash(investorId, fundId){
  let c=0;
  for(const t of confirmedTxns()){
    if(t.InvestorID!==investorId||t.FundID!==fundId||!isTradeType(t.Type)) continue;
    c += t.Type===TXN_TYPE.SUB ? num(t.NetAmount) : -num(t.NetAmount);
  }
  return c;
}
/** outstanding units of a fund */
function fundUnitsOutstanding(fundId){
  let u=0;
  for(const t of confirmedTxns()){
    if(t.FundID!==fundId||!isTradeType(t.Type)) continue;
    u += t.Type===TXN_TYPE.SUB ? num(t.Units) : -num(t.Units);
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
/** unit price recorded on or before a given date (fallback nominal) */
function unitPriceAsOf(fundId,dateStr){
  const f=fundById(fundId);
  const rows=state.data.NAVHistory.filter(n=>n.FundID===fundId&&num(n.UnitPrice)>0&&String(n.Date)<=dateStr)
    .sort((a,b)=>String(b.Date).localeCompare(String(a.Date)));
  return rows.length ? num(rows[0].UnitPrice) : (f?num(f.NominalPrice):0);
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
   Bonds — analytics (all auto-computed, never entered by hand)
   Bond prices are quoted as a % of par (face value).
   ============================================================ */
function bondById(id){ return (state.data.Bonds||[]).find(b=>b.BondID===id); }
function fundBonds(fundId){ return (state.data.Bonds||[]).filter(b=>b.FundID===fundId); }
/** coupon payments per year from the frequency label */
function couponFreqPerYear(s){
  s=String(s||'').toLowerCase();
  if(s.includes('сар')||s.includes('month'))   return 12;
  if(s.includes('улирал')||s.includes('quart')) return 4;
  return 1; // annually
}
function monthsPerPeriod(s){ return Math.round(12/couponFreqPerYear(s)); }
/** add whole months to an ISO date, returning YYYY-MM-DD */
function addMonths(iso,m){ const d=new Date(iso); if(isNaN(d))return iso;
  d.setMonth(d.getMonth()+m); return d.toISOString().slice(0,10); }
function daysBetween(a,b){ const d=(new Date(b)-new Date(a))/864e5; return isFinite(d)?Math.round(d):0; }

/** market quote (% of par); falls back to purchase price when no market price yet */
function bondQuote(b){ return num(b.MarketPrice)>0 ? num(b.MarketPrice) : num(b.PurchasePrice); }
/** Market value = quote × face / 100 */
function bondMarketValue(b){ return bondQuote(b)*num(b.FaceValue)/100; }
/** Cost = purchase price × face / 100 */
function bondCostValue(b){ return num(b.PurchasePrice)*num(b.FaceValue)/100; }
/** Unrealized gain/loss = market value − cost */
function bondUnrealizedPL(b){ return bondMarketValue(b)-bondCostValue(b); }
/** coupon amount paid each period */
function bondPeriodCoupon(b){ const f=couponFreqPerYear(b.CouponFreq); return f? (num(b.CouponRate)/100)*num(b.FaceValue)/f : 0; }

/** all coupon payment dates from first coupon after issue through maturity (inclusive) */
function bondCouponDates(b){
  const mpp=monthsPerPeriod(b.CouponFreq);
  if(!b.IssueDate||!b.MaturityDate||!mpp) return [];
  const out=[]; let d=addMonths(b.IssueDate,mpp), guard=0;
  while(d<b.MaturityDate && guard<4000){ out.push(d); d=addMonths(d,mpp); guard++; }
  out.push(b.MaturityDate);
  return out;
}

/** Accrued interest = period coupon × (days since last coupon ÷ days in coupon period) */
function bondAccruedInterest(b){
  const f=couponFreqPerYear(b.CouponFreq), mpp=monthsPerPeriod(b.CouponFreq);
  if(!f||!mpp||!b.IssueDate||!num(b.FaceValue)) return 0;
  const asOf=b.MarketPriceDate||todayISO();
  if(asOf<=b.IssueDate) return 0;
  if(b.MaturityDate && asOf>=b.MaturityDate) return 0; // redeemed
  let last=b.IssueDate, next=addMonths(b.IssueDate,mpp), guard=0;
  while(next<=asOf && guard<4000){ last=next; next=addMonths(next,mpp); guard++; }
  const daysInPeriod=daysBetween(last,next)||1;
  const daysSince=Math.max(0,daysBetween(last,asOf));
  return bondPeriodCoupon(b)*(daysSince/daysInPeriod);
}

/** Yield to maturity (annual %), solved from the current market quote by bisection */
function bondYTM(b){
  const face=num(b.FaceValue), f=couponFreqPerYear(b.CouponFreq), quote=bondQuote(b);
  if(!face||!f||quote<=0||!b.MaturityDate) return 0;
  const asOf=b.MarketPriceDate||todayISO();
  const years=daysBetween(asOf,b.MaturityDate)/365.25;
  if(years<=0) return 0;
  const n=Math.max(1,Math.round(years*f));     // remaining coupon periods
  const price=quote/100*face;                   // present value (clean)
  const c=bondPeriodCoupon(b);                   // coupon per period
  const pv=i=>{ let s=0; for(let k=1;k<=n;k++) s+=c/Math.pow(1+i,k); return s+face/Math.pow(1+i,n)-price; };
  let lo=-0.9999, hi=1, flo=pv(lo), fhi=pv(hi), guard=0;
  while(flo*fhi>0 && hi<1000){ hi*=2; fhi=pv(hi); if(++guard>60) break; }
  if(flo*fhi>0) return 0;                         // no sign change → cannot solve
  for(let k=0;k<200;k++){ const mid=(lo+hi)/2, fm=pv(mid);
    if(Math.abs(fm)<1e-7){ lo=hi=mid; break; }
    if(flo*fm<0){ hi=mid; fhi=fm; } else { lo=mid; flo=fm; } }
  return ((lo+hi)/2)*f*100;
}

/** snapshot the computed metrics onto each bond record (called before save) */
function refreshBondComputed(){
  for(const b of state.data.Bonds||[]){
    b.YTM             = +bondYTM(b).toFixed(4);
    b.AccruedInterest = +bondAccruedInterest(b).toFixed(2);
    b.MarketValue     = +bondMarketValue(b).toFixed(2);
    b.UnrealizedPL    = +bondUnrealizedPL(b).toFixed(2);
  }
}

/* ============================================================
   Investment instruments — analytics (all auto-computed)
   Prices are money-per-unit; quantity is the number held.
   ============================================================ */
function instrumentById(id){ return (state.data.Instruments||[]).find(x=>x.InsID===id); }
function fundInstruments(fundId){ return (state.data.Instruments||[]).filter(x=>x.FundID===fundId); }
function insIsBond(t){ return /бонд|bond/i.test(String(t)); }
function insIsEquity(t){ return /хувьцаа|equity|stock/i.test(String(t)); }
function insIsFundUnit(t){ return /сангийн нэгж|fund unit/i.test(String(t)); }
function insIsMoneyMarket(t){ return /мөнгөн|money market/i.test(String(t)); }
/** current price (money per unit); falls back to purchase price until a market price is entered */
function insCurPrice(ins){ return num(ins.CurrentPrice)>0 ? num(ins.CurrentPrice) : num(ins.PurchasePrice); }
/** Market value = current price × quantity */
function insMarketValue(ins){ return insCurPrice(ins)*num(ins.Quantity); }
/** Cost = purchase price × quantity */
function insCostValue(ins){ return num(ins.PurchasePrice)*num(ins.Quantity); }
/** Unrealized gain/loss = market value − cost */
function insUnrealizedPL(ins){ return insMarketValue(ins)-insCostValue(ins); }
/** Gain/loss % = (current − purchase) / purchase × 100 */
function insGainPct(ins){ const p=num(ins.PurchasePrice); return p? (insCurPrice(ins)-p)/p*100 : 0; }
/** shape an instrument-bond into the object the bond helpers expect (price as % of par) */
function instrToBond(ins){
  const face=num(ins.FaceValue);
  const toQuote=v=> face>0 ? num(v)/face*100 : num(v); // money/unit → % of par
  return {FaceValue:face, CouponRate:ins.CouponRate, CouponFreq:ins.CouponFreq,
    IssueDate:ins.IssueDate, MaturityDate:ins.MaturityDate,
    PurchasePrice:toQuote(ins.PurchasePrice), MarketPrice:toQuote(insCurPrice(ins)),
    MarketPriceDate:ins.CurrentPriceDate};
}
/** YTM (annual %) for an instrument-bond, 0 for non-bonds */
function insYTM(ins){ return insIsBond(ins.InsType)? bondYTM(instrToBond(ins)) : 0; }
/** total accrued interest across the holding (per-bond accrual × quantity) */
function insAccrued(ins){ return insIsBond(ins.InsType)? bondAccruedInterest(instrToBond(ins))*(num(ins.Quantity)||1) : 0; }

/** snapshot the computed metrics onto each instrument record (called before save) */
function refreshInstrumentComputed(){
  for(const ins of state.data.Instruments||[]){
    ins.MarketValue  = +insMarketValue(ins).toFixed(2);
    ins.UnrealizedPL = +insUnrealizedPL(ins).toFixed(2);
    ins.GainPct      = +insGainPct(ins).toFixed(4);
    ins.YTM          = +insYTM(ins).toFixed(4);
    ins.AccruedInterest = +insAccrued(ins).toFixed(2);
  }
}

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
      el('button',{class:'btn btn-ghost btn-sm',onclick:()=>statementForm(i.InvestorID)},'Тайлан'),
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
  const i=existing||{InvestorID:nextId('INV',state.data.Investors,'InvestorID'),RegDate:todayISO(),InvestorType:'Хувь хүн',Residency:INV_RESIDENCY[0],Status:'Идэвхтэй'};
  const f=(name,labelMN,labelEN,opts={})=>fieldHTML(name,labelMN,labelEN,i[name],opts);
  const body=el('div',{},
    el('div',{class:'form-grid'},
      f('InvestorID','Дугаар','ID',{readonly:true}),
      f('RegDate','Бүртгэсэн огноо','Reg. date',{type:'date'}),
      f('InvestorType','Төрөл','Type',{type:'select',options:INV_TYPES}),
      f('Residency','Харьяалал','Residency',{type:'select',options:INV_RESIDENCY,hint:'Татварын хувь: Дотоодын 10%, Гадаадын 20%'}),
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
  add('Дугаар',i.InvestorID); add('Төрөл',i.InvestorType);
  add('Харьяалал',`${i.Residency||INV_RESIDENCY[0]} (татвар ${investorTaxRate(i)*100}%)`); add('Регистр',i.RegNo);
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
   Investor statement (printable, A4) — reads Investors /
   Transactions / Funds / NAVHistory only; no extra storage.
   ============================================================ */
const STMT_PERIODS = {        // label -> length in months
  'Сар · Monthly':      1,
  'Улирал · Quarterly': 3,
  'Жил · Annually':     12,
};
const TXN_LABEL = {
  [TXN_TYPE.SUB]:    'Худалдан авалт · Subscription',
  [TXN_TYPE.RED]:    'Буцаан худалдалт · Redemption',
  [TXN_TYPE.MGMT]:   'Удирдлагын шимтгэл · Management fee',
  [TXN_TYPE.SUCCESS]:'Гүйцэтгэлийн шимтгэл · Success fee',
  [TXN_TYPE.TAX]:    'Татвар · Withholding tax',
};

/** small modal to choose the reporting period before generating */
function statementForm(investorId){
  const inv=investorById(investorId);
  if(!inv){ toast('Хөрөнгө оруулагч олдсонгүй.','err'); return; }
  const periodSel=optSelect('period',Object.keys(STMT_PERIODS),'Улирал · Quarterly');
  const asOfIn=el('input',{type:'date',value:todayISO()});
  const body=el('div',{},
    el('div',{class:'small muted',style:'margin-bottom:10px'},'Хөрөнгө оруулагч: ',
      el('strong',{},inv.NameMN||inv.InvestorID)),
    el('div',{class:'form-grid'},
      wrapField('Тайлант хугацаа','Period',periodSel,true),
      wrapField('Тайлангийн огноо','Statement date',asOfIn,true)),
    el('div',{class:'small muted',style:'margin-top:6px'},
      'Үнэлгээ нь сонгосон огнооны байдлаар тооцоологдоно. Огнооны формат: YYYY-MM-DD.'));
  modal('Тайлан үүсгэх · Generate statement', body, {okText:'Үүсгэх · Generate', onOk:()=>{
    const end=asOfIn.value||todayISO();
    const start=addMonths(end, -(STMT_PERIODS[periodSel.value]||3));
    openInvestorStatement(investorId, periodSel.value, start, end);
  }});
}

/** assemble all statement figures as of `endDate` from the existing sheets */
function computeStatementData(investorId, endDate){
  const conf=state.data.Transactions.filter(t=>t.InvestorID===investorId
    && t.Status==='Баталгаажсан' && String(t.TradeDate)<=endDate);
  const sumType=(type,fid)=>conf.filter(t=>t.Type===type && t.FundID===fid)
    .reduce((s,t)=>s+num(t.NetAmount),0);
  const fundIds=[...new Set(conf.map(t=>t.FundID))];

  const funds=[];
  for(const fid of fundIds){
    const f=fundById(fid); if(!f) continue;
    const subs=conf.filter(t=>t.Type===TXN_TYPE.SUB && t.FundID===fid)
      .sort((a,b)=>String(a.TradeDate).localeCompare(String(b.TradeDate)));
    const trades=conf.filter(t=>isTradeType(t.Type) && t.FundID===fid);
    const units=trades.reduce((s,t)=>s+(t.Type===TXN_TYPE.SUB?num(t.Units):-num(t.Units)),0);
    const netInvested=trades.reduce((s,t)=>s+(t.Type===TXN_TYPE.SUB?num(t.NetAmount):-num(t.NetAmount)),0);
    const initialGross=subs.reduce((s,t)=>s+num(t.GrossAmount),0);
    const price=unitPriceAsOf(fid,endDate);
    const value=units*price;
    const mgmtFee=sumType(TXN_TYPE.MGMT,fid), successFee=sumType(TXN_TYPE.SUCCESS,fid), tax=sumType(TXN_TYPE.TAX,fid);
    const grossReturn=value-netInvested;
    const hurdle=num(f.HurdleRatePct);
    funds.push({fund:f, investDate:subs.length?subs[0].TradeDate:'', initialGross, mgmtFee, netInvested,
      units, price, value, hurdle, targetReturn:netInvested*hurdle/100,
      actualRetPct: netInvested>0?grossReturn/netInvested*100:0,
      successFee, tax, grossReturn, netReturn:grossReturn-successFee-tax});
  }
  funds.sort((a,b)=>String(a.fund.FundID).localeCompare(String(b.fund.FundID)));

  // full ledger with per-fund running unit balance
  const bal={};
  const rows=[...conf]
    .sort((a,b)=>String(a.TradeDate).localeCompare(String(b.TradeDate))||String(a.TxnID).localeCompare(String(b.TxnID)))
    .map(t=>{
      const u=t.Type===TXN_TYPE.SUB?num(t.Units): t.Type===TXN_TYPE.RED?-num(t.Units):0;
      bal[t.FundID]=(bal[t.FundID]||0)+u;
      return {t, units:u, balance:bal[t.FundID]};
    });

  const totals={
    invested:    funds.reduce((s,r)=>s+r.initialGross,0),
    netInvested: funds.reduce((s,r)=>s+r.netInvested,0),
    mgmt:        funds.reduce((s,r)=>s+r.mgmtFee,0),
    success:     funds.reduce((s,r)=>s+r.successFee,0),
    tax:         funds.reduce((s,r)=>s+r.tax,0),
    value:       funds.reduce((s,r)=>s+r.value,0),
    netReturn:   funds.reduce((s,r)=>s+r.netReturn,0),
  };
  totals.feesPaid=totals.mgmt+totals.success;
  totals.returnPct= totals.netInvested>0 ? totals.netReturn/totals.netInvested*100 : 0;
  return {funds, rows, totals};
}

/** signed cash amount of a ledger row from the investor's perspective */
function stmtLedgerAmount(t){
  if(t.Type===TXN_TYPE.SUB) return num(t.GrossAmount);   // contribution in
  if(t.Type===TXN_TYPE.RED) return -num(t.NetAmount);    // redemption out
  return -num(t.NetAmount);                              // fees & tax deducted
}

function openInvestorStatement(investorId, periodLabel, start, end){
  const inv=investorById(investorId); if(!inv) return;
  const data=computeStatementData(investorId, end);
  const logoUrl=new URL('assets/logo.png', window.location.href).href;
  const html=buildStatementHTML(inv, periodLabel, start, end, data, todayISO(), logoUrl);
  const w=window.open('', '_blank');
  if(!w){ toast('Попап хаагдсан тул тайлан нээгдсэнгүй. Хөтчийн попап зөвшөөрлийг идэвхжүүлнэ үү.','err','Тайлан'); return; }
  w.document.open(); w.document.write(html); w.document.close();
}

function buildStatementHTML(inv, periodLabel, start, end, data, genDate, logoUrl){
  const m  =(v,dec=0)=>'₮'+fmtMoney(v,dec);
  const sm =(v,dec=0)=>(num(v)<0?'−₮':'₮')+fmtMoney(Math.abs(v),dec);
  const cls=v=>num(v)<0?'neg':'pos';
  const kv =(mn,en,val,total)=>`<div class="kv-row${total?' kv-total':''}"><span class="kv-l">${mn} <i>· ${en}</i></span><span class="kv-v">${val}</span></div>`;

  const fundsHTML = data.funds.length ? data.funds.map(r=>`
    <section class="fund">
      <h3><span class="fdot" style="background:${esc(r.fund.Color||'#1366c4')}"></span>
        ${esc(r.fund.NameMN||r.fund.ShortName||r.fund.FundID)}${r.fund.NameEN?` <i>· ${esc(r.fund.NameEN)}</i>`:''}</h3>
      <div class="kv">
        ${kv('Хөрөнгө оруулсан огноо','Investment date', fmtDate(r.investDate)||'—')}
        ${kv('Анхны хөрөнгө оруулалт','Initial investment', m(r.initialGross))}
        ${kv('Удирдлагын шимтгэл','Management fee deducted', '−'+m(r.mgmtFee))}
        ${kv('Цэвэр оруулсан дүн','Net invested amount', m(r.netInvested))}
        ${kv('Нэгж эрх','Units / shares', fmtUnits(r.units))}
        ${kv('Нэгж эрхийн үнэ (тайлант огноонд)','Unit NAV at statement date', m(r.price,2))}
        ${kv('Одоогийн үнэлгээ','Current value', m(r.value))}
        ${kv('Босго хүү','Hurdle rate', r.hurdle.toFixed(2)+'%')}
        ${kv('Зорилтот өгөөж','Target return', m(r.targetReturn))}
        ${kv('Бодит өгөөж','Actual return', `<span class="${cls(r.actualRetPct)}">${r.actualRetPct.toFixed(2)}%</span> (${sm(r.grossReturn)})`)}
        ${kv('Гүйцэтгэлийн шимтгэл','Success fee', '−'+m(r.successFee))}
        ${kv('Суутгасан татвар','Tax withheld', '−'+m(r.tax))}
        ${kv('Цэвэр өгөөж','Net return to investor', `<strong class="${cls(r.netReturn)}">${sm(r.netReturn)}</strong>`, true)}
      </div>
    </section>`).join('') :
    `<p class="muted center">Энэ огнооны байдлаар идэвхтэй эзэмшил / гүйлгээ алга.<br><i>No holdings or activity as of this date.</i></p>`;

  const ledgerHTML = data.rows.length ? data.rows.map(({t,units,balance})=>{
    const amt=stmtLedgerAmount(t), f=fundById(t.FundID);
    const desc=[f?f.ShortName:t.FundID, t.Notes||''].filter(Boolean).join(' — ');
    return `<tr>
      <td>${fmtDate(t.TradeDate)}</td>
      <td>${esc(TXN_LABEL[t.Type]||t.Type)}</td>
      <td>${esc(desc)}</td>
      <td class="num ${amt<0?'neg':''}">${sm(amt)}</td>
      <td class="num">${units?fmtUnits(units):'—'}</td>
      <td class="num">${fmtUnits(balance)}</td>
    </tr>`;
  }).join('') : `<tr><td colspan="6" class="muted center">Гүйлгээ алга · No transactions</td></tr>`;

  const pr=(mn,en,val,extra='')=>`<tr><td>${mn} <i>· ${en}</i></td><td class="num ${extra}">${val}</td></tr>`;
  const summaryHTML=`
    ${pr('Нийт хөрөнгө оруулалт','Total invested', m(data.totals.invested))}
    ${pr('Нийт шимтгэл (удирдлага + гүйцэтгэл)','Total fees paid', m(data.totals.feesPaid))}
    ${pr('Суутгасан татвар','Total tax withheld', m(data.totals.tax))}
    ${pr('Багцын одоогийн үнэлгээ','Current portfolio value', m(data.totals.value))}
    ${pr('Нийт цэвэр өгөөж','Total net return', `<strong>${sm(data.totals.netReturn)}</strong>`, cls(data.totals.netReturn))}
    ${pr('Өгөөжийн хувь','Return percentage', `<strong>${data.totals.returnPct.toFixed(2)}%</strong>`, cls(data.totals.returnPct))}`;

  const name=esc(inv.NameMN||'')+(inv.NameEN?` · ${esc(inv.NameEN)}`:'');

  return `<!doctype html><html lang="mn"><head><meta charset="utf-8">
<title>Хөрөнгө оруулагчийн тайлан — ${esc(inv.InvestorID)}</title>
<style>
  :root{--navy:#0b2e4f;--blue:#1366c4;--gold:#c8a14b;--green:#1c8a4a;--red:#c0392b;--muted:#6b7785;--line:#dce3ea}
  *{box-sizing:border-box}
  body{margin:0;background:#eef1f5;color:#1b2733;font:13px/1.5 "Segoe UI",Arial,sans-serif}
  .toolbar{position:sticky;top:0;display:flex;gap:10px;justify-content:flex-end;padding:12px 18px;background:var(--navy)}
  .toolbar button{font:600 13px/1 "Segoe UI",Arial;padding:9px 18px;border:0;border-radius:7px;cursor:pointer}
  .btn-print{background:var(--gold);color:var(--navy)}
  .btn-close{background:transparent;color:#fff;border:1px solid rgba(255,255,255,.4)!important}
  .sheet{background:#fff;max-width:820px;margin:18px auto;padding:32px 36px;box-shadow:0 2px 14px rgba(0,0,0,.12)}
  header.stmt{display:flex;align-items:center;gap:16px;border-bottom:3px solid var(--navy);padding-bottom:16px}
  header.stmt .logo{width:54px;height:54px;border-radius:10px;object-fit:contain}
  header.stmt .logo-fb{width:54px;height:54px;border-radius:10px;background:linear-gradient(135deg,var(--gold),#e0c074);
    color:var(--navy);font-weight:800;display:none;align-items:center;justify-content:center;font-size:18px}
  header.stmt .co{font-size:18px;font-weight:800;color:var(--navy)}
  header.stmt .co small{display:block;font-weight:600;color:var(--muted);font-size:11px;letter-spacing:.3px}
  .title{margin:18px 0 4px;font-size:20px;font-weight:800;color:var(--navy)}
  .title i{font-weight:600;color:var(--blue);font-style:normal}
  .meta{display:grid;grid-template-columns:1fr 1fr;gap:4px 24px;margin:14px 0 4px;font-size:12.5px}
  .meta .row{display:flex;justify-content:space-between;border-bottom:1px dotted var(--line);padding:3px 0}
  .meta .row span:first-child{color:var(--muted)}
  .meta .row span:last-child{font-weight:600;text-align:right}
  h2.sec{margin:26px 0 10px;font-size:14px;color:#fff;background:var(--navy);padding:7px 12px;border-radius:6px;letter-spacing:.3px}
  h2.sec i{font-weight:500;opacity:.85;font-style:normal}
  section.fund{border:1px solid var(--line);border-radius:8px;padding:12px 14px;margin-bottom:12px}
  section.fund h3{margin:0 0 8px;font-size:14px;color:var(--navy);display:flex;align-items:center;gap:8px}
  section.fund h3 i{font-weight:500;color:var(--muted);font-style:normal}
  .fdot{width:11px;height:11px;border-radius:50%;display:inline-block}
  .kv{display:grid;grid-template-columns:1fr 1fr;gap:0 26px}
  .kv-row{display:flex;justify-content:space-between;gap:10px;padding:4px 0;border-bottom:1px dotted var(--line);font-size:12.5px}
  .kv-l{color:var(--muted)} .kv-l i{font-style:normal;opacity:.8}
  .kv-v{font-weight:600;text-align:right;white-space:nowrap}
  .kv-total{grid-column:1/-1;border-bottom:0;border-top:2px solid var(--navy);margin-top:4px;padding-top:7px;font-size:13.5px}
  table{width:100%;border-collapse:collapse;font-size:12px}
  thead th{background:#f3f6fa;color:var(--navy);text-align:left;padding:7px 9px;border-bottom:2px solid var(--line);font-size:11px;text-transform:uppercase;letter-spacing:.3px}
  tbody td{padding:6px 9px;border-bottom:1px solid var(--line)}
  td.num,th.num{text-align:right;white-space:nowrap}
  table.summary td:first-child{color:#33414f} table.summary td.num{font-weight:700}
  table.summary tr:last-child td{border-bottom:0}
  .pos{color:var(--green)} .neg{color:var(--red)} .muted{color:var(--muted)} .center{text-align:center}
  footer.stmt{margin-top:28px;border-top:2px solid var(--line);padding-top:14px;font-size:11px;color:var(--muted)}
  footer.stmt .conf{background:#f3f6fa;border-left:3px solid var(--gold);padding:8px 11px;border-radius:4px;margin-bottom:10px}
  footer.stmt strong{color:#33414f}
  @page{size:A4;margin:13mm}
  @media print{
    body{background:#fff}
    .toolbar{display:none!important}
    .sheet{box-shadow:none;margin:0;max-width:none;padding:0}
    section.fund{break-inside:avoid} tr{break-inside:avoid}
    .pos,.neg,h2.sec,.btn-print,header.stmt{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  }
</style></head><body>
  <div class="toolbar">
    <button class="btn-print" onclick="window.print()">🖨 Хэвлэх · Print</button>
    <button class="btn-close" onclick="window.close()">Хаах · Close</button>
  </div>
  <div class="sheet">
    <header class="stmt">
      <img class="logo" src="${esc(logoUrl)}" alt="GAM" onerror="this.style.display='none';this.nextElementSibling.style.display='flex';">
      <div class="logo-fb">GAM</div>
      <div class="co">Голомт Ассет Менежмент<small>Golomt Asset Management · «Голомт Ассет Менежмент ҮЦК» ХХК</small></div>
    </header>

    <div class="title">Хөрөнгө оруулагчийн тайлан <i>· Investor Statement</i></div>

    <div class="meta">
      <div class="row"><span>Хөрөнгө оруулагч · Investor</span><span>${name||'—'}</span></div>
      <div class="row"><span>Дугаар · ID</span><span>${esc(inv.InvestorID||'—')}</span></div>
      <div class="row"><span>Регистр · Registration no.</span><span>${esc(inv.RegNo||'—')}</span></div>
      <div class="row"><span>Харьяалал · Residency</span><span>${esc(inv.Residency||'—')}</span></div>
      <div class="row"><span>Тайлант хугацаа · Period</span><span>${esc(periodLabel)}</span></div>
      <div class="row"><span>Огнооны хязгаар · Date range</span><span>${fmtDate(start)} — ${fmtDate(end)}</span></div>
      <div class="row"><span>Тайлант огноо · Statement date</span><span>${fmtDate(end)}</span></div>
      <div class="row"><span>Үүсгэсэн огноо · Generated</span><span>${fmtDate(genDate)}</span></div>
    </div>

    <h2 class="sec">Сан тус бүрийн мэдээлэл <i>· Holdings by fund</i></h2>
    ${fundsHTML}

    <h2 class="sec">Гүйлгээний түүх <i>· Transaction history</i></h2>
    <table>
      <thead><tr><th>Огноо<br>Date</th><th>Төрөл<br>Type</th><th>Тайлбар<br>Description</th>
        <th class="num">Дүн<br>Amount</th><th class="num">Нэгж<br>Units</th><th class="num">Үлдэгдэл<br>Balance</th></tr></thead>
      <tbody>${ledgerHTML}</tbody>
    </table>

    <h2 class="sec">Гүйцэтгэлийн нэгдсэн дүн <i>· Performance summary</i></h2>
    <table class="summary"><tbody>${summaryHTML}</tbody></table>

    <footer class="stmt">
      <div class="conf"><strong>Нууцлалын мэдэгдэл:</strong> Энэхүү тайлан нь зөвхөн дээр нэрлэгдсэн хөрөнгө оруулагчид
        зориулагдсан бөгөөд нууц мэдээлэл агуулна. Зөвшөөрөлгүйгээр хуулбарлах, тараахыг хориглоно.<br>
        <i><strong>Confidentiality notice:</strong> This statement is intended solely for the named investor and contains
        confidential information. Unauthorised copying or distribution is prohibited.</i></div>
      <div>Энэ тайланг <strong>GAM Back Office System</strong>-ээр үүсгэв · Generated by GAM Back Office System.</div>
      <div>Холбоо барих · Contact: «Голомт Ассет Менежмент ҮЦК» ХХК · [хаяг / address] · [утас / phone] · [и-мэйл / email]</div>
    </footer>
  </div>
</body></html>`;
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
  /** on subscription, default the fee % to the fund's annual management fee rate */
  function syncMgmtFee(){
    if(!isSub) return;
    const f=fundById(fundSel.value);
    if(f) feePctIn.value = num(f.MgmtFeePct)||0;
  }
  function updateHolding(){
    if(isSub){ holdingNote.textContent=''; return; }
    const u=investorUnits(invSel.value,fundSel.value);
    holdingNote.textContent=`Эзэмшиж буй нэгж эрх: ${fmtUnits(u)}`;
  }
  function recalc(){
    const price=num(priceIn.value), feePct=num(feePctIn.value);
    let units, gross, fee;
    if(isSub){ // amount-driven: management fee (1 yr advance) deducted before units
      gross=num(amtIn.value); fee=gross*feePct/100;
      units= price? (gross-fee)/price : 0; unitsIn.value= units? units.toFixed(4):'';
    }else{ // units-driven
      units=num(unitsIn.value); gross=units*price; fee=gross*feePct/100;
      amtIn.value= gross? gross.toFixed(2):'';
    }
    const net=gross-fee;
    calc.innerHTML='';
    calc.append(
      crow('Нэгж эрхийн үнэ · Unit price', '₮'+fmtMoney(price,2)),
      crow(isSub?'Оруулсан дүн · Amount':'Дүн (gross)', '₮'+fmtMoney(gross,2)),
      crow(isSub?`Удирдлагын шимтгэл (${feePct||0}% · 1 жил урьдчилгаа)`:`Шимтгэл (${feePct||0}%)`, '−₮'+fmtMoney(fee,2)),
      crow(isSub?'Хөрөнгө оруулалт (нэгж эрх тооцох дүн)':'Хөрөнгө оруулагчид төлөх', '₮'+fmtMoney(net,2)),
      crow('Нэгж эрх · Units', fmtUnits(units), true),
    );
  }
  fundSel.addEventListener('change',()=>{priceIn.value='';syncMgmtFee();syncPrice();updateHolding();});
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
    isSub? wrapField('Удирдлагын шимтгэл % (жилийн)','Mgmt fee % (annual)',feePctIn) : wrapField('Шимтгэл %','Fee %',feePctIn),
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
  syncMgmtFee(); syncPrice(); updateHolding(); recalc();

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
    commit(()=>{
      state.data.Transactions.push(rec);
      // On subscription, book the management fee (1 year in advance) as its own line.
      if(isSub && fee>0){
        state.data.Transactions.push({
          TxnID:nextId('TXN',state.data.Transactions,'TxnID'), TradeDate:dateIn.value, Type:TXN_TYPE.MGMT,
          FundID:fundId, InvestorID:invId, Units:0, UnitPrice:0, GrossAmount:fee,
          FeePct:feePct, FeeAmount:fee, NetAmount:fee, SettlementDate:settleIn.value,
          Status:'Баталгаажсан', PaymentRef:rec.TxnID,
          Notes:'Удирдлагын шимтгэл (1 жилийн урьдчилгаа) · '+rec.TxnID, CreatedAt:new Date().toISOString(),
        });
      }
    }, (isSub?'Худалдан авалт':'Буцаан худалдалт')+' бүртгэгдлээ ('+rec.TxnID+')'
       +(isSub&&fee>0?' · Удирдлагын шимтгэл ₮'+fmtMoney(fee,0):''));
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
/** select with no placeholder; options are plain strings used as value+label */
function optSelect(name,options,value){
  const s=el('select',{name});
  for(const o of options) s.append(el('option',{value:o,selected:o===value?'':null},o));
  return s;
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
  ['',TXN_TYPE.SUB,TXN_TYPE.RED,TXN_TYPE.MGMT,TXN_TYPE.SUCCESS,TXN_TYPE.TAX]
    .forEach(t=>typeSel.append(el('option',{value:t},t||'Бүх төрөл')));
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
    const typeColor={[TXN_TYPE.SUB]:'badge-blue',[TXN_TYPE.RED]:'badge-amber',
      [TXN_TYPE.MGMT]:'badge-gray',[TXN_TYPE.SUCCESS]:'badge-gray',[TXN_TYPE.TAX]:'badge-red'};
    const typeBadge=el('span',{class:'badge '+(typeColor[x.Type]||'badge-gray')},x.Type);
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
    add(dl2,'Гүйцэтгэлийн шимтгэл',f.PerfFeePct+'%'); add(dl2,'Босго өгөөж · Hurdle',(num(f.HurdleRatePct)||0)+'%');
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
    g('MgmtFeePct','Удирдлагын шимтгэл % (жилийн)','Annual mgmt fee %',{type:'number'}),
    g('PerfFeePct','Гүйцэтгэлийн шимтгэл %','Success fee %',{type:'number'}),
    g('HurdleRatePct','Босго өгөөж %','Hurdle rate %',{type:'number',hint:'Гүйцэтгэлийн шимтгэл зөвхөн энэ хувиас давсан өгөөжид'}),
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
   VIEW: Bonds — bond portfolio tracking
   ============================================================ */
const beforeParen = s => String(s||'').split('(')[0].trim();
function ccySym(c){ return (!c||c==='MNT') ? '₮' : c+' '; }
/** signed money element, coloured green (gain) / red (loss) */
function plMoney(v,ccy){
  const s=num(v); const col= s>0?'var(--green)': s<0?'var(--red)':'var(--muted)';
  return el('span',{style:'color:'+col+';font-weight:600'},(s<0?'−':'')+ccySym(ccy)+fmtMoney(Math.abs(s),0));
}

VIEWS.bonds=function(){
  const wrap=el('div',{});
  wrap.append(viewHead('Бондын багц','Bond Portfolio',
    el('button',{class:'btn btn-primary',onclick:()=>bondForm()},'＋ Шинэ бонд · New bond')));

  wrap.append(el('div',{class:'panel'},el('div',{class:'panel-body'},
    el('p',{class:'muted small',style:'margin:0'},
      'Зах зээлийн үнийг (par-ийн %) гараар шинэчилнэ. ',
      el('strong',{},'YTM, хуримтлагдсан хүү, зах зээлийн үнэлгээ, боломжит ашиг/алдагдал'),
      ' автоматаар тооцоологдоно. Market value = Зах зээлийн үнэ × Нэрлэсэн үнэ ÷ 100. Огнооны формат: YYYY-MM-DD.'))));

  const bonds=state.data.Bonds||[];

  // ---- portfolio KPIs ----
  const totMV=bonds.reduce((s,b)=>s+bondMarketValue(b),0);
  const totCost=bonds.reduce((s,b)=>s+bondCostValue(b),0);
  const totAccr=bonds.reduce((s,b)=>s+bondAccruedInterest(b),0);
  const plPct= totCost? (totMV-totCost)/totCost*100 : 0;
  wrap.append(el('div',{class:'cards'},
    kpi('Нийт бонд · Holdings', fmtInt(bonds.length), 'bonds'),
    kpi('Зах зээлийн үнэлгээ', '₮'+fmtMoney(totMV,0), 'өртөг ₮'+fmtMoney(totCost,0)),
    kpi('Боломжит ашиг/алдагдал', (totMV-totCost<0?'−₮':'₮')+fmtMoney(Math.abs(totMV-totCost),0), plPct.toFixed(2)+'%'),
    kpi('Хуримтлагдсан хүү', '₮'+fmtMoney(totAccr,0), 'accrued interest')));

  // ---- per-fund exposure ----
  const fundsWithBonds=investableFunds().filter(f=>fundBonds(f.FundID).length);
  if(fundsWithBonds.length){
    const sp=el('div',{class:'panel'});
    sp.append(el('div',{class:'panel-head'},el('h3',{},'Сан тус бүрийн бондын дүн · Per-fund exposure')));
    const st=el('table',{class:'grid'});
    st.innerHTML=`<thead><tr><th>Сан</th><th class="num">Бонд</th><th class="num">Нэрлэсэн нийт</th>
      <th class="num">Өртөг</th><th class="num">Зах зээлийн үнэлгээ</th><th class="num">Ашиг/Алдагдал</th><th class="num">Хуримтлагдсан хүү</th></tr></thead>`;
    const stb=el('tbody');
    for(const f of fundsWithBonds){
      const list=fundBonds(f.FundID);
      const face=list.reduce((s,b)=>s+num(b.FaceValue),0);
      const cost=list.reduce((s,b)=>s+bondCostValue(b),0);
      const mv=list.reduce((s,b)=>s+bondMarketValue(b),0);
      const accr=list.reduce((s,b)=>s+bondAccruedInterest(b),0);
      stb.append(el('tr',{},
        el('td',{},fundChip(f.FundID)),
        el('td',{class:'num'},fmtInt(list.length)),
        el('td',{class:'num'},'₮'+fmtMoney(face,0)),
        el('td',{class:'num'},'₮'+fmtMoney(cost,0)),
        el('td',{class:'num'},'₮'+fmtMoney(mv,0)),
        el('td',{class:'num'},plMoney(mv-cost)),
        el('td',{class:'num'},'₮'+fmtMoney(accr,0))));
    }
    st.append(stb); sp.append(el('div',{class:'table-wrap'},st)); wrap.append(sp);
  }

  // ---- bond list ----
  const panel=el('div',{class:'panel'});
  panel.append(el('div',{class:'panel-head'},el('h3',{},'Бондын жагсаалт · All bonds')));
  const t=el('table',{class:'grid'});
  t.innerHTML=`<thead><tr><th>Гаргагч · Issuer</th><th>Сан</th><th>Төрөл</th><th class="num">Нэрлэсэн үнэ</th>
    <th class="num">Купон</th><th>Дуусах огноо</th><th class="num">Зах зээл (%)</th><th class="num">Зах зээлийн үнэлгээ</th>
    <th class="num">YTM</th><th class="num">Хуримтлагдсан</th><th class="num">Ашиг/Алдагдал</th><th></th></tr></thead>`;
  const tb=el('tbody');
  if(!bonds.length) tb.append(emptyRow(12,'Бонд бүртгэгдээгүй байна. «Шинэ бонд» дарж эхлүүлнэ үү.'));
  for(const b of [...bonds].sort((a,b)=>String(a.MaturityDate).localeCompare(String(b.MaturityDate)))){
    const ccy=b.Currency;
    tb.append(el('tr',{},
      el('td',{},el('div',{},el('strong',{},b.IssuerMN||''), b.IssuerEN?el('div',{class:'small muted'},b.IssuerEN):null)),
      el('td',{},fundChip(b.FundID)),
      el('td',{},el('span',{class:'small'},beforeParen(b.BondType))),
      el('td',{class:'num'},ccySym(ccy)+fmtMoney(b.FaceValue,0)),
      el('td',{class:'num'},(num(b.CouponRate)).toFixed(2)+'%',el('div',{class:'small muted'},beforeParen(b.CouponFreq))),
      el('td',{},fmtDate(b.MaturityDate)),
      el('td',{class:'num'},num(b.MarketPrice)?fmtMoney(b.MarketPrice,2):el('span',{class:'muted'},'—')),
      el('td',{class:'num'},ccySym(ccy)+fmtMoney(bondMarketValue(b),0)),
      el('td',{class:'num'},bondYTM(b).toFixed(2)+'%'),
      el('td',{class:'num'},ccySym(ccy)+fmtMoney(bondAccruedInterest(b),0)),
      el('td',{class:'num'},plMoney(bondUnrealizedPL(b),ccy)),
      el('td',{},el('div',{class:'row-actions'},
        el('button',{class:'btn btn-ghost btn-sm',onclick:()=>bondDetail(b.BondID)},'Дэлгэрэнгүй'),
        el('button',{class:'btn btn-ghost btn-sm',onclick:()=>bondForm(b)},'Засах')))));
  }
  t.append(tb); panel.append(el('div',{class:'table-wrap'},t)); wrap.append(panel);
  return wrap;
};

function bondForm(existing){
  const isEdit=!!existing;
  const b=existing||{BondID:nextId('BOND',state.data.Bonds,'BondID'),Currency:'MNT',
    BondType:BOND_TYPES[0],CouponFreq:COUPON_FREQ[1],Status:BOND_STATUS[0],
    IssueDate:todayISO(),PurchaseDate:todayISO()};

  const issuerMN=el('input',{type:'text',value:b.IssuerMN||''});
  const issuerEN=el('input',{type:'text',value:b.IssuerEN||''});
  const fundSel=selectEl('FundID',investableFunds().map(f=>({v:f.FundID,t:`${f.ShortName} (${f.FundID})`}))); fundSel.value=b.FundID||'';
  const typeSel=optSelect('BondType',BOND_TYPES,b.BondType);
  const faceIn=el('input',{type:'number',step:'0.01',min:'0',value:b.FaceValue??''});
  const rateIn=el('input',{type:'number',step:'0.001',min:'0',value:b.CouponRate??''});
  const freqSel=optSelect('CouponFreq',COUPON_FREQ,b.CouponFreq);
  const issueIn=el('input',{type:'date',value:b.IssueDate||''});
  const matIn=el('input',{type:'date',value:b.MaturityDate||''});
  const purchPriceIn=el('input',{type:'number',step:'0.0001',min:'0',value:b.PurchasePrice??'',placeholder:'par-ийн %, ж: 100'});
  const purchDateIn=el('input',{type:'date',value:b.PurchaseDate||''});
  const mktIn=el('input',{type:'number',step:'0.0001',min:'0',value:b.MarketPrice??'',placeholder:'par-ийн %'});
  const mktDateIn=el('input',{type:'date',value:b.MarketPriceDate||''});
  const curSel=optSelect('Currency',BOND_CCY,b.Currency||'MNT');
  const statusSel=optSelect('Status',BOND_STATUS,b.Status||BOND_STATUS[0]);
  const noteIn=el('input',{type:'text',value:b.Notes||''});
  const calc=el('div',{class:'calc-box'});

  function readBond(){ return {FaceValue:faceIn.value,CouponRate:rateIn.value,CouponFreq:freqSel.value,
    IssueDate:issueIn.value,MaturityDate:matIn.value,PurchasePrice:purchPriceIn.value,
    MarketPrice:mktIn.value,MarketPriceDate:mktDateIn.value,Currency:curSel.value}; }
  function recalc(){
    const t=readBond(), ccy=t.Currency;
    calc.innerHTML='';
    calc.append(
      crow('Зах зээлийн үнэлгээ · Market value', ccySym(ccy)+fmtMoney(bondMarketValue(t),2)),
      crow('Өртөг · Cost', ccySym(ccy)+fmtMoney(bondCostValue(t),2)),
      crow('Хуримтлагдсан хүү · Accrued interest', ccySym(ccy)+fmtMoney(bondAccruedInterest(t),2)),
      crow('YTM (жилийн) · Yield to maturity', bondYTM(t).toFixed(3)+'%'),
      crow('Боломжит ашиг/алдагдал · Unrealized P/L', (bondUnrealizedPL(t)<0?'−':'')+ccySym(ccy)+fmtMoney(Math.abs(bondUnrealizedPL(t)),2), true),
    );
  }
  [faceIn,rateIn,purchPriceIn,mktIn].forEach(x=>x.addEventListener('input',recalc));
  [freqSel,issueIn,matIn,mktDateIn,curSel].forEach(x=>x.addEventListener('change',recalc));

  const body=el('div',{},
    el('div',{class:'form-grid'},
      wrapField('Гаргагч (Монгол)','Issuer (MN)',issuerMN,true),
      wrapField('Гаргагч (English)','Issuer (EN)',issuerEN),
      wrapField('Сан','Fund',fundSel,true),
      wrapField('Бондын төрөл','Bond type',typeSel),
      wrapField('Нэрлэсэн үнэ','Face value',faceIn,true),
      wrapField('Купон хүү %','Coupon rate %',rateIn),
      wrapField('Купон давтамж','Coupon frequency',freqSel),
      wrapField('Валют','Currency',curSel),
      wrapField('Гаргасан огноо','Issue date',issueIn),
      wrapField('Дуусах огноо','Maturity date',matIn,true),
      wrapField('Худалдан авсан үнэ (%)','Purchase price (%)',purchPriceIn),
      wrapField('Худалдан авсан огноо','Purchase date',purchDateIn),
      wrapField('Зах зээлийн үнэ (%)','Market price (%)',mktIn),
      wrapField('Зах зээлийн үнийн огноо','Market price date',mktDateIn),
      wrapField('Төлөв','Status',statusSel),
      wrapField('Тэмдэглэл','Notes',noteIn),
    ),
    el('div',{class:'small muted',style:'margin:6px 2px'},'Доорх үзүүлэлтүүд автоматаар тооцоологдоно · auto-computed:'),
    calc);
  recalc();

  modal(isEdit?'Бонд засах':'Шинэ бонд', body, {okText:isEdit?'Хадгалах':'Бүртгэх', wide:true, onOk:()=>{
    if(!issuerMN.value.trim()){ toast('Гаргагчийн нэр заавал.','err'); return false; }
    if(!fundSel.value){ toast('Сан сонгоно уу.','err'); return false; }
    if(num(faceIn.value)<=0){ toast('Нэрлэсэн үнэ оруулна уу.','err'); return false; }
    if(!matIn.value){ toast('Дуусах огноо оруулна уу.','err'); return false; }
    const rec={
      BondID:b.BondID, IssuerMN:issuerMN.value.trim(), IssuerEN:issuerEN.value.trim(),
      FundID:fundSel.value, BondType:typeSel.value, FaceValue:num(faceIn.value),
      CouponRate:num(rateIn.value), CouponFreq:freqSel.value, IssueDate:issueIn.value,
      MaturityDate:matIn.value, PurchasePrice:num(purchPriceIn.value), PurchaseDate:purchDateIn.value,
      MarketPrice:num(mktIn.value), MarketPriceDate:mktDateIn.value, Currency:curSel.value,
      Status:statusSel.value, Notes:noteIn.value,
    };
    rec.YTM=+bondYTM(rec).toFixed(4); rec.AccruedInterest=+bondAccruedInterest(rec).toFixed(2);
    rec.MarketValue=+bondMarketValue(rec).toFixed(2); rec.UnrealizedPL=+bondUnrealizedPL(rec).toFixed(2);
    commit(()=>{ if(isEdit) Object.assign(existing,rec); else state.data.Bonds.push(rec); },
      isEdit?'Бонд шинэчлэгдлээ':'Бонд бүртгэгдлээ ('+rec.BondID+')');
  }});
}

function bondDetail(id){
  const b=bondById(id); if(!b)return;
  const ccy=b.Currency;
  const body=el('div',{});
  body.append(el('h3',{style:'margin:0 0 2px;color:var(--gam-navy)'},b.IssuerMN||''),
    el('div',{class:'muted small',style:'margin-bottom:12px'},
      (b.IssuerEN||'')+' · '+b.BondID+' · '+beforeParen(b.BondType)));

  // facts
  const two=el('div',{class:'panel-body two-col',style:'padding:0'});
  const dl1=el('dl',{class:'def-list'}), dl2=el('dl',{class:'def-list'});
  const add=(dl,k,v)=>{dl.append(el('dt',{},k),el('dd',{},v||'—'));};
  add(dl1,'Сан · Fund', (fundById(b.FundID)||{}).ShortName||b.FundID);
  add(dl1,'Нэрлэсэн үнэ · Face', ccySym(ccy)+fmtMoney(b.FaceValue,2));
  add(dl1,'Купон · Coupon', num(b.CouponRate).toFixed(3)+'% · '+beforeParen(b.CouponFreq));
  add(dl1,'Гаргасан · Issued', fmtDate(b.IssueDate));
  add(dl1,'Дуусах · Maturity', fmtDate(b.MaturityDate));
  add(dl1,'Валют · Currency', b.Currency||'MNT');
  add(dl2,'Худалдан авсан үнэ · Purchase', num(b.PurchasePrice).toFixed(4)+'%  ('+fmtDate(b.PurchaseDate)+')');
  add(dl2,'Зах зээлийн үнэ · Market', (num(b.MarketPrice)?num(b.MarketPrice).toFixed(4)+'%':'—')+(b.MarketPriceDate?'  ('+fmtDate(b.MarketPriceDate)+')':''));
  add(dl2,'Зах зээлийн үнэлгээ · Market value', ccySym(ccy)+fmtMoney(bondMarketValue(b),2));
  add(dl2,'YTM', bondYTM(b).toFixed(3)+'%');
  add(dl2,'Хуримтлагдсан хүү · Accrued', ccySym(ccy)+fmtMoney(bondAccruedInterest(b),2));
  two.append(dl1,dl2); body.append(two);
  body.append(el('div',{class:'calc-box',style:'margin-top:12px'},
    crow('Боломжит ашиг/алдагдал · Unrealized P/L',
      (bondUnrealizedPL(b)<0?'−':'')+ccySym(ccy)+fmtMoney(Math.abs(bondUnrealizedPL(b)),2), true)));

  // cashflow schedule
  body.append(el('h4',{style:'margin:18px 0 8px;color:var(--gam-navy)'},'Мөнгөн урсгалын хуваарь · Cashflow schedule'));
  const coupon=bondPeriodCoupon(b);
  const dates=bondCouponDates(b);
  const today=todayISO();
  const t=el('table',{class:'grid'});
  t.innerHTML=`<thead><tr><th>Огноо · Date</th><th class="num">Купон · Coupon</th><th class="num">Үндсэн төлбөр · Principal</th>
    <th class="num">Нийт · Total</th><th>Төлөв · Status</th></tr></thead>`;
  const tb=el('tbody');
  if(!dates.length) tb.append(emptyRow(5,'Гаргасан/дуусах огноо болон давтамжаа оруулна уу.'));
  let futureCoupons=0;
  dates.forEach((d,i)=>{
    const isLast=i===dates.length-1;
    const principal=isLast?num(b.FaceValue):0;
    const total=coupon+principal;
    const paid=d<=today;
    if(!paid) futureCoupons+=coupon;
    tb.append(el('tr',{},
      el('td',{},fmtDate(d)),
      el('td',{class:'num'},ccySym(ccy)+fmtMoney(coupon,2)),
      el('td',{class:'num'},principal?ccySym(ccy)+fmtMoney(principal,2):'—'),
      el('td',{class:'num'},ccySym(ccy)+fmtMoney(total,2)),
      el('td',{},el('span',{class:'badge '+(paid?'badge-gray':'badge-blue')},paid?'Төлөгдсөн':'Хүлээгдэж буй'))));
  });
  t.append(tb); body.append(el('div',{class:'table-wrap'},t));
  if(dates.length) body.append(el('div',{class:'small muted',style:'margin-top:6px'},
    'Үлдсэн купон төлбөр · remaining coupons: ',el('strong',{},ccySym(ccy)+fmtMoney(futureCoupons,2)),
    ' · Дуусахад эргэн төлөгдөх үндсэн төлбөр · principal at maturity: ',el('strong',{},ccySym(ccy)+fmtMoney(num(b.FaceValue),2))));

  modal('Бондын дэлгэрэнгүй · Bond detail', body, {okText:'Хаах', cancel:false, wide:true});
}

/* ============================================================
   VIEW: Investment Instruments — master registry of every
   financial instrument the GAM funds invest in.
   ============================================================ */
const INS_BUCKETS = ['Бонд','Хувьцаа','Сангийн нэгж','Мөнгөний ЗЗ','Бусад'];
function insBucket(t){
  if(insIsBond(t))       return 'Бонд';
  if(insIsEquity(t))     return 'Хувьцаа';
  if(insIsFundUnit(t))   return 'Сангийн нэгж';
  if(insIsMoneyMarket(t))return 'Мөнгөний ЗЗ';
  return 'Бусад';
}

VIEWS.instruments=function(){
  const wrap=el('div',{});
  wrap.append(viewHead('Хөрөнгө оруулалтын хэрэгсэл','Investment Instruments — master registry',
    el('button',{class:'btn btn-primary',onclick:()=>instrumentForm()},'＋ Шинэ хэрэгсэл · New instrument')));

  wrap.append(el('div',{class:'panel'},el('div',{class:'panel-body'},
    el('p',{class:'muted small',style:'margin:0'},
      'Бүх санхүүгийн хэрэгслийг энд бүртгэж, санд хуваарилна. ',
      el('strong',{},'Зах зээлийн үнэлгээ, ашиг/алдагдал, жин (%)'),
      ' автоматаар тооцоологдоно. Бондын YTM ба хуримтлагдсан хүү автомат. Огнооны формат: YYYY-MM-DD.'))));

  const all=state.data.Instruments||[];

  // ---- portfolio KPIs ----
  const totMV=all.reduce((s,x)=>s+insMarketValue(x),0);
  const totCost=all.reduce((s,x)=>s+insCostValue(x),0);
  const plPct= totCost? (totMV-totCost)/totCost*100 : 0;
  wrap.append(el('div',{class:'cards'},
    kpi('Нийт хэрэгсэл · Instruments', fmtInt(all.length), ''),
    kpi('Зах зээлийн үнэлгээ', '₮'+fmtMoney(totMV,0), 'өртөг ₮'+fmtMoney(totCost,0)),
    kpi('Боломжит ашиг/алдагдал', (totMV-totCost<0?'−₮':'₮')+fmtMoney(Math.abs(totMV-totCost),0), plPct.toFixed(2)+'%'),
    kpi('Төрлийн тоо · Types', fmtInt(new Set(all.map(x=>insBucket(x.InsType))).size), 'instrument classes')));

  // ---- per-fund summary, broken down by instrument type ----
  const fundIds=[...new Set(all.map(x=>x.FundID||''))];
  if(fundIds.length){
    const sp=el('div',{class:'panel'});
    sp.append(el('div',{class:'panel-head'},el('h3',{},'Сан тус бүрийн дүн (төрлөөр) · Per-fund value by type')));
    const st=el('table',{class:'grid'});
    st.innerHTML=`<thead><tr><th>Сан</th>${INS_BUCKETS.map(b=>`<th class="num">${b}</th>`).join('')}<th class="num">Нийт · Total</th></tr></thead>`;
    const stb=el('tbody');
    for(const fid of fundIds){
      const list=all.filter(x=>(x.FundID||'')===fid);
      const cells=INS_BUCKETS.map(b=>{
        const v=list.filter(x=>insBucket(x.InsType)===b).reduce((s,x)=>s+insMarketValue(x),0);
        return el('td',{class:'num'}, v? '₮'+fmtMoney(v,0) : el('span',{class:'muted'},'—'));
      });
      const tot=list.reduce((s,x)=>s+insMarketValue(x),0);
      stb.append(el('tr',{},
        el('td',{}, fid? fundChip(fid) : el('span',{class:'muted'},'Хуваарилаагүй · Unassigned')),
        ...cells, el('td',{class:'num'},el('strong',{},'₮'+fmtMoney(tot,0)))));
    }
    st.append(stb); sp.append(el('div',{class:'table-wrap'},st)); wrap.append(sp);
  }

  // ---- filters + list ----
  let fType='',fFund='',fStatus='',fSearch='';
  const typeSel=el('select',{onchange:e=>{fType=e.target.value;draw();}});
  ['',...INS_TYPES].forEach(t=>typeSel.append(el('option',{value:t},t||'Бүх төрөл')));
  const fundSel=el('select',{onchange:e=>{fFund=e.target.value;draw();}});
  fundSel.append(el('option',{value:''},'Бүх сан'));
  investableFunds().forEach(f=>fundSel.append(el('option',{value:f.FundID},f.ShortName)));
  const statusSel=el('select',{onchange:e=>{fStatus=e.target.value;draw();}});
  ['',...INS_STATUS].forEach(s=>statusSel.append(el('option',{value:s},s||'Бүх төлөв')));
  const searchIn=el('input',{type:'text',placeholder:'Нэр, гаргагч, ISIN, тикэрээр хайх…',
    oninput:e=>{fSearch=e.target.value.toLowerCase().trim();draw();}});
  wrap.append(el('div',{class:'toolbar'},typeSel,fundSel,statusSel,el('div',{class:'search',style:'margin-left:auto'},searchIn)));

  const panel=el('div',{class:'panel'});
  const host=el('div',{class:'table-wrap'});
  panel.append(host); wrap.append(panel);

  function draw(){
    let list=[...all];
    if(fType)   list=list.filter(x=>x.InsType===fType);
    if(fFund)   list=list.filter(x=>x.FundID===fFund);
    if(fStatus) list=list.filter(x=>x.Status===fStatus);
    if(fSearch) list=list.filter(x=>`${x.InsID} ${x.NameMN} ${x.NameEN} ${x.Issuer} ${x.ISIN} ${x.Ticker}`.toLowerCase().includes(fSearch));
    list.sort((a,b)=>String(a.InsID).localeCompare(String(b.InsID)));
    const t=el('table',{class:'grid'});
    t.innerHTML=`<thead><tr><th>ID</th><th>Төрөл</th><th>Нэр · Name</th><th>Гаргагч</th><th>Сан</th>
      <th class="num">Авсан үнэ</th><th class="num">Одоогийн үнэ</th><th class="num">Ашиг/Алдагдал %</th><th>Төлөв</th><th></th></tr></thead>`;
    const tb=el('tbody');
    if(!list.length) tb.append(emptyRow(10, all.length?'Шүүлтэд тохирох хэрэгсэл алга.':'Хэрэгсэл бүртгэгдээгүй. «Шинэ хэрэгсэл» дарж эхэлнэ үү.'));
    for(const x of list){
      const ccy=x.Currency, g=insGainPct(x);
      tb.append(el('tr',{},
        el('td',{},x.InsID),
        el('td',{},el('span',{class:'small'},beforeParen(x.InsType))),
        el('td',{},el('div',{},el('strong',{},x.NameMN||''), x.NameEN?el('div',{class:'small muted'},x.NameEN):null)),
        el('td',{},x.Issuer||''),
        el('td',{}, x.FundID? fundChip(x.FundID) : el('span',{class:'muted'},'—')),
        el('td',{class:'num'},ccySym(ccy)+fmtMoney(x.PurchasePrice,2)),
        el('td',{class:'num'},ccySym(ccy)+fmtMoney(insCurPrice(x),2)),
        el('td',{class:'num'},el('span',{style:'color:'+(g<0?'var(--red)':g>0?'var(--green)':'var(--muted)')+';font-weight:600'},(g<0?'−':'')+Math.abs(g).toFixed(2)+'%')),
        el('td',{},statusBadge(beforeParen(x.Status))),
        el('td',{},el('div',{class:'row-actions'},
          el('button',{class:'btn btn-ghost btn-sm',onclick:()=>instrumentDetail(x.InsID)},'Дэлгэрэнгүй'),
          el('button',{class:'btn btn-ghost btn-sm',onclick:()=>instrumentForm(x)},'Засах')))));
    }
    t.append(tb); host.innerHTML=''; host.append(t);
  }
  draw();
  return wrap;
};

function instrumentForm(existing){
  const isEdit=!!existing;
  const x=existing||{InsID:nextId('INS',state.data.Instruments,'InsID'),InsType:INS_TYPES[0],
    Currency:'MNT',Status:INS_STATUS[0],PurchaseDate:todayISO()};

  const typeSel=optSelect('InsType',INS_TYPES,x.InsType);
  const statusSel=optSelect('Status',INS_STATUS,x.Status||INS_STATUS[0]);
  const nameMN=el('input',{type:'text',value:x.NameMN||''});
  const nameEN=el('input',{type:'text',value:x.NameEN||''});
  const issuer=el('input',{type:'text',value:x.Issuer||''});
  const isin=el('input',{type:'text',value:x.ISIN||''});
  const ccySel=optSelect('Currency',BOND_CCY,x.Currency||'MNT');
  const fundSel=selectEl('FundID',investableFunds().map(f=>({v:f.FundID,t:`${f.ShortName} (${f.FundID})`}))); fundSel.value=x.FundID||'';
  const purchDate=el('input',{type:'date',value:x.PurchaseDate||''});
  const purchPrice=el('input',{type:'number',step:'0.0001',min:'0',value:x.PurchasePrice??''});
  const qty=el('input',{type:'number',step:'0.0001',min:'0',value:x.Quantity??''});
  const curPrice=el('input',{type:'number',step:'0.0001',min:'0',value:x.CurrentPrice??'',placeholder:'гараар шинэчилнэ'});
  const curDate=el('input',{type:'date',value:x.CurrentPriceDate||''});
  const noteIn=el('input',{type:'text',value:x.Notes||''});
  // bond
  const faceIn=el('input',{type:'number',step:'0.01',min:'0',value:x.FaceValue??''});
  const rateIn=el('input',{type:'number',step:'0.001',min:'0',value:x.CouponRate??''});
  const freqSel=optSelect('CouponFreq',COUPON_FREQ,x.CouponFreq||COUPON_FREQ[1]);
  const issueIn=el('input',{type:'date',value:x.IssueDate||''});
  const matIn=el('input',{type:'date',value:x.MaturityDate||''});
  // equity
  const exchSel=optSelect('Exchange',INS_EXCHANGES,x.Exchange||INS_EXCHANGES[0]);
  const tickerIn=el('input',{type:'text',value:x.Ticker||''});
  const divIn=el('input',{type:'number',step:'0.01',min:'0',value:x.DividendYield??''});
  // fund unit
  const fmgrIn=el('input',{type:'text',value:x.FundManager||''});
  const futypeIn=el('input',{type:'text',value:x.FundUnitType||''});

  const calc=el('div',{class:'calc-box'});

  // wrap type-specific fields so we can show/hide them as a group
  const bondFields=[
    wrapField('Нэрлэсэн үнэ','Face value',faceIn),
    wrapField('Купон хүү %','Coupon rate %',rateIn),
    wrapField('Купон давтамж','Coupon frequency',freqSel),
    wrapField('Гаргасан огноо','Issue date',issueIn),
    wrapField('Дуусах огноо','Maturity date',matIn)];
  const equityFields=[
    wrapField('Бирж','Stock exchange',exchSel),
    wrapField('Тикэр','Ticker symbol',tickerIn),
    wrapField('Ногдол ашгийн өгөөж %','Dividend yield %',divIn)];
  const fundUnitFields=[
    wrapField('Сангийн менежер','Fund manager',fmgrIn),
    wrapField('Сангийн төрөл','Fund type',futypeIn)];

  function readIns(){ return {InsType:typeSel.value,FaceValue:faceIn.value,CouponRate:rateIn.value,
    CouponFreq:freqSel.value,IssueDate:issueIn.value,MaturityDate:matIn.value,
    PurchasePrice:purchPrice.value,Quantity:qty.value,CurrentPrice:curPrice.value,CurrentPriceDate:curDate.value}; }
  function recalc(){
    const t=readIns(), ccy=ccySel.value;
    const rows=[
      crow('Зах зээлийн үнэлгээ · Market value', ccySym(ccy)+fmtMoney(insMarketValue(t),2)),
      crow('Өртөг · Cost', ccySym(ccy)+fmtMoney(insCostValue(t),2)),
      crow('Ашиг/Алдагдал % · Gain/Loss %', insGainPct(t).toFixed(2)+'%')];
    if(insIsBond(t.InsType)) rows.push(
      crow('YTM (жилийн)', insYTM(t).toFixed(3)+'%'),
      crow('Хуримтлагдсан хүү · Accrued', ccySym(ccy)+fmtMoney(insAccrued(t),2)));
    rows.push(crow('Боломжит ашиг/алдагдал · Unrealized P/L',
      (insUnrealizedPL(t)<0?'−':'')+ccySym(ccy)+fmtMoney(Math.abs(insUnrealizedPL(t)),2), true));
    calc.innerHTML=''; calc.append(...rows);
  }
  function syncType(){
    const t=typeSel.value;
    bondFields.forEach(f=>f.style.display=insIsBond(t)?'':'none');
    equityFields.forEach(f=>f.style.display=insIsEquity(t)?'':'none');
    fundUnitFields.forEach(f=>f.style.display=insIsFundUnit(t)?'':'none');
    recalc();
  }
  typeSel.addEventListener('change',syncType);
  [purchPrice,qty,curPrice,faceIn,rateIn].forEach(i=>i.addEventListener('input',recalc));
  [freqSel,issueIn,matIn,curDate,ccySel].forEach(i=>i.addEventListener('change',recalc));

  const body=el('div',{},
    el('div',{class:'form-grid'},
      wrapField('Хэрэгслийн төрөл','Instrument type',typeSel,true),
      wrapField('Төлөв','Status',statusSel),
      wrapField('Нэр (Монгол)','Name (MN)',nameMN,true),
      wrapField('Нэр (English)','Name (EN)',nameEN),
      wrapField('Гаргагч / Компани','Issuer / Company',issuer),
      wrapField('Бүртгэл / ISIN','Registration / ISIN',isin),
      wrapField('Валют','Currency',ccySel),
      wrapField('Сан','Fund assignment',fundSel),
      wrapField('Худалдан авсан огноо','Purchase date',purchDate),
      wrapField('Худалдан авсан үнэ','Purchase price',purchPrice,true),
      wrapField('Тоо ширхэг / нэгж','Quantity / units',qty,true),
      wrapField('Одоогийн үнэ','Current price',curPrice),
      wrapField('Одоогийн үнийн огноо','Current price date',curDate),
      ...bondFields, ...equityFields, ...fundUnitFields,
      wrapField('Тэмдэглэл','Notes',noteIn),
    ),
    el('div',{class:'small muted',style:'margin:6px 2px'},'Доорх үзүүлэлтүүд автоматаар тооцоологдоно · auto-computed:'),
    calc);
  syncType();

  modal(isEdit?'Хэрэгсэл засах':'Шинэ хэрэгсэл', body, {okText:isEdit?'Хадгалах':'Бүртгэх', wide:true, onOk:()=>{
    if(!nameMN.value.trim()){ toast('Нэр заавал бөглөнө.','err'); return false; }
    if(num(purchPrice.value)<=0){ toast('Худалдан авсан үнэ оруулна уу.','err'); return false; }
    if(num(qty.value)<=0){ toast('Тоо ширхэг оруулна уу.','err'); return false; }
    const rec={
      InsID:x.InsID, InsType:typeSel.value, NameMN:nameMN.value.trim(), NameEN:nameEN.value.trim(),
      Issuer:issuer.value.trim(), ISIN:isin.value.trim(), Currency:ccySel.value, FundID:fundSel.value,
      Status:statusSel.value, PurchaseDate:purchDate.value, PurchasePrice:num(purchPrice.value),
      Quantity:num(qty.value), CurrentPrice:num(curPrice.value), CurrentPriceDate:curDate.value,
      FaceValue:num(faceIn.value), CouponRate:num(rateIn.value), CouponFreq:freqSel.value,
      IssueDate:issueIn.value, MaturityDate:matIn.value, Exchange:exchSel.value, Ticker:tickerIn.value.trim(),
      DividendYield:num(divIn.value), FundManager:fmgrIn.value.trim(), FundUnitType:futypeIn.value.trim(),
      Notes:noteIn.value,
    };
    rec.MarketValue=+insMarketValue(rec).toFixed(2); rec.UnrealizedPL=+insUnrealizedPL(rec).toFixed(2);
    rec.GainPct=+insGainPct(rec).toFixed(4); rec.YTM=+insYTM(rec).toFixed(4); rec.AccruedInterest=+insAccrued(rec).toFixed(2);
    commit(()=>{ if(isEdit) Object.assign(existing,rec); else state.data.Instruments.push(rec); },
      isEdit?'Хэрэгсэл шинэчлэгдлээ':'Хэрэгсэл бүртгэгдлээ ('+rec.InsID+')');
  }});
}

function instrumentDetail(id){
  const x=instrumentById(id); if(!x)return;
  const ccy=x.Currency;
  const body=el('div',{});
  body.append(el('h3',{style:'margin:0 0 2px;color:var(--gam-navy)'},x.NameMN||''),
    el('div',{class:'muted small',style:'margin-bottom:12px'},
      (x.NameEN||'')+' · '+x.InsID+' · '+beforeParen(x.InsType)));

  // metric cards
  const mv=insMarketValue(x), pl=insUnrealizedPL(x);
  const fundList=x.FundID?fundInstruments(x.FundID):[];
  const fundTotal=fundList.reduce((s,i)=>s+insMarketValue(i),0);
  const weight= fundTotal? mv/fundTotal*100 : 0;
  body.append(el('div',{class:'cards',style:'margin-bottom:6px'},
    kpi('Зах зээлийн үнэлгээ',ccySym(ccy)+fmtMoney(mv,0),''),
    kpi('Ашиг/Алдагдал %',insGainPct(x).toFixed(2)+'%',(pl<0?'−':'')+ccySym(ccy)+fmtMoney(Math.abs(pl),0)),
    kpi('Сангийн жин · Weight', x.FundID?weight.toFixed(2)+'%':'—', x.FundID?(fundById(x.FundID)||{}).ShortName||'':'хуваарилаагүй'),
    insIsBond(x.InsType)?kpi('YTM',insYTM(x).toFixed(3)+'%','хуримтлагдсан '+ccySym(ccy)+fmtMoney(insAccrued(x),0))
      :kpi('Тоо ширхэг · Quantity',fmtUnits(x.Quantity),'')));

  // facts
  const two=el('div',{class:'panel-body two-col',style:'padding:0'});
  const dl1=el('dl',{class:'def-list'}), dl2=el('dl',{class:'def-list'});
  const add=(dl,k,v)=>{dl.append(el('dt',{},k),el('dd',{},(v===''||v==null)?'—':v));};
  add(dl1,'Төрөл · Type', beforeParen(x.InsType));
  add(dl1,'Гаргагч · Issuer', x.Issuer);
  add(dl1,'Бүртгэл / ISIN', x.ISIN);
  add(dl1,'Валют · Currency', x.Currency||'MNT');
  add(dl1,'Сан · Fund', x.FundID?((fundById(x.FundID)||{}).ShortName||x.FundID):'Хуваарилаагүй');
  add(dl1,'Төлөв · Status', beforeParen(x.Status));
  add(dl2,'Худалдан авсан · Purchased', fmtDate(x.PurchaseDate)+' @ '+ccySym(ccy)+fmtMoney(x.PurchasePrice,4));
  add(dl2,'Тоо ширхэг · Quantity', fmtUnits(x.Quantity));
  add(dl2,'Одоогийн үнэ · Current price', ccySym(ccy)+fmtMoney(insCurPrice(x),4)+(x.CurrentPriceDate?'  ('+fmtDate(x.CurrentPriceDate)+')':''));
  add(dl2,'Өртөг · Cost', ccySym(ccy)+fmtMoney(insCostValue(x),2));
  add(dl2,'Зах зээлийн үнэлгээ · Market value', ccySym(ccy)+fmtMoney(mv,2));
  // type-specific facts
  if(insIsBond(x.InsType)){
    add(dl2,'Нэрлэсэн үнэ · Face', ccySym(ccy)+fmtMoney(x.FaceValue,2));
    add(dl2,'Купон · Coupon', num(x.CouponRate).toFixed(3)+'% · '+beforeParen(x.CouponFreq));
    add(dl2,'Гаргасан/Дуусах · Issue/Maturity', fmtDate(x.IssueDate)+' → '+fmtDate(x.MaturityDate));
  } else if(insIsEquity(x.InsType)){
    add(dl2,'Бирж / Тикэр · Exchange / Ticker', beforeParen(x.Exchange)+(x.Ticker?' · '+x.Ticker:''));
    add(dl2,'Ногдол ашгийн өгөөж · Dividend yield', num(x.DividendYield).toFixed(2)+'%');
  } else if(insIsFundUnit(x.InsType)){
    add(dl2,'Сангийн менежер · Manager', x.FundManager);
    add(dl2,'Сангийн төрөл · Fund type', x.FundUnitType);
  }
  if(x.Notes) add(dl1,'Тэмдэглэл · Notes', x.Notes);
  two.append(dl1,dl2); body.append(two);

  // price history
  body.append(el('div',{style:'display:flex;justify-content:space-between;align-items:center;margin:18px 0 8px'},
    el('h4',{style:'margin:0;color:var(--gam-navy)'},'Үнийн түүх · Price history'),
    el('button',{class:'btn btn-outline btn-sm',onclick:()=>pricePointForm(x.InsID)},'＋ Үнэ нэмэх')));
  const hist=(state.data.PriceHistory||[]).filter(p=>p.InsID===x.InsID)
    .sort((a,b)=>String(b.Date).localeCompare(String(a.Date)));
  const t=el('table',{class:'grid'});
  t.innerHTML=`<thead><tr><th>Огноо · Date</th><th class="num">Үнэ · Price</th><th>Эх сурвалж · Source</th><th>Тэмдэглэл</th></tr></thead>`;
  const tb=el('tbody');
  if(!hist.length) tb.append(emptyRow(4,'Үнийн бүртгэл алга. «Үнэ нэмэх» дарж оруулна уу.'));
  for(const p of hist) tb.append(el('tr',{},
    el('td',{},fmtDate(p.Date)), el('td',{class:'num'},ccySym(ccy)+fmtMoney(p.Price,4)),
    el('td',{},p.Source||''), el('td',{},p.Notes||'')));
  t.append(tb); body.append(el('div',{class:'table-wrap'},t));

  modal('Хэрэгслийн дэлгэрэнгүй · Instrument detail', body, {okText:'Хаах', cancel:false, wide:true});
}

function pricePointForm(insId){
  const x=instrumentById(insId); if(!x)return;
  const dateIn=el('input',{type:'date',value:todayISO()});
  const priceIn=el('input',{type:'number',step:'0.0001',min:'0',value:insCurPrice(x)||''});
  const srcIn=el('input',{type:'text',value:'Гараар · Manual'});
  const noteIn=el('input',{type:'text'});
  const body=el('div',{},
    el('div',{class:'small muted',style:'margin-bottom:8px'},'Хэрэгсэл: ',el('strong',{},x.NameMN||x.InsID)),
    el('div',{class:'form-grid'},
      wrapField('Огноо','Date',dateIn,true),
      wrapField('Үнэ','Price',priceIn,true),
      wrapField('Эх сурвалж','Source',srcIn),
      wrapField('Тэмдэглэл','Notes',noteIn)),
    el('div',{class:'small muted',style:'margin-top:6px'},'Хамгийн сүүлийн огнооны үнэ нь «одоогийн үнэ» болж шинэчлэгдэнэ.'));
  modal('Үнэ нэмэх · Add price', body, {okText:'Нэмэх', onOk:()=>{
    if(!dateIn.value){ toast('Огноо оруулна уу.','err'); return false; }
    if(num(priceIn.value)<=0){ toast('Үнэ оруулна уу.','err'); return false; }
    commit(()=>{
      state.data.PriceHistory.push({PriceID:nextId('PRC',state.data.PriceHistory,'PriceID'),
        InsID:x.InsID, Date:dateIn.value, Price:num(priceIn.value), Source:srcIn.value, Notes:noteIn.value});
      // newest-dated price becomes the current price
      if(!x.CurrentPriceDate || dateIn.value>=x.CurrentPriceDate){
        x.CurrentPrice=num(priceIn.value); x.CurrentPriceDate=dateIn.value;
      }
    },'Үнэ нэмэгдлээ');
    instrumentDetail(x.InsID); return false;
  }});
}

/* ============================================================
   VIEW: Fees — management & performance fee calculation
   ============================================================ */
VIEWS.fees=function(){
  const wrap=el('div',{});
  wrap.append(viewHead('Шимтгэл ба татвар','Fees & withholding tax',
    el('div',{},
      el('button',{class:'btn btn-outline btn-sm',onclick:()=>feeForm()},'＋ ЦХ-ийн шимтгэл (сан)'),' ',
      el('button',{class:'btn btn-primary',onclick:calculateSuccessFees},'⟳ Гүйцэтгэлийн шимтгэл тооцох · Calculate Fees'))));

  wrap.append(el('div',{class:'panel'},el('div',{class:'panel-body'},
    el('div',{class:'def-list',style:'grid-template-columns:auto 1fr'},
      el('dt',{},'Удирдлагын төлбөр'),
      el('dd',{class:'small',style:'font-weight:400'},'Худалдан авалт хийх үед урьдчилан суутгана = Хөрөнгө оруулалт × жилийн удирдлагын шимтгэл %. Тусдаа гүйлгээний мөрөөр бүртгэгдэнэ.'),
      el('dt',{},'Гүйцэтгэлийн төлбөр'),
      el('dd',{class:'small',style:'font-weight:400'},'Хөрөнгө оруулалт хийгдсэнээс хойш 1 жилийн дараа, өгөөж нь сангийн босго өгөөжөөс давсан тохиолдолд: (Бодит өгөөж − Босго) × Хөрөнгө оруулалт × Гүйцэтгэлийн шимтгэл %.'),
      el('dt',{},'Татвар · Tax'),
      el('dd',{class:'small',style:'font-weight:400'},'Өгөөжид ногдох суутгал: Дотоодын (Монгол) 10%, Гадаадын 20%. Гүйцэтгэлийн шимтгэлийн дараа тооцно.'))) ));

  // ---- Summary (collected fees & tax across all funds) ----
  const totMgmt=sumTxnType(TXN_TYPE.MGMT), totSucc=sumTxnType(TXN_TYPE.SUCCESS), totTax=sumTxnType(TXN_TYPE.TAX);
  const breakdown=feeBreakdown();
  const totGross=breakdown.reduce((s,r)=>s+r.grossReturn,0);
  const netDist=totGross-totSucc-totTax;
  wrap.append(el('div',{class:'cards'},
    kpi('Нийт удирдлагын шимтгэл','₮'+fmtMoney(totMgmt,0),'Total mgmt fees collected'),
    kpi('Нийт гүйцэтгэлийн шимтгэл','₮'+fmtMoney(totSucc,0),'Total success fees collected'),
    kpi('Нийт суутгасан татвар','₮'+fmtMoney(totTax,0),'Total tax withheld'),
    kpi('Хөрөнгө оруулагчид цэвэр','₮'+fmtMoney(netDist,0),'Net distributions to investors')));

  // ---- Per-fund, per-investor breakdown ----
  const bp=el('div',{class:'panel'});
  bp.append(el('div',{class:'panel-head'},el('h3',{},'Хөрөнгө оруулагч тус бүрийн задаргаа · Per-investor breakdown')));
  const bt=el('table',{class:'grid'});
  bt.innerHTML=`<thead><tr><th>Сан</th><th>Хөрөнгө оруулагч</th><th class="num">Нийт өгөөж<br><span class="small muted">Gross return</span></th>
    <th class="num">Удирдлага<br><span class="small muted">Mgmt fee</span></th><th class="num">Гүйцэтгэл<br><span class="small muted">Success fee</span></th>
    <th class="num">Татвар<br><span class="small muted">Tax</span></th><th class="num">Цэвэр өгөөж<br><span class="small muted">Net to investor</span></th></tr></thead>`;
  const btb=el('tbody');
  if(!breakdown.length) btb.append(emptyRow(7,'Идэвхтэй эзэмшил алга.'));
  for(const r of breakdown){
    const inv=investorById(r.investorId);
    btb.append(el('tr',{},
      el('td',{},fundChip(r.fundId)),
      el('td',{},inv?inv.NameMN:r.investorId),
      el('td',{class:'num'},(r.grossReturn<0?'−₮':'₮')+fmtMoney(Math.abs(r.grossReturn),0)),
      el('td',{class:'num'},'₮'+fmtMoney(r.mgmtFee,0)),
      el('td',{class:'num'},'₮'+fmtMoney(r.successFee,0)),
      el('td',{class:'num'},'₮'+fmtMoney(r.tax,0)),
      el('td',{class:'num'},el('strong',{},(r.netReturn<0?'−₮':'₮')+fmtMoney(Math.abs(r.netReturn),0)))));
  }
  bt.append(btb); bp.append(el('div',{class:'table-wrap'},bt)); wrap.append(bp);

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

/* ---------- Per-investor fee / tax breakdown ----------
   For every investable fund, list each investor that currently holds units
   or has any fee/tax line booked. Gross return = current value − net invested.
   Management / success fees and tax are summed from the recorded ledger lines. */
function feeBreakdown(){
  const out=[];
  for(const f of investableFunds()){
    const ids=new Set();
    for(const inv of state.data.Investors)
      if(investorUnits(inv.InvestorID,f.FundID)>0.00001) ids.add(inv.InvestorID);
    for(const t of state.data.Transactions)
      if(t.FundID===f.FundID && t.Status!=='Цуцлагдсан'
        && [TXN_TYPE.MGMT,TXN_TYPE.SUCCESS,TXN_TYPE.TAX].includes(t.Type)) ids.add(t.InvestorID);
    for(const id of ids){
      const units=investorUnits(id,f.FundID);
      const invested=investorNetCash(id,f.FundID);
      const grossReturn=units*latestUnitPrice(f.FundID)-invested;
      const mgmtFee=sumTxnType(TXN_TYPE.MGMT,id,f.FundID);
      const successFee=sumTxnType(TXN_TYPE.SUCCESS,id,f.FundID);
      const tax=sumTxnType(TXN_TYPE.TAX,id,f.FundID);
      out.push({fundId:f.FundID,investorId:id,units,invested,grossReturn,mgmtFee,successFee,tax,
        netReturn:grossReturn-successFee-tax});
    }
  }
  return out.sort((a,b)=>String(a.fundId).localeCompare(String(b.fundId))
    ||String(a.investorId).localeCompare(String(b.investorId)));
}

/* ---------- Calculate pending success fees & tax ----------
   For every confirmed subscription lot at least one year old that has not yet
   been processed, charge a success fee when the lot's return exceeds the fund
   hurdle, then withhold tax on the net return. Each booking references its
   source subscription (PaymentRef) so re-running never double-charges a lot. */
function calculateSuccessFees(){
  const today=todayISO();
  const processed=new Set(state.data.Transactions
    .filter(t=>t.Type===TXN_TYPE.SUCCESS||t.Type===TXN_TYPE.TAX).map(t=>t.PaymentRef));
  const subs=state.data.Transactions.filter(t=>t.Type===TXN_TYPE.SUB && t.Status==='Баталгаажсан'
    && !processed.has(t.TxnID) && isOneYearOld(t.TradeDate,today));
  if(!subs.length){
    toast('Тооцоолох шинэ хөрөнгө оруулалт алга (1 жил болоогүй эсвэл аль хэдийн тооцсон).','ok','Мэдээлэл');
    return;
  }
  const plan=[]; let nSucc=0, nTax=0, sumSucc=0, sumTax=0;
  for(const s of subs){
    const f=fundById(s.FundID); if(!f) continue;
    const inv=investorById(s.InvestorID);
    const buyPrice=num(s.UnitPrice), units=num(s.Units), investment=num(s.NetAmount);
    if(buyPrice<=0||units<=0) continue;
    const curPrice=latestUnitPrice(s.FundID);
    const actualReturn=(curPrice-buyPrice)/buyPrice;
    const hurdle=num(f.HurdleRatePct)/100;
    const successRate=num(f.PerfFeePct)/100;
    const grossReturn=(curPrice-buyPrice)*units;
    const successFee= actualReturn>hurdle ? (actualReturn-hurdle)*investment*successRate : 0;
    const taxRate=investorTaxRate(inv);
    const tax=Math.max(0,grossReturn-successFee)*taxRate;
    if(successFee>0){
      plan.push({type:TXN_TYPE.SUCCESS,s,amount:successFee,feePct:num(f.PerfFeePct),
        note:`Гүйцэтгэлийн шимтгэл · өгөөж ${(actualReturn*100).toFixed(1)}% > босго ${(hurdle*100)}% · ${s.TxnID}`});
      sumSucc+=successFee; nSucc++;
    }
    if(tax>0){
      plan.push({type:TXN_TYPE.TAX,s,amount:tax,feePct:taxRate*100,
        note:`Өгөөжийн татвар ${(taxRate*100)}% · ${s.TxnID}`});
      sumTax+=tax; nTax++;
    }
  }
  if(!plan.length){
    toast('1 жил болсон хөрөнгө оруулалт байна, гэвч босго өгөөжид хүрээгүй тул шимтгэл/татвар үүсээгүй.','ok','Мэдээлэл');
    return;
  }
  commit(()=>{
    for(const p of plan){
      state.data.Transactions.push({
        TxnID:nextId('TXN',state.data.Transactions,'TxnID'), TradeDate:today, Type:p.type,
        FundID:p.s.FundID, InvestorID:p.s.InvestorID, Units:0, UnitPrice:0, GrossAmount:p.amount,
        FeePct:p.feePct, FeeAmount:p.amount, NetAmount:p.amount, SettlementDate:today,
        Status:'Баталгаажсан', PaymentRef:p.s.TxnID, Notes:p.note, CreatedAt:new Date().toISOString(),
      });
    }
  }, `${nSucc} гүйцэтгэлийн шимтгэл (₮${fmtMoney(sumSucc,0)}) · ${nTax} татвар (₮${fmtMoney(sumTax,0)}) бүртгэгдлээ`);
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
  const f=existing||{FirmID:nextId('FIRM',state.data.ECIFFirms,'FirmID'),JoinDate:todayISO(),PortfolioVariant:'Хувилбар 1 — Бага эрсдэл (Бонд 85 / Бэлэн мөнгө 15)',Status:'Идэвхтэй'};
  const g=(name,mn,en,opts={})=>fieldHTML(name,mn,en,f[name],opts);
  const body=el('div',{},el('div',{class:'form-grid'},
    g('FirmID','Дугаар','ID',{readonly:true}),
    g('JoinDate','Элссэн огноо','Join date',{type:'date'}),
    g('NameMN','Байгууллагын нэр (Монгол)','Name (MN)',{required:true,full:true}),
    g('NameEN','Нэр (English)','Name (EN)',{full:true}),
    g('RegNo','Улсын бүртгэлийн дугаар','Registration no.'),
    g('EmployeeCount','Нийт ажилтан (тоо)','Headcount',{type:'number'}),
    g('PortfolioVariant','Багцын хувилбар','Portfolio variant',{type:'select',
      options:['Хувилбар 1 — Бага эрсдэл (Бонд 85 / Бэлэн мөнгө 15)','Хувилбар 2 — Дунд эрсдэл (Бонд 40 / Олон улсын сангийн нэгж 25 / Хувьцаа 20 / Бэлэн мөнгө 15)'],full:true}),
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
