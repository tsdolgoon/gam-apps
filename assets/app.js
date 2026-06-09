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
  data: { Funds:[], Investors:[], Transactions:[], NAVHistory:[], Bonds:[], Fees:[], ECIFFirms:[], ECIFEmployees:[], ECIFContributions:[], Meta:[] },
  dirty: false,
  view: 'dashboard',
  savedAt: null,
};

/* ---------- Bond reference data ---------- */
const BOND_TYPES   = ['Засгийн газрын (Government)','Корпорацийн (Corporate)','Банкны (Bank)'];
const COUPON_FREQ  = ['Сар бүр (Monthly)','Улирал бүр (Quarterly)','Жилд нэг (Annually)'];
const BOND_CCY     = ['MNT','USD','CNY','EUR','JPY'];
const BOND_STATUS  = ['Идэвхтэй','Дууссан','Зарсан']; // Active / Matured / Sold

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
  state.data = { Funds:[], Investors:[], Transactions:[], NAVHistory:[], Bonds:[], Fees:[],
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
    Investors:[], Transactions:[], NAVHistory:[], Bonds:[], Fees:[],
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
  refreshBondComputed(); // snapshot auto-computed bond metrics into the workbook
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
