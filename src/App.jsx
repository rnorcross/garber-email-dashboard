import { useState, useMemo, useEffect, useCallback, useRef } from "react";
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Cell, ResponsiveContainer, Legend } from "recharts";
import Papa from "papaparse";
import html2canvas from "html2canvas";
import jsPDF from "jspdf";
import JSZip from "jszip";

/* ───────────────── CONSTANTS ───────────────── */
const SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTYVu1c2AZGFprO4Qk2sgvY6GDl1PuBAxW-7J5xg4xjIrz-ZCTaxn2oC2vVfCxECbGqOt6e9KkgAjHs/pub?output=csv";
const SHEET_HTML_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTYVu1c2AZGFprO4Qk2sgvY6GDl1PuBAxW-7J5xg4xjIrz-ZCTaxn2oC2vVfCxECbGqOt6e9KkgAjHs/pubhtml";

const MN = ["","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const MONTH_MAP = {January:1,February:2,March:3,April:4,May:5,June:6,July:7,August:8,September:9,October:10,November:11,December:12,
  Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};

const C = { main:"#072a60", sec:"#0d6eff", acc1:"#1e88e5", green:"#0a8754", amber:"#c9710d", red:"#c0392b", bg:"#edf1f7", purple:"#7c3aed", teal:"#0d9488" };

const TABS = [
  { id:"groupSales", label:"Group Sales Email" },
  { id:"groupService", label:"Group Service Email" },
  { id:"groupAds", label:"Group Advertising" },
  { id:"dealerSales", label:"Dealership Sales" },
  { id:"dealerService", label:"Dealership Service + Google Ads" },
  { id:"dealerFB", label:"Dealership Facebook Ads" },
  { id:"customerData", label:"Customer Data" },
];

/* ───────────────── CSV PARSING ───────────────── */
function findCol(headers, ...needles) {
  return headers.findIndex(h => {
    const lc = (h||"").toLowerCase().replace(/[^a-z0-9%$]/g,"");
    return needles.some(n => lc.includes(n.toLowerCase().replace(/[^a-z0-9%$]/g,"")));
  });
}

function num(v) {
  if (v === undefined || v === null || v === "") return 0;
  const s = String(v).replace(/[$,\s%]/g, "");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function parseOverallCSV(text) {
  const result = Papa.parse(text, { header: false, skipEmptyLines: true });
  if (!result.data || result.data.length < 2) return [];
  const hdr = result.data[0].map(h => (h||"").trim());
  const ci = {
    month: findCol(hdr, "month"),
    year: findCol(hdr, "year"),
    location: findCol(hdr, "location"),
    monthlyCost: findCol(hdr, "monthlycost", "cost"),
    salesEngaged: findCol(hdr, "salesengaged", "engaged"),
    salesShoppers: findCol(hdr, "salesshoppers"),
    salesLeads: findCol(hdr, "salesleads"),
    salesInfluenced: findCol(hdr, "salesinfluenced", "influenced"),
    salesWinback: findCol(hdr, "saleswinback", "winback"),
    salesGross: findCol(hdr, "salesgross", "gross"),
    salesWinbackProfit: findCol(hdr, "winbackprofit", "saleswinbackprofit"),
    salesROI: findCol(hdr, "salesroi", "roi%"),
    winbackSalesPct: findCol(hdr, "winbacksales%", "winbacksalespct"),
    serviceShoppers: findCol(hdr, "serviceshoppers"),
    serviceLeads: findCol(hdr, "serviceleads"),
    serviceROs: findCol(hdr, "servicero", "serviceROs"),
    serviceWinbackROs: findCol(hdr, "servicewinbackro"),
    roValue: findCol(hdr, "rovalue"),
    serviceWinbackROValue: findCol(hdr, "servicewinbackrovalue", "winbackrovalue"),
    winbackServicePct: findCol(hdr, "winbackservice%", "winbackservicepct"),
    clicksGoogle: findCol(hdr, "clicksgoogle", "clicksgoog"),
    pageViewsGoogle: findCol(hdr, "pageviewsgoogle", "viewsgoogle"),
    leadsGoogle: findCol(hdr, "leadsgoogle"),
    phoneCallsGoogle: findCol(hdr, "phonecallsgoogle", "callsgoogle"),
    spendGoogle: findCol(hdr, "spendgoogle"),
    clicksFB: findCol(hdr, "clicksfacebook", "clicksfb"),
    pageViewsFB: findCol(hdr, "pageviewsfacebook", "viewsfacebook"),
    leadsFB: findCol(hdr, "leadsfacebook", "leadsfb"),
    phoneCallsFB: findCol(hdr, "phonecallsfacebook", "callsfacebook"),
    spendFB: findCol(hdr, "spendfacebook", "spendfb"),
  };
  // If column-index search didn't work well, try positional (columns in order described)
  // Fallback: just use index positions based on user's described column order
  if (ci.month < 0) ci.month = 0;
  if (ci.year < 0) ci.year = 1;
  if (ci.location < 0) ci.location = 2;
  if (ci.monthlyCost < 0) ci.monthlyCost = 3;

  return result.data.slice(1).map(r => {
    const monthRaw = (r[ci.month]||"").trim();
    const m = MONTH_MAP[monthRaw] || parseInt(monthRaw) || 0;
    const y = parseInt(r[ci.year]) || 0;
    if (!m || !y) return null;
    return {
      month: m, year: y,
      location: (r[ci.location]||"").trim(),
      monthlyCost: num(r[ci.monthlyCost]),
      salesEngaged: num(r[ci.salesEngaged]),
      salesShoppers: num(r[ci.salesShoppers]),
      salesLeads: num(r[ci.salesLeads]),
      salesInfluenced: num(r[ci.salesInfluenced]),
      salesWinback: num(r[ci.salesWinback]),
      salesGross: num(r[ci.salesGross]),
      salesWinbackProfit: num(r[ci.salesWinbackProfit]),
      salesROI: num(r[ci.salesROI]),
      winbackSalesPct: num(r[ci.winbackSalesPct]),
      serviceShoppers: num(r[ci.serviceShoppers]),
      serviceLeads: num(r[ci.serviceLeads]),
      serviceROs: num(r[ci.serviceROs]),
      serviceWinbackROs: num(r[ci.serviceWinbackROs]),
      roValue: num(r[ci.roValue]),
      serviceWinbackROValue: num(r[ci.serviceWinbackROValue]),
      winbackServicePct: num(r[ci.winbackServicePct]),
      clicksGoogle: num(r[ci.clicksGoogle]),
      pageViewsGoogle: num(r[ci.pageViewsGoogle]),
      leadsGoogle: num(r[ci.leadsGoogle]),
      phoneCallsGoogle: num(r[ci.phoneCallsGoogle]),
      spendGoogle: num(r[ci.spendGoogle]),
      clicksFB: num(r[ci.clicksFB]),
      pageViewsFB: num(r[ci.pageViewsFB]),
      leadsFB: num(r[ci.leadsFB]),
      phoneCallsFB: num(r[ci.phoneCallsFB]),
      spendFB: num(r[ci.spendFB]),
    };
  }).filter(Boolean);
}

/* ───────────────── HELPERS ───────────────── */
function getPeriods(data) {
  const set = new Set(data.map(r => `${r.year}-${String(r.month).padStart(2,"0")}`));
  return [...set].sort().map(p => {
    const [y,m] = p.split("-").map(Number);
    return { year:y, month:m, label:`${MN[m]} ${y}`, short:`${MN[m]} '${String(y).slice(2)}`, key:p };
  });
}
function getPrev(periods, key) { const i = periods.findIndex(p => p.key === key); return i > 0 ? periods[i-1] : null; }
function fmtMoney(v) { if (v >= 1000000) return `$${(v/1000000).toFixed(1)}M`; if (v >= 1000) return `$${(v/1000).toFixed(1)}K`; return `$${v.toLocaleString()}`; }
function fmtNum(v) { return v.toLocaleString(); }
function fmtPct(v) { return `${v.toFixed(1)}%`; }

function filterPeriod(data, p) { return data.filter(r => r.year === p.year && r.month === p.month); }

function aggSales(rows) {
  const o = { monthlyCost:0, salesEngaged:0, salesShoppers:0, salesLeads:0, salesInfluenced:0, salesWinback:0, salesGross:0, salesWinbackProfit:0 };
  rows.forEach(r => { for (const k in o) o[k] += r[k]; });
  o.salesROI = o.monthlyCost > 0 ? ((o.salesGross / o.monthlyCost) * 100) : 0;
  o.winbackSalesPct = o.salesInfluenced > 0 ? ((o.salesWinback / o.salesInfluenced) * 100) : 0;
  return o;
}
function aggService(rows) {
  const o = { serviceShoppers:0, serviceLeads:0, serviceROs:0, serviceWinbackROs:0, roValue:0, serviceWinbackROValue:0 };
  rows.forEach(r => { for (const k in o) o[k] += r[k]; });
  o.winbackServicePct = o.serviceROs > 0 ? ((o.serviceWinbackROs / o.serviceROs) * 100) : 0;
  return o;
}
function aggGoogle(rows) {
  const o = { clicksGoogle:0, pageViewsGoogle:0, leadsGoogle:0, phoneCallsGoogle:0, spendGoogle:0 };
  rows.forEach(r => { for (const k in o) o[k] += r[k]; });
  o.cpl = o.leadsGoogle > 0 ? (o.spendGoogle / o.leadsGoogle) : 0;
  return o;
}
function aggFB(rows) {
  const o = { clicksFB:0, pageViewsFB:0, leadsFB:0, phoneCallsFB:0, spendFB:0 };
  rows.forEach(r => { for (const k in o) o[k] += r[k]; });
  o.cpl = o.leadsFB > 0 ? (o.spendFB / o.leadsFB) : 0;
  return o;
}

/* ───────────────── REUSABLE COMPONENTS ───────────────── */
const Badge = ({ cur, prev }) => {
  if (prev === null || prev === undefined) return <span style={{color:"#9aa",fontSize:14}}>—</span>;
  const d = cur - prev;
  if (d === 0) return <span style={{color:"#8896a4",fontSize:14,fontWeight:600}}>—</span>;
  const up = d > 0;
  return <span style={{fontSize:14,fontWeight:700,color:up?C.green:C.red,whiteSpace:"nowrap"}}>{up?"▲":"▼"} {up?"+":""}{typeof cur==="number"&&cur%1!==0?d.toFixed(1):d.toLocaleString()}</span>;
};

const MoneyBadge = ({ cur, prev }) => {
  if (prev === null || prev === undefined) return <span style={{color:"#9aa",fontSize:14}}>—</span>;
  const d = cur - prev;
  if (d === 0) return <span style={{color:"#8896a4",fontSize:14,fontWeight:600}}>—</span>;
  const up = d > 0;
  return <span style={{fontSize:14,fontWeight:700,color:up?C.green:C.red,whiteSpace:"nowrap"}}>{up?"▲":"▼"} {up?"+":""}{fmtMoney(d)}</span>;
};

const KPI = ({ label, value, fmt="num", color=C.main, sub }) => (
  <div style={{background:"white",borderRadius:14,padding:"22px 24px",minWidth:150,flex:1,boxShadow:"0 2px 8px rgba(0,0,0,0.06)",borderTop:`4px solid ${color}`}}>
    <div style={{fontSize:12,color:"#7a8a9a",fontWeight:600,textTransform:"uppercase",letterSpacing:1.1,marginBottom:6}}>{label}</div>
    <div style={{fontSize:32,fontWeight:800,color,lineHeight:1.1}}>
      {fmt==="money" ? fmtMoney(value) : fmt==="pct" ? fmtPct(value) : fmtNum(value)}
    </div>
    {sub && <div style={{fontSize:13,color:"#7a8a9a",marginTop:6}}>{sub}</div>}
  </div>
);

const SortHeader = ({ label, sortKey, sortState, onSort, align="right", first, last }) => (
  <th onClick={() => onSort(sortKey)} style={{padding:"11px 12px",fontWeight:700,fontSize:13,textAlign:align,cursor:"pointer",userSelect:"none",whiteSpace:"nowrap",background:C.main,color:"white",borderRadius:first?"10px 0 0 0":last?"0 10px 0 0":undefined,position:"sticky",top:0,zIndex:2}}>
    {label}{sortState.key===sortKey?(sortState.dir==="asc"?" ↑":" ↓"):""}
  </th>
);

function useSort(defaultKey, defaultDir="desc") {
  const [s, setS] = useState({ key:defaultKey, dir:defaultDir });
  const onSort = useCallback(key => setS(prev => prev.key===key?{key,dir:prev.dir==="desc"?"asc":"desc"}:{key,dir:"desc"}), []);
  const doSort = useCallback(arr => [...arr].sort((a,b) => {
    const av=a[s.key]??0, bv=b[s.key]??0;
    if (typeof av==="string") return s.dir==="asc"?av.localeCompare(bv):bv.localeCompare(av);
    return s.dir==="asc"?av-bv:bv-av;
  }), [s]);
  return { sortState:s, onSort, doSort };
}

const PeriodSelector = ({ periods, sp, setSp }) => (
  <select value={sp} onChange={e=>setSp(e.target.value)} style={{padding:"10px 16px",borderRadius:10,border:`2px solid ${C.sec}`,fontSize:15,fontWeight:700,color:C.main,background:"white",cursor:"pointer",outline:"none"}}>
    {periods.map(p => <option key={p.key} value={p.key}>{p.label}</option>)}
  </select>
);

const TrendChart = ({ data, periods, valueKey, label, color, fmt="num", aggFn }) => {
  const cd = periods.map(p => {
    const rows = filterPeriod(data, p);
    const a = aggFn(rows);
    return { name: p.short, value: a[valueKey] || 0 };
  });
  return (
    <div style={{background:"white",borderRadius:14,padding:"20px 22px",boxShadow:"0 2px 8px rgba(0,0,0,0.06)",flex:1,minWidth:340}}>
      <div style={{fontSize:15,fontWeight:700,color:C.main,marginBottom:12}}>{label}</div>
      <ResponsiveContainer width="100%" height={220}>
        <LineChart data={cd} margin={{top:5,right:20,bottom:5,left:10}}>
          <CartesianGrid strokeDasharray="3 3" stroke="#e8ecf0"/>
          <XAxis dataKey="name" tick={{fontSize:12,fill:C.main,fontWeight:600}} interval={0}/>
          <YAxis tick={{fontSize:11,fill:"#8896a4"}} tickFormatter={v=>fmt==="money"?fmtMoney(v):fmt==="pct"?`${v.toFixed(0)}%`:v>=1000?`${(v/1000).toFixed(1)}k`:v} width={60}/>
          <Tooltip formatter={v=>fmt==="money"?fmtMoney(Number(v)):fmt==="pct"?fmtPct(Number(v)):Number(v).toLocaleString()} contentStyle={{borderRadius:10,border:`1px solid ${C.sec}`,fontSize:13}}/>
          <Line type="monotone" dataKey="value" stroke={color} strokeWidth={3} dot={{r:4,fill:color,stroke:"white",strokeWidth:2}} activeDot={{r:6}}/>
        </LineChart>
      </ResponsiveContainer>
    </div>
  );
};

const BarChartPanel = ({ data, periods, keys, labels, colors, title, fmt="num" }) => {
  const cd = periods.map(p => {
    const rows = filterPeriod(data, p);
    const obj = { name: p.short };
    keys.forEach((k, i) => {
      const a = { ...aggGoogle(rows), ...aggFB(rows), ...aggSales(rows), ...aggService(rows) };
      obj[k] = a[k] || 0;
    });
    return obj;
  });
  return (
    <div style={{background:"white",borderRadius:14,padding:"20px 22px",boxShadow:"0 2px 8px rgba(0,0,0,0.06)",flex:1,minWidth:340}}>
      <div style={{fontSize:15,fontWeight:700,color:C.main,marginBottom:12}}>{title}</div>
      <ResponsiveContainer width="100%" height={240}>
        <BarChart data={cd} margin={{top:5,right:20,bottom:5,left:10}}>
          <CartesianGrid strokeDasharray="3 3" stroke="#e8ecf0"/>
          <XAxis dataKey="name" tick={{fontSize:12,fill:C.main,fontWeight:600}} interval={0}/>
          <YAxis tick={{fontSize:11,fill:"#8896a4"}} tickFormatter={v=>fmt==="money"?fmtMoney(v):v>=1000?`${(v/1000).toFixed(1)}k`:v} width={60}/>
          <Tooltip formatter={v=>fmt==="money"?fmtMoney(Number(v)):Number(v).toLocaleString()} contentStyle={{borderRadius:10,border:`1px solid ${C.sec}`,fontSize:13}}/>
          <Legend wrapperStyle={{fontSize:12}}/>
          {keys.map((k,i) => <Bar key={k} dataKey={k} name={labels[i]} fill={colors[i]} radius={[4,4,0,0]}/>)}
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
};

const Section = ({ title, desc, children }) => (
  <div style={{background:"white",borderRadius:14,padding:24,boxShadow:"0 2px 8px rgba(0,0,0,0.06)",marginBottom:20}}>
    <div style={{fontSize:17,fontWeight:700,marginBottom:4,color:C.main}}>{title}</div>
    {desc && <div style={{fontSize:13,color:"#7a8a9a",marginBottom:16}}>{desc}</div>}
    {children}
  </div>
);

/* ──────────── TAB 1: GROUP SALES EMAIL ──────────── */
function GroupSalesTab({ data, periods, sp, setSp }) {
  const p = periods.find(pp => pp.key === sp);
  const prev = getPrev(periods, sp);
  const cur = p ? aggSales(filterPeriod(data, p)) : null;
  const prv = prev ? aggSales(filterPeriod(data, prev)) : null;
  if (!cur) return <div style={{padding:40,color:"#999",textAlign:"center"}}>No data for selected period.</div>;
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:16,marginBottom:20,flexWrap:"wrap"}}>
        <PeriodSelector periods={periods} sp={sp} setSp={setSp}/>
        <span style={{fontSize:14,color:"#7a8a9a"}}>Showing group-level email marketing results for <strong>Sales</strong></span>
      </div>
      <div style={{display:"flex",gap:14,flexWrap:"wrap",marginBottom:20}}>
        <KPI label="Sales Engaged" value={cur.salesEngaged} color={C.main} sub={prv?<Badge cur={cur.salesEngaged} prev={prv.salesEngaged}/>:null}/>
        <KPI label="Sales Shoppers" value={cur.salesShoppers} color={C.acc1} sub={prv?<Badge cur={cur.salesShoppers} prev={prv.salesShoppers}/>:null}/>
        <KPI label="Sales Leads" value={cur.salesLeads} color={C.sec} sub={prv?<Badge cur={cur.salesLeads} prev={prv.salesLeads}/>:null}/>
        <KPI label="Sales Influenced" value={cur.salesInfluenced} color={C.green} sub={prv?<Badge cur={cur.salesInfluenced} prev={prv.salesInfluenced}/>:null}/>
        <KPI label="Sales Winback" value={cur.salesWinback} color={C.purple} sub={prv?<Badge cur={cur.salesWinback} prev={prv.salesWinback}/>:null}/>
      </div>
      <div style={{display:"flex",gap:14,flexWrap:"wrap",marginBottom:20}}>
        <KPI label="Sales Gross" value={cur.salesGross} fmt="money" color={C.green} sub={prv?<MoneyBadge cur={cur.salesGross} prev={prv.salesGross}/>:null}/>
        <KPI label="Winback Profit" value={cur.salesWinbackProfit} fmt="money" color={C.purple} sub={prv?<MoneyBadge cur={cur.salesWinbackProfit} prev={prv.salesWinbackProfit}/>:null}/>
        <KPI label="Monthly Cost" value={cur.monthlyCost} fmt="money" color={C.amber}/>
        <KPI label="Sales ROI %" value={cur.salesROI} fmt="pct" color={cur.salesROI>=100?C.green:C.red}/>
        <KPI label="Winback Sales %" value={cur.winbackSalesPct} fmt="pct" color={C.teal}/>
      </div>
      <div style={{display:"flex",gap:16,flexWrap:"wrap",marginBottom:20}}>
        <TrendChart data={data} periods={periods} valueKey="salesInfluenced" label="Sales Influenced — Trend" color={C.green} aggFn={aggSales}/>
        <TrendChart data={data} periods={periods} valueKey="salesGross" label="Sales Gross — Trend" color={C.sec} fmt="money" aggFn={aggSales}/>
      </div>
      <div style={{display:"flex",gap:16,flexWrap:"wrap"}}>
        <TrendChart data={data} periods={periods} valueKey="salesLeads" label="Sales Leads — Trend" color={C.acc1} aggFn={aggSales}/>
        <TrendChart data={data} periods={periods} valueKey="salesROI" label="Sales ROI % — Trend" color={C.amber} fmt="pct" aggFn={aggSales}/>
      </div>
    </div>
  );
}

/* ──────────── TAB 2: GROUP SERVICE EMAIL ──────────── */
function GroupServiceTab({ data, periods, sp, setSp }) {
  const p = periods.find(pp => pp.key === sp);
  const prev = getPrev(periods, sp);
  const cur = p ? aggService(filterPeriod(data, p)) : null;
  const prv = prev ? aggService(filterPeriod(data, prev)) : null;
  if (!cur) return <div style={{padding:40,color:"#999",textAlign:"center"}}>No data for selected period.</div>;
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:16,marginBottom:20,flexWrap:"wrap"}}>
        <PeriodSelector periods={periods} sp={sp} setSp={setSp}/>
        <span style={{fontSize:14,color:"#7a8a9a"}}>Showing group-level email marketing results for <strong>Service</strong></span>
      </div>
      <div style={{display:"flex",gap:14,flexWrap:"wrap",marginBottom:20}}>
        <KPI label="Service Shoppers" value={cur.serviceShoppers} color={C.main} sub={prv?<Badge cur={cur.serviceShoppers} prev={prv.serviceShoppers}/>:null}/>
        <KPI label="Service Leads" value={cur.serviceLeads} color={C.sec} sub={prv?<Badge cur={cur.serviceLeads} prev={prv.serviceLeads}/>:null}/>
        <KPI label="Service RO's" value={cur.serviceROs} color={C.green} sub={prv?<Badge cur={cur.serviceROs} prev={prv.serviceROs}/>:null}/>
        <KPI label="Service Winback RO's" value={cur.serviceWinbackROs} color={C.purple} sub={prv?<Badge cur={cur.serviceWinbackROs} prev={prv.serviceWinbackROs}/>:null}/>
      </div>
      <div style={{display:"flex",gap:14,flexWrap:"wrap",marginBottom:20}}>
        <KPI label="RO Value" value={cur.roValue} fmt="money" color={C.green} sub={prv?<MoneyBadge cur={cur.roValue} prev={prv.roValue}/>:null}/>
        <KPI label="Winback RO Value" value={cur.serviceWinbackROValue} fmt="money" color={C.purple} sub={prv?<MoneyBadge cur={cur.serviceWinbackROValue} prev={prv.serviceWinbackROValue}/>:null}/>
        <KPI label="Winback Service %" value={cur.winbackServicePct} fmt="pct" color={C.teal}/>
      </div>
      <div style={{display:"flex",gap:16,flexWrap:"wrap",marginBottom:20}}>
        <TrendChart data={data} periods={periods} valueKey="serviceROs" label="Service RO's — Trend" color={C.green} aggFn={aggService}/>
        <TrendChart data={data} periods={periods} valueKey="roValue" label="RO Value — Trend" color={C.sec} fmt="money" aggFn={aggService}/>
      </div>
      <div style={{display:"flex",gap:16,flexWrap:"wrap"}}>
        <TrendChart data={data} periods={periods} valueKey="serviceLeads" label="Service Leads — Trend" color={C.acc1} aggFn={aggService}/>
        <TrendChart data={data} periods={periods} valueKey="winbackServicePct" label="Winback Service % — Trend" color={C.purple} fmt="pct" aggFn={aggService}/>
      </div>
    </div>
  );
}

/* ──────────── TAB 3: GROUP ADVERTISING ──────────── */
function GroupAdsTab({ data, periods, sp, setSp }) {
  const p = periods.find(pp => pp.key === sp);
  const prev = getPrev(periods, sp);
  const gCur = p ? aggGoogle(filterPeriod(data, p)) : null;
  const gPrv = prev ? aggGoogle(filterPeriod(data, prev)) : null;
  const fCur = p ? aggFB(filterPeriod(data, p)) : null;
  const fPrv = prev ? aggFB(filterPeriod(data, prev)) : null;
  if (!gCur || !fCur) return <div style={{padding:40,color:"#999",textAlign:"center"}}>No data for selected period.</div>;
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:16,marginBottom:20,flexWrap:"wrap"}}>
        <PeriodSelector periods={periods} sp={sp} setSp={setSp}/>
        <span style={{fontSize:14,color:"#7a8a9a"}}>Showing group-level advertising results — Google Ads (Service) & Facebook Ads (Sales)</span>
      </div>
      <Section title="Google Ads — Service" desc="Performance metrics for Google search ads promoting service.">
        <div style={{display:"flex",gap:14,flexWrap:"wrap",marginBottom:16}}>
          <KPI label="Clicks" value={gCur.clicksGoogle} color={C.main} sub={gPrv?<Badge cur={gCur.clicksGoogle} prev={gPrv.clicksGoogle}/>:null}/>
          <KPI label="Page Views" value={gCur.pageViewsGoogle} color={C.acc1} sub={gPrv?<Badge cur={gCur.pageViewsGoogle} prev={gPrv.pageViewsGoogle}/>:null}/>
          <KPI label="Leads" value={gCur.leadsGoogle} color={C.sec} sub={gPrv?<Badge cur={gCur.leadsGoogle} prev={gPrv.leadsGoogle}/>:null}/>
          <KPI label="Phone Calls" value={gCur.phoneCallsGoogle} color={C.green} sub={gPrv?<Badge cur={gCur.phoneCallsGoogle} prev={gPrv.phoneCallsGoogle}/>:null}/>
          <KPI label="Spend" value={gCur.spendGoogle} fmt="money" color={C.amber}/>
          <KPI label="Cost/Lead" value={gCur.cpl} fmt="money" color={gCur.cpl<=50?C.green:gCur.cpl<=100?C.amber:C.red}/>
        </div>
      </Section>
      <Section title="Facebook Ads — Sales" desc="Performance metrics for Facebook ads promoting sales.">
        <div style={{display:"flex",gap:14,flexWrap:"wrap",marginBottom:16}}>
          <KPI label="Clicks" value={fCur.clicksFB} color={C.main} sub={fPrv?<Badge cur={fCur.clicksFB} prev={fPrv.clicksFB}/>:null}/>
          <KPI label="Page Views" value={fCur.pageViewsFB} color={C.acc1} sub={fPrv?<Badge cur={fCur.pageViewsFB} prev={fPrv.pageViewsFB}/>:null}/>
          <KPI label="Leads" value={fCur.leadsFB} color={C.sec} sub={fPrv?<Badge cur={fCur.leadsFB} prev={fPrv.leadsFB}/>:null}/>
          <KPI label="Phone Calls" value={fCur.phoneCallsFB} color={C.green} sub={fPrv?<Badge cur={fCur.phoneCallsFB} prev={fPrv.phoneCallsFB}/>:null}/>
          <KPI label="Spend" value={fCur.spendFB} fmt="money" color={C.amber}/>
          <KPI label="Cost/Lead" value={fCur.cpl} fmt="money" color={fCur.cpl<=50?C.green:fCur.cpl<=100?C.amber:C.red}/>
        </div>
      </Section>
      <div style={{display:"flex",gap:16,flexWrap:"wrap",marginBottom:20}}>
        <TrendChart data={data} periods={periods} valueKey="clicksGoogle" label="Google Clicks — Trend" color={C.main} aggFn={aggGoogle}/>
        <TrendChart data={data} periods={periods} valueKey="clicksFB" label="Facebook Clicks — Trend" color={C.sec} aggFn={aggFB}/>
      </div>
      <div style={{display:"flex",gap:16,flexWrap:"wrap"}}>
        <TrendChart data={data} periods={periods} valueKey="spendGoogle" label="Google Spend — Trend" color={C.amber} fmt="money" aggFn={aggGoogle}/>
        <TrendChart data={data} periods={periods} valueKey="spendFB" label="Facebook Spend — Trend" color={C.red} fmt="money" aggFn={aggFB}/>
      </div>
    </div>
  );
}

/* ──────────── TAB 4: DEALERSHIP SALES EMAIL ──────────── */
function DealerSalesTab({ data, periods, sp, setSp, locations }) {
  const { sortState, onSort, doSort } = useSort("salesInfluenced","desc");
  const p = periods.find(pp => pp.key === sp);
  const prev = getPrev(periods, sp);
  const rows = locations.map(loc => {
    const cur = p ? aggSales(filterPeriod(data, p).filter(r => r.location === loc)) : { salesEngaged:0,salesShoppers:0,salesLeads:0,salesInfluenced:0,salesWinback:0,salesGross:0,salesWinbackProfit:0,salesROI:0,winbackSalesPct:0,monthlyCost:0 };
    const prv = prev ? aggSales(filterPeriod(data, prev).filter(r => r.location === loc)) : null;
    return { location:loc, ...cur, prevInfluenced: prv?.salesInfluenced ?? null };
  });
  const sorted = doSort(rows);
  const tot = { salesEngaged:0,salesShoppers:0,salesLeads:0,salesInfluenced:0,salesWinback:0,salesGross:0,salesWinbackProfit:0 };
  rows.forEach(r => { for (const k in tot) tot[k] += r[k]; });
  const td = i => ({padding:"10px 12px",fontSize:13,background:i%2===0?"#f4f7fc":"white"});
  const totStyle = {padding:"11px 12px",fontSize:13,fontWeight:800,background:"#e8edf5",borderTop:`2px solid ${C.main}`};
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:16,marginBottom:20,flexWrap:"wrap"}}>
        <PeriodSelector periods={periods} sp={sp} setSp={setSp}/>
        <span style={{fontSize:14,color:"#7a8a9a"}}>Email marketing <strong>Sales</strong> results by dealership</span>
      </div>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"separate",borderSpacing:0}}>
          <thead><tr>
            <th style={{padding:"11px 12px",fontWeight:700,fontSize:13,textAlign:"left",background:C.main,color:"white",borderRadius:"10px 0 0 0"}}>#</th>
            <SortHeader label="Dealership" sortKey="location" sortState={sortState} onSort={onSort} align="left"/>
            <SortHeader label="Engaged" sortKey="salesEngaged" sortState={sortState} onSort={onSort}/>
            <SortHeader label="Shoppers" sortKey="salesShoppers" sortState={sortState} onSort={onSort}/>
            <SortHeader label="Leads" sortKey="salesLeads" sortState={sortState} onSort={onSort}/>
            <SortHeader label="Sales" sortKey="salesInfluenced" sortState={sortState} onSort={onSort}/>
            <SortHeader label="vs PM" sortKey="prevInfluenced" sortState={sortState} onSort={onSort}/>
            <SortHeader label="Winback" sortKey="salesWinback" sortState={sortState} onSort={onSort}/>
            <SortHeader label="Gross" sortKey="salesGross" sortState={sortState} onSort={onSort}/>
            <SortHeader label="ROI %" sortKey="salesROI" sortState={sortState} onSort={onSort}/>
            <SortHeader label="Winback %" sortKey="winbackSalesPct" sortState={sortState} onSort={onSort} last/>
          </tr></thead>
          <tbody>
            {sorted.map((r,i) => (
              <tr key={r.location} onMouseEnter={e=>e.currentTarget.style.background="#e4edff"} onMouseLeave={e=>e.currentTarget.style.background=i%2===0?"#f4f7fc":"white"}>
                <td style={{...td(i),fontWeight:800,color:i<3?C.sec:"#8896a4"}}>{i+1}</td>
                <td style={{...td(i),fontWeight:600,color:C.main,maxWidth:240,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{r.location}</td>
                <td style={{...td(i),textAlign:"right"}}>{fmtNum(r.salesEngaged)}</td>
                <td style={{...td(i),textAlign:"right"}}>{fmtNum(r.salesShoppers)}</td>
                <td style={{...td(i),textAlign:"right"}}>{fmtNum(r.salesLeads)}</td>
                <td style={{...td(i),textAlign:"right",fontWeight:700,color:C.sec}}>{fmtNum(r.salesInfluenced)}</td>
                <td style={{...td(i),textAlign:"right"}}><Badge cur={r.salesInfluenced} prev={r.prevInfluenced}/></td>
                <td style={{...td(i),textAlign:"right",color:C.purple,fontWeight:600}}>{fmtNum(r.salesWinback)}</td>
                <td style={{...td(i),textAlign:"right",fontWeight:600}}>{fmtMoney(r.salesGross)}</td>
                <td style={{...td(i),textAlign:"right"}}>
                  <span style={{background:r.salesROI>=200?"#d4edda":r.salesROI>=100?"#fff3cd":"#f8d7da",color:r.salesROI>=200?"#155724":r.salesROI>=100?"#856404":"#721c24",padding:"3px 10px",borderRadius:20,fontWeight:700,fontSize:12}}>{fmtPct(r.salesROI)}</span>
                </td>
                <td style={{...td(i),textAlign:"right"}}>{fmtPct(r.winbackSalesPct)}</td>
              </tr>
            ))}
            <tr>
              <td style={{...totStyle,borderRadius:"0 0 0 10px"}}></td>
              <td style={{...totStyle,textAlign:"left",color:C.main,fontSize:14}}>GROUP TOTALS</td>
              <td style={{...totStyle,textAlign:"right"}}>{fmtNum(tot.salesEngaged)}</td>
              <td style={{...totStyle,textAlign:"right"}}>{fmtNum(tot.salesShoppers)}</td>
              <td style={{...totStyle,textAlign:"right"}}>{fmtNum(tot.salesLeads)}</td>
              <td style={{...totStyle,textAlign:"right",color:C.sec}}>{fmtNum(tot.salesInfluenced)}</td>
              <td style={{...totStyle,textAlign:"right"}}>—</td>
              <td style={{...totStyle,textAlign:"right",color:C.purple}}>{fmtNum(tot.salesWinback)}</td>
              <td style={{...totStyle,textAlign:"right"}}>{fmtMoney(tot.salesGross)}</td>
              <td style={{...totStyle,textAlign:"right"}}>—</td>
              <td style={{...totStyle,textAlign:"right",borderRadius:"0 0 10px 0"}}>—</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ──────────── TAB 5: DEALERSHIP SERVICE + GOOGLE ADS ──────────── */
function DealerServiceGoogleTab({ data, periods, sp, setSp, locations }) {
  const { sortState, onSort, doSort } = useSort("serviceROs","desc");
  const p = periods.find(pp => pp.key === sp);
  const prev = getPrev(periods, sp);
  const rows = locations.map(loc => {
    const pRows = p ? filterPeriod(data, p).filter(r => r.location === loc) : [];
    const prRows = prev ? filterPeriod(data, prev).filter(r => r.location === loc) : [];
    const svc = aggService(pRows);
    const ggl = aggGoogle(pRows);
    const prvSvc = prRows.length ? aggService(prRows) : null;
    return { location:loc, ...svc, ...ggl, prevROs: prvSvc?.serviceROs ?? null };
  });
  const sorted = doSort(rows);
  const td = i => ({padding:"10px 12px",fontSize:13,background:i%2===0?"#f4f7fc":"white"});
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:16,marginBottom:20,flexWrap:"wrap"}}>
        <PeriodSelector periods={periods} sp={sp} setSp={setSp}/>
        <span style={{fontSize:14,color:"#7a8a9a"}}>Email <strong>Service</strong> + <strong>Google Ads</strong> results by dealership</span>
      </div>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"separate",borderSpacing:0}}>
          <thead><tr>
            <th style={{padding:"11px 12px",fontWeight:700,fontSize:13,textAlign:"left",background:C.main,color:"white",borderRadius:"10px 0 0 0"}}>#</th>
            <SortHeader label="Dealership" sortKey="location" sortState={sortState} onSort={onSort} align="left"/>
            <SortHeader label="Svc Leads" sortKey="serviceLeads" sortState={sortState} onSort={onSort}/>
            <SortHeader label="RO's" sortKey="serviceROs" sortState={sortState} onSort={onSort}/>
            <SortHeader label="vs PM" sortKey="prevROs" sortState={sortState} onSort={onSort}/>
            <SortHeader label="WB RO's" sortKey="serviceWinbackROs" sortState={sortState} onSort={onSort}/>
            <SortHeader label="RO Value" sortKey="roValue" sortState={sortState} onSort={onSort}/>
            <SortHeader label="WB %" sortKey="winbackServicePct" sortState={sortState} onSort={onSort}/>
            <th style={{padding:"11px 5px",background:C.main,color:"white",fontSize:10}}>│</th>
            <SortHeader label="G Clicks" sortKey="clicksGoogle" sortState={sortState} onSort={onSort}/>
            <SortHeader label="G Leads" sortKey="leadsGoogle" sortState={sortState} onSort={onSort}/>
            <SortHeader label="G Calls" sortKey="phoneCallsGoogle" sortState={sortState} onSort={onSort}/>
            <SortHeader label="G Spend" sortKey="spendGoogle" sortState={sortState} onSort={onSort} last/>
          </tr></thead>
          <tbody>
            {sorted.map((r,i) => (
              <tr key={r.location} onMouseEnter={e=>e.currentTarget.style.background="#e4edff"} onMouseLeave={e=>e.currentTarget.style.background=i%2===0?"#f4f7fc":"white"}>
                <td style={{...td(i),fontWeight:800,color:i<3?C.sec:"#8896a4"}}>{i+1}</td>
                <td style={{...td(i),fontWeight:600,color:C.main,maxWidth:220,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{r.location}</td>
                <td style={{...td(i),textAlign:"right"}}>{fmtNum(r.serviceLeads)}</td>
                <td style={{...td(i),textAlign:"right",fontWeight:700,color:C.green}}>{fmtNum(r.serviceROs)}</td>
                <td style={{...td(i),textAlign:"right"}}><Badge cur={r.serviceROs} prev={r.prevROs}/></td>
                <td style={{...td(i),textAlign:"right",color:C.purple}}>{fmtNum(r.serviceWinbackROs)}</td>
                <td style={{...td(i),textAlign:"right",fontWeight:600}}>{fmtMoney(r.roValue)}</td>
                <td style={{...td(i),textAlign:"right"}}>{fmtPct(r.winbackServicePct)}</td>
                <td style={{...td(i),textAlign:"center",color:"#d0d5dd"}}>│</td>
                <td style={{...td(i),textAlign:"right"}}>{fmtNum(r.clicksGoogle)}</td>
                <td style={{...td(i),textAlign:"right",fontWeight:600,color:C.sec}}>{fmtNum(r.leadsGoogle)}</td>
                <td style={{...td(i),textAlign:"right"}}>{fmtNum(r.phoneCallsGoogle)}</td>
                <td style={{...td(i),textAlign:"right",color:C.amber,fontWeight:600}}>{fmtMoney(r.spendGoogle)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ──────────── TAB 6: DEALERSHIP FACEBOOK ADS ──────────── */
function DealerFBTab({ data, periods, sp, setSp, locations }) {
  const { sortState, onSort, doSort } = useSort("leadsFB","desc");
  const p = periods.find(pp => pp.key === sp);
  const prev = getPrev(periods, sp);
  const rows = locations.map(loc => {
    const pRows = p ? filterPeriod(data, p).filter(r => r.location === loc) : [];
    const prRows = prev ? filterPeriod(data, prev).filter(r => r.location === loc) : [];
    const fb = aggFB(pRows);
    const prvFb = prRows.length ? aggFB(prRows) : null;
    return { location:loc, ...fb, prevLeads: prvFb?.leadsFB ?? null };
  });
  const sorted = doSort(rows);
  const td = i => ({padding:"10px 12px",fontSize:13,background:i%2===0?"#f4f7fc":"white"});
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:16,marginBottom:20,flexWrap:"wrap"}}>
        <PeriodSelector periods={periods} sp={sp} setSp={setSp}/>
        <span style={{fontSize:14,color:"#7a8a9a"}}><strong>Facebook Ads</strong> results by dealership</span>
      </div>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"separate",borderSpacing:0}}>
          <thead><tr>
            <th style={{padding:"11px 12px",fontWeight:700,fontSize:13,textAlign:"left",background:C.main,color:"white",borderRadius:"10px 0 0 0"}}>#</th>
            <SortHeader label="Dealership" sortKey="location" sortState={sortState} onSort={onSort} align="left"/>
            <SortHeader label="Clicks" sortKey="clicksFB" sortState={sortState} onSort={onSort}/>
            <SortHeader label="Page Views" sortKey="pageViewsFB" sortState={sortState} onSort={onSort}/>
            <SortHeader label="Leads" sortKey="leadsFB" sortState={sortState} onSort={onSort}/>
            <SortHeader label="vs PM" sortKey="prevLeads" sortState={sortState} onSort={onSort}/>
            <SortHeader label="Phone Calls" sortKey="phoneCallsFB" sortState={sortState} onSort={onSort}/>
            <SortHeader label="Spend" sortKey="spendFB" sortState={sortState} onSort={onSort}/>
            <SortHeader label="Cost/Lead" sortKey="cpl" sortState={sortState} onSort={onSort} last/>
          </tr></thead>
          <tbody>
            {sorted.map((r,i) => (
              <tr key={r.location} onMouseEnter={e=>e.currentTarget.style.background="#e4edff"} onMouseLeave={e=>e.currentTarget.style.background=i%2===0?"#f4f7fc":"white"}>
                <td style={{...td(i),fontWeight:800,color:i<3?C.sec:"#8896a4"}}>{i+1}</td>
                <td style={{...td(i),fontWeight:600,color:C.main,maxWidth:240,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{r.location}</td>
                <td style={{...td(i),textAlign:"right"}}>{fmtNum(r.clicksFB)}</td>
                <td style={{...td(i),textAlign:"right"}}>{fmtNum(r.pageViewsFB)}</td>
                <td style={{...td(i),textAlign:"right",fontWeight:700,color:C.sec}}>{fmtNum(r.leadsFB)}</td>
                <td style={{...td(i),textAlign:"right"}}><Badge cur={r.leadsFB} prev={r.prevLeads}/></td>
                <td style={{...td(i),textAlign:"right"}}>{fmtNum(r.phoneCallsFB)}</td>
                <td style={{...td(i),textAlign:"right",color:C.amber,fontWeight:600}}>{fmtMoney(r.spendFB)}</td>
                <td style={{...td(i),textAlign:"right"}}>
                  <span style={{background:r.cpl<=50?"#d4edda":r.cpl<=100?"#fff3cd":"#f8d7da",color:r.cpl<=50?"#155724":r.cpl<=100?"#856404":"#721c24",padding:"3px 10px",borderRadius:20,fontWeight:700,fontSize:12}}>{fmtMoney(r.cpl)}</span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ──────────── TAB 7: CUSTOMER DATA ──────────── */
function CustomerDataTab({ sheetTabs }) {
  const [selectedTab, setSelectedTab] = useState("");
  const [custData, setCustData] = useState(null);
  const [custHeaders, setCustHeaders] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const { sortState, onSort, doSort } = useSort("col0","asc");

  const loadTab = useCallback(async (tab) => {
    if (!tab) return;
    setLoading(true); setError(""); setCustData(null);
    try {
      const url = SHEET_CSV_URL + `&gid=${tab.gid}`;
      const resp = await fetch(url);
      const text = await resp.text();
      const result = Papa.parse(text, { header: true, skipEmptyLines: true });
      if (result.data && result.data.length > 0) {
        const headers = Object.keys(result.data[0]);
        setCustHeaders(headers);
        const mapped = result.data.map((row, idx) => {
          const obj = { _idx: idx };
          headers.forEach((h, ci) => { obj[`col${ci}`] = row[h] || ""; });
          return obj;
        });
        setCustData(mapped);
      } else {
        setError("No data found in this tab.");
      }
    } catch (e) { setError("Failed to load tab data. " + e.message); }
    setLoading(false);
  }, []);

  useEffect(() => {
    if (selectedTab && sheetTabs.length) {
      const tab = sheetTabs.find(t => t.name === selectedTab);
      if (tab) loadTab(tab);
    }
  }, [selectedTab, sheetTabs, loadTab]);

  const sorted = custData ? doSort(custData) : [];

  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:16,marginBottom:20,flexWrap:"wrap"}}>
        <select value={selectedTab} onChange={e => setSelectedTab(e.target.value)}
          style={{padding:"10px 16px",borderRadius:10,border:`2px solid ${C.sec}`,fontSize:15,fontWeight:700,color:C.main,background:"white",cursor:"pointer",outline:"none"}}>
          <option value="">— Select Dealership —</option>
          {sheetTabs.map(t => <option key={t.gid} value={t.name}>{t.name}</option>)}
        </select>
        <span style={{fontSize:14,color:"#7a8a9a"}}>Customer-specific data from dealership email marketing for <strong>Sales</strong></span>
      </div>
      {loading && <div style={{padding:40,textAlign:"center",color:C.sec}}>Loading data...</div>}
      {error && <div style={{padding:40,textAlign:"center",color:C.red}}>{error}</div>}
      {!selectedTab && !loading && <div style={{padding:60,textAlign:"center",color:"#aab"}}>
        <div style={{fontSize:48,marginBottom:12}}>📊</div>
        <div style={{fontSize:16,fontWeight:600}}>Select a dealership above to view customer-specific data</div>
      </div>}
      {sorted.length > 0 && (
        <div style={{overflowX:"auto",maxHeight:700,overflowY:"auto"}}>
          <table style={{width:"100%",borderCollapse:"separate",borderSpacing:0}}>
            <thead><tr>
              {custHeaders.map((h,ci) => (
                <SortHeader key={ci} label={h} sortKey={`col${ci}`} sortState={sortState} onSort={onSort} align={ci===0?"left":"right"} first={ci===0} last={ci===custHeaders.length-1}/>
              ))}
            </tr></thead>
            <tbody>
              {sorted.map((row,i) => (
                <tr key={row._idx} onMouseEnter={e=>e.currentTarget.style.background="#e4edff"} onMouseLeave={e=>e.currentTarget.style.background=i%2===0?"#f4f7fc":"white"}>
                  {custHeaders.map((h,ci) => (
                    <td key={ci} style={{padding:"9px 12px",fontSize:13,textAlign:ci===0?"left":"right",background:i%2===0?"#f4f7fc":"white",whiteSpace:"nowrap",maxWidth:300,overflow:"hidden",textOverflow:"ellipsis",fontWeight:ci===0?600:400,color:ci===0?C.main:undefined}}>
                      {row[`col${ci}`]}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

/* ──────────── EXPORT UTILITIES ──────────── */
async function captureTab(ref) {
  return html2canvas(ref, { backgroundColor:"#edf1f7", scale:2, useCORS:true, logging:false });
}

/* ──────────── MAIN APP ──────────── */
export default function App() {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [activeTab, setActiveTab] = useState("groupSales");
  const [sp, setSp] = useState("");
  const [sheetTabs, setSheetTabs] = useState([]);
  const contentRef = useRef(null);

  // Fetch main data
  const fetchData = useCallback(async () => {
    setLoading(true); setError("");
    try {
      const resp = await fetch(SHEET_CSV_URL);
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const text = await resp.text();
      const parsed = parseOverallCSV(text);
      if (parsed.length === 0) throw new Error("No valid data rows found in the spreadsheet.");
      setData(parsed);
      const periods = getPeriods(parsed);
      if (periods.length > 0) setSp(periods[periods.length - 1].key);
    } catch (e) { setError(e.message); }
    setLoading(false);
  }, []);

  // Discover sheet tabs for customer data
  const discoverTabs = useCallback(async () => {
    try {
      const resp = await fetch(SHEET_HTML_URL);
      const html = await resp.text();
      // Parse sheet tabs from the HTML — Google Sheets embeds them as list items
      const tabRegex = /id="sheet-button-(\d+)"[^>]*>([^<]+)</g;
      const tabs = [];
      let match;
      while ((match = tabRegex.exec(html)) !== null) {
        const gid = match[1];
        const name = match[2].trim();
        // Skip the "Overall Results" tab (gid=0 or first tab)
        if (name.toLowerCase().includes("overall")) continue;
        tabs.push({ gid, name });
      }
      setSheetTabs(tabs);
    } catch (e) { console.warn("Could not discover sheet tabs:", e); }
  }, []);

  useEffect(() => { fetchData(); discoverTabs(); }, [fetchData, discoverTabs]);

  const periods = useMemo(() => getPeriods(data), [data]);
  const locations = useMemo(() => [...new Set(data.map(r => r.location))].sort(), [data]);

  // Export
  const [exporting, setExporting] = useState(false);
  const exportJPG = async () => {
    if (!contentRef.current) return;
    setExporting(true);
    try {
      const canvas = await captureTab(contentRef.current);
      const link = document.createElement("a");
      link.download = `Garber_Email_${TABS.find(t=>t.id===activeTab)?.label.replace(/\s+/g,"_")}_${sp}.jpg`;
      link.href = canvas.toDataURL("image/jpeg", 0.95);
      link.click();
    } catch(e) { alert("Export failed: " + e.message); }
    setExporting(false);
  };
  const exportPDF = async () => {
    if (!contentRef.current) return;
    setExporting(true);
    try {
      const canvas = await captureTab(contentRef.current);
      const imgData = canvas.toDataURL("image/jpeg", 0.92);
      const pdf = new jsPDF({ orientation: canvas.width > canvas.height ? "landscape" : "portrait", unit: "px", format: [canvas.width, canvas.height] });
      pdf.addImage(imgData, "JPEG", 0, 0, canvas.width, canvas.height);
      pdf.save(`Garber_Email_${TABS.find(t=>t.id===activeTab)?.label.replace(/\s+/g,"_")}_${sp}.pdf`);
    } catch(e) { alert("Export failed: " + e.message); }
    setExporting(false);
  };
  const exportZIP = async () => {
    setExporting(true);
    try {
      const zip = new JSZip();
      const origTab = activeTab;
      for (const tab of TABS) {
        setActiveTab(tab.id);
        await new Promise(r => setTimeout(r, 300));
        if (contentRef.current) {
          const canvas = await captureTab(contentRef.current);
          const blob = await new Promise(resolve => canvas.toBlob(resolve, "image/jpeg", 0.92));
          zip.file(`${tab.label.replace(/\s+/g,"_")}_${sp}.jpg`, blob);
        }
      }
      setActiveTab(origTab);
      const content = await zip.generateAsync({ type: "blob" });
      const link = document.createElement("a");
      link.download = `Garber_Email_Marketing_Report_${sp}.zip`;
      link.href = URL.createObjectURL(content);
      link.click();
    } catch(e) { alert("Export failed: " + e.message); }
    setExporting(false);
  };

  // Loading / Error states
  if (loading) return (
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:C.bg}}>
      <div style={{textAlign:"center"}}>
        <div style={{width:56,height:56,border:`5px solid ${C.sec}`,borderTopColor:"transparent",borderRadius:"50%",animation:"spin 1s linear infinite",margin:"0 auto 20px"}}/>
        <div style={{fontSize:18,fontWeight:700,color:C.main}}>Loading Email Marketing Data...</div>
        <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
      </div>
    </div>
  );
  if (error) return (
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:C.bg}}>
      <div style={{textAlign:"center",maxWidth:500,padding:40}}>
        <div style={{fontSize:48,marginBottom:16}}>⚠️</div>
        <div style={{fontSize:20,fontWeight:700,color:C.red,marginBottom:12}}>Error Loading Data</div>
        <div style={{fontSize:15,color:"#666",marginBottom:24}}>{error}</div>
        <button onClick={fetchData} style={{padding:"12px 28px",borderRadius:10,background:C.sec,color:"white",fontWeight:700,fontSize:15,border:"none",cursor:"pointer"}}>Retry</button>
      </div>
    </div>
  );

  return (
    <div style={{minHeight:"100vh",background:C.bg}}>
      {/* HEADER */}
      <div style={{background:`linear-gradient(135deg, ${C.main} 0%, #0f3d7a 100%)`,padding:"28px 36px 20px",boxShadow:"0 4px 20px rgba(7,42,96,0.25)"}}>
        <div style={{maxWidth:1500,margin:"0 auto"}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:16}}>
            <div>
              <div style={{fontSize:28,fontWeight:800,color:"white",letterSpacing:0.5}}>Garber Automotive Group</div>
              <div style={{fontSize:15,color:"rgba(255,255,255,0.7)",marginTop:4}}>Email Marketing & Advertising Dashboard</div>
            </div>
            <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
              <button onClick={fetchData} style={{padding:"9px 18px",borderRadius:8,background:"rgba(255,255,255,0.15)",color:"white",fontWeight:700,fontSize:14,border:"1px solid rgba(255,255,255,0.25)",cursor:"pointer",transition:"all .2s"}}
                onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,0.25)"} onMouseLeave={e=>e.currentTarget.style.background="rgba(255,255,255,0.15)"}>
                🔄 Refresh Data
              </button>
              <button onClick={exportJPG} disabled={exporting} style={{padding:"9px 18px",borderRadius:8,background:"rgba(255,255,255,0.15)",color:"white",fontWeight:700,fontSize:14,border:"1px solid rgba(255,255,255,0.25)",cursor:"pointer"}}>📸 JPG</button>
              <button onClick={exportPDF} disabled={exporting} style={{padding:"9px 18px",borderRadius:8,background:"rgba(255,255,255,0.15)",color:"white",fontWeight:700,fontSize:14,border:"1px solid rgba(255,255,255,0.25)",cursor:"pointer"}}>📄 PDF</button>
              <button onClick={exportZIP} disabled={exporting} style={{padding:"9px 18px",borderRadius:8,background:"rgba(255,255,255,0.15)",color:"white",fontWeight:700,fontSize:14,border:"1px solid rgba(255,255,255,0.25)",cursor:"pointer"}}>📦 ZIP All</button>
            </div>
          </div>
          {/* TAB NAV */}
          <div style={{display:"flex",gap:6,marginTop:18,flexWrap:"wrap"}}>
            {TABS.map(t => (
              <button key={t.id} onClick={() => setActiveTab(t.id)}
                style={{padding:"10px 20px",borderRadius:"10px 10px 0 0",fontSize:14,fontWeight:700,border:"none",cursor:"pointer",transition:"all .15s",
                  background:activeTab===t.id?"white":"rgba(255,255,255,0.08)",
                  color:activeTab===t.id?C.main:"rgba(255,255,255,0.75)",
                  borderBottom:activeTab===t.id?`3px solid ${C.sec}`:"3px solid transparent",
                }}
                onMouseEnter={e=>{if(activeTab!==t.id)e.currentTarget.style.background="rgba(255,255,255,0.15)"}}
                onMouseLeave={e=>{if(activeTab!==t.id)e.currentTarget.style.background="rgba(255,255,255,0.08)"}}>
                {t.label}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* CONTENT */}
      <div style={{maxWidth:1500,margin:"0 auto",padding:"24px 36px 60px"}} ref={contentRef}>
        {exporting && <div style={{textAlign:"center",padding:12,background:"#fff3cd",borderRadius:10,marginBottom:16,color:"#856404",fontWeight:600}}>Exporting... please wait</div>}
        {activeTab==="groupSales" && <GroupSalesTab data={data} periods={periods} sp={sp} setSp={setSp}/>}
        {activeTab==="groupService" && <GroupServiceTab data={data} periods={periods} sp={sp} setSp={setSp}/>}
        {activeTab==="groupAds" && <GroupAdsTab data={data} periods={periods} sp={sp} setSp={setSp}/>}
        {activeTab==="dealerSales" && <DealerSalesTab data={data} periods={periods} sp={sp} setSp={setSp} locations={locations}/>}
        {activeTab==="dealerService" && <DealerServiceGoogleTab data={data} periods={periods} sp={sp} setSp={setSp} locations={locations}/>}
        {activeTab==="dealerFB" && <DealerFBTab data={data} periods={periods} sp={sp} setSp={setSp} locations={locations}/>}
        {activeTab==="customerData" && <CustomerDataTab sheetTabs={sheetTabs}/>}
      </div>
    </div>
  );
}
