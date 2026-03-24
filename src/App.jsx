import { useState, useMemo, useEffect, useCallback, useRef } from "react";
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from "recharts";
import Papa from "papaparse";
import html2canvas from "html2canvas";
import jsPDF from "jspdf";
import JSZip from "jszip";

const SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTYVu1c2AZGFprO4Qk2sgvY6GDl1PuBAxW-7J5xg4xjIrz-ZCTaxn2oC2vVfCxECbGqOt6e9KkgAjHs/pub?output=csv";
const SHEETS_API = "/api/sheets";
const MN=["","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const MONTH_MAP={January:1,February:2,March:3,April:4,May:5,June:6,July:7,August:8,September:9,October:10,November:11,December:12,Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};
const C={main:"#072a60",sec:"#0d6eff",acc1:"#1e88e5",green:"#0a8754",amber:"#c9710d",red:"#c0392b",bg:"#edf1f7",purple:"#7c3aed",teal:"#0d9488"};
const LOC_NORMALIZE={"Volvo Cars Rochester":"Volvo Cars of Rochester"};
const TABS=[
  {id:"groupSales",label:"Sales Email Marketing",icon:"\ud83d\udcca"},
  {id:"groupService",label:"Service Email Marketing",icon:"\ud83d\udd27"},
  {id:"groupAds",label:"Advertising Results",icon:"\ud83c\udfaf"},
  {id:"dealerSales",label:"Sales Rankings",icon:"\ud83c\udfc6"},
  {id:"dealerService",label:"Service Rankings",icon:"\ud83c\udfc6"},
  {id:"customerData",label:"Customer Data",icon:"\ud83d\udc64"},
];

function num(v){if(v===undefined||v===null||v==="")return 0;const s=String(v).replace(/[$,\s%]/g,"");const n=parseFloat(s);return isNaN(n)?0:n;}

function parseOverallCSV(text){
  const result=Papa.parse(text,{header:true,skipEmptyLines:true});
  if(!result.data||result.data.length<2)return[];
  const headers=Object.keys(result.data[0]);
  const fc=(...needles)=>headers.find(h=>{const lc=h.toLowerCase().replace(/[^a-z0-9%$]/g,"");return needles.some(n=>lc.includes(n.toLowerCase().replace(/[^a-z0-9%$]/g,"")));})||"";
  const colMonth=fc("month"),colYear=fc("year"),colLoc=fc("location"),colCost=fc("monthlycost","cost");
  const colSalesEngaged=fc("salesengaged","engaged"),colSalesShoppers=fc("salesshoppers"),colSalesLeads=fc("salesleads");
  const colSalesInfluenced=fc("salesinfluenced","influenced"),colSalesWinback=fc("saleswinback");
  const colSalesGross=fc("salesgross","gross"),colSalesWinbackProfit=fc("winbackprofit","saleswinbackprofit");
  const colSalesROI=fc("salesroi","roi%"),colWinbackSalesPct=fc("winbacksales%","winbacksalespct");
  const colServiceShoppers=fc("serviceshoppers"),colServiceLeads=fc("serviceleads");
  const colServiceROs=fc("servicero"),colServiceWinbackROs=fc("servicewinbackro");
  const colROValue=fc("rovalue"),colServiceWinbackROValue=fc("servicewinbackrovalue","winbackrovalue");
  const colWinbackServicePct=fc("winbackservice%","winbackservicepct");
  const colClicksGoogle=fc("clicksgoogle"),colPageViewsGoogle=fc("pageviewsgoogle","viewsgoogle");
  const colLeadsGoogle=fc("leadsgoogle"),colPhoneCallsGoogle=fc("phonecallsgoogle","callsgoogle"),colSpendGoogle=fc("spendgoogle");
  const colClicksFB=fc("clicksfacebook","clicksfb"),colPageViewsFB=fc("pageviewsfacebook","viewsfacebook");
  const colLeadsFB=fc("leadsfacebook","leadsfb"),colPhoneCallsFB=fc("phonecallsfacebook","callsfacebook"),colSpendFB=fc("spendfacebook","spendfb");
  return result.data.map(r=>{
    const monthRaw=(r[colMonth]||"").trim();const m=MONTH_MAP[monthRaw]||parseInt(monthRaw)||0;const y=parseInt(r[colYear])||0;
    if(!m||!y)return null;
    const rawLoc=(r[colLoc]||"").trim();
    return{month:m,year:y,location:LOC_NORMALIZE[rawLoc]||rawLoc,monthlyCost:num(r[colCost]),
      salesEngaged:num(r[colSalesEngaged]),salesShoppers:num(r[colSalesShoppers]),salesLeads:num(r[colSalesLeads]),
      salesInfluenced:num(r[colSalesInfluenced]),salesWinback:num(r[colSalesWinback]),salesGross:num(r[colSalesGross]),
      salesWinbackProfit:num(r[colSalesWinbackProfit]),salesROI:num(r[colSalesROI]),winbackSalesPct:num(r[colWinbackSalesPct]),
      serviceShoppers:num(r[colServiceShoppers]),serviceLeads:num(r[colServiceLeads]),serviceROs:num(r[colServiceROs]),
      serviceWinbackROs:num(r[colServiceWinbackROs]),roValue:num(r[colROValue]),serviceWinbackROValue:num(r[colServiceWinbackROValue]),
      winbackServicePct:num(r[colWinbackServicePct]),
      clicksGoogle:num(r[colClicksGoogle]),pageViewsGoogle:num(r[colPageViewsGoogle]),leadsGoogle:num(r[colLeadsGoogle]),
      phoneCallsGoogle:num(r[colPhoneCallsGoogle]),spendGoogle:num(r[colSpendGoogle]),
      clicksFB:num(r[colClicksFB]),pageViewsFB:num(r[colPageViewsFB]),leadsFB:num(r[colLeadsFB]),
      phoneCallsFB:num(r[colPhoneCallsFB]),spendFB:num(r[colSpendFB])};
  }).filter(Boolean);
}

function getPeriods(data){const set=new Set(data.map(r=>`${r.year}-${String(r.month).padStart(2,"0")}`));return[...set].sort().map(p=>{const[y,m]=p.split("-").map(Number);return{year:y,month:m,label:`${MN[m]} ${y}`,short:`${MN[m]} '${String(y).slice(2)}`,key:p};});}
function getPrev(periods,key){const i=periods.findIndex(p=>p.key===key);return i>0?periods[i-1]:null;}
function fmtMoney(v){if(Math.abs(v)>=1e6)return`$${(v/1e6).toFixed(1)}M`;if(Math.abs(v)>=1000)return`$${(v/1000).toFixed(1)}K`;return`$${v.toLocaleString(undefined,{maximumFractionDigits:0})}`;}
function fmtNum(v){return v.toLocaleString();}
function fmtPct(v){return`${v.toFixed(1)}%`;}
function filterPeriod(data,p){return data.filter(r=>r.year===p.year&&r.month===p.month);}
function getLast12(periods,sp){const idx=periods.findIndex(p=>p.key===sp);if(idx<0)return periods.slice(-12);const start=Math.max(0,idx-11);return periods.slice(start,idx+1);}

function aggSales(rows){const o={monthlyCost:0,salesEngaged:0,salesShoppers:0,salesLeads:0,salesInfluenced:0,salesWinback:0,salesGross:0,salesWinbackProfit:0};rows.forEach(r=>{for(const k in o)o[k]+=r[k];});o.salesROI=o.monthlyCost>0?((o.salesGross/o.monthlyCost)*100):0;o.winbackSalesPct=o.salesInfluenced>0?((o.salesWinback/o.salesInfluenced)*100):0;return o;}
function aggService(rows){const o={serviceShoppers:0,serviceLeads:0,serviceROs:0,serviceWinbackROs:0,roValue:0,serviceWinbackROValue:0};rows.forEach(r=>{for(const k in o)o[k]+=r[k];});o.winbackServicePct=o.serviceROs>0?((o.serviceWinbackROs/o.serviceROs)*100):0;return o;}
function aggGoogle(rows){const o={clicksGoogle:0,pageViewsGoogle:0,leadsGoogle:0,phoneCallsGoogle:0,spendGoogle:0};rows.forEach(r=>{for(const k in o)o[k]+=r[k];});return o;}
function aggFB(rows){const o={clicksFB:0,pageViewsFB:0,leadsFB:0,phoneCallsFB:0,spendFB:0};rows.forEach(r=>{for(const k in o)o[k]+=r[k];});return o;}

const Badge=({cur,prev})=>{if(prev===null||prev===undefined)return<span style={{color:"#9aa",fontSize:14}}>{"\u2014"}</span>;const d=cur-prev;if(d===0)return<span style={{color:"#8896a4",fontSize:14,fontWeight:600}}>{"\u2014"} <span style={{color:"#99AAAA"}}>vs PM</span></span>;const up=d>0;return<span style={{fontSize:14,fontWeight:700,color:up?C.green:C.red,whiteSpace:"nowrap"}}>{up?"\u25B2":"\u25BC"} {up?"+":""}{typeof cur==="number"&&cur%1!==0?d.toFixed(1):d.toLocaleString()} <span style={{fontWeight:600,fontSize:12,color:"#99AAAA"}}>vs PM</span></span>;};
const MoneyBadge=({cur,prev})=>{if(prev===null||prev===undefined)return<span style={{color:"#9aa",fontSize:14}}>{"\u2014"}</span>;const d=cur-prev;if(d===0)return<span style={{color:"#8896a4",fontSize:14,fontWeight:600}}>{"\u2014"} <span style={{color:"#99AAAA"}}>vs PM</span></span>;const up=d>0;return<span style={{fontSize:14,fontWeight:700,color:up?C.green:C.red,whiteSpace:"nowrap"}}>{up?"\u25B2":"\u25BC"} {up?"+":""}{fmtMoney(d)} <span style={{fontWeight:600,fontSize:12,color:"#99AAAA"}}>vs PM</span></span>;};

const KPI=({label,value,fmt="num",color=C.main,sub})=>(
  <div style={{background:"white",borderRadius:14,padding:"22px 24px",minWidth:150,flex:1,boxShadow:"0 2px 8px rgba(0,0,0,0.06)",borderTop:`4px solid ${color}`}}>
    <div style={{fontSize:12,color:"#7a8a9a",fontWeight:600,textTransform:"uppercase",letterSpacing:1.1,marginBottom:6}}>{label}</div>
    <div style={{fontSize:32,fontWeight:800,color,lineHeight:1.1}}>{fmt==="money"?fmtMoney(value):fmt==="pct"?fmtPct(value):fmtNum(value)}</div>
    {sub&&<div style={{fontSize:13,color:"#7a8a9a",marginTop:6}}>{sub}</div>}
  </div>
);

const SortHeader=({label,sortKey,sortState,onSort,align="right",first,last})=>(
  <th onClick={()=>onSort(sortKey)} style={{padding:"11px 12px",fontWeight:700,fontSize:13,textAlign:align,cursor:"pointer",userSelect:"none",whiteSpace:"nowrap",background:C.main,color:"white",borderRadius:first?"10px 0 0 0":last?"0 10px 0 0":undefined,position:"sticky",top:0,zIndex:2}}>
    {label}{sortState.key===sortKey?(sortState.dir==="asc"?" \u2191":" \u2193"):""}
  </th>
);

function useSort(defaultKey,defaultDir="desc"){const[s,setS]=useState({key:defaultKey,dir:defaultDir});const onSort=useCallback(key=>setS(prev=>prev.key===key?{key,dir:prev.dir==="desc"?"asc":"desc"}:{key,dir:"desc"}),[]);const doSort=useCallback(arr=>[...arr].sort((a,b)=>{const av=a[s.key]??0,bv=b[s.key]??0;if(typeof av==="string")return s.dir==="asc"?av.localeCompare(bv):bv.localeCompare(av);return s.dir==="asc"?av-bv:bv-av;}),[s]);return{sortState:s,onSort,doSort};}

const selStyle={padding:"9px 16px",borderRadius:10,border:`2px solid ${C.sec}`,fontSize:14,fontWeight:700,color:C.main,background:"white",cursor:"pointer",outline:"none"};

const TrendChart=({data,periods,valueKey,label,color,fmt="num",aggFn,filterLoc})=>{
  const cd=periods.map(p=>{let rows=filterPeriod(data,p);if(filterLoc)rows=rows.filter(r=>r.location===filterLoc);const a=aggFn(rows);return{name:p.short,value:a[valueKey]||0};});
  return(
    <div style={{background:"white",borderRadius:14,padding:"20px 22px",boxShadow:"0 2px 8px rgba(0,0,0,0.06)",flex:1,minWidth:340}}>
      <div style={{fontSize:15,fontWeight:700,color:C.main,marginBottom:12}}>{label}</div>
      <ResponsiveContainer width="100%" height={220}>
        <LineChart data={cd} margin={{top:5,right:20,bottom:5,left:10}}>
          <CartesianGrid strokeDasharray="3 3" stroke="#e8ecf0"/>
          <XAxis dataKey="name" tick={{fontSize:11,fill:C.main,fontWeight:600}} interval={0} angle={periods.length>8?-30:0} textAnchor={periods.length>8?"end":"middle"} height={periods.length>8?50:30}/>
          <YAxis tick={{fontSize:11,fill:"#8896a4"}} tickFormatter={v=>fmt==="money"?fmtMoney(v):fmt==="pct"?`${v.toFixed(0)}%`:v>=1000?`${(v/1000).toFixed(1)}k`:v} width={60}/>
          <Tooltip formatter={v=>fmt==="money"?fmtMoney(Number(v)):fmt==="pct"?fmtPct(Number(v)):Number(v).toLocaleString()} contentStyle={{borderRadius:10,border:`1px solid ${C.sec}`,fontSize:13}}/>
          <Line type="monotone" dataKey="value" stroke={color} strokeWidth={3} dot={{r:4,fill:color,stroke:"white",strokeWidth:2}} activeDot={{r:6}}/>
        </LineChart>
      </ResponsiveContainer>
    </div>
  );
};

const Section=({title,desc,children})=>(<div style={{background:"white",borderRadius:14,padding:24,boxShadow:"0 2px 8px rgba(0,0,0,0.06)",marginBottom:20}}><div style={{fontSize:17,fontWeight:700,marginBottom:4,color:C.main}}>{title}</div>{desc&&<div style={{fontSize:13,color:"#7a8a9a",marginBottom:16}}>{desc}</div>}{children}</div>);

const FilterBar=({periods,sp,setSp,locations,selectedLoc,setSelectedLoc})=>(
  <div style={{padding:"16px 0",display:"flex",gap:14,alignItems:"center",flexWrap:"wrap",marginBottom:16}}>
    <span style={{fontSize:14,fontWeight:700,color:"#7a8a9a"}}>PERIOD:</span>
    <select value={sp} onChange={e=>setSp(e.target.value)} style={selStyle}>{periods.map(p=><option key={p.key} value={p.key}>{p.label}</option>)}</select>
    {locations&&(<><span style={{fontSize:14,fontWeight:700,color:"#7a8a9a",marginLeft:12}}>DEALERSHIP:</span>
      <select value={selectedLoc} onChange={e=>setSelectedLoc(e.target.value)} style={{...selStyle,maxWidth:340}}><option value="">All Dealerships</option>{locations.map(l=><option key={l} value={l}>{l}</option>)}</select></>)}
  </div>
);

function GroupSalesTab({data,periods,sp,setSp,locations,captureMode,captureLoc}){
  const[internalLoc,setInternalLoc]=useState("");
  const loc=captureMode&&captureLoc?captureLoc:internalLoc;
  const[newCustData,setNewCustData]=useState(null);
  const[prevNewCustData,setPrevNewCustData]=useState(null);
  const p=periods.find(pp=>pp.key===sp);const prev=getPrev(periods,sp);
  const filt=loc?data.filter(r=>r.location===loc):data;
  const cur=p?aggSales(filterPeriod(filt,p)):null;const prv=prev?aggSales(filterPeriod(filt,prev)):null;
  const last12=getLast12(periods,sp);

  // Fetch new customer counts + condition data for current and previous period
  useEffect(()=>{
    if(!p)return;
    fetch(SHEETS_API+"?action=newcustomer&month="+p.month+"&year="+p.year)
      .then(r=>r.json()).then(d=>setNewCustData(d)).catch(()=>setNewCustData(null));
    if(prev){
      fetch(SHEETS_API+"?action=newcustomer&month="+prev.month+"&year="+prev.year)
        .then(r=>r.json()).then(d=>setPrevNewCustData(d)).catch(()=>setPrevNewCustData(null));
    }else{setPrevNewCustData(null);}
  },[p,prev]);

  const newCustCount=useMemo(()=>{
    if(!newCustData||!newCustData.counts)return 0;
    if(loc)return newCustData.counts[loc]||0;
    return newCustData.total||0;
  },[newCustData,loc]);
  const prevNewCustCount=useMemo(()=>{
    if(!prevNewCustData||!prevNewCustData.counts)return null;
    if(loc)return prevNewCustData.counts[loc]||0;
    return prevNewCustData.total||0;
  },[prevNewCustData,loc]);

  // New vs Used condition percentages
  const{newPct,usedPct}=useMemo(()=>{
    if(!newCustData||!newCustData.condCounts)return{newPct:0,usedPct:0};
    let n=0,u=0;
    if(loc){
      const cc=newCustData.condCounts[loc];
      if(cc){n=cc.new||0;u=cc.used||0;}
    }else{
      n=newCustData.condTotal?.new||0;u=newCustData.condTotal?.used||0;
    }
    const tot=n+u;
    return{newPct:tot>0?(n/tot*100):0,usedPct:tot>0?(u/tot*100):0};
  },[newCustData,loc]);

  if(!cur)return<div style={{padding:40,color:"#999",textAlign:"center"}}>No data for selected period.</div>;
  return(<div>
    {!captureMode&&<FilterBar periods={periods} sp={sp} setSp={setSp} locations={locations} selectedLoc={loc} setSelectedLoc={setInternalLoc}/>}
    {captureMode&&<div style={{padding:"8px 0 12px",fontSize:18,fontWeight:800,color:C.main}}>Sales Email Marketing {"\u2014"} {loc} {"\u2014"} {periods.find(pp=>pp.key===sp)?.label}</div>}
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:14,marginBottom:14}}>
      <KPI label="Sales Engaged" value={cur.salesEngaged} color={C.main} sub={prv?<Badge cur={cur.salesEngaged} prev={prv.salesEngaged}/>:null}/>
      <KPI label="Sales Shoppers" value={cur.salesShoppers} color={C.acc1} sub={prv?<Badge cur={cur.salesShoppers} prev={prv.salesShoppers}/>:null}/>
      <KPI label="Sales Leads" value={cur.salesLeads} color={C.sec} sub={prv?<Badge cur={cur.salesLeads} prev={prv.salesLeads}/>:null}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:14,marginBottom:14}}>
      <KPI label="Sales Influenced" value={cur.salesInfluenced} color={C.green} sub={prv?<Badge cur={cur.salesInfluenced} prev={prv.salesInfluenced}/>:null}/>
      <KPI label="Sales Gross" value={cur.salesGross} fmt="money" color={C.green} sub={prv?<MoneyBadge cur={cur.salesGross} prev={prv.salesGross}/>:null}/>
      <KPI label="Sales Winback" value={cur.salesWinback} color={C.purple} sub={prv?<Badge cur={cur.salesWinback} prev={prv.salesWinback}/>:null}/>
      <KPI label="Winback Profit" value={cur.salesWinbackProfit} fmt="money" color={C.purple} sub={prv?<MoneyBadge cur={cur.salesWinbackProfit} prev={prv.salesWinbackProfit}/>:null}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:14,marginBottom:14}}>
      <KPI label="Monthly Cost" value={cur.monthlyCost} fmt="money" color={C.amber}/>
      <KPI label="Sales ROI %" value={cur.salesROI} fmt="pct" color={cur.salesROI>=100?C.green:C.red}/>
      <KPI label="Winback Sales %" value={cur.winbackSalesPct} fmt="pct" color={C.teal}/>
      <KPI label="New Customer Sales" value={newCustCount} color={C.teal} sub={prevNewCustCount!==null?<Badge cur={newCustCount} prev={prevNewCustCount}/>:null}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:20}}>
      <KPI label="New % Sold" value={newPct} fmt="pct" color={C.sec} sub={<span style={{fontSize:12,color:"#7a8a9a"}}>of influenced sales</span>}/>
      <KPI label="Used % Sold" value={usedPct} fmt="pct" color={C.amber} sub={<span style={{fontSize:12,color:"#7a8a9a"}}>of influenced sales</span>}/>
    </div>
    {!captureMode&&<><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
      <TrendChart data={filt} periods={last12} valueKey="salesShoppers" label="Sales Shoppers" color={C.acc1} aggFn={aggSales} filterLoc={loc}/>
      <TrendChart data={filt} periods={last12} valueKey="salesLeads" label="Sales Leads" color={C.sec} aggFn={aggSales} filterLoc={loc}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
      <TrendChart data={filt} periods={last12} valueKey="salesInfluenced" label="Sales Influenced" color={C.green} aggFn={aggSales} filterLoc={loc}/>
      <TrendChart data={filt} periods={last12} valueKey="salesWinback" label="Sales Winback" color={C.purple} aggFn={aggSales} filterLoc={loc}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
      <TrendChart data={filt} periods={last12} valueKey="salesGross" label="Sales Gross" color={C.green} fmt="money" aggFn={aggSales} filterLoc={loc}/>
      <TrendChart data={filt} periods={last12} valueKey="salesWinbackProfit" label="Winback Profit" color={C.purple} fmt="money" aggFn={aggSales} filterLoc={loc}/>
    </div></>}
  </div>);
}

function GroupServiceTab({data,periods,sp,setSp,locations,captureMode,captureLoc}){
  const[internalLoc,setInternalLoc]=useState("");
  const loc=captureMode&&captureLoc?captureLoc:internalLoc;
  const p=periods.find(pp=>pp.key===sp);const prev=getPrev(periods,sp);
  const filt=loc?data.filter(r=>r.location===loc):data;
  const cur=p?aggService(filterPeriod(filt,p)):null;const prv=prev?aggService(filterPeriod(filt,prev)):null;
  const last12=getLast12(periods,sp);
  if(!cur)return<div style={{padding:40,color:"#999",textAlign:"center"}}>No data for selected period.</div>;
  return(<div>
    {!captureMode&&<FilterBar periods={periods} sp={sp} setSp={setSp} locations={locations} selectedLoc={loc} setSelectedLoc={setInternalLoc}/>}
    {captureMode&&<div style={{padding:"8px 0 12px",fontSize:18,fontWeight:800,color:C.main}}>Service Email Marketing {"\u2014"} {loc} {"\u2014"} {periods.find(pp=>pp.key===sp)?.label}</div>}
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
    {!captureMode&&<><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
      <TrendChart data={filt} periods={last12} valueKey="serviceShoppers" label="Service Shoppers" color={C.main} aggFn={aggService} filterLoc={loc}/>
      <TrendChart data={filt} periods={last12} valueKey="serviceLeads" label="Service Leads" color={C.sec} aggFn={aggService} filterLoc={loc}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
      <TrendChart data={filt} periods={last12} valueKey="serviceROs" label="Service RO's" color={C.green} aggFn={aggService} filterLoc={loc}/>
      <TrendChart data={filt} periods={last12} valueKey="roValue" label="RO Value" color={C.amber} fmt="money" aggFn={aggService} filterLoc={loc}/>
    </div></>}
  </div>);
}

function GroupAdsTab({data,periods,sp,setSp,locations,captureMode,captureLoc}){
  const[internalLoc,setInternalLoc]=useState("");
  const loc=captureMode&&captureLoc?captureLoc:internalLoc;
  const p=periods.find(pp=>pp.key===sp);const prev=getPrev(periods,sp);
  const filt=loc?data.filter(r=>r.location===loc):data;
  const gCur=p?aggGoogle(filterPeriod(filt,p)):null;const gPrv=prev?aggGoogle(filterPeriod(filt,prev)):null;
  const fCur=p?aggFB(filterPeriod(filt,p)):null;const fPrv=prev?aggFB(filterPeriod(filt,prev)):null;
  const last12=getLast12(periods,sp);
  // Compute which dealerships have ANY ad data across all periods
  const locsWithAdData=useMemo(()=>{const s=new Set();data.forEach(r=>{if(r.clicksGoogle>0||r.leadsGoogle>0||r.spendGoogle>0||r.pageViewsGoogle>0||r.phoneCallsGoogle>0||r.clicksFB>0||r.leadsFB>0||r.spendFB>0||r.pageViewsFB>0||r.phoneCallsFB>0)s.add(r.location);});return s;},[data]);
  if(!gCur||!fCur)return<div style={{padding:40,color:"#999",textAlign:"center"}}>No data for selected period.</div>;
  const hasG=gCur.clicksGoogle>0||gCur.leadsGoogle>0||gCur.spendGoogle>0||gCur.pageViewsGoogle>0||gCur.phoneCallsGoogle>0;
  const hasF=fCur.clicksFB>0||fCur.leadsFB>0||fCur.spendFB>0||fCur.pageViewsFB>0||fCur.phoneCallsFB>0;
  return(<div>
    {!captureMode&&<div style={{padding:"16px 0",display:"flex",gap:14,alignItems:"center",flexWrap:"wrap",marginBottom:16}}>
      <span style={{fontSize:14,fontWeight:700,color:"#7a8a9a"}}>PERIOD:</span>
      <select value={sp} onChange={e=>setSp(e.target.value)} style={selStyle}>{periods.map(p=><option key={p.key} value={p.key}>{p.label}</option>)}</select>
      <span style={{fontSize:14,fontWeight:700,color:"#7a8a9a",marginLeft:12}}>DEALERSHIP:</span>
      <select value={loc} onChange={e=>setInternalLoc(e.target.value)} style={{...selStyle,maxWidth:340}}>
        <option value="">All Dealerships</option>
        {locations.map(l=><option key={l} value={l} disabled={!locsWithAdData.has(l)} style={{color:locsWithAdData.has(l)?C.main:"#ccc"}}>{l}{locsWithAdData.has(l)?"":" (no data)"}</option>)}
      </select>
    </div>}
    {captureMode&&<div style={{padding:"8px 0 12px",fontSize:18,fontWeight:800,color:C.main}}>{[hasG&&"Google Ads",hasF&&"Facebook Ads"].filter(Boolean).join(" & ")} {"\u2014"} {loc} {"\u2014"} {periods.find(pp=>pp.key===sp)?.label}</div>}
    {hasG&&<Section title="Google Ads" desc="Performance metrics for Google search ads promoting service.">
      <div style={{display:"flex",gap:14,flexWrap:"wrap"}}>
        <KPI label="Clicks" value={gCur.clicksGoogle} color={C.main} sub={gPrv?<Badge cur={gCur.clicksGoogle} prev={gPrv.clicksGoogle}/>:null}/>
        <KPI label="Page Views" value={gCur.pageViewsGoogle} color={C.acc1} sub={gPrv?<Badge cur={gCur.pageViewsGoogle} prev={gPrv.pageViewsGoogle}/>:null}/>
        <KPI label="Leads" value={gCur.leadsGoogle} color={C.sec} sub={gPrv?<Badge cur={gCur.leadsGoogle} prev={gPrv.leadsGoogle}/>:null}/>
        <KPI label="Phone Calls" value={gCur.phoneCallsGoogle} color={C.green} sub={gPrv?<Badge cur={gCur.phoneCallsGoogle} prev={gPrv.phoneCallsGoogle}/>:null}/>
        <KPI label="Spend" value={gCur.spendGoogle} fmt="money" color={C.amber}/>
      </div>
    </Section>}
    {hasF&&<Section title="Facebook Ads" desc="Performance metrics for Facebook ads promoting sales.">
      <div style={{display:"flex",gap:14,flexWrap:"wrap"}}>
        <KPI label="Clicks" value={fCur.clicksFB} color={C.main} sub={fPrv?<Badge cur={fCur.clicksFB} prev={fPrv.clicksFB}/>:null}/>
        <KPI label="Page Views" value={fCur.pageViewsFB} color={C.acc1} sub={fPrv?<Badge cur={fCur.pageViewsFB} prev={fPrv.pageViewsFB}/>:null}/>
        <KPI label="Leads" value={fCur.leadsFB} color={C.sec} sub={fPrv?<Badge cur={fCur.leadsFB} prev={fPrv.leadsFB}/>:null}/>
        <KPI label="Phone Calls" value={fCur.phoneCallsFB} color={C.green} sub={fPrv?<Badge cur={fCur.phoneCallsFB} prev={fPrv.phoneCallsFB}/>:null}/>
        <KPI label="Spend" value={fCur.spendFB} fmt="money" color={C.amber}/>
      </div>
    </Section>}
    {!captureMode&&<>{hasG&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
      <TrendChart data={filt} periods={last12} valueKey="clicksGoogle" label="Google Clicks" color={C.main} aggFn={aggGoogle} filterLoc={loc}/>
      <TrendChart data={filt} periods={last12} valueKey="leadsGoogle" label="Google Leads" color={C.sec} aggFn={aggGoogle} filterLoc={loc}/>
    </div>}
    {hasG&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
      <TrendChart data={filt} periods={last12} valueKey="pageViewsGoogle" label="Google Views" color={C.acc1} aggFn={aggGoogle} filterLoc={loc}/>
      <TrendChart data={filt} periods={last12} valueKey="phoneCallsGoogle" label="Google Phone Calls" color={C.green} aggFn={aggGoogle} filterLoc={loc}/>
    </div>}
    {hasF&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
      <TrendChart data={filt} periods={last12} valueKey="clicksFB" label="Facebook Clicks" color={C.main} aggFn={aggFB} filterLoc={loc}/>
      <TrendChart data={filt} periods={last12} valueKey="leadsFB" label="Facebook Leads" color={C.sec} aggFn={aggFB} filterLoc={loc}/>
    </div>}
    {hasF&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
      <TrendChart data={filt} periods={last12} valueKey="pageViewsFB" label="Facebook Views" color={C.acc1} aggFn={aggFB} filterLoc={loc}/>
      <TrendChart data={filt} periods={last12} valueKey="phoneCallsFB" label="Facebook Phone Calls" color={C.green} aggFn={aggFB} filterLoc={loc}/>
    </div>}</>}
    {!hasG&&!hasF&&<div style={{padding:40,textAlign:"center",color:"#999"}}>No advertising data for this selection.</div>}
  </div>);
}

function DealerSalesTab({data,periods,sp,setSp,locations}){
  const{sortState,onSort,doSort}=useSort("salesInfluenced","desc");
  const[newCustData,setNewCustData]=useState(null);
  const[prevNewCustData,setPrevNewCustData]=useState(null);
  const p=periods.find(pp=>pp.key===sp);const prev=getPrev(periods,sp);

  // Fetch new customer counts
  useEffect(()=>{
    if(!p)return;
    fetch(SHEETS_API+"?action=newcustomer&month="+p.month+"&year="+p.year)
      .then(r=>r.json()).then(d=>setNewCustData(d)).catch(()=>setNewCustData(null));
    if(prev){
      fetch(SHEETS_API+"?action=newcustomer&month="+prev.month+"&year="+prev.year)
        .then(r=>r.json()).then(d=>setPrevNewCustData(d)).catch(()=>setPrevNewCustData(null));
    }else{setPrevNewCustData(null);}
  },[p,prev]);

  const ncCounts=newCustData?.counts||{};
  const pncCounts=prevNewCustData?.counts||{};

  const rows=locations.map(loc=>{
    const cur=p?aggSales(filterPeriod(data,p).filter(r=>r.location===loc)):{salesEngaged:0,salesShoppers:0,salesLeads:0,salesInfluenced:0,salesWinback:0,salesGross:0,salesWinbackProfit:0,salesROI:0,winbackSalesPct:0,monthlyCost:0};
    const prv=prev?aggSales(filterPeriod(data,prev).filter(r=>r.location===loc)):null;
    // Match dealership name to sheet tab name (sheet tabs may not exactly match location names)
    // Try exact match first, then partial
    let nc=ncCounts[loc]??null;
    if(nc===null){const lcLoc=loc.toLowerCase();const match=Object.keys(ncCounts).find(k=>k.toLowerCase()===lcLoc||lcLoc.includes(k.toLowerCase())||k.toLowerCase().includes(lcLoc));if(match)nc=ncCounts[match];}
    let pnc=Object.keys(pncCounts).length>0?(pncCounts[loc]??null):null;
    if(pnc===null&&Object.keys(pncCounts).length>0){const lcLoc=loc.toLowerCase();const match=Object.keys(pncCounts).find(k=>k.toLowerCase()===lcLoc||lcLoc.includes(k.toLowerCase())||k.toLowerCase().includes(lcLoc));if(match)pnc=pncCounts[match];}
    return{location:loc,...cur,prevInfluenced:prv?.salesInfluenced??null,prevWinback:prv?.salesWinback??null,newCust:nc??0,prevNewCust:pnc};
  });
  const sorted=doSort(rows);
  const tot={salesEngaged:0,salesShoppers:0,salesLeads:0,salesInfluenced:0,salesWinback:0,salesGross:0,newCust:0};
  rows.forEach(r=>{for(const k in tot)tot[k]+=r[k];});
  const totPrevInf=rows.every(r=>r.prevInfluenced===null)?null:rows.reduce((a,r)=>a+(r.prevInfluenced??0),0);
  const totPrevWb=rows.every(r=>r.prevWinback===null)?null:rows.reduce((a,r)=>a+(r.prevWinback??0),0);
  const totPrevNC=rows.every(r=>r.prevNewCust===null)?null:rows.reduce((a,r)=>a+(r.prevNewCust??0),0);
  const activeRows=rows.filter(r=>r.salesInfluenced>0||r.salesEngaged>0);
  const avgROI=activeRows.length>0?(activeRows.reduce((a,r)=>a+r.salesROI,0)/activeRows.length):0;
  const avgWbPct=activeRows.length>0?(activeRows.reduce((a,r)=>a+r.winbackSalesPct,0)/activeRows.length):0;
  const td=i=>({padding:"10px 12px",fontSize:13,background:i%2===0?"#f4f7fc":"white"});
  const ts={padding:"11px 12px",fontSize:13,fontWeight:800,background:"#e8edf5",borderTop:`2px solid ${C.main}`};
  return(<div>
    <div style={{padding:"16px 0",display:"flex",gap:14,alignItems:"center",flexWrap:"wrap",marginBottom:16}}>
      <span style={{fontSize:14,fontWeight:700,color:"#7a8a9a"}}>PERIOD:</span>
      <select value={sp} onChange={e=>setSp(e.target.value)} style={selStyle}>{periods.map(p=><option key={p.key} value={p.key}>{p.label}</option>)}</select>
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
          <SortHeader label="WB vs PM" sortKey="prevWinback" sortState={sortState} onSort={onSort}/>
          <SortHeader label="New Cust." sortKey="newCust" sortState={sortState} onSort={onSort}/>
          <SortHeader label="NC vs PM" sortKey="prevNewCust" sortState={sortState} onSort={onSort}/>
          <SortHeader label="Gross" sortKey="salesGross" sortState={sortState} onSort={onSort}/>
          <SortHeader label="ROI %" sortKey="salesROI" sortState={sortState} onSort={onSort}/>
          <SortHeader label="WB %" sortKey="winbackSalesPct" sortState={sortState} onSort={onSort} last/>
        </tr></thead>
        <tbody>
          {sorted.map((r,i)=>(<tr key={r.location} onMouseEnter={e=>e.currentTarget.style.background="#e4edff"} onMouseLeave={e=>e.currentTarget.style.background=i%2===0?"#f4f7fc":"white"}>
            <td style={{...td(i),fontWeight:800,color:i<3?C.sec:"#8896a4"}}>{i+1}</td>
            <td style={{...td(i),fontWeight:600,color:C.main,maxWidth:240,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{r.location}</td>
            <td style={{...td(i),textAlign:"right"}}>{fmtNum(r.salesEngaged)}</td>
            <td style={{...td(i),textAlign:"right"}}>{fmtNum(r.salesShoppers)}</td>
            <td style={{...td(i),textAlign:"right"}}>{fmtNum(r.salesLeads)}</td>
            <td style={{...td(i),textAlign:"right",fontWeight:700,color:C.sec}}>{fmtNum(r.salesInfluenced)}</td>
            <td style={{...td(i),textAlign:"right"}}><Badge cur={r.salesInfluenced} prev={r.prevInfluenced}/></td>
            <td style={{...td(i),textAlign:"right",color:C.purple,fontWeight:600}}>{fmtNum(r.salesWinback)}</td>
            <td style={{...td(i),textAlign:"right"}}><Badge cur={r.salesWinback} prev={r.prevWinback}/></td>
            <td style={{...td(i),textAlign:"right",color:C.teal,fontWeight:700}}>{fmtNum(r.newCust)}</td>
            <td style={{...td(i),textAlign:"right"}}><Badge cur={r.newCust} prev={r.prevNewCust}/></td>
            <td style={{...td(i),textAlign:"right",fontWeight:600}}>{fmtMoney(r.salesGross)}</td>
            <td style={{...td(i),textAlign:"right"}}><span style={{background:r.salesROI>=200?"#d4edda":r.salesROI>=100?"#fff3cd":"#f8d7da",color:r.salesROI>=200?"#155724":r.salesROI>=100?"#856404":"#721c24",padding:"3px 10px",borderRadius:20,fontWeight:700,fontSize:12}}>{fmtPct(r.salesROI)}</span></td>
            <td style={{...td(i),textAlign:"right"}}>{fmtPct(r.winbackSalesPct)}</td>
          </tr>))}
          <tr>
            <td style={{...ts,borderRadius:"0 0 0 10px"}}></td>
            <td style={{...ts,textAlign:"left",color:C.main,fontSize:14}}>GROUP TOTALS</td>
            <td style={{...ts,textAlign:"right"}}>{fmtNum(tot.salesEngaged)}</td>
            <td style={{...ts,textAlign:"right"}}>{fmtNum(tot.salesShoppers)}</td>
            <td style={{...ts,textAlign:"right"}}>{fmtNum(tot.salesLeads)}</td>
            <td style={{...ts,textAlign:"right",color:C.sec}}>{fmtNum(tot.salesInfluenced)}</td>
            <td style={{...ts,textAlign:"right"}}><Badge cur={tot.salesInfluenced} prev={totPrevInf}/></td>
            <td style={{...ts,textAlign:"right",color:C.purple}}>{fmtNum(tot.salesWinback)}</td>
            <td style={{...ts,textAlign:"right"}}><Badge cur={tot.salesWinback} prev={totPrevWb}/></td>
            <td style={{...ts,textAlign:"right",color:C.teal}}>{fmtNum(tot.newCust)}</td>
            <td style={{...ts,textAlign:"right"}}><Badge cur={tot.newCust} prev={totPrevNC}/></td>
            <td style={{...ts,textAlign:"right"}}>{fmtMoney(tot.salesGross)}</td>
            <td style={{...ts,textAlign:"right"}}><span style={{background:avgROI>=200?"#d4edda":avgROI>=100?"#fff3cd":"#f8d7da",color:avgROI>=200?"#155724":avgROI>=100?"#856404":"#721c24",padding:"3px 10px",borderRadius:20,fontWeight:800,fontSize:13}}>{fmtPct(avgROI)}</span></td>
            <td style={{...ts,textAlign:"right",borderRadius:"0 0 10px 0"}}>{fmtPct(avgWbPct)}</td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>);
}

function DealerServiceTab({data,periods,sp,setSp,locations}){
  const{sortState,onSort,doSort}=useSort("serviceROs","desc");
  const p=periods.find(pp=>pp.key===sp);const prev=getPrev(periods,sp);
  const rows=locations.map(loc=>{
    const pRows=p?filterPeriod(data,p).filter(r=>r.location===loc):[];const prRows=prev?filterPeriod(data,prev).filter(r=>r.location===loc):[];
    const svc=aggService(pRows);const prvSvc=prRows.length?aggService(prRows):null;
    return{location:loc,...svc,prevROs:prvSvc?.serviceROs??null,prevWinbackROs:prvSvc?.serviceWinbackROs??null};
  });
  const sorted=doSort(rows);
  const tot={serviceShoppers:0,serviceLeads:0,serviceROs:0,serviceWinbackROs:0,roValue:0,serviceWinbackROValue:0};
  rows.forEach(r=>{for(const k in tot)tot[k]+=r[k];});
  const totWbPct=tot.serviceROs>0?((tot.serviceWinbackROs/tot.serviceROs)*100):0;
  const totPrevROs=rows.every(r=>r.prevROs===null)?null:rows.reduce((a,r)=>a+(r.prevROs??0),0);
  const totPrevWbROs=rows.every(r=>r.prevWinbackROs===null)?null:rows.reduce((a,r)=>a+(r.prevWinbackROs??0),0);
  const td=i=>({padding:"10px 12px",fontSize:13,background:i%2===0?"#f4f7fc":"white"});
  const ts={padding:"11px 12px",fontSize:13,fontWeight:800,background:"#e8edf5",borderTop:`2px solid ${C.main}`};
  return(<div>
    <div style={{padding:"16px 0",display:"flex",gap:14,alignItems:"center",flexWrap:"wrap",marginBottom:16}}>
      <span style={{fontSize:14,fontWeight:700,color:"#7a8a9a"}}>PERIOD:</span>
      <select value={sp} onChange={e=>setSp(e.target.value)} style={selStyle}>{periods.map(p=><option key={p.key} value={p.key}>{p.label}</option>)}</select>
    </div>
    <div style={{overflowX:"auto"}}>
      <table style={{width:"100%",borderCollapse:"separate",borderSpacing:0}}>
        <thead><tr>
          <th style={{padding:"11px 12px",fontWeight:700,fontSize:13,textAlign:"left",background:C.main,color:"white",borderRadius:"10px 0 0 0"}}>#</th>
          <SortHeader label="Dealership" sortKey="location" sortState={sortState} onSort={onSort} align="left"/>
          <SortHeader label="Shoppers" sortKey="serviceShoppers" sortState={sortState} onSort={onSort}/>
          <SortHeader label="Svc Leads" sortKey="serviceLeads" sortState={sortState} onSort={onSort}/>
          <SortHeader label="RO's" sortKey="serviceROs" sortState={sortState} onSort={onSort}/>
          <SortHeader label="vs PM" sortKey="prevROs" sortState={sortState} onSort={onSort}/>
          <SortHeader label="WB RO's" sortKey="serviceWinbackROs" sortState={sortState} onSort={onSort}/>
          <SortHeader label="WB vs PM" sortKey="prevWinbackROs" sortState={sortState} onSort={onSort}/>
          <SortHeader label="RO Value" sortKey="roValue" sortState={sortState} onSort={onSort}/>
          <SortHeader label="WB RO Value" sortKey="serviceWinbackROValue" sortState={sortState} onSort={onSort}/>
          <SortHeader label="WB %" sortKey="winbackServicePct" sortState={sortState} onSort={onSort} last/>
        </tr></thead>
        <tbody>
          {sorted.map((r,i)=>(<tr key={r.location} onMouseEnter={e=>e.currentTarget.style.background="#e4edff"} onMouseLeave={e=>e.currentTarget.style.background=i%2===0?"#f4f7fc":"white"}>
            <td style={{...td(i),fontWeight:800,color:i<3?C.sec:"#8896a4"}}>{i+1}</td>
            <td style={{...td(i),fontWeight:600,color:C.main,maxWidth:220,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{r.location}</td>
            <td style={{...td(i),textAlign:"right"}}>{fmtNum(r.serviceShoppers)}</td>
            <td style={{...td(i),textAlign:"right"}}>{fmtNum(r.serviceLeads)}</td>
            <td style={{...td(i),textAlign:"right",fontWeight:700,color:C.green}}>{fmtNum(r.serviceROs)}</td>
            <td style={{...td(i),textAlign:"right"}}><Badge cur={r.serviceROs} prev={r.prevROs}/></td>
            <td style={{...td(i),textAlign:"right",color:C.purple}}>{fmtNum(r.serviceWinbackROs)}</td>
            <td style={{...td(i),textAlign:"right"}}><Badge cur={r.serviceWinbackROs} prev={r.prevWinbackROs}/></td>
            <td style={{...td(i),textAlign:"right",fontWeight:600}}>{fmtMoney(r.roValue)}</td>
            <td style={{...td(i),textAlign:"right",color:C.purple,fontWeight:600}}>{fmtMoney(r.serviceWinbackROValue)}</td>
            <td style={{...td(i),textAlign:"right"}}>{fmtPct(r.winbackServicePct)}</td>
          </tr>))}
          <tr>
            <td style={{...ts,borderRadius:"0 0 0 10px"}}></td>
            <td style={{...ts,textAlign:"left",color:C.main,fontSize:14}}>GROUP TOTALS</td>
            <td style={{...ts,textAlign:"right"}}>{fmtNum(tot.serviceShoppers)}</td>
            <td style={{...ts,textAlign:"right"}}>{fmtNum(tot.serviceLeads)}</td>
            <td style={{...ts,textAlign:"right",color:C.green}}>{fmtNum(tot.serviceROs)}</td>
            <td style={{...ts,textAlign:"right"}}><Badge cur={tot.serviceROs} prev={totPrevROs}/></td>
            <td style={{...ts,textAlign:"right",color:C.purple}}>{fmtNum(tot.serviceWinbackROs)}</td>
            <td style={{...ts,textAlign:"right"}}><Badge cur={tot.serviceWinbackROs} prev={totPrevWbROs}/></td>
            <td style={{...ts,textAlign:"right"}}>{fmtMoney(tot.roValue)}</td>
            <td style={{...ts,textAlign:"right",color:C.purple}}>{fmtMoney(tot.serviceWinbackROValue)}</td>
            <td style={{...ts,textAlign:"right",borderRadius:"0 0 10px 0"}}><span style={{background:totWbPct>=30?"#d4edda":totWbPct>=15?"#fff3cd":"#f8d7da",color:totWbPct>=30?"#155724":totWbPct>=15?"#856404":"#721c24",padding:"3px 10px",borderRadius:20,fontWeight:800,fontSize:13}}>{fmtPct(totWbPct)}</span></td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>);
}

function CustomerDataTab({sheetTabs,locations}){
  const[selectedTab,setSelectedTab]=useState("");
  const[custData,setCustData]=useState(null);const[custHeaders,setCustHeaders]=useState([]);
  const[loading,setLoading]=useState(false);const[error,setError]=useState("");
  const[monthFilter,setMonthFilter]=useState("");
  const{sortState,onSort,doSort}=useSort("_sortYear","desc");

  const apiFailed=sheetTabs.length===1&&(sheetTabs[0].name==="_API_ERROR_"||sheetTabs[0].name==="_DISCOVERY_FAILED_");
  const validTabs=useMemo(()=>{
    const tabs=apiFailed?[]:sheetTabs;
    return[...tabs].sort((a,b)=>a.name.localeCompare(b.name));
  },[sheetTabs,apiFailed]);

  const HIDDEN_COLS=["front end gross","back end gross","frontend gross","backend gross","front-end gross","back-end gross"];
  const DOLLAR_COLS=["total gross","winback profit","gross profit","profit"];
  const isHidden=(h)=>{const lc=h.toLowerCase();return HIDDEN_COLS.some(hc=>lc.includes(hc));};
  const isDollar=(h)=>{const lc=h.toLowerCase();return DOLLAR_COLS.some(dc=>lc.includes(dc));};
  const fmtCell=(val,h)=>{
    if(!isDollar(h))return val;
    const n=parseFloat(String(val).replace(/[$,\s]/g,""));
    if(isNaN(n))return val;
    return"$"+n.toLocaleString(undefined,{minimumFractionDigits:0,maximumFractionDigits:0});
  };

  const loadSheet=useCallback(async(sheetName)=>{
    if(!sheetName)return;setLoading(true);setError("");setCustData(null);setMonthFilter("");
    try{
      const resp=await fetch(SHEETS_API+"?action=data&sheetName="+encodeURIComponent(sheetName));
      if(!resp.ok){const e=await resp.json().catch(()=>({}));throw new Error(e.error||"HTTP "+resp.status);}
      const{headers,data:records}=await resp.json();
      if(records&&records.length>0){
        const visibleHdrs=(headers||[]).filter(h=>h.trim()!==""&&!isHidden(h));
        setCustHeaders(visibleHdrs);
        const allHdrs=(headers||[]).filter(h=>h.trim()!=="");
        const yearCol=allHdrs.findIndex(h=>{const lc=h.toLowerCase();return lc==="year";});
        const monthCol=allHdrs.findIndex(h=>{const lc=h.toLowerCase();return lc==="month";});
        const dateCol=allHdrs.findIndex(h=>{const lc=h.toLowerCase();return lc.includes("date");});
        const MMAP={january:1,february:2,march:3,april:4,may:5,june:6,july:7,august:8,september:9,october:10,november:11,december:12,jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};
        const mapped=records.map((row,idx)=>{
          const obj={_idx:idx,_sortYear:0,_sortMonth:0};
          visibleHdrs.forEach((h,ci)=>{obj["col"+ci]=row[h]||"";});
          if(yearCol>=0)obj._sortYear=parseInt(row[allHdrs[yearCol]])||0;
          if(monthCol>=0){const mv=(row[allHdrs[monthCol]]||"").trim().toLowerCase();obj._sortMonth=MMAP[mv]||parseInt(mv)||0;}
          if(dateCol>=0&&!obj._sortYear){
            const ds=row[allHdrs[dateCol]]||"";const pd=new Date(ds);
            if(!isNaN(pd.getTime())){obj._sortYear=pd.getFullYear();obj._sortMonth=pd.getMonth()+1;}
          }
          return obj;
        });
        mapped.sort((a,b)=>b._sortYear-a._sortYear||b._sortMonth-a._sortMonth);
        setCustData(mapped);
      }else{setError("No data found in this sheet.");}
    }catch(e){setError(e.message);}
    setLoading(false);
  },[]);

  useEffect(()=>{
    if(selectedTab&&validTabs.length){loadSheet(selectedTab);}
  },[selectedTab,validTabs,loadSheet]);

  // Build available month/year options from loaded data
  const monthOptions=useMemo(()=>{
    if(!custData)return[];
    const set=new Set();
    custData.forEach(r=>{if(r._sortYear&&r._sortMonth)set.add(`${r._sortYear}-${String(r._sortMonth).padStart(2,"0")}`);});
    return[...set].sort().reverse().map(k=>{const[y,m]=k.split("-").map(Number);return{key:k,label:`${MN[m]} ${y}`};});
  },[custData]);

  // Filter and sort
  const filtered=useMemo(()=>{
    if(!custData)return[];
    if(!monthFilter)return custData;
    const[fy,fm]=monthFilter.split("-").map(Number);
    return custData.filter(r=>r._sortYear===fy&&r._sortMonth===fm);
  },[custData,monthFilter]);
  const sorted=doSort(filtered);

  return(<div>
    {validTabs.length>0&&(<div style={{padding:"16px 0",display:"flex",gap:14,alignItems:"center",flexWrap:"wrap",marginBottom:16}}>
      <span style={{fontSize:14,fontWeight:700,color:"#7a8a9a"}}>DEALERSHIP:</span>
      <select value={selectedTab} onChange={e=>setSelectedTab(e.target.value)} style={{...selStyle,maxWidth:400}}>
        <option value="">{"\u2014"} Select Dealership {"\u2014"}</option>
        {validTabs.map(t=><option key={t.sheetId} value={t.name}>{t.name}</option>)}
      </select>
      {monthOptions.length>0&&(<>
        <span style={{fontSize:14,fontWeight:700,color:"#7a8a9a",marginLeft:12}}>MONTH:</span>
        <select value={monthFilter} onChange={e=>setMonthFilter(e.target.value)} style={selStyle}>
          <option value="">All Months</option>
          {monthOptions.map(o=><option key={o.key} value={o.key}>{o.label}</option>)}
        </select>
      </>)}
      <span style={{fontSize:13,color:"#7a8a9a"}}>Customer-specific data from email marketing for <strong>Sales</strong></span>
    </div>)}

    {apiFailed&&!loading&&!custData&&(<div style={{background:"white",borderRadius:14,padding:32,boxShadow:"0 2px 8px rgba(0,0,0,0.06)",maxWidth:700,margin:"20px auto",textAlign:"left"}}>
      <div style={{fontSize:20,fontWeight:800,color:C.main,marginBottom:16}}>{"\u26A0\uFE0F"} API Setup Required</div>
      <p style={{fontSize:14,color:"#555",lineHeight:1.7,marginBottom:16}}>The Customer Data tab needs two environment variables set in your Vercel project settings. Check the deployment guide for step-by-step instructions:</p>
      <div style={{background:"#f0f4ff",borderRadius:10,padding:20,marginBottom:12}}>
        <div style={{fontWeight:700,color:C.main,fontSize:13,marginBottom:4}}>GOOGLE_API_KEY</div>
        <div style={{fontSize:13,color:"#555"}}>A Google Cloud API key with Google Sheets API enabled</div>
      </div>
      <div style={{background:"#f0f4ff",borderRadius:10,padding:20}}>
        <div style={{fontWeight:700,color:C.main,fontSize:13,marginBottom:4}}>GOOGLE_SHEET_ID</div>
        <div style={{fontSize:13,color:"#555"}}>The spreadsheet ID from your Google Sheet URL</div>
      </div>
    </div>)}

    {sheetTabs.length===0&&!loading&&!custData&&(<div style={{padding:60,textAlign:"center",color:"#aab"}}>
      <div style={{width:40,height:40,border:"4px solid "+C.sec,borderTopColor:"transparent",borderRadius:"50%",animation:"spin 1s linear infinite",margin:"0 auto 16px"}}/>
      <div style={{fontSize:14,color:"#888"}}>Loading dealership list...</div>
    </div>)}

    {validTabs.length>0&&!selectedTab&&!loading&&!custData&&(<div style={{padding:60,textAlign:"center",color:"#aab"}}>
      <div style={{fontSize:48,marginBottom:12}}>{"\ud83d\udcca"}</div>
      <div style={{fontSize:16,fontWeight:600}}>Select a dealership above to view customer-specific data</div>
    </div>)}

    {loading&&<div style={{padding:40,textAlign:"center",color:C.sec,fontSize:16,fontWeight:600}}>
      <div style={{width:40,height:40,border:"4px solid "+C.sec,borderTopColor:"transparent",borderRadius:"50%",animation:"spin 1s linear infinite",margin:"0 auto 16px"}}/>
      Loading customer data...
    </div>}
    {error&&<div style={{padding:30,textAlign:"center"}}><div style={{background:"#fff3f3",border:"1px solid #fecaca",borderRadius:12,padding:20,display:"inline-block",maxWidth:600,textAlign:"left"}}>
      <div style={{fontWeight:700,color:C.red,marginBottom:8}}>{"\u26A0\uFE0F"} Error Loading Data</div>
      <div style={{fontSize:13,color:"#666",lineHeight:1.6}}>{error}</div>
    </div></div>}

    {sorted.length>0&&(<div style={{overflowX:"auto",maxHeight:700,overflowY:"auto"}}>
      <table style={{width:"100%",borderCollapse:"separate",borderSpacing:0}}>
        <thead><tr>{custHeaders.map((h,ci)=>(<SortHeader key={ci} label={h} sortKey={"col"+ci} sortState={sortState} onSort={onSort} align={ci===0?"left":"right"} first={ci===0} last={ci===custHeaders.length-1}/>))}</tr></thead>
        <tbody>{sorted.map((row,i)=>(<tr key={row._idx} onMouseEnter={e=>e.currentTarget.style.background="#e4edff"} onMouseLeave={e=>e.currentTarget.style.background=i%2===0?"#f4f7fc":"white"}>
          {custHeaders.map((h,ci)=>(<td key={ci} style={{padding:"9px 12px",fontSize:13,textAlign:ci===0?"left":"right",background:i%2===0?"#f4f7fc":"white",whiteSpace:"nowrap",maxWidth:300,overflow:"hidden",textOverflow:"ellipsis",fontWeight:ci===0?600:isDollar(h)?600:400,color:ci===0?C.main:isDollar(h)?C.green:undefined}}>{fmtCell(row["col"+ci],h)}</td>))}
        </tr>))}</tbody>
      </table>
      <div style={{padding:"12px 0",fontSize:13,color:"#7a8a9a",textAlign:"right"}}>{sorted.length} customer records{monthFilter?" (filtered)":""}</div>
    </div>)}
    {custData&&sorted.length===0&&monthFilter&&(<div style={{padding:40,textAlign:"center",color:"#999"}}>No records found for {monthOptions.find(o=>o.key===monthFilter)?.label||monthFilter}.</div>)}
  </div>);
}

async function captureTab(ref,opts={}){return html2canvas(ref,{backgroundColor:opts.bg||"#edf1f7",scale:opts.scale||2,useCORS:true,logging:false,width:opts.width||undefined});}

export default function App(){
  const[data,setData]=useState([]);const[loading,setLoading]=useState(true);const[error,setError]=useState("");
  const[activeTab,setActiveTab]=useState("groupSales");const[sp,setSp]=useState("");
  const[sheetTabs,setSheetTabs]=useState([]);const[refreshing,setRefreshing]=useState(false);
  const[downloading,setDownloading]=useState(false);const[dlProgress,setDlProgress]=useState("");
  const[captureMode,setCaptureMode]=useState(false);const[captureLoc,setCaptureLoc]=useState("");
  const contentRef=useRef(null);

  const fetchData=useCallback(async(isRefresh)=>{
    if(isRefresh)setRefreshing(true);else setLoading(true);setError("");
    try{const resp=await fetch(SHEET_CSV_URL);if(!resp.ok)throw new Error(`HTTP ${resp.status}`);
      const text=await resp.text();const parsed=parseOverallCSV(text);
      if(parsed.length===0)throw new Error("No valid data rows found.");
      setData(parsed);const periods=getPeriods(parsed);
      if(periods.length>0)setSp(prev=>prev&&periods.find(p=>p.key===prev)?prev:periods[periods.length-1].key);
    }catch(e){setError(e.message);}
    setLoading(false);setRefreshing(false);
  },[]);

  const discoverTabs=useCallback(async()=>{
    try{
      const resp=await fetch(SHEETS_API+"?action=list");
      if(!resp.ok){const e=await resp.json().catch(()=>({}));console.warn("[TabDiscovery] API error:",e.error||resp.status);
        setSheetTabs([{gid:"_NONE_",name:"_API_ERROR_"}]);return;}
      const{sheets}=await resp.json();
      // Filter out the "Overall Results" tab — keep only dealership tabs
      const dealerTabs=(sheets||[]).filter(s=>{const lc=s.title.toLowerCase();return!lc.includes("overall")&&!lc.includes("sheet1")&&!lc.includes("template")&&s.title.length>1;})
        .map(s=>({sheetId:s.sheetId,name:s.title}));
      console.log("[TabDiscovery] Found",dealerTabs.length,"dealership tabs:",dealerTabs.map(t=>t.name));
      if(dealerTabs.length>0)setSheetTabs(dealerTabs);
      else setSheetTabs([{sheetId:"_NONE_",name:"_DISCOVERY_FAILED_"}]);
    }catch(e){console.warn("[TabDiscovery] Failed:",e);setSheetTabs([{sheetId:"_NONE_",name:"_API_ERROR_"}]);}
  },[]);

  useEffect(()=>{fetchData(false);discoverTabs();},[fetchData,discoverTabs]);
  const periods=useMemo(()=>getPeriods(data),[data]);
  const locations=useMemo(()=>[...new Set(data.map(r=>r.location))].filter(l=>l).sort(),[data]);
  const locsWithAdData=useMemo(()=>{const s=new Set();data.forEach(r=>{if(r.clicksGoogle>0||r.leadsGoogle>0||r.spendGoogle>0||r.pageViewsGoogle>0||r.phoneCallsGoogle>0||r.clicksFB>0||r.leadsFB>0||r.spendFB>0||r.pageViewsFB>0||r.phoneCallsFB>0)s.add(r.location);});return s;},[data]);

  const downloadJPG=async()=>{if(!contentRef.current||downloading)return;setDownloading(true);
    try{const canvas=await captureTab(contentRef.current);const link=document.createElement("a");
      link.download=`Garber_Fullpath_${TABS.find(t=>t.id===activeTab)?.label.replace(/\s+/g,"_")}_${sp}.jpg`;
      link.href=canvas.toDataURL("image/jpeg",0.95);link.click();}catch(e){alert("Export failed: "+e.message);}
    setDownloading(false);};
  const downloadPDF=async()=>{if(!contentRef.current||downloading)return;setDownloading(true);
    try{const canvas=await captureTab(contentRef.current);const imgData=canvas.toDataURL("image/jpeg",0.92);
      const pdf=new jsPDF({orientation:canvas.width>canvas.height?"landscape":"portrait",unit:"px",format:[canvas.width,canvas.height]});
      pdf.addImage(imgData,"JPEG",0,0,canvas.width,canvas.height);
      pdf.save(`Garber_Fullpath_${TABS.find(t=>t.id===activeTab)?.label.replace(/\s+/g,"_")}_${sp}.pdf`);}catch(e){alert("Export failed: "+e.message);}
    setDownloading(false);};
  const downloadAll=async()=>{if(downloading)return;setDownloading(true);setCaptureMode(true);
    try{const zip=new JSZip();const origTab=activeTab;
      const captureTabs=[
        {id:"groupSales",label:"Sales_Email",locs:locations},
        {id:"groupService",label:"Service_Email",locs:locations},
        {id:"groupAds",label:"Advertising",locs:locations.filter(l=>locsWithAdData.has(l))},
      ];
      let count=0;const total=captureTabs.reduce((a,t)=>a+t.locs.length,0);
      for(const tab of captureTabs){
        setActiveTab(tab.id);
        for(const loc of tab.locs){
          count++;setDlProgress(`Capturing ${tab.label} - ${loc} (${count}/${total})...`);
          setCaptureLoc(loc);
          await new Promise(r=>setTimeout(r,500));
          if(contentRef.current){
            const canvas=await captureTab(contentRef.current,{scale:2,bg:"#edf1f7"});
            const blob=await new Promise(resolve=>canvas.toBlob(resolve,"image/jpeg",0.95));
            const safeLoc=loc.replace(/[^a-zA-Z0-9 ]/g,"").replace(/\s+/g,"_");
            zip.file(`${safeLoc}_${tab.label}_${sp}.jpg`,blob);
          }
        }
      }
      setActiveTab(origTab);setCaptureMode(false);setCaptureLoc("");
      setDlProgress("Creating ZIP...");
      const content=await zip.generateAsync({type:"blob"});const link=document.createElement("a");
      link.download=`Garber_Fullpath_Report_${sp}.zip`;link.href=URL.createObjectURL(content);link.click();
    }catch(e){alert("Export failed: "+e.message);setCaptureMode(false);setCaptureLoc("");}
    setDlProgress("");setDownloading(false);};

  if(loading)return(<div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",background:C.bg,fontFamily:"'DM Sans',system-ui,sans-serif"}}>
    <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800;900&display=swap" rel="stylesheet"/>
    <div style={{textAlign:"center"}}><div style={{fontSize:48,marginBottom:16}}>{"\ud83d\udcca"}</div>
    <div style={{fontSize:20,fontWeight:700,color:C.main}}>Loading dashboard data...</div></div></div>);

  if(error)return(<div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:C.bg,fontFamily:"'DM Sans',system-ui,sans-serif"}}>
    <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800;900&display=swap" rel="stylesheet"/>
    <div style={{textAlign:"center",maxWidth:500,padding:40}}><div style={{fontSize:48,marginBottom:16}}>{"\u26A0\uFE0F"}</div>
    <div style={{fontSize:20,fontWeight:700,color:C.red,marginBottom:12}}>Error Loading Data</div>
    <div style={{fontSize:15,color:"#666",marginBottom:24}}>{error}</div>
    <button onClick={()=>fetchData(false)} style={{padding:"12px 28px",borderRadius:10,background:C.sec,color:"white",fontWeight:700,fontSize:15,border:"none",cursor:"pointer"}}>Retry</button></div></div>);

  return(<div style={{fontFamily:"'DM Sans',system-ui,sans-serif",background:C.bg,minHeight:"100vh",color:C.main,fontSize:16}}>
    <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800;900&display=swap" rel="stylesheet"/>
    {dlProgress&&(<div style={{position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(7,42,96,0.85)",zIndex:9999,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:20}}>
      <div style={{fontSize:48,animation:"spin 2s linear infinite"}}>{"\ud83d\udce6"}</div>
      <div style={{fontSize:22,fontWeight:700,color:"white"}}>{dlProgress}</div>
      <div style={{fontSize:14,color:"rgba(255,255,255,0.6)"}}>Please wait while all pages are captured...</div></div>)}

    <div style={{background:`linear-gradient(135deg,${C.main} 0%,#0a3d7a 50%,${C.sec} 100%)`,padding:"28px 36px 20px",color:"white"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:16}}>
        <div>
          <div style={{fontSize:12,fontWeight:700,letterSpacing:3,opacity:0.5,textTransform:"uppercase"}}>Garber Automotive Group</div>
          <h1 style={{margin:"4px 0 0",fontSize:28,fontWeight:900}}>Fullpath Results Dashboard</h1>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <button onClick={downloadJPG} disabled={downloading} style={{cursor:downloading?"wait":"pointer",background:"rgba(255,255,255,0.15)",padding:"10px 16px",borderRadius:10,fontSize:13,fontWeight:700,border:"1px solid rgba(255,255,255,0.25)",color:"white",transition:"all 0.2s",opacity:downloading?0.6:1}}>{"\ud83d\udcf7"} JPG</button>
          <button onClick={downloadPDF} disabled={downloading} style={{cursor:downloading?"wait":"pointer",background:"rgba(255,255,255,0.15)",padding:"10px 16px",borderRadius:10,fontSize:13,fontWeight:700,border:"1px solid rgba(255,255,255,0.25)",color:"white",transition:"all 0.2s",opacity:downloading?0.6:1}}>{"\ud83d\udcc4"} PDF</button>
          <button onClick={downloadAll} disabled={downloading} style={{cursor:downloading?"wait":"pointer",background:"rgba(255,255,255,0.25)",padding:"10px 16px",borderRadius:10,fontSize:13,fontWeight:700,border:"1px solid rgba(255,255,255,0.35)",color:"white",transition:"all 0.2s",opacity:downloading?0.6:1}}>{"\ud83d\udce6"} All</button>
          <div style={{width:1,height:28,background:"rgba(255,255,255,0.2)",margin:"0 4px"}}></div>
          <button onClick={()=>fetchData(true)} disabled={refreshing||downloading} style={{cursor:refreshing?"wait":"pointer",background:"rgba(255,255,255,0.15)",padding:"10px 20px",borderRadius:10,fontSize:14,fontWeight:700,border:"1px solid rgba(255,255,255,0.25)",color:"white",display:"flex",alignItems:"center",gap:8,transition:"all 0.2s",opacity:refreshing?0.6:1}}>
            <span style={{display:"inline-block",animation:refreshing?"spin 1s linear infinite":"none"}}>{"\u21BB"}</span>
            {refreshing?"Refreshing...":"Refresh Data"}
          </button>
          <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
        </div>
      </div>
      <div style={{display:"flex",gap:6,marginTop:20}}>
        {TABS.map(t=>(<button key={t.id} onClick={()=>setActiveTab(t.id)} style={{padding:"10px 22px",border:"none",borderRadius:"10px 10px 0 0",cursor:"pointer",fontSize:14,fontWeight:700,background:activeTab===t.id?"white":"rgba(255,255,255,0.08)",color:activeTab===t.id?C.main:"rgba(255,255,255,0.6)",transition:"all 0.2s"}}>{t.icon} {t.label}</button>))}
      </div>
    </div>

    <div ref={contentRef} style={{padding:captureMode?"24px 28px 20px":"24px 36px 48px",maxWidth:captureMode?1200:undefined}}>
      {activeTab==="groupSales"&&<GroupSalesTab data={data} periods={periods} sp={sp} setSp={setSp} locations={locations} captureMode={captureMode} captureLoc={captureLoc}/>}
      {activeTab==="groupService"&&<GroupServiceTab data={data} periods={periods} sp={sp} setSp={setSp} locations={locations} captureMode={captureMode} captureLoc={captureLoc}/>}
      {activeTab==="groupAds"&&<GroupAdsTab data={data} periods={periods} sp={sp} setSp={setSp} locations={locations} captureMode={captureMode} captureLoc={captureLoc}/>}
      {activeTab==="dealerSales"&&<DealerSalesTab data={data} periods={periods} sp={sp} setSp={setSp} locations={locations}/>}
      {activeTab==="dealerService"&&<DealerServiceTab data={data} periods={periods} sp={sp} setSp={setSp} locations={locations}/>}
      {activeTab==="customerData"&&<CustomerDataTab sheetTabs={sheetTabs} locations={locations}/>}
    </div>
  </div>);
}
