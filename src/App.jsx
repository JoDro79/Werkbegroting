import { useState, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import PagePlanning from "./PagePlanning.jsx";
import PageDashboard from "./PageDashboard.jsx";

const TARIEF_RAW = [["202010","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Aanbrengen en verwijderen gronddepot.","st",7746.17,31513.11,0,0,8646.44,22866.67,0,0,0,0,0,0,0,0,0],["202101","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Transporteren materialen uit Logistieke Depot","uur",129.53,90,0,0,0,0,0,0,90,0,0,0,0,0,0],["202210","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Inrichten en opruimen werkterrein.","st",21645.54,26447,0,26447,0,0,0,0,0,0,0,0,0,0,0],["202220","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Toepassen beveiliging materiaaldepot.","st*wk",286.76,112.66,0,112.66,0,0,0,0,0,0,0,0,0,0,0],["202230","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Toepassen beveiliging sleuf.","st*wk",286.76,170.13,0,170.13,0,0,0,0,0,0,0,0,0,0,0],["202240","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Toepassen tijdelijke afrastering bouwhekken.","m",9.26,11.87,0,0.55,11.32,0,0,0,0,0,0,0,0,0,0],["202330","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Nul- en eindopname eenmalige kosten","keer",2573.5,1280,0,0,0,0,0,1280,0,0,0,0,0,0,0],["202340","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Nul- en eindopname trace.","km",807.45,145.56,0,0,0,0,0,145.56,0,0,0,0,0,0,0],["202410","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Toepassen boombescherming.","st",70.22,26.47,0,0,2.94,0,0,0,0,0,23.53,0,0,0,0],["207010","VOORBEREIDENDE WERKZAAMHEDEN","Straatmeubilair","Opnemen en terugplaatsen verkeersbordpaal.","st",125.97,117.2,0,0,117.2,0,0,0,0,0,0,0,0,0,0],["207040","VOORBEREIDENDE WERKZAAMHEDEN","Straatmeubilair","Opnemen en terugplaatsen lichtmasten.","st",746.72,331.36,0,0,331.36,0,0,0,0,0,0,0,0,0,0],["210010","VERHARDINGEN","Elementenverhardingen","Opnemen en aanbrengen betonstraatstenen.","m2",28.37,30.28,0,0,30.28,0,0,0,0,0,0,0,0,0,0],["210020","VERHARDINGEN","Elementenverhardingen","Opnemen en aanbrengen straatbakstenen.","m2",34.92,21.05,0,0,21.05,0,0,0,0,0,0,0,0,0,0],["210040","VERHARDINGEN","Elementenverhardingen","Opnemen en aanbrengen betontegels.","m2",23.47,21.77,0,0,21.77,0,0,0,0,0,0,0,0,0,0],["210060","VERHARDINGEN","Elementenverhardingen","Opnemen en aanbrengen bedrijfsvloerplaten.","m2",17.91,7.1,0,0,7.1,0,0,0,0,0,0,0,0,0,0],["210210","VERHARDINGEN","Elementenverhardingen","Opnemen en aanbrengen opsluitbanden zand.","m",23.93,26.05,0,0,26.05,0,0,0,0,0,0,0,0,0,0],["210230","VERHARDINGEN","Elementenverhardingen","Opnemen en aanbrengen trottoirbanden zand.","m",31.12,38.01,0,0,38.01,0,0,0,0,0,0,0,0,0,0],["213010","VERHARDINGEN","Asfaltverhardingen","Zagen asfaltverharding.","m",7.68,6.35,0,0,6.35,0,0,0,0,0,0,0,0,0,0],["213020","VERHARDINGEN","Asfaltverhardingen","Opbreken asfaltverharding.","m2",138.66,14.74,0,0,14.74,0,0,0,0,0,0,0,0,0,0],["220070","GRONDWERK SLEUVEN","Openbare grond","Realiseren sleuf MS (hinder).","m",13.97,5.98,0,0,0,0,0,5.98,0,0,0,0,0,0,0],["220080","GRONDWERK SLEUVEN","Openbare grond","Realiseren sleuf MS (veel hinder).","m",29,10.06,0,0,0,0,0,10.06,0,0,0,0,0,0,0],["222040","GRONDWERK SLEUVEN","Bredere sleuf","Extra breedte sleuf 50cm hinder.","m",5.44,11.34,0,0,0,0,0,11.34,0,0,0,0,0,0,0],["222060","GRONDWERK SLEUVEN","Bredere sleuf","Extra breedte sleuf 50cm veel hinder.","m",12.52,17.18,0,0,0,0,0,17.18,0,0,0,0,0,0,0],["223010","GRONDWERK SLEUVEN","Overdiepte","Grotere gronddekking (per 10cm).","m",0.33,1.57,0,0,0,0,0,1.57,0,0,0,0,0,0,0],["226040","GRONDWERK SLEUVEN","Grondverbetering","Aanvulzand zandbed in sleuf.","m3",34.17,27.78,0,0,0,0,0,27.78,0,0,0,0,0,0,0],["227010","GRONDWERK SLEUVEN","Grond vervoeren","Grond vervoeren naar/van gronddepot.","m3",14.95,6.98,0,0,0,0,0,6.98,0,0,0,0,0,0,0],["261010","GROENVOORZIENINGEN","Grondwerk","Frezen bermen/werkstrook.","are",4.94,0,0,0,0,0,0,0,0,0,0,0,0,0,0],["262010","GROENVOORZIENINGEN","Inzaaien","Inzaaien van bermen/werkstroken.","are",9.44,0,0,0,0,0,0,0,0,0,0,0,0,0,0],["301010","KRUISINGEN","Mantelbuizen","Aanbrengen mantelbuizen in sleuf.","m",4.43,8.47,0,0,0,0,0,0,8.47,0,0,0,0,0,0],["302010","KRUISINGEN","Persingen","Aanbrengen perslocatie.","st",2126.33,175,0,0,0,0,0,0,0,0,0,0,175,0,0],["302040","KRUISINGEN","Persingen","Boring stalen buis DN250.","m",165.78,120,0,0,0,0,0,0,0,0,0,0,120,0,0],["304010","KRUISINGEN","HDD-boringen","Boorlocatie hdd L<101m.","st",1743.65,343.47,0,0,0,20.38,0,150.36,0,0,0,0,172.73,0,0],["304020","KRUISINGEN","HDD-boringen","Boorlocatie hdd L=101-250m.","st",28530.16,647.68,0,0,0,38.42,0,283.54,0,0,0,0,325.71,0,0],["304120","KRUISINGEN","HDD-boringen","Hdd-boring 1x HDPE 160mm.","m",96.23,21.35,0,0,0,0,0,0,0,0,0,0,21.35,0,0],["304150","KRUISINGEN","HDD-boringen","Hdd-boring 3x HDPE 110mm.","m",99.74,148.95,0,0,0,0,0,0,0,0,0,0,148.95,0,0],["304170","KRUISINGEN","HDD-boringen","Hdd-boring 6x HDPE 110mm.","m",112.04,191.71,0,0,0,0,0,0,0,0,0,0,191.71,0,0],["304190","KRUISINGEN","HDD-boringen","Hdd-boring 9x HDPE 110mm.","m",129.67,38.12,0,0,0,0,0,0,0,0,0,0,38.12,0,0],["304320","KRUISINGEN","HDD-boringen","Meerprijs HDPE 160mm.","m",24.2,24.2,0,0,0,0,0,0,0,0,0,0,24.2,0,0],["309040","KRUISINGEN","Mantelbuizen PVC/PE","Leveren HDPE 110mm SDR11.","m",6.68,7.4,0,0,0,0,0,0,0,0,0,0,0,0,7.4],["309060","KRUISINGEN","Mantelbuizen PVC/PE","Leveren HDPE 160mm SDR11.","m",14.19,3.83,0,0,0,0,0,0,0,0,0,0,0,0,3.83],["340010","KRUISINGEN","Bemaling","Plaatsen en verwijderen peilbuis.","st",229.41,303.75,0,0,0,0,0,0,0,0,0,303.75,0,0,0],["345010","KRUISINGEN","Open bemaling","Aanbrengen en verwijderen open bemaling.","m",4.62,1.27,0,0,0,0,0,0,0,0,0,1.27,0,0,0],["345020","KRUISINGEN","Open bemaling","Instandhouden open bemaling.","dag",118.08,302.16,0,0,0,0,0,0,0,265.39,0,36.77,0,0,0],["350010","KRUISINGEN","Bronbemaling","Bronbemaling puntlocatie.","keer",1233.07,516.31,0,0,0,0,0,0,0,503.31,0,13,0,0,0],["350030","KRUISINGEN","Bronbemaling","Bronbemaling langs sleuf.","m",19.44,4.62,0,0,0,0,0,0,0,4.54,0,0.08,0,0,0],["350040","KRUISINGEN","Bronbemaling","Instandhouden bronbemalingspompen.","st*wk",539.1,274.58,0,0,0,0,0,0,0,0,0,274.58,0,0,0],["400020","KABEL/LEIDINGWERK ELEKTRA","Leggen kabels","Aanbrengen MS kabel 1x3x240mm2.","m",3.94,1.14,0,0,0,0,0,0,1.14,0,0,0,0,0,0],["400040","KABEL/LEIDINGWERK ELEKTRA","Leggen kabels","Aanbrengen MS kabel 3x1x240mm2 gebundeld.","m",11.67,1.06,0,0,0,0,0,0,1.06,0,0,0,0,0,0],["400050","KABEL/LEIDINGWERK ELEKTRA","Leggen kabels","Aanbrengen MS kabel 3x1x630mm2 gebundeld.","m",11.89,10.51,0,0,0,0,0,0,10.51,0,0,0,0,0,0],["400070","KABEL/LEIDINGWERK ELEKTRA","Leggen kabels","Aanbrengen kabel in mantelbuis.","m",4.86,2.08,0,0,0,0,0,0,2.08,0,0,0,0,0,0],["410040","MONTAGEWERKZAAMHEDEN ELEKTRA","Moffen","Monteren MS verbindingsmof 3-fase 1x3x240mm2.","st",779.99,42.67,0,0,0,0,0,0,42.67,0,0,0,0,0,0],["410050","MONTAGEWERKZAAMHEDEN ELEKTRA","Moffen","Monteren MS overgangmof.","st",820.14,14.99,0,0,0,0,0,0,14.99,0,0,0,0,0,0],["410060","MONTAGEWERKZAAMHEDEN ELEKTRA","Moffen","Monteren MS verbindingsmof 1-fase 3x1x630mm2.","set",1043.83,51.2,0,0,0,0,0,0,51.2,0,0,0,0,0,0],["416010","MONTAGEWERKZAAMHEDEN ELEKTRA","Verwijderen","Kabel verwijderen.","m",4.83,0.72,0,0,0,0,0,0,0.72,0,0,0,0,0,0],["418020","MONTAGEWERKZAAMHEDEN ELEKTRA","Invoeren","Invoeren 1-fase MS kabels in stations.","set",527.54,68.57,0,0,0,0,0,0,68.57,0,0,0,0,0,0],["418030","MONTAGEWERKZAAMHEDEN ELEKTRA","Invoeren","Gat boren fundering voor kabels.","st",68.36,13.39,0,0,0,0,0,0,13.39,0,0,0,0,0,0],["420010","AARDING CS - OS (op het veld)","Aarding","Aanbrengen aardelektrode 25mm2 Cu vertind.","m",123.63,1.93,0,0,0,0,0,0,1.93,0,0,0,0,0,0],["420020","AARDING CS - OS (op het veld)","Aarding","Aanbrengen aansluitdraad Cu 25mm2.","m",41.75,3.96,0,0,0,0,0,0,3.96,0,0,0,0,0,0],["441010","KABEL/LEIDINGWERK INFORMATIE DISTRIBUTIE MS","Glasvezel","Aanbrengen HDPE glasvezelbuis 40mm in sleuf.","m",1.74,0,0,0,0,0,0,0,0,0,0,0,0,0,0],["500020","STATIONS","Civiel","Aanbrengen betontegels rondom stations.","m2",31.48,8.45,0,0,8.45,0,0,0,0,0,0,0,0,0,0],["500030","STATIONS","Civiel","Aanbrengen grasbetontegels nabij stations.","m2",25.39,5.82,0,0,5.82,0,0,0,0,0,0,0,0,0,0],["500040","STATIONS","Civiel","Coordinatie plaatsen station.","st",279.88,0,0,0,0,0,0,0,0,0,0,0,0,0,0],["800010","BIJKOMENDE WERKZAAMHEDEN","Bijkomende werkzaamheden","Digitaal uitzetten trace (GPS).","m",0.81,0.37,0.37,0,0,0,0,0,0,0,0,0,0,0,0],["810010","DIRECTIEBEHOEFTEN","Directiebehoeften","Gebruik directieverblijf.","week",1318.28,1737.68,0,0,1737.68,0,0,0,0,0,0,0,0,0,0],["820010","OPLEVERING EN REVISIE","Oplevering","Uitvoeren eindmeting/fingerprintmeting.","circuit",4588.22,1733.33,0,0,0,0,0,0,1733.33,0,0,0,0,0,0],["820040","OPLEVERING EN REVISIE","Oplevering","Opstellen en aanleveren opleverdossier.","st",1119.53,0,0,0,0,0,0,0,0,0,0,0,0,0,0],["871010","TER BESCHIKKING STELLEN","T.b.s. Materieel","T.b.s. mobiele hydraulische graafmachine.","uur",94.63,0,0,0,0,0,0,0,0,0,0,0,0,0,0],["871060","TER BESCHIKKING STELLEN","T.b.s. Materieel","T.b.s. vrachtauto.","uur",78.56,0,0,0,0,0,0,0,0,0,0,0,0,0,0]];

const TARIEF_MAP = {};
TARIEF_RAW.forEach(r => {
  TARIEF_MAP[r[0]] = r;
  if (r[0] === "202101") TARIEF_MAP["202120"] = r;
});

const CAT_LABELS = ["Stafkosten","Projectkosten","Werkterrein","Bouwwegen","Sleufloos","Grondverzet","Kabelwerk","Bemaling","Civiel (derden)","Bemaling (derden)","Sleufloos (derden)","Kabelwerk (derden)","Leveranties"];
const CHAP_COLORS = {"VOORBEREIDENDE WERKZAAMHEDEN":{acc:"#66bb6a"},"VERHARDINGEN":{acc:"#64b5f6"},"GRONDWERK SLEUVEN":{acc:"#ffa726"},"GROENVOORZIENINGEN":{acc:"#66bb6a"},"KRUISINGEN":{acc:"#ba68c8"},"KABEL/LEIDINGWERK ELEKTRA":{acc:"#ffd54f"},"MONTAGEWERKZAAMHEDEN ELEKTRA":{acc:"#ffd54f"},"AARDING CS - OS (op het veld)":{acc:"#ffd54f"},"KABEL/LEIDINGWERK INFORMATIE DISTRIBUTIE MS":{acc:"#ffd54f"},"STATIONS":{acc:"#f48fb1"},"BIJKOMENDE WERKZAAMHEDEN":{acc:"#90a4ae"},"DIRECTIEBEHOEFTEN":{acc:"#90a4ae"},"OPLEVERING EN REVISIE":{acc:"#90a4ae"},"TER BESCHIKKING STELLEN":{acc:"#90a4ae"}};

const euro = v => new Intl.NumberFormat("nl-NL",{style:"currency",currency:"EUR",maximumFractionDigits:0}).format(v||0);
const fmtP = v => (!isFinite(v)||v==null)?"—":(v*100).toFixed(1)+"%";
const num = v => new Intl.NumberFormat("nl-NL",{maximumFractionDigits:2}).format(v||0);

// Normalize any cell value to a clean postnr string or ""
function toPostnr(val) {
  if (val == null) return "";
  const n = typeof val === "number" ? Math.round(val) : parseInt(String(val).replace(/[^0-9]/g, ""), 10);
  // Accept 6-digit HELIX postnrs AND 7-digit project-specific postnrs (8800xxx etc.)
  if (!isNaN(n) && n >= 100000 && n <= 9999999) return String(n);
  return "";
}

function parseInschrijfstaat(buf) {
  const wb=XLSX.read(new Uint8Array(buf),{type:"array",raw:true});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const data=XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
  const posts=[],sks=[];
  data.forEach(row=>{
    const v = toPostnr(row[1]);
    if (!v) return;
    const isStaart = v.startsWith("9") && v.length === 6;
    if (!isStaart) {
      // col9 = pre-calculated totaal (includes indexation, correct per post)
      // col4 = hoeveelheid, col7 = prijs per eenheid
      const col9 = parseFloat(row[9]) || 0;
      const h = parseFloat(row[4]) || 0;
      const p = parseFloat(row[7]) || 0;
      // Use col9 when available — it reflects the correct (possibly indexed) total.
      // Derive inschrijfprijs: if col9>0 use col9/h, else use col7 directly.
      // Only include the post when it has an actual amount.
      const totaal = col9 !== 0 ? col9 : (h > 0 ? h * p : 0);
      const inschrijfprijs = totaal > 0 && h > 0 ? totaal / h : (p || 0);
      if (totaal > 0) posts.push({postnr:v, hoeveelheid:h||1, inschrijfprijs});
    } else {
      sks.push({postnr:v, omschrijving:row[2]?String(row[2]).split(".")[0].trim():v, pct:parseFloat(row[7])||0, interneKosten:0});
    }
  });
  const enriched=posts.map(p=>{
    const tar=TARIEF_MAP[p.postnr];
    return{postnr:p.postnr,omschrijving:tar?tar[3]:"Post "+p.postnr,eenheid:tar?tar[4]:"",
      hoofdstuk:tar?tar[1]:"OVERIG",hoeveelheid:p.hoeveelheid,
      inschrijfprijs:p.inschrijfprijs||(tar?tar[5]:0),kostprijs:tar?tar[6]:0,
      cats:tar?tar.slice(7):Array(13).fill(0),inLib:!!tar};
  });
  return{posts:enriched,staartkosten:sks.length>0?sks:[
    {postnr:"920010",omschrijving:"Uitvoeringskosten",pct:19.61,interneKosten:0},
    {postnr:"930010",omschrijving:"Algemene kosten",pct:11,interneKosten:0},
    {postnr:"940010",omschrijving:"Winst",pct:2,interneKosten:0},
    {postnr:"940020",omschrijving:"Risico",pct:2,interneKosten:0},
  ]};
}

function Badge({n,color}){
  if(!n)return null;
  return <span style={{background:color,color:"#fff",fontSize:9,fontWeight:700,padding:"1px 6px",borderRadius:8,marginLeft:6}}>{n}</span>;
}

export default function App() {
  const [page,setPage]=useState("begroting");
  const [posts,setPosts]=useState([]);
  const [staartkosten,setStaartkosten]=useState([]);
  const [projectName,setProjectName]=useState("");
  const [tasks,setTasks]=useState([]);
  const [koppelingen,setKoppelingen]=useState({});
  const [expandedChaps,setExpandedChaps]=useState({});
  const fileRef=useRef();

  const handleInschrijfstaat=useCallback(async(file)=>{
    const buf=await file.arrayBuffer();
    const{posts:p,staartkosten:sk}=parseInschrijfstaat(buf);
    setPosts(p);setStaartkosten(sk);
    setProjectName(file.name.replace(/\.xlsx$/i,"").replace(/_/g," ").substring(0,60));
    const chaps={};p.forEach(x=>{chaps[x.hoofdstuk]=true;});setExpandedChaps(chaps);
    setPage("begroting");
  },[]);

  const totals=useMemo(()=>{
    let si=0,sk=0;
    posts.forEach(p=>{si+=p.hoeveelheid*p.inschrijfprijs;sk+=p.hoeveelheid*p.kostprijs;});
    const skI=staartkosten.reduce((s,x)=>s+x.pct/100*si,0);
    const skK=staartkosten.reduce((s,x)=>s+(x.interneKosten||0),0);
    const ti=si+skI,tk=sk+skK;
    return{si,sk,ti,tk,m:ti-tk,mp:ti>0?(ti-tk)/ti:0};
  },[posts,staartkosten]);

  const grouped=useMemo(()=>{
    const m={};posts.forEach(p=>{if(!m[p.hoofdstuk])m[p.hoofdstuk]=[];m[p.hoofdstuk].push(p);});return m;
  },[posts]);

  const catTotals=useMemo(()=>{const t=Array(13).fill(0);posts.forEach(p=>p.cats.forEach((c,i)=>{t[i]+=p.hoeveelheid*c;}));return t;},[posts]);
  const missing=posts.filter(p=>!p.inLib);
  const koppeldeTasks=tasks.filter(t=>koppelingen[t.id]?.length>0);
  const completeTasks=koppeldeTasks.filter(t=>(koppelingen[t.id]||[]).reduce((s,k)=>s+k.percentage,0)===100);
  const [activeTab,setActiveTab]=useState("begroting");

  const exportXlsx=()=>{
    const wb=XLSX.utils.book_new();
    const rows=[["WERKBEGROTING"],[""],["Postnr","Omschrijving","Eenh.","Hoeveelheid","Inschrijfprijs","Inschrijfsom","Kostprijs","Int.kosten","Marge","Marge%",...CAT_LABELS]];
    let lh="";
    posts.forEach(p=>{
      if(p.hoofdstuk!==lh){rows.push(["",p.hoofdstuk]);lh=p.hoofdstuk;}
      const ins=p.hoeveelheid*p.inschrijfprijs,kst=p.hoeveelheid*p.kostprijs;
      rows.push([p.postnr,p.omschrijving,p.eenheid,p.hoeveelheid,p.inschrijfprijs,ins,p.kostprijs,kst,ins-kst,ins>0?(ins-kst)/ins:0,...p.cats.map(c=>p.hoeveelheid*c)]);
    });
    rows.push([],["SUBTOTAAL UITVOERING","","","","",totals.si,"",totals.sk,totals.si-totals.sk]);
    rows.push([],["STAARTKOSTEN"]);
    staartkosten.forEach(sk=>{const b=sk.pct/100*totals.si;rows.push([sk.postnr,sk.omschrijving,"%",sk.pct,"",b,"",sk.interneKosten||0,b-(sk.interneKosten||0)]);});
    rows.push(["TOTAAL","","","","",totals.ti,"",totals.tk,totals.m]);
    const ws=XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"]=[8,12,50,8,6,14,14,14,14,10,...Array(13).fill(12)].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb,ws,"Werkbegroting");
    XLSX.writeFile(wb,"Werkbegroting_"+projectName.replace(/\s+/g,"_")+".xlsx");
  };

  return (
    <div style={{minHeight:"100vh",background:"#0f1923",fontFamily:"'IBM Plex Mono','Courier New',monospace",color:"#e8eaed"}}>
      <header style={{background:"linear-gradient(135deg,#0f1923,#1a2d40,#0f1923)",borderBottom:"2px solid #1e4976",padding:"12px 22px",display:"flex",alignItems:"center",gap:14,position:"sticky",top:0,zIndex:100}}>
        <div style={{width:32,height:32,background:"linear-gradient(135deg,#1976d2,#0d47a1)",borderRadius:7,display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,fontWeight:900,color:"#fff",boxShadow:"0 0 12px rgba(25,118,210,.5)",flexShrink:0}}>W</div>
        <div>
          <div style={{fontSize:11,fontWeight:700,letterSpacing:3,color:"#90caf9"}}>A.HAK ELECTRON</div>
          <div style={{fontSize:9,color:"#546e7a",letterSpacing:2}}>WERKBEGROTING + PLANNING</div>
        </div>
        <nav style={{marginLeft:24,display:"flex",gap:2}}>
          {[["begroting","BEGROTING",posts.length,"#1976d2"],["planning","PLANNING",tasks.length,"#00897b"],["dashboard","DASHBOARD",completeTasks.length||null,"#7b1fa2"]].map(([id,lbl,badge,bc])=>(
            <button key={id} onClick={()=>setPage(id)} style={{background:page===id?"rgba(25,118,210,.15)":"transparent",border:"none",borderBottom:`2px solid ${page===id?"#1976d2":"transparent"}`,color:page===id?"#90caf9":"#37474f",padding:"6px 14px",cursor:"pointer",fontSize:10,fontWeight:700,letterSpacing:1.5,fontFamily:"inherit",transition:"all .15s",display:"flex",alignItems:"center"}}>
              {lbl}<Badge n={badge} color={bc}/>
            </button>
          ))}
        </nav>
        <div style={{marginLeft:"auto",display:"flex",gap:8,alignItems:"center"}}>
          {posts.length>0&&<div style={{fontSize:10,color:"#37474f",textAlign:"right"}}><div style={{color:"#546e7a"}}>{projectName.substring(0,30)}</div><div style={{color:"#1976d2",fontWeight:700}}>{euro(totals.ti)}</div></div>}
          <button onClick={()=>fileRef.current.click()} style={{background:"linear-gradient(135deg,#1976d2,#0d47a1)",border:"none",color:"#fff",padding:"7px 14px",borderRadius:6,cursor:"pointer",fontSize:10,fontWeight:700,letterSpacing:1,fontFamily:"inherit"}}>+ INSCHRIJFSTAAT</button>
        </div>
        <input ref={fileRef} type="file" accept=".xlsx" style={{display:"none"}} onChange={e=>{ if(e.target.files[0]) handleInschrijfstaat(e.target.files[0]); }}/>
      </header>

      {page==="begroting"&&(posts.length===0?(
        <div style={{maxWidth:520,margin:"80px auto",padding:"0 20px",textAlign:"center"}}>
          <div style={{fontSize:40,marginBottom:14}}>📋</div>
          <h1 style={{fontSize:20,fontWeight:900,background:"linear-gradient(90deg,#90caf9,#42a5f5)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",marginBottom:8}}>WERKBEGROTING GENERATOR</h1>
          <p style={{color:"#546e7a",fontSize:12,marginBottom:28}}>Laad een inschrijfstaat · Koppel de planning · Zie de cashflow</p>
          <div
            onClick={()=>{ fileRef.current.value=""; fileRef.current.click(); }}
            onDragOver={e=>{ e.preventDefault(); e.stopPropagation(); e.currentTarget.style.borderColor="#42a5f5"; e.currentTarget.style.background="rgba(25,118,210,.1)"; }}
            onDragLeave={e=>{ e.currentTarget.style.borderColor="#1e4976"; e.currentTarget.style.background="rgba(25,118,210,.03)"; }}
            onDrop={e=>{ e.preventDefault(); e.stopPropagation(); e.currentTarget.style.borderColor="#1e4976"; e.currentTarget.style.background="rgba(25,118,210,.03)"; const f=e.dataTransfer.files[0]; if(f)handleInschrijfstaat(f); }}
            style={{border:"2px dashed #1e4976",borderRadius:14,padding:"44px 28px",cursor:"pointer",background:"rgba(25,118,210,.03)",transition:"all .2s"}}>
            <div style={{fontSize:13,color:"#90caf9",fontWeight:700,marginBottom:6}}>Sleep inschrijfstaat hier of klik</div>
            <div style={{fontSize:11,color:"#37474f"}}>.xlsx uitvoer van de bestekposter</div>
          </div>
        </div>
      ):(
        <div style={{maxWidth:1360,margin:"0 auto",padding:"18px 18px 60px"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
            <input value={projectName} onChange={e=>setProjectName(e.target.value)} style={{background:"transparent",border:"none",borderBottom:"1px solid #1e4976",color:"#90caf9",fontSize:16,fontWeight:700,letterSpacing:2,padding:"3px 0",width:"60%",outline:"none",textTransform:"uppercase",fontFamily:"inherit"}}/>
            <button onClick={exportXlsx} style={{background:"linear-gradient(135deg,#1976d2,#0d47a1)",border:"none",color:"#fff",padding:"7px 16px",borderRadius:6,cursor:"pointer",fontSize:10,fontWeight:700,letterSpacing:1,fontFamily:"inherit"}}>EXPORTEER XLSX</button>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:18}}>
            {[{lbl:"INSCHRIJFSOM",val:euro(totals.ti),sub:"Excl. staartkost: "+euro(totals.si),c:"#1976d2"},{lbl:"INTERNE KOSTEN",val:euro(totals.tk),sub:"Excl. staartkost: "+euro(totals.sk),c:"#7b1fa2"},{lbl:"BRUTO MARGE",val:euro(totals.m),sub:fmtP(totals.mp)+" van inschrijfsom",c:totals.m>=0?"#2e7d32":"#c62828"},{lbl:"POSTEN",val:posts.length,sub:missing.length>0?"? "+missing.length+" zonder kostprijs":"? Alle gekoppeld",c:missing.length>0?"#e65100":"#2e7d32"}].map((k,i)=>(
              <div key={i} style={{background:"linear-gradient(135deg,#0d1f2d,#132436)",border:`1px solid ${k.c}30`,borderLeft:`3px solid ${k.c}`,borderRadius:9,padding:"12px 14px"}}>
                <div style={{fontSize:8,color:"#546e7a",letterSpacing:2,marginBottom:4}}>{k.lbl}</div>
                <div style={{fontSize:17,fontWeight:900,color:k.c}}>{k.val}</div>
                <div style={{fontSize:9,color:"#37474f",marginTop:3}}>{k.sub}</div>
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:3,borderBottom:"1px solid #1e4976",marginBottom:14}}>
            {[["begroting","BEGROTING"],["staartkosten","STAARTKOSTEN"],["kostensoorten","KOSTENSOORTEN"]].map(([id,lbl])=>(
              <button key={id} onClick={()=>setActiveTab(id)} style={{background:activeTab===id?"rgba(25,118,210,.15)":"transparent",border:"none",borderBottom:`2px solid ${activeTab===id?"#1976d2":"transparent"}`,color:activeTab===id?"#90caf9":"#37474f",padding:"6px 14px",cursor:"pointer",fontSize:10,fontWeight:700,letterSpacing:2,fontFamily:"inherit"}}>{lbl}</button>
            ))}
          </div>
          {activeTab==="begroting"&&(
            <div>
              {Object.entries(grouped).map(([hfdst,hposts])=>{
                const col=CHAP_COLORS[hfdst]||{acc:"#546e7a"};
                const ci=hposts.reduce((s,p)=>s+p.hoeveelheid*p.inschrijfprijs,0);
                const ck=hposts.reduce((s,p)=>s+p.hoeveelheid*p.kostprijs,0);
                const open=expandedChaps[hfdst]!==false;
                return(
                  <div key={hfdst} style={{marginBottom:6,borderRadius:9,overflow:"hidden",border:"1px solid #1e4976"}}>
                    <div onClick={()=>setExpandedChaps(p=>({...p,[hfdst]:!open}))} style={{background:`${col.acc}12`,borderLeft:`4px solid ${col.acc}`,padding:"8px 14px",cursor:"pointer",display:"flex",alignItems:"center",gap:10}}>
                      <span style={{color:col.acc,fontSize:11}}>{open?"?":"?"}</span>
                      <span style={{fontWeight:700,fontSize:11,letterSpacing:1,color:col.acc,flex:1}}>{hfdst}</span>
                      <span style={{fontSize:10,color:"#90caf9"}}>{euro(ci)}</span>
                      <span style={{fontSize:10,color:"#78909c",marginLeft:8}}>intern: {euro(ck)}</span>
                      <span style={{fontSize:10,fontWeight:700,marginLeft:8,color:ci-ck>=0?"#4caf50":"#ef5350"}}>{euro(ci-ck)}</span>
                    </div>
                    {open&&(
                      <div style={{overflowX:"auto"}}>
                        <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                          <thead><tr style={{background:"#0d1520"}}>{["Postnr","Omschrijving","Eenh.","Hoeveelheid","Inschr.prijs","Inschrijfsom","Kostprijs","Int. kosten","Marge"].map(h=><th key={h} style={{padding:"5px 9px",color:"#37474f",fontWeight:700,textAlign:h.match(/prijs|som|kosten|Marge/)?"right":"left",borderBottom:"1px solid #1e4976",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                          <tbody>
                            {hposts.map((p,i)=>{
                              const ins=p.hoeveelheid*p.inschrijfprijs,kst=p.hoeveelheid*p.kostprijs,mar=ins-kst;
                              return(<tr key={p.postnr} style={{background:i%2?"rgba(255,255,255,.02)":"transparent",borderBottom:"1px solid #0a1218"}}>
                                <td style={{padding:"4px 9px",color:"#546e7a",fontFamily:"monospace",fontSize:10}}>{p.postnr}</td>
                                <td style={{padding:"4px 9px",color:p.inLib?"#cfd8dc":"#ff8a65",maxWidth:280}}>{p.omschrijving}</td>
                                <td style={{padding:"4px 9px",color:"#546e7a",textAlign:"center"}}>{p.eenheid}</td>
                                <td style={{padding:"4px 9px",textAlign:"right",color:"#90a4ae"}}>{num(p.hoeveelheid)}</td>
                                <td style={{padding:"4px 9px",textAlign:"right",color:"#546e7a"}}>{euro(p.inschrijfprijs)}</td>
                                <td style={{padding:"4px 9px",textAlign:"right",color:"#90caf9",fontWeight:600}}>{euro(ins)}</td>
                                <td style={{padding:"4px 9px",textAlign:"right",color:p.kostprijs>0?"#78909c":"#1e4976"}}>{p.kostprijs>0?euro(p.kostprijs):"?"}</td>
                                <td style={{padding:"4px 9px",textAlign:"right",color:p.kostprijs>0?"#ce93d8":"#1e4976"}}>{p.kostprijs>0?euro(kst):"?"}</td>
                                <td style={{padding:"4px 9px",textAlign:"right",fontWeight:700,color:mar>=0?"#4caf50":"#ef5350"}}>{p.kostprijs>0?euro(mar):"?"}</td>
                              </tr>);
                            })}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                );
              })}
              <div style={{marginTop:8,background:"rgba(25,118,210,.1)",border:"1px solid #1976d2",borderRadius:7,padding:"10px 16px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <span style={{fontWeight:700,fontSize:11,letterSpacing:2,color:"#90caf9"}}>SUBTOTAAL UITVOERING</span>
                <div style={{display:"flex",gap:32}}>
                  {[["INSCHRIJFSOM",euro(totals.si),"#90caf9"],["INTERNE KOSTEN",euro(totals.sk),"#ce93d8"],["MARGE",euro(totals.si-totals.sk),totals.si-totals.sk>=0?"#4caf50":"#ef5350"]].map(([l,v,c])=>(
                    <div key={l} style={{textAlign:"right"}}><div style={{fontSize:8,color:"#546e7a"}}>{l}</div><div style={{fontSize:14,color:c,fontWeight:700}}>{v}</div></div>
                  ))}
                </div>
              </div>
            </div>
          )}
          {activeTab==="staartkosten"&&(
            <div style={{maxWidth:820}}>
              <div style={{marginBottom:10,fontSize:11,color:"#546e7a"}}>Percentage over subtotaal uitvoering ({euro(totals.si)}). Kolom Interne kosten = eigen personeel (projectleider, uitvoerder, enz.)</div>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                <thead><tr style={{background:"#0d1520",borderBottom:"2px solid #1e4976"}}>{["Post","Omschrijving","% (inschr.)","Inschrijfsom","Interne kosten","Marge"].map(h=><th key={h} style={{padding:"7px 12px",color:"#37474f",fontWeight:700,textAlign:h.match(/Post|Omschr/)?"left":"right"}}>{h}</th>)}</tr></thead>
                <tbody>
                  {staartkosten.map((sk,i)=>{
                    const b=sk.pct/100*totals.si,mar=b-(sk.interneKosten||0);
                    return(<tr key={sk.postnr} style={{borderBottom:"1px solid #0d1520",background:i%2?"rgba(255,255,255,.02)":"transparent"}}>
                      <td style={{padding:"7px 12px",color:"#546e7a",fontFamily:"monospace"}}>{sk.postnr}</td>
                      <td style={{padding:"7px 12px",color:"#cfd8dc"}}>{sk.omschrijving}</td>
                      <td style={{padding:"7px 12px",textAlign:"right"}}><input type="number" value={sk.pct} step={0.01} onChange={e=>setStaartkosten(prev=>prev.map((s,j)=>j===i?{...s,pct:parseFloat(e.target.value)||0}:s))} style={{background:"rgba(25,118,210,.1)",border:"1px solid #1e4976",color:"#90caf9",padding:"3px 7px",borderRadius:4,fontSize:11,width:68,textAlign:"right",outline:"none",fontFamily:"inherit"}}/><span style={{color:"#37474f",marginLeft:4}}>%</span></td>
                      <td style={{padding:"7px 12px",textAlign:"right",color:"#90caf9"}}>{euro(b)}</td>
                      <td style={{padding:"7px 12px",textAlign:"right"}}><input type="number" value={sk.interneKosten||""} placeholder="0" onChange={e=>setStaartkosten(prev=>prev.map((s,j)=>j===i?{...s,interneKosten:parseFloat(e.target.value)||0}:s))} style={{background:"rgba(106,27,154,.1)",border:"1px solid #6a1b9a",color:"#ce93d8",padding:"3px 7px",borderRadius:4,fontSize:11,width:110,textAlign:"right",outline:"none",fontFamily:"inherit"}}/></td>
                      <td style={{padding:"7px 12px",textAlign:"right",fontWeight:700,color:mar>=0?"#4caf50":"#ef5350"}}>{euro(mar)}</td>
                    </tr>);
                  })}
                </tbody>
              </table>
              <div style={{marginTop:14,background:"rgba(25,118,210,.1)",border:"1px solid #1976d2",borderRadius:7,padding:"12px 16px",display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:16}}>
                {[["TOTAAL INSCHRIJFSOM",euro(totals.ti),"#90caf9"],["TOTAAL INTERNE KOSTEN",euro(totals.tk),"#ce93d8"],["BRUTO MARGE",euro(totals.m)+" ? "+fmtP(totals.mp),totals.m>=0?"#4caf50":"#ef5350"]].map(([l,v,c])=>(
                  <div key={l}><div style={{fontSize:8,color:"#546e7a",letterSpacing:2}}>{l}</div><div style={{fontSize:15,fontWeight:900,color:c,marginTop:3}}>{v}</div></div>
                ))}
              </div>
            </div>
          )}
          {activeTab==="kostensoorten"&&(
            <div style={{maxWidth:640}}>
              {(()=>{const total=catTotals.reduce((s,v)=>s+v,0);return CAT_LABELS.map((lbl,i)=>{const v=catTotals[i],bp=total>0?v/total:0;return(<div key={lbl} style={{marginBottom:9}}><div style={{display:"flex",justifyContent:"space-between",fontSize:10,marginBottom:3}}><span style={{color:"#90a4ae"}}>{lbl}</span><span style={{color:v>0?"#ce93d8":"#37474f"}}>{v>0?euro(v)+" ? "+fmtP(bp):"?"}</span></div><div style={{height:5,background:"#0d1520",borderRadius:3,overflow:"hidden"}}><div style={{height:"100%",width:bp*100+"%",background:"hsl("+(200+i*15)+",70%,55%)",borderRadius:3}}/></div></div>);});})()}
              <div style={{marginTop:18,padding:"10px 14px",background:"rgba(255,255,255,.03)",borderRadius:7,border:"1px solid #1e4976",display:"flex",justifyContent:"space-between"}}>
                <span style={{fontWeight:700,fontSize:11,color:"#90caf9"}}>TOTAAL INTERNE KOSTEN</span>
                <span style={{fontSize:13,fontWeight:900,color:"#ce93d8"}}>{euro(catTotals.reduce((s,v)=>s+v,0))}</span>
              </div>
            </div>
          )}
        </div>
      ))}
      {page==="planning"&&<PagePlanning tasks={tasks} setTasks={setTasks} koppelingen={koppelingen} setKoppelingen={setKoppelingen} posts={posts}/>}
      {page==="dashboard"&&<PageDashboard posts={posts} tasks={tasks} koppelingen={koppelingen} staartkosten={staartkosten} projectName={projectName}/>}
    </div>
  );
}
