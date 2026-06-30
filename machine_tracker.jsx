import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

const SERVICE_INTERVAL = 400;
const MACHINE_TYPES = ["Excavator","Backhoe Loader","Combat Lorry","Shovel","Vibratory Roller","Back Pusher","Forklift","Motor Grader","Asphalt","Others"];
const STATUSES = ["Work","Idle","Repair"];
const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const SS = {
  Work:   { pill:"bg-green-100 text-green-700 border-green-300", dot:"bg-green-500", btn:"bg-green-50 border-green-400 text-green-700" },
  Idle:   { pill:"bg-yellow-100 text-yellow-700 border-yellow-300", dot:"bg-yellow-400", btn:"bg-yellow-50 border-yellow-400 text-yellow-700" },
  Repair: { pill:"bg-red-100 text-red-700 border-red-300", dot:"bg-red-500", btn:"bg-red-50 border-red-400 text-red-700" }
};
const inp = "w-full border-2 border-red-200 rounded-xl px-3 py-2 mt-0.5 text-sm font-bold focus:outline-none focus:ring-2 focus:ring-red-400 bg-white text-gray-800";

function today(){ return new Date().toISOString().slice(0,10); }

function Field(props){
  return (
    <div>
      <label className="text-xs font-black text-red-600">{props.label}</label>
      {props.children}
    </div>
  );
}

function ServiceBadge(props){
  const hours = props.hours;
  const pct = (hours / SERVICE_INTERVAL) * 100;
  const s = hours >= SERVICE_INTERVAL ? "overdue" : pct >= 80 ? "warning" : "ok";
  const st = { overdue:"border-red-500 text-red-700", warning:"border-yellow-400 text-yellow-700", ok:"border-green-400 text-green-700" };
  const bg = { overdue:"#ffe0e0", warning:"#fffde4", ok:"#e4ffec" };
  const lb = { overdue:"⚠️ Overdue", warning:"🔔 Due Soon", ok:"✅ OK" };
  return (
    <span className={"text-xs font-black px-3 py-1 rounded-full border-2 " + st[s]} style={{background:bg[s]}}>
      {lb[s]}
    </span>
  );
}

function ProgressBar(props){
  const hours = props.hours;
  const pct = Math.min((hours / SERVICE_INTERVAL) * 100, 100);
  const c = hours >= SERVICE_INTERVAL ? "bg-red-600" : pct >= 80 ? "bg-yellow-400" : "bg-red-400";
  return (
    <div className="w-full bg-red-900 bg-opacity-20 rounded-full h-2.5 mt-1.5">
      <div className={c + " h-2.5 rounded-full transition-all duration-500"} style={{width: pct + "%"}}/>
    </div>
  );
}

const FILTER_TYPES = ["All Types"].concat(MACHINE_TYPES);
const FILTER_STATUS = ["All Status"].concat(STATUSES);

function htmlToDocxBlob(htmlContent, title){
  const html =
    "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>" +
    "<head><meta charset='utf-8'><title>" + title + "</title>" +
    "<style>" +
    "body { font-family: Arial, sans-serif; font-size: 10pt; }" +
    "h2 { color: #b91c1c; font-size: 14pt; margin-bottom: 4pt; }" +
    "p.sub { color: #888; font-size: 8pt; margin-bottom: 10pt; }" +
    "table { border-collapse: collapse; width: 100%; }" +
    "th { background: #dc2626; color: white; padding: 6pt 8pt; font-size: 9pt; border: 1pt solid #f87171; text-align: center; }" +
    "td { padding: 5pt 8pt; font-size: 9pt; border: 1pt solid #fecaca; text-align: center; }" +
    "tr:nth-child(even) td { background: #fef2f2; }" +
    ".left { text-align: left; }" +
    ".bold { font-weight: bold; }" +
    ".pink { color: #b91c1c; font-weight: bold; }" +
    "</style></head><body>" + htmlContent + "</body></html>";
  return new Blob([html], {type:"application/msword"});
}

const SUPABASE_URL = "https://krvwtfbnzunyacstibiq.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imtydnd0ZmJuenVueWFjc3RpYmlxIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODE4ODc3NTgsImV4cCI6MjA5NzQ2Mzc1OH0.hSQMlGcmsHdQjKd5X3OmqxI6IODDcgLHDV6_q0FnghU";

function sbFetch(path, options){
  options = options || {};
  const headers = Object.assign({
    "apikey": SUPABASE_ANON_KEY,
    "Authorization": "Bearer " + SUPABASE_ANON_KEY,
    "Content-Type": "application/json",
    "Prefer": options.prefer || "return=representation"
  }, options.headers || {});
  return fetch(SUPABASE_URL + "/rest/v1/" + path, Object.assign({}, options, {headers:headers}))
    .then(function(res){
      if(!res.ok){
        return res.text().then(function(text){
          throw new Error("Supabase error " + res.status + ": " + text);
        });
      }
      return res.text().then(function(text){
        return text ? JSON.parse(text) : null;
      });
    });
}

function loadMachinesFromDb(){
  return sbFetch("machines?select=*&order=created_at.asc").then(function(rows){
    return rows.map(function(r){
      return {
        id:r.id, name:r.name, type:r.type, model:r.model||"", dateMade:r.date_made||"",
        operator:r.operator||"", operatorStart:r.operator_start||"", condition:r.condition||"Work",
        hourLogs:r.hour_logs||[], serviceLogs:r.service_logs||[]
      };
    });
  });
}
function loadSettingsFromDb(){
  return sbFetch("app_settings?select=*&id=eq.main").then(function(rows){
    if(rows && rows.length>0) return {operators:rows[0].operators||[], sites:rows[0].sites||[]};
    return {operators:[],sites:[]};
  });
}
function insertMachineDb(m){
  return sbFetch("machines",{
    method:"POST",
    body:JSON.stringify({
      id:m.id,name:m.name,type:m.type,model:m.model,date_made:m.dateMade,
      operator:m.operator,operator_start:m.operatorStart,condition:m.condition,
      hour_logs:m.hourLogs,service_logs:m.serviceLogs
    })
  });
}
function updateMachineDb(m){
  return sbFetch("machines?id=eq." + m.id,{
    method:"PATCH",
    body:JSON.stringify({
      name:m.name,type:m.type,model:m.model,date_made:m.dateMade,
      operator:m.operator,operator_start:m.operatorStart,condition:m.condition,
      hour_logs:m.hourLogs,service_logs:m.serviceLogs
    })
  });
}
function deleteMachineDb(id){
  return sbFetch("machines?id=eq." + id,{method:"DELETE"});
}
function bulkInsertMachinesDb(list){
  return sbFetch("machines",{
    method:"POST",
    body:JSON.stringify(list.map(function(m){
      return {
        id:m.id,name:m.name,type:m.type,model:m.model,date_made:m.dateMade,
        operator:m.operator,operator_start:m.operatorStart,condition:m.condition,
        hour_logs:m.hourLogs,service_logs:m.serviceLogs
      };
    }))
  });
}
function updateSettingsDb(s){
  return sbFetch("app_settings?id=eq.main",{
    method:"PATCH",
    body:JSON.stringify({operators:s.operators,sites:s.sites})
  });
}

// ── Local backup (safety net in case Supabase sync fails — e.g. RLS blocking writes,
// the published artifact's network being restricted, or no internet) ──
const LS_MACHINES_KEY="machine_tracker_machines_backup";
const LS_SETTINGS_KEY="machine_tracker_settings_backup";
function lsSave(key,data){
  try{ if(typeof localStorage!=="undefined") localStorage.setItem(key,JSON.stringify(data)); }catch(e){}
}
function lsLoad(key){
  try{
    if(typeof localStorage==="undefined") return null;
    const raw=localStorage.getItem(key);
    return raw?JSON.parse(raw):null;
  }catch(e){ return null; }
}

export default function App(){
  useEffect(function(){
    if(typeof document!=="undefined" && !document.getElementById("no-spinner-style")){
      const style=document.createElement("style");
      style.id="no-spinner-style";
      style.innerHTML =
        "input[type=number]::-webkit-inner-spin-button," +
        "input[type=number]::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }" +
        "input[type=number] { -moz-appearance: textfield; appearance: textfield; }";
      document.head.appendChild(style);
    }
  },[]);

  const [machines,setMachines]=useState([]);
  const [settings,setSettings]=useState({operators:[],sites:[]});
  const [loaded,setLoaded]=useState(false);
  const [syncing,setSyncing]=useState(false);
  const [syncError,setSyncError]=useState("");
  const [usingLocalBackup,setUsingLocalBackup]=useState(false);
  const [view,setView]=useState("dashboard");
  const [selId,setSelId]=useState(null);
  const [form,setForm]=useState({name:"",type:"Excavator",dateMade:"",model:"",operator:"",operatorStart:today(),condition:"Work"});
  const [editLog,setEditLog]=useState(null);
  const [showSvc,setShowSvc]=useState(false);
  const [svcSite,setSvcSite]=useState("");
  const [svcDoNo,setSvcDoNo]=useState("");
  const [search,setSearch]=useState("");
  const [showSearch,setShowSearch]=useState(false);
  const [fType,setFType]=useState("All Types");
  const [fStatus,setFStatus]=useState("All Status");
  const [viewMonth,setViewMonth]=useState(new Date().getMonth());
  const [viewYear,setViewYear]=useState(new Date().getFullYear());
  const [printMonth,setPrintMonth]=useState(new Date().getMonth());
  const [printYear,setPrintYear]=useState(new Date().getFullYear());
  const [printType,setPrintType]=useState("All Types");
  const [printMachine,setPrintMachine]=useState("All Machines");
  const [printReportType,setPrintReportType]=useState("monthly");
  const [importError,setImportError]=useState("");
  const [importSuccess,setImportSuccess]=useState("");
  const [editMachine,setEditMachine]=useState(null);
  const [newOp,setNewOp]=useState("");
  const [newSite,setNewSite]=useState("");
  const [editOpIdx,setEditOpIdx]=useState(null);
  const [editOpVal,setEditOpVal]=useState("");
  const [editSiteIdx,setEditSiteIdx]=useState(null);
  const [editSiteVal,setEditSiteVal]=useState("");
  const [dayPopup,setDayPopup]=useState(null);
  const [quickHrs,setQuickHrs]=useState({});
  const hrsInputRefs=useRef({});
  const csvRef=useRef();

  useEffect(function(){
    let mounted=true;
    function loadAll(){
      loadMachinesFromDb().then(function(m){
        if(!mounted) return;
        setMachines(m);
        lsSave(LS_MACHINES_KEY,m);
        setUsingLocalBackup(false);
        setSyncError("");
      }).catch(function(e){
        console.error("Load machines failed:",e);
        if(!mounted) return;
        const backup=lsLoad(LS_MACHINES_KEY);
        if(backup){ setMachines(backup); setUsingLocalBackup(true); }
        setSyncError("Could not reach the database. " + (backup?"Showing your last locally saved data — it will sync once connection is restored.":"Your data may not be visible until connection is restored."));
      });
      loadSettingsFromDb().then(function(s){
        if(!mounted) return;
        setSettings(s);
        lsSave(LS_SETTINGS_KEY,s);
      }).catch(function(e){
        console.error("Load settings failed:",e);
        if(!mounted) return;
        const backup=lsLoad(LS_SETTINGS_KEY);
        if(backup) setSettings(backup);
      }).then(function(){ if(mounted) setLoaded(true); });
    }
    loadAll();
    return function(){ mounted=false; };
  },[]);

  useEffect(function(){
    if(!loaded) return;
    const interval=setInterval(function(){
      loadMachinesFromDb().then(function(m){ setMachines(m); lsSave(LS_MACHINES_KEY,m); setUsingLocalBackup(false); setSyncError(""); }).catch(function(e){ console.error("Poll machines failed:",e); setSyncError("Connection to database lost — changes may not be saving. Retrying..."); });
      loadSettingsFromDb().then(function(s){ setSettings(s); lsSave(LS_SETTINGS_KEY,s); }).catch(function(){});
    },5000);
    return function(){ clearInterval(interval); };
  },[loaded]);

  function upd(fn){
    setMachines(function(prev){
      const raw=fn(prev);
      const next=raw.map(function(m){
        const byDate={};
        for(let i=0;i<m.hourLogs.length;i++){
          const l=m.hourLogs[i];
          if(!byDate[l.date] || l.id>byDate[l.date].id) byDate[l.date]=l;
        }
        return Object.assign({},m,{hourLogs:Object.values(byDate)});
      });
      lsSave(LS_MACHINES_KEY,next);
      setSyncing(true);
      (function(){
        const prevIds={}; prev.forEach(function(m){ prevIds[m.id]=true; });
        const nextIds={}; next.forEach(function(m){ nextIds[m.id]=true; });
        const added=next.filter(function(m){ return !prevIds[m.id]; });
        const removed=prev.filter(function(m){ return !nextIds[m.id]; });
        const common=next.filter(function(m){ return prevIds[m.id]; });
        let chain=Promise.resolve();
        if(added.length===1) chain=chain.then(function(){ return insertMachineDb(added[0]); });
        else if(added.length>1) chain=chain.then(function(){ return bulkInsertMachinesDb(added); });
        removed.forEach(function(m){ chain=chain.then(function(){ return deleteMachineDb(m.id); }); });
        common.forEach(function(m){
          const old=prev.find(function(p){ return p.id===m.id; });
          if(old && JSON.stringify(old)!==JSON.stringify(m)){
            chain=chain.then(function(){ return updateMachineDb(m); });
          }
        });
        chain.then(function(){ setSyncError(""); }).catch(function(e){
          console.error("Sync failed:",e);
          setSyncError("⚠️ Could not save to the database (changes kept locally only — they will NOT appear for other people until this is fixed). " + (e && e.message ? e.message : ""));
        }).then(function(){ setSyncing(false); });
      })();
      return next;
    });
  }
  function updSettings(fn){
    setSettings(function(prev){
      const next=fn(prev);
      lsSave(LS_SETTINGS_KEY,next);
      setSyncing(true);
      updateSettingsDb(next).then(function(){ setSyncError(""); }).catch(function(e){
        console.error("Settings sync failed:",e);
        setSyncError("⚠️ Could not save settings to the database (kept locally only). " + (e && e.message ? e.message : ""));
      }).then(function(){ setSyncing(false); });
      return next;
    });
  }

  const sel=machines.find(function(m){ return m.id===selId; });

  function totalHrs(m){
    return m.hourLogs.filter(function(l){ return l.status==="Work"; }).reduce(function(s,l){ return s+l.hours; },0);
  }
  function sinceServiceHrs(m){
    if(m.serviceLogs.length===0) return totalHrs(m);
    const lastSvcDate=m.serviceLogs[0].date;
    return m.hourLogs.filter(function(l){ return l.status==="Work" && l.date>lastSvcDate; }).reduce(function(s,l){ return s+l.hours; },0);
  }
  function getDaysInMonth(y,mo){ return new Date(y,mo+1,0).getDate(); }
  function logSummary(logs){
    const work=logs.filter(function(l){ return l.status==="Work"; }).reduce(function(s,l){ return s+l.hours; },0);
    const idle=logs.filter(function(l){ return l.status==="Idle"; }).reduce(function(s,l){ return s+l.hours; },0);
    const repair=logs.filter(function(l){ return l.status==="Repair"; }).reduce(function(s,l){ return s+l.hours; },0);
    return {work:work,idle:idle,repair:repair,total:work+idle+repair};
  }
  function getDateRows(hourLogs,y,mo){
    const days=getDaysInMonth(y,mo);
    const rows=[];
    for(let d=1;d<=days;d++){
      const ds=y+"-"+String(mo+1).padStart(2,'0')+"-"+String(d).padStart(2,'0');
      const dayEntries=hourLogs.filter(function(l){ return l.date===ds; });
      rows.push({date:ds,entries:dayEntries});
    }
    return rows;
  }
  function cleanupDuplicates(){
    if(!sel) return;
    upd(function(p){ return p.map(function(m){ return m.id!==selId?m:Object.assign({},m); }); });
  }
  function prevMonth(){ if(viewMonth===0){setViewMonth(11);setViewYear(function(y){return y-1;});}else setViewMonth(function(m){return m-1;}); }
  function nextMonth(){ if(viewMonth===11){setViewMonth(0);setViewYear(function(y){return y+1;});}else setViewMonth(function(m){return m+1;}); }

  function downloadTemplate(){
    const csv="Machine Name,Machine Type,Model,Date Made,Operator Name,Operator Start Date\nExcavator Unit 1,Excavator,Komatsu PC200,2020-01-15,Ahmad bin Ali,2023-06-01";
    const blob=new Blob([csv],{type:"text/csv"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a"); a.href=url; a.download="machine_import_template.csv"; a.click();
    URL.revokeObjectURL(url);
  }

  function handleCSVUpload(e){
    setImportError(""); setImportSuccess("");
    const file=e.target.files[0]; if(!file) return;
    const reader=new FileReader();
    reader.onload=function(ev){
      try{
        const lines=ev.target.result.split("\n").map(function(l){return l.trim();}).filter(function(l){return l;});
        if(lines.length<2){ setImportError("CSV is empty or has no data rows."); return; }
        const header=lines[0].split(",").map(function(h){return h.trim().toLowerCase();});
        const nameIdx=header.findIndex(function(h){return h.indexOf("machine name")>=0;});
        const typeIdx=header.findIndex(function(h){return h.indexOf("machine type")>=0;});
        const modelIdx=header.findIndex(function(h){return h.indexOf("model")>=0;});
        const dateMadeIdx=header.findIndex(function(h){return h.indexOf("date made")>=0;});
        const operatorIdx=header.findIndex(function(h){return h.indexOf("operator name")>=0;});
        const opStartIdx=header.findIndex(function(h){return h.indexOf("operator start")>=0;});
        if(nameIdx<0){ setImportError("CSV must have a 'Machine Name' column."); return; }
        const newMachines=[];
        for(let i=1;i<lines.length;i++){
          const cols=lines[i].split(",").map(function(c){return c.trim();});
          const name=cols[nameIdx]||""; if(!name) continue;
          const rawType=typeIdx>=0?(cols[typeIdx]||""):"";
          const matchedType=MACHINE_TYPES.find(function(t){return t.toLowerCase()===rawType.toLowerCase();})||"Others";
          newMachines.push({
            id:Date.now()+i,name:name,type:matchedType,
            model:modelIdx>=0?(cols[modelIdx]||""):"",dateMade:dateMadeIdx>=0?(cols[dateMadeIdx]||""):"",
            operator:operatorIdx>=0?(cols[operatorIdx]||""):"",operatorStart:opStartIdx>=0?(cols[opStartIdx]||today()):today(),
            condition:"Work",serviceLogs:[],hourLogs:[]
          });
        }
        if(newMachines.length===0){ setImportError("No valid rows found in CSV."); return; }
        upd(function(p){ return p.concat(newMachines); });
        setImportSuccess("✅ Imported " + newMachines.length + " machine(s) successfully!");
        csvRef.current.value="";
      }catch(err){ setImportError("Failed to parse CSV."); }
    };
    reader.readAsText(file);
  }

  function addMachine(){
    if(!form.name.trim()) return;
    upd(function(p){
      return p.concat([{
        id:Date.now(),name:form.name.trim(),type:form.type,dateMade:form.dateMade,model:form.model.trim(),
        operator:form.operator.trim(),operatorStart:form.operatorStart,condition:form.condition,
        serviceLogs:[],hourLogs:[]
      }]);
    });
    setForm({name:"",type:"Excavator",dateMade:"",model:"",operator:"",operatorStart:today(),condition:"Work"});
    setView("dashboard");
  }

  function delLog(lid){
    upd(function(p){ return p.map(function(m){ return m.id!==selId?m:Object.assign({},m,{hourLogs:m.hourLogs.filter(function(l){return l.id!==lid;})}); }); });
  }
  function startEdit(l){ setEditLog(Object.assign({},l)); }
  function saveEdit(){
    if(!editLog) return;
    upd(function(p){
      return p.map(function(m){
        if(m.id!==selId) return m;
        return Object.assign({},m,{hourLogs:m.hourLogs.map(function(l){
          return l.id===editLog.id ? Object.assign({},editLog,{hours:parseFloat(editLog.hours)||0}) : l;
        })});
      });
    });
    setEditLog(null);
  }
  function saveDayPopup(){
    if(!dayPopup) return;
    const h=parseFloat(dayPopup.hours);
    if(!dayPopup.hours||h<=0) return;
    const existing=sel.hourLogs.find(function(l){ return l.date===dayPopup.date; });
    if(existing){
      upd(function(p){
        return p.map(function(m){
          if(m.id!==selId) return m;
          return Object.assign({},m,{hourLogs:m.hourLogs.map(function(l){
            return l.date===dayPopup.date ? Object.assign({},l,{hours:h,site:dayPopup.site,doNo:dayPopup.doNo,note:dayPopup.note}) : l;
          })});
        });
      });
    }else{
      upd(function(p){
        return p.map(function(m){
          if(m.id!==selId) return m;
          const newEntry={id:Date.now(),date:dayPopup.date,hours:h,note:dayPopup.note||"",status:"Work",site:dayPopup.site||"",doNo:dayPopup.doNo||""};
          return Object.assign({},m,{hourLogs:[newEntry].concat(m.hourLogs)});
        });
      });
    }
    setDayPopup(null);
  }
  function deleteDayPopupLog(){
    if(!dayPopup) return;
    const existing=sel.hourLogs.find(function(l){ return l.date===dayPopup.date; });
    if(existing) delLog(existing.id);
    setDayPopup(null);
  }
  function saveQuickHrs(dateStr,rawVal){
    const val=rawVal!==undefined?rawVal:quickHrs[dateStr];
    if(val===undefined||val===""){
      upd(function(p){ return p.map(function(m){ return m.id!==selId?m:Object.assign({},m,{hourLogs:m.hourLogs.filter(function(l){return l.date!==dateStr;})}); }); });
      return;
    }
    const h=parseFloat(val);
    if(isNaN(h)||h<=0) return;
    upd(function(p){
      return p.map(function(m){
        if(m.id!==selId) return m;
        const existing=m.hourLogs.find(function(l){ return l.date===dateStr; });
        const others=m.hourLogs.filter(function(l){ return l.date!==dateStr; });
        const newEntry=existing ? Object.assign({},existing,{hours:h}) : {id:Date.now(),date:dateStr,hours:h,note:"",status:"Work",site:"",doNo:""};
        return Object.assign({},m,{hourLogs:[newEntry].concat(others)});
      });
    });
  }
  function handleQuickHrsKeyDown(e,dateStr,allDates){
    if(e.key==="Enter"){
      e.preventDefault();
      e.target.blur();
      const idx=allDates.indexOf(dateStr);
      const nextDate=allDates[idx+1];
      if(nextDate){
        setTimeout(function(){
          const nextInput=hrsInputRefs.current[nextDate];
          if(nextInput){ nextInput.focus(); nextInput.select(); }
        },0);
      }
    }
  }
  // Handle pasting multiple values (from Excel) into the Hrs row — auto-fills consecutive days
  function handleHrsPaste(e,dateStr,allDates){
    const text=e.clipboardData.getData("text");
    const parts=text.split(/[\r\n\t,]+/).map(function(p){return p.trim();}).filter(function(p){return p!=="";});
    if(parts.length<=1) return;
    e.preventDefault();
    const startIdx=allDates.indexOf(dateStr);
    if(startIdx<0) return;
    parts.forEach(function(val,i){
      const targetDate=allDates[startIdx+i];
      if(!targetDate) return;
      const num=parseFloat(val);
      if(isNaN(num)) return;
      setQuickHrs(function(q){ const n=Object.assign({},q); n[targetDate]=String(num); return n; });
      saveQuickHrs(targetDate,String(num));
    });
  }

  function markServiced(){
    upd(function(p){
      return p.map(function(m){
        if(m.id!==selId) return m;
        const lastSvcDate=m.serviceLogs.length>0?m.serviceLogs[0].date:null;
        const workSinceSvc=m.hourLogs.filter(function(l){ return l.status==="Work" && (!lastSvcDate||l.date>lastSvcDate); }).reduce(function(s,l){return s+l.hours;},0);
        const sorted=m.hourLogs.slice().sort(function(a,b){return b.date.localeCompare(a.date);});
        const latestDate=sorted.length>0?sorted[0].date:today();
        const newSvc={date:latestDate,hoursAtService:Math.min(workSinceSvc,SERVICE_INTERVAL),serviceCount:m.serviceLogs.length+1,site:svcSite,doNo:svcDoNo};
        return Object.assign({},m,{serviceLogs:[newSvc].concat(m.serviceLogs)});
      });
    });
    setShowSvc(false); setSvcSite(""); setSvcDoNo("");
  }

  function updateCondition(id,cond){
    upd(function(p){ return p.map(function(m){ return m.id!==id?m:Object.assign({},m,{condition:cond}); }); });
  }
  function delMachine(id){
    upd(function(p){ return p.filter(function(m){ return m.id!==id; }); });
    setView("dashboard");
  }
  function delServiceLog(idx){
    upd(function(p){
      return p.map(function(m){
        if(m.id!==selId) return m;
        return Object.assign({},m,{serviceLogs:m.serviceLogs.filter(function(_,i){return i!==idx;})});
      });
    });
  }
  function saveEditMachine(){
    if(!editMachine||!editMachine.name.trim()) return;
    upd(function(p){
      return p.map(function(m){
        if(m.id!==editMachine.id) return m;
        return Object.assign({},m,{
          name:editMachine.name.trim(),type:editMachine.type,model:editMachine.model,
          dateMade:editMachine.dateMade,operator:editMachine.operator,operatorStart:editMachine.operatorStart
        });
      });
    });
    setEditMachine(null);
  }

  function getPrintData(){
    let mList=printType==="All Types"?machines:machines.filter(function(m){return m.type===printType;});
    if(printMachine!=="All Machines") mList=mList.filter(function(m){return m.name===printMachine;});
    const moStr=String(printMonth+1).padStart(2,'0');
    const thisMonthPrefix=printYear+"-"+moStr;
    const prevMonthEnd=printYear+"-"+moStr+"-01";
    return mList.map(function(m){
      const monthLogs=m.hourLogs.filter(function(l){return l.date.indexOf(thisMonthPrefix)===0;});
      const monthSum=logSummary(monthLogs);
      const cumulativeHrs=m.hourLogs.filter(function(l){return l.status==="Work" && l.date<prevMonthEnd;}).reduce(function(s,l){return s+l.hours;},0);
      const svcThisMonth=m.serviceLogs.filter(function(s){return s.date.indexOf(thisMonthPrefix)===0;}).sort(function(a,b){return b.date.localeCompare(a.date);});
      const svcPrevious=m.serviceLogs.filter(function(s){return s.date<prevMonthEnd;}).sort(function(a,b){return b.date.localeCompare(a.date);});
      const displaySvc=svcThisMonth.length>0?svcThisMonth[0]:(svcPrevious.length>0?svcPrevious[0]:null);
      return Object.assign({},m,{monthLogs:monthLogs,monthSum:monthSum,cumulativeHrs:cumulativeHrs,displaySvc:displaySvc});
    });
  }

  function handlePrintMonthly(){
    const printData=getPrintData();
    const rows=printData.map(function(m){
      const sd=m.displaySvc?(m.displaySvc.date.slice(8)+"-"+m.displaySvc.date.slice(5,7)+"-"+m.displaySvc.date.slice(2,4)):"—";
      const svcHrs=m.displaySvc?m.displaySvc.hoursAtService:0;
      const mTotal=m.monthSum.work; const total=m.cumulativeHrs+mTotal;
      return "<tr><td class='left bold'>"+m.name+"</td><td>"+(m.operator||"—")+"</td><td class='pink'>"+sd+"</td><td>"+(svcHrs>0?svcHrs.toFixed(1):"—")+"</td><td>"+m.cumulativeHrs.toFixed(1)+"</td><td>"+mTotal.toFixed(1)+"</td><td class='bold'>"+total.toFixed(1)+"</td></tr>";
    }).join("");
    const html="<!DOCTYPE html><html><head><meta charset='utf-8'><style>body{font-family:Arial;font-size:11px;padding:14px;}h2{margin-bottom:4px;}p{color:#888;font-size:9px;margin-bottom:10px;}table{width:100%;border-collapse:collapse;border:1.5px solid #dc2626;}thead tr{background:#dc2626;}th{color:white;padding:8px;text-align:center;font-size:10px;border:1px solid #f87171;}td{padding:7px 8px;text-align:center;border:1px solid #fecaca;}tr:nth-child(even) td{background:#fef2f2;}.left{text-align:left;padding-left:10px;}.bold{font-weight:bold;}.pink{color:#b91c1c;font-weight:bold;}@media print{@page{size:A4;margin:12mm;}}</style></head><body><h2>Monthly Machine Report — "+MONTHS[printMonth]+" "+printYear+"</h2><p>Generated: "+today()+" | Machines: "+printData.length+"</p><table><thead><tr><th style='text-align:left;padding-left:10px;'>Machine Name</th><th>Operator</th><th>Last Svc Date</th><th>Last Svc Hrs</th><th>Cumulative</th><th>Month Total</th><th>Total Hrs</th></tr></thead><tbody>"+rows+"</tbody></table><script>window.onload=function(){window.print();}<\/script></body></html>";
    const w=window.open('','_blank'); w.document.write(html); w.document.close();
  }

  function handlePrintYearlyService(){
    let mList=printType==="All Types"?machines:machines.filter(function(m){return m.type===printType;});
    if(printMachine!=="All Machines") mList=mList.filter(function(m){return m.name===printMachine;});
    const maxSvc=10;
    let svcHeaders="";
    for(let i=0;i<maxSvc;i++){
      svcHeaders+="<th style='background:#dc2626;color:white;border:1px solid #f87171;padding:8px 10px;font-size:11px;min-width:160px;'>Service #"+(i+1)+"</th>";
    }
    const dataRows=mList.map(function(m,ri){
      const allSvcs=m.serviceLogs.slice().reverse();
      const yearSvcs=allSvcs.filter(function(s){return s.date.indexOf(String(printYear))===0;});
      let cells="";
      for(let i=0;i<maxSvc;i++){
        const s=yearSvcs[i];
        if(!s){ cells+="<td style='border:1px solid #fecaca;padding:8px 10px;text-align:center;color:#fca5a5;'>—</td>"; continue; }
        const fmtDate=s.date.slice(8)+"-"+s.date.slice(5,7)+"-"+s.date.slice(2,4);
        const idxInAll=allSvcs.indexOf(s);
        const cumulative=allSvcs.slice(0,idxInAll+1).reduce(function(sum,x){return sum+x.hoursAtService;},0);
        cells+="<td style='border:1px solid #fecaca;padding:8px 10px;text-align:left;font-size:11px;line-height:1.6;'>"+
          "<div>📅 "+fmtDate+"</div>"+
          "<div>📍 "+(s.site||"—")+"</div>"+
          "<div>⏱ <strong>"+cumulative.toFixed(1)+"</strong> hrs</div>"+
          "<div style='color:#b91c1c;font-weight:bold;'>Do No: "+(s.doNo||"—")+"</div>"+
        "</td>";
      }
      return "<tr style='background:"+(ri%2===0?"#fff":"#fef2f2")+";'><td style='font-weight:bold;padding:8px 10px;border:1px solid #fecaca;white-space:nowrap;color:#7f1d1d;'>"+m.name+"</td>"+cells+"</tr>";
    }).join("");
    const html="<!DOCTYPE html><html><head><title>SERVICE RECORD LIST FOR THE YEAR "+printYear+"</title><style>*{box-sizing:border-box;margin:0;padding:0;}body{font-family:Arial;padding:14px;font-size:10px;}h2{margin-bottom:4px;font-size:13px;color:#7f1d1d;}p{color:#888;font-size:9px;margin-bottom:10px;}table{width:100%;border-collapse:collapse;border:2px solid #dc2626;}@media print{@page{size:A3 landscape;margin:8mm;}}</style></head><body><h2>SERVICE RECORD LIST FOR THE YEAR "+printYear+"</h2><p>Type: "+printType+" | Generated: "+today()+" | Machines: "+mList.length+"</p><table><thead><tr><th style='background:#7f1d1d;color:white;border:1px solid #f87171;padding:8px 10px;text-align:left;min-width:110px;'>No. Machine</th>"+svcHeaders+"</tr></thead><tbody>"+dataRows+"</tbody></table><script>window.onload=function(){window.print();}<\/script></body></html>";
    const w=window.open('','_blank'); w.document.write(html); w.document.close();
  }

  function exportExcelMonthly(){
    const printData=getPrintData();
    const wb=XLSX.utils.book_new();
    const rows=[["Machine Name","Operator","Last Svc Date","Last Svc Hrs","Cumulative Total","Month Total","Total Hrs"]];
    printData.forEach(function(m){
      const sd=m.displaySvc?(m.displaySvc.date.slice(8)+"-"+m.displaySvc.date.slice(5,7)+"-"+m.displaySvc.date.slice(2,4)):"—";
      const svcHrs=m.displaySvc?m.displaySvc.hoursAtService:0;
      const mTotal=m.monthSum.work;
      rows.push([m.name,m.operator||"—",sd,svcHrs>0?svcHrs:0,m.cumulativeHrs,mTotal,m.cumulativeHrs+mTotal]);
    });
    const ws=XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"]=[{wch:24},{wch:20},{wch:14},{wch:14},{wch:16},{wch:14},{wch:12}];
    XLSX.utils.book_append_sheet(wb,ws,MONTHS[printMonth]+" "+printYear);
    XLSX.writeFile(wb,"Monthly_Report_"+MONTHS[printMonth]+"_"+printYear+".xlsx");
  }

  function exportExcelYearly(){
    const wb=XLSX.utils.book_new();
    let mList=printType==="All Types"?machines:machines.filter(function(m){return m.type===printType;});
    if(printMachine!=="All Machines") mList=mList.filter(function(m){return m.name===printMachine;});
    const maxSvc=10;
    const h1=["No. Machine"];
    for(let i=1;i<=maxSvc;i++) h1.push("Service #"+i);
    const rows=[h1];
    mList.forEach(function(m){
      const allSvcs=m.serviceLogs.slice().reverse();
      const yearSvcs=allSvcs.filter(function(s){return s.date.indexOf(String(printYear))===0;});
      const row=[m.name];
      for(let i=0;i<maxSvc;i++){
        const s=yearSvcs[i];
        if(!s){ row.push("—"); continue; }
        const fmtDate=s.date.slice(8)+"-"+s.date.slice(5,7)+"-"+s.date.slice(2,4);
        const idxInAll=allSvcs.indexOf(s);
        const cumulative=allSvcs.slice(0,idxInAll+1).reduce(function(sum,x){return sum+x.hoursAtService;},0);
        row.push("Date: "+fmtDate+" | Site: "+(s.site||"—")+" | Hrs: "+cumulative.toFixed(1)+" | Do No: "+(s.doNo||"—"));
      }
      rows.push(row);
    });
    const ws=XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"]=[{wch:16}].concat(Array(maxSvc).fill({wch:32}));
    XLSX.utils.book_append_sheet(wb,ws,"Service Record "+printYear);
    XLSX.writeFile(wb,"Service_Record_"+printYear+".xlsx");
  }

  function exportWordMonthly(){
    const printData=getPrintData();
    const rows=printData.map(function(m,i){
      const sd=m.displaySvc?(m.displaySvc.date.slice(8)+"-"+m.displaySvc.date.slice(5,7)+"-"+m.displaySvc.date.slice(2,4)):"—";
      const svcHrs=m.displaySvc?m.displaySvc.hoursAtService:0;
      const mTotal=m.monthSum.work; const total=m.cumulativeHrs+mTotal;
      const bg=i%2===0?"":"background:#fef2f2;";
      return "<tr style='"+bg+"'><td class='left bold'>"+m.name+"</td><td>"+(m.operator||"—")+"</td><td class='pink'>"+sd+"</td><td>"+(svcHrs>0?svcHrs.toFixed(1):"—")+"</td><td>"+m.cumulativeHrs.toFixed(1)+"</td><td>"+mTotal.toFixed(1)+"</td><td class='bold'>"+total.toFixed(1)+"</td></tr>";
    }).join("");
    const body="<h2>Monthly Machine Report — "+MONTHS[printMonth]+" "+printYear+"</h2><p class='sub'>Generated: "+today()+" | Machines: "+printData.length+"</p><table><thead><tr><th class='left'>Machine Name</th><th>Operator</th><th>Last Svc Date</th><th>Last Svc Hrs</th><th>Cumulative</th><th>Month Total</th><th>Total Hrs</th></tr></thead><tbody>"+rows+"</tbody></table>";
    const blob=htmlToDocxBlob(body,"Monthly Report "+MONTHS[printMonth]+" "+printYear);
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a"); a.href=url; a.download="Monthly_Report_"+MONTHS[printMonth]+"_"+printYear+".doc"; a.click();
    URL.revokeObjectURL(url);
  }

  function exportWordYearly(){
    let mList=printType==="All Types"?machines:machines.filter(function(m){return m.type===printType;});
    if(printMachine!=="All Machines") mList=mList.filter(function(m){return m.name===printMachine;});
    const maxSvc=10;
    let svcTh="";
    for(let i=0;i<maxSvc;i++) svcTh+="<th>Service #"+(i+1)+"</th>";
    const dataRows=mList.map(function(m,ri){
      const allSvcs=m.serviceLogs.slice().reverse();
      const yearSvcs=allSvcs.filter(function(s){return s.date.indexOf(String(printYear))===0;});
      let cells="";
      for(let i=0;i<maxSvc;i++){
        const s=yearSvcs[i];
        if(!s){ cells+="<td>—</td>"; continue; }
        const fmtDate=s.date.slice(8)+"-"+s.date.slice(5,7)+"-"+s.date.slice(2,4);
        const idxInAll=allSvcs.indexOf(s);
        const cumulative=allSvcs.slice(0,idxInAll+1).reduce(function(sum,x){return sum+x.hoursAtService;},0);
        cells+="<td style='text-align:left;font-size:9pt;'>📅 "+fmtDate+"<br/>📍 "+(s.site||"—")+"<br/>⏱ <strong>"+cumulative.toFixed(1)+"</strong> hrs<br/><span class='pink'>Do No: "+(s.doNo||"—")+"</span></td>";
      }
      return "<tr style='"+(ri%2===0?"":"background:#fef2f2;")+"'><td class='left bold'>"+m.name+"</td>"+cells+"</tr>";
    }).join("");
    const body="<h2>SERVICE RECORD LIST FOR THE YEAR "+printYear+"</h2><p class='sub'>Type: "+printType+" | Generated: "+today()+" | Machines: "+mList.length+"</p><table><thead><tr><th class='left'>No. Machine</th>"+svcTh+"</tr></thead><tbody>"+dataRows+"</tbody></table>";
    const blob=htmlToDocxBlob(body,"Service Record "+printYear);
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a"); a.href=url; a.download="Service_Record_"+printYear+".doc"; a.click();
    URL.revokeObjectURL(url);
  }

  const filtered=machines.filter(function(m){
    const ms=(m.name+m.type+(m.operator||"")+(m.model||"")).toLowerCase();
    return ms.indexOf(search.toLowerCase())>=0 && (fType==="All Types"||m.type===fType) && (fStatus==="All Status"||m.condition===fStatus);
  });

  function exportFullBackup(){
    const wb=XLSX.utils.book_new();
    const sr=[["Machine Name","Type","Model","Condition","Hours Since Svc","Total Hours","Last Service Date"]];
    machines.forEach(function(m){
      sr.push([m.name,m.type,m.model||"",m.condition,sinceServiceHrs(m),totalHrs(m),m.serviceLogs.length>0?m.serviceLogs[0].date:"—"]);
    });
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sr),"Machine Summary");
    const lr=[["Machine Name","Date","Status","Hours","Site","Do No.","Note"]];
    machines.forEach(function(m){
      m.hourLogs.forEach(function(l){ lr.push([m.name,l.date,l.status,l.hours,l.site||"",l.doNo||"",l.note||""]); });
    });
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(lr),"Daily Hour Logs");
    const svcr=[["Machine Name","Service #","Date","Site","Do No.","Hours at Service"]];
    machines.forEach(function(m){
      const revLogs=m.serviceLogs.slice();
      revLogs.forEach(function(l,i){ svcr.push([m.name,m.serviceLogs.length-i,l.date,l.site||"",l.doNo||"",l.hoursAtService]); });
    });
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(svcr),"Service History");
    XLSX.writeFile(wb,"MachineTracker_Backup_"+today()+".xlsx");
  }

  // ── SETTINGS VIEW ──
  if(view==="settings"){
    return (
      <div className="min-h-screen pb-10" style={{background:"#fff5f5"}}>
        <div className="p-4 max-w-md mx-auto">
          <button onClick={function(){setView("dashboard");}} className="font-bold text-red-700 mb-4 flex items-center gap-1 text-sm opacity-70 hover:opacity-100">← Back</button>
          <h2 className="text-xl font-black text-red-900 mb-4">⚙️ Settings</h2>
          <div className="rounded-2xl shadow p-4 mb-4 bg-white border border-red-100">
            <h3 className="font-black text-red-900 mb-3">👷 Operator List</h3>
            <div className="flex gap-2 mb-3">
              <input className={inp+" flex-1"} placeholder="Add operator name..." value={newOp} onChange={function(e){setNewOp(e.target.value);}}
                onKeyDown={function(e){
                  if(e.key==="Enter" && newOp.trim()){
                    updSettings(function(s){ return Object.assign({},s,{operators:s.operators.concat([newOp.trim()])}); });
                    setNewOp("");
                  }
                }}/>
              <button onClick={function(){
                if(newOp.trim()){
                  updSettings(function(s){ return Object.assign({},s,{operators:s.operators.concat([newOp.trim()])}); });
                  setNewOp("");
                }
              }} className="text-white px-4 py-2 rounded-xl font-bold text-sm" style={{background:"linear-gradient(90deg,#dc2626,#b91c1c)"}}>+ Add</button>
            </div>
            {settings.operators.length===0?(
              <p className="text-gray-400 text-sm text-center py-3">No operators added yet.</p>
            ):(
              <div className="space-y-2">
                {settings.operators.map(function(op,i){
                  return (
                    <div key={i} className="flex items-center bg-red-50 rounded-xl px-3 py-2 gap-2">
                      {editOpIdx===i?(
                        <>
                          <input className="flex-1 border border-red-400 rounded-lg px-2 py-1 text-sm bg-white text-gray-800 focus:outline-none" value={editOpVal} onChange={function(e){setEditOpVal(e.target.value);}}/>
                          <button onClick={function(){
                            if(editOpVal.trim()){
                              updSettings(function(s){ return Object.assign({},s,{operators:s.operators.map(function(x,j){return j===i?editOpVal.trim():x;})}); });
                              setEditOpIdx(null);
                            }
                          }} className="text-white text-xs px-3 py-1 rounded-lg font-bold" style={{background:"#dc2626"}}>Save</button>
                          <button onClick={function(){setEditOpIdx(null);}} className="text-gray-400 text-xs px-2 py-1 rounded-lg border border-gray-300">✕</button>
                        </>
                      ):(
                        <>
                          <span className="text-red-900 text-sm font-semibold flex-1">👷 {op}</span>
                          <button onClick={function(){setEditOpIdx(i);setEditOpVal(op);}} className="text-xs text-red-600 border border-red-400 rounded-lg px-2 py-1 hover:bg-red-100">✏️</button>
                          <button onClick={function(){
                            updSettings(function(s){ return Object.assign({},s,{operators:s.operators.filter(function(_,j){return j!==i;})}); });
                          }} className="text-xs text-red-500 border border-red-400 rounded-lg px-2 py-1 hover:bg-red-50">🗑</button>
                        </>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
          <div className="rounded-2xl shadow p-4 bg-white border border-red-100">
            <h3 className="font-black text-red-900 mb-3">📍 Project Site List</h3>
            <div className="flex gap-2 mb-3">
              <input className={inp+" flex-1"} placeholder="Add project site..." value={newSite} onChange={function(e){setNewSite(e.target.value);}}
                onKeyDown={function(e){
                  if(e.key==="Enter" && newSite.trim()){
                    updSettings(function(s){ return Object.assign({},s,{sites:s.sites.concat([newSite.trim()])}); });
                    setNewSite("");
                  }
                }}/>
              <button onClick={function(){
                if(newSite.trim()){
                  updSettings(function(s){ return Object.assign({},s,{sites:s.sites.concat([newSite.trim()])}); });
                  setNewSite("");
                }
              }} className="text-white px-4 py-2 rounded-xl font-bold text-sm" style={{background:"linear-gradient(90deg,#dc2626,#b91c1c)"}}>+ Add</button>
            </div>
            {settings.sites.length===0?(
              <p className="text-gray-400 text-sm text-center py-3">No sites added yet.</p>
            ):(
              <div className="space-y-2">
                {settings.sites.map(function(site,i){
                  return (
                    <div key={i} className="flex items-center bg-red-50 rounded-xl px-3 py-2 gap-2">
                      {editSiteIdx===i?(
                        <>
                          <input className="flex-1 border border-red-400 rounded-lg px-2 py-1 text-sm bg-white text-gray-800 focus:outline-none" value={editSiteVal} onChange={function(e){setEditSiteVal(e.target.value);}}/>
                          <button onClick={function(){
                            if(editSiteVal.trim()){
                              updSettings(function(s){ return Object.assign({},s,{sites:s.sites.map(function(x,j){return j===i?editSiteVal.trim():x;})}); });
                              setEditSiteIdx(null);
                            }
                          }} className="text-white text-xs px-3 py-1 rounded-lg font-bold" style={{background:"#dc2626"}}>Save</button>
                          <button onClick={function(){setEditSiteIdx(null);}} className="text-gray-400 text-xs px-2 py-1 rounded-lg border border-gray-300">✕</button>
                        </>
                      ):(
                        <>
                          <span className="text-red-900 text-sm font-semibold flex-1">📍 {site}</span>
                          <button onClick={function(){setEditSiteIdx(i);setEditSiteVal(site);}} className="text-xs text-red-600 border border-red-400 rounded-lg px-2 py-1 hover:bg-red-100">✏️</button>
                          <button onClick={function(){
                            updSettings(function(s){ return Object.assign({},s,{sites:s.sites.filter(function(_,j){return j!==i;})}); });
                          }} className="text-xs text-red-500 border border-red-400 rounded-lg px-2 py-1 hover:bg-red-50">🗑</button>
                        </>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </div>
      </div>
    );
  }

  // ── ADD VIEW ──
  if(view==="add"){
    return (
      <div className="min-h-screen pb-10" style={{background:"#fff5f5"}}>
        <div className="p-4 max-w-md mx-auto">
          <button onClick={function(){setView("dashboard");}} className="font-bold text-red-700 mb-4 flex items-center gap-1 text-sm opacity-70 hover:opacity-100">← Back</button>
          <h2 className="text-xl font-black text-red-900 mb-4">Add New Machine</h2>
          <div className="rounded-2xl shadow p-4 space-y-3 bg-white border border-red-100">
            <Field label="Machine Name *"><input className={inp} placeholder="e.g. Excavator Unit 1" value={form.name} onChange={function(e){setForm(function(f){return Object.assign({},f,{name:e.target.value});});}}/></Field>
            <Field label="Machine Type">
              <select className={inp} value={form.type} onChange={function(e){setForm(function(f){return Object.assign({},f,{type:e.target.value});});}}>
                {MACHINE_TYPES.map(function(t){return <option key={t}>{t}</option>;})}
              </select>
            </Field>
            <Field label="Model"><input className={inp} placeholder="e.g. Komatsu PC200" value={form.model} onChange={function(e){setForm(function(f){return Object.assign({},f,{model:e.target.value});});}}/></Field>
            <Field label="Date Made"><input type="date" className={inp} value={form.dateMade} onChange={function(e){setForm(function(f){return Object.assign({},f,{dateMade:e.target.value});});}}/></Field>
            <div className="border-t border-red-100 pt-3">
              <p className="text-xs font-black text-red-900 uppercase mb-2">Operator Info</p>
              <Field label="Operator Name">
                {settings.operators.length>0?(
                  <select className={inp} value={form.operator} onChange={function(e){setForm(function(f){return Object.assign({},f,{operator:e.target.value});});}}>
                    <option value="">— Select Operator —</option>
                    {settings.operators.map(function(op){return <option key={op}>{op}</option>;})}
                  </select>
                ):(
                  <input className={inp} placeholder="e.g. Ahmad bin Ali" value={form.operator} onChange={function(e){setForm(function(f){return Object.assign({},f,{operator:e.target.value});});}}/>
                )}
              </Field>
              <div className="mt-2">
                <Field label="Start Working Date"><input type="date" className={inp} value={form.operatorStart} onChange={function(e){setForm(function(f){return Object.assign({},f,{operatorStart:e.target.value});});}}/></Field>
              </div>
            </div>
            <div className="border-t border-red-100 pt-3">
              <p className="text-xs font-black text-red-900 uppercase mb-2">Machine Condition</p>
              <div className="flex gap-2">
                {STATUSES.map(function(s){
                  return (
                    <button key={s} onClick={function(){setForm(function(f){return Object.assign({},f,{condition:s});});}}
                      className={"flex-1 py-2 rounded-xl text-sm font-semibold border transition " + (form.condition===s?SS[s].btn:"bg-gray-50 text-gray-400 border-gray-200")}>
                      {s==="Work"?"🟢":s==="Idle"?"🟡":"🔴"} {s}
                    </button>
                  );
                })}
              </div>
            </div>
            <button onClick={addMachine} className="w-full text-white rounded-xl py-2.5 font-bold" style={{background:"linear-gradient(90deg,#dc2626,#b91c1c)"}}>Add Machine</button>
          </div>
        </div>
      </div>
    );
  }

  // ── IMPORT VIEW ──
  if(view==="import"){
    return (
      <div className="min-h-screen pb-10" style={{background:"#fff5f5"}}>
        <div className="p-4 max-w-md mx-auto">
          <button onClick={function(){setView("dashboard");}} className="font-bold text-red-700 mb-4 flex items-center gap-1 text-sm opacity-70 hover:opacity-100">← Back</button>
          <h2 className="text-xl font-black text-red-900 mb-4">📥 Data Import</h2>
          <div className="bg-white rounded-2xl shadow p-5 space-y-4 border border-red-100">
            <div className="flex gap-3 items-start">
              <div className="w-7 h-7 rounded-full flex items-center justify-center text-white text-sm font-bold shrink-0 mt-0.5" style={{background:"#dc2626"}}>1</div>
              <div className="flex-1">
                <p className="font-bold text-gray-800 text-sm">Download Template</p>
                <button onClick={downloadTemplate} className="mt-1 text-white text-sm px-4 py-2 rounded-xl font-semibold" style={{background:"linear-gradient(90deg,#dc2626,#b91c1c)"}}>⬇️ Download Template</button>
              </div>
            </div>
            <div className="border-t border-gray-200"/>
            <div className="flex gap-3 items-start">
              <div className="w-7 h-7 rounded-full flex items-center justify-center text-white text-sm font-bold shrink-0 mt-0.5" style={{background:"#dc2626"}}>2</div>
              <div className="flex-1">
                <p className="font-bold text-gray-800 text-sm">Upload CSV</p>
                <label className="block w-full border-2 border-dashed border-red-300 rounded-xl p-4 text-center cursor-pointer hover:bg-red-50 transition mt-1">
                  <p className="text-red-600 font-semibold text-sm">📂 Click to select CSV file</p>
                  <input ref={csvRef} type="file" accept=".csv" className="hidden" onChange={handleCSVUpload}/>
                </label>
                {importError&&<p className="text-red-600 text-xs mt-2 font-semibold">⚠️ {importError}</p>}
                {importSuccess&&(
                  <div className="mt-2 bg-green-50 border border-green-200 rounded-xl p-3">
                    <p className="text-green-700 text-xs font-semibold">{importSuccess}</p>
                    <button onClick={function(){setView("dashboard");}} className="mt-2 text-white text-xs px-3 py-1.5 rounded-lg font-bold" style={{background:"#dc2626"}}>Go to Dashboard →</button>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ── PRINT VIEW ──
  if(view==="print"){
    const printData=getPrintData();
    return (
      <div className="min-h-screen pb-10" style={{background:"#fff5f5"}}>
        <div className="p-4 bg-white border-b-2 border-red-500">
          <div className="max-w-2xl mx-auto flex justify-between items-center">
            <div><h2 className="text-xl font-black text-red-900">🖨️ Print</h2><p className="text-red-500 text-xs">Choose report type</p></div>
            <button onClick={function(){setView("dashboard");}} className="text-red-600 text-sm hover:text-red-900">← Back</button>
          </div>
        </div>
        <div className="p-4 max-w-2xl mx-auto">
          <div className="flex gap-2 mb-4">
            {[["monthly","📅 Monthly Report"],["yearly","📋 Yearly Service Record"]].map(function(item){
              const val=item[0], lb=item[1];
              return (
                <button key={val} onClick={function(){setPrintReportType(val);}}
                  className="flex-1 py-2.5 rounded-xl font-bold text-sm transition"
                  style={printReportType===val?{background:"#dc2626",color:"white"}:{background:"#fee2e2",color:"#b91c1c",border:"1px solid #fca5a5"}}>
                  {lb}
                </button>
              );
            })}
          </div>

          <div className="bg-white rounded-2xl shadow p-4 mb-4 border border-red-100">
            <h3 className="font-black text-red-900 mb-3">Filter</h3>
            <div className="grid grid-cols-2 gap-3 mb-4">
              {printReportType==="monthly" && (
                <div>
                  <label className="text-xs text-red-500 font-medium">Month</label>
                  <select className={inp} value={printMonth} onChange={function(e){setPrintMonth(parseInt(e.target.value));}}>
                    {MONTHS.map(function(m,i){return <option key={m} value={i}>{m}</option>;})}
                  </select>
                </div>
              )}
              <div>
                <label className="text-xs text-red-500 font-medium">Year</label>
                <select className={inp} value={printYear} onChange={function(e){setPrintYear(parseInt(e.target.value));}}>
                  {[2023,2024,2025,2026,2027].map(function(y){return <option key={y}>{y}</option>;})}
                </select>
              </div>
              <div>
                <label className="text-xs text-red-500 font-medium">Machine Type</label>
                <select className={inp} value={printType} onChange={function(e){setPrintType(e.target.value);setPrintMachine("All Machines");}}>
                  {FILTER_TYPES.map(function(t){return <option key={t}>{t}</option>;})}
                </select>
              </div>
              <div>
                <label className="text-xs text-red-500 font-medium">Machine Name</label>
                <select className={inp} value={printMachine} onChange={function(e){setPrintMachine(e.target.value);}}>
                  <option>All Machines</option>
                  {(printType==="All Types"?machines:machines.filter(function(m){return m.type===printType;})).map(function(m){return <option key={m.id}>{m.name}</option>;})}
                </select>
              </div>
            </div>
            <div className="grid grid-cols-3 gap-2">
              <button onClick={printReportType==="monthly"?handlePrintMonthly:handlePrintYearlyService}
                className="text-white rounded-xl py-2 font-bold text-xs" style={{background:"linear-gradient(90deg,#dc2626,#b91c1c)"}}>
                🖨️ Print / PDF
              </button>
              <button onClick={printReportType==="monthly"?exportExcelMonthly:exportExcelYearly}
                className="text-white rounded-xl py-2 font-bold text-xs" style={{background:"linear-gradient(90deg,#16a34a,#15803d)"}}>
                📊 Excel
              </button>
              <button onClick={printReportType==="monthly"?exportWordMonthly:exportWordYearly}
                className="text-white rounded-xl py-2 font-bold text-xs" style={{background:"linear-gradient(90deg,#ea580c,#c2410c)"}}>
                📝 Word
              </button>
            </div>
          </div>

          <div className="rounded-2xl shadow overflow-hidden bg-white border border-red-100">
            <div className="px-4 py-3 border-b border-red-100">
              <p className="text-red-700 font-bold text-sm">
                {printReportType==="monthly"?("📅 "+MONTHS[printMonth]+" "+printYear):("📋 SERVICE RECORD "+printYear)}
              </p>
              <p className="text-red-400 text-xs">{(printReportType==="monthly"?printData.length:machines.length)} machine(s)</p>
            </div>
            {printReportType==="monthly"?(
              printData.length===0?(
                <div className="p-8 text-center text-gray-400">No machines found.</div>
              ):(
                <div className="overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead>
                      <tr style={{background:"#dc2626"}}>
                        {["Machine Name","Operator","Last Svc","Svc Hrs","Cumulative","Month","Total"].map(function(h){
                          return <th key={h} className="px-3 py-2 text-white font-bold text-left">{h}</th>;
                        })}
                      </tr>
                    </thead>
                    <tbody>
                      {printData.map(function(m,i){
                        const sd=m.displaySvc?(m.displaySvc.date.slice(8)+"-"+m.displaySvc.date.slice(5,7)+"-"+m.displaySvc.date.slice(2,4)):"—";
                        const svcHrs=m.displaySvc?m.displaySvc.hoursAtService:0;
                        const mTotal=m.monthSum.work; const total=m.cumulativeHrs+mTotal;
                        return (
                          <tr key={m.id} style={{background:i%2===0?"#fff":"#fef2f2"}}>
                            <td className="px-3 py-2 font-bold text-red-900">{m.name}</td>
                            <td className="px-3 py-2 text-gray-500">{m.operator||"—"}</td>
                            <td className="px-3 py-2 text-red-600 font-semibold">{sd}</td>
                            <td className="px-3 py-2 text-center text-gray-600">{svcHrs>0?svcHrs.toFixed(1):"—"}</td>
                            <td className="px-3 py-2 text-center text-gray-600">{m.cumulativeHrs.toFixed(1)}</td>
                            <td className="px-3 py-2 text-center text-gray-600">{mTotal.toFixed(1)}</td>
                            <td className="px-3 py-2 text-center font-black text-red-900">{total.toFixed(1)}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )
            ):(
              <div className="overflow-x-auto p-3">
                {machines.length===0?(
                  <div className="p-4 text-center text-gray-400">No machines.</div>
                ):(
                  <table className="border-collapse text-xs" style={{minWidth:"800px"}}>
                    <thead>
                      <tr>
                        <th className="px-3 py-2 text-white font-bold text-left sticky left-0 z-10" style={{background:"#7f1d1d",minWidth:"110px"}}>No. Machine</th>
                        {Array.from({length:10},function(_,i){
                          return <th key={i} className="px-3 py-2 text-white font-bold text-center" style={{background:"#dc2626",minWidth:"170px",border:"1px solid #f87171"}}>Service #{i+1}</th>;
                        })}
                      </tr>
                    </thead>
                    <tbody>
                      {(printType==="All Types"?machines:machines.filter(function(m){return m.type===printType;}))
                        .filter(function(m){return printMachine==="All Machines"||m.name===printMachine;})
                        .map(function(m,ri){
                          const allSvcs=m.serviceLogs.slice().reverse();
                          const yearSvcs=allSvcs.filter(function(s){return s.date.indexOf(String(printYear))===0;});
                          return (
                            <tr key={m.id} style={{background:ri%2===0?"#fff":"#fef2f2"}}>
                              <td className="px-3 py-2 font-black text-red-900 sticky left-0 z-10" style={{background:ri%2===0?"#fff":"#fef2f2",border:"1px solid #fecaca"}}>{m.name}</td>
                              {Array.from({length:10},function(_,i){
                                const s=yearSvcs[i];
                                if(!s) return <td key={i} className="px-3 py-2 text-center text-gray-300" style={{border:"1px solid #fecaca"}}>—</td>;
                                const fmtDate=s.date.slice(8)+"-"+s.date.slice(5,7)+"-"+s.date.slice(2,4);
                                const idxInAll=allSvcs.indexOf(s);
                                const cumulative=allSvcs.slice(0,idxInAll+1).reduce(function(sum,x){return sum+x.hoursAtService;},0);
                                return (
                                  <td key={i} className="px-3 py-2 text-left" style={{border:"1px solid #fecaca",fontSize:"11px",lineHeight:"1.5"}}>
                                    <div className="text-gray-600">📅 {fmtDate}</div>
                                    <div className="text-gray-600">📍 {s.site||"—"}</div>
                                    <div className="text-gray-700">⏱ <strong className="text-red-900">{cumulative.toFixed(1)}</strong> hrs</div>
                                    <div className="font-bold" style={{color:"#b91c1c"}}>Do No: {s.doNo||"—"}</div>
                                  </td>
                                );
                              })}
                            </tr>
                          );
                        })}
                    </tbody>
                  </table>
                )}
              </div>
            )}
          </div>
        </div>
      </div>
    );
  }

  // ── DETAIL VIEW ──
  if(view==="detail" && sel){
    const tot=totalHrs(sel);
    const svcNo=sel.serviceLogs.length+1;
    const sinceSvc=sinceServiceHrs(sel);
    const dateRows=getDateRows(sel.hourLogs,viewYear,viewMonth);
    const monthEntries=dateRows.reduce(function(acc,r){return acc.concat(r.entries);},[]);
    const summary=logSummary(monthEntries);
    const canService=sinceSvc>0;

    return (
      <div className="min-h-screen pb-10" style={{background:"#fff5f5"}}>
        {syncing && (
          <div className="fixed top-0 left-0 right-0 z-50 flex items-center justify-center gap-2 py-2 text-white text-sm font-bold"
            style={{background:"linear-gradient(90deg,#16a34a,#15803d)",boxShadow:"0 2px 10px rgba(0,0,0,0.2)"}}>
            <span className="inline-block w-2 h-2 rounded-full bg-white" style={{animation:"pulse 1s infinite"}}/>
            ● Syncing data...
          </div>
        )}
        <div className="p-4">
          <button onClick={function(){setView("dashboard");}} className="font-bold text-red-700 mb-3 flex items-center gap-1 text-sm opacity-70 hover:opacity-100">← Back</button>

          <div className="grid gap-4 mb-4" style={{gridTemplateColumns:"1fr 1fr 1.2fr"}}>
            <div className="rounded-2xl shadow p-4 bg-white border border-red-100">
              <div className="flex justify-between items-center mb-3">
                <h3 className="font-black text-red-900">🚜 Machine Info</h3>
                <button onClick={function(){
                  setEditMachine({id:sel.id,name:sel.name,type:sel.type,model:sel.model||"",dateMade:sel.dateMade||"",operator:sel.operator||"",operatorStart:sel.operatorStart||today()});
                }} className="text-xs text-red-600 border border-red-300 bg-white rounded-lg px-2 py-1 font-semibold hover:bg-red-50 transition">✏️ Edit</button>
              </div>
              <div className="grid grid-cols-2 gap-2 text-sm">
                <div><p className="text-xs font-bold text-red-400">Type</p><p className="font-semibold text-gray-700">{sel.type}</p></div>
                <div><p className="text-xs font-bold text-red-400">Model</p><p className="font-semibold text-gray-700">{sel.model||"—"}</p></div>
              </div>
              <div className="mt-2">
                <p className="text-xs font-bold text-red-400">Date Made</p>
                <p className="font-semibold text-gray-700">{sel.dateMade?(sel.dateMade.slice(8)+"-"+sel.dateMade.slice(5,7)+"-"+sel.dateMade.slice(2,4)):"—"}</p>
              </div>
              <div className="border-t border-red-100 mt-2 pt-2 grid grid-cols-2 gap-2 text-sm">
                <div><p className="text-xs font-bold text-red-400">👷 Operator</p><p className="font-semibold text-gray-700">{sel.operator||"—"}</p></div>
                <div><p className="text-xs font-bold text-red-400">Start Working</p><p className="font-semibold text-gray-700">{sel.operatorStart?(sel.operatorStart.slice(8)+"-"+sel.operatorStart.slice(5,7)+"-"+sel.operatorStart.slice(2,4)):"—"}</p></div>
              </div>
            </div>

            <div className="rounded-2xl shadow p-4 bg-white border border-red-100">
              <div className="flex justify-between items-start mb-2">
                <h2 className="text-xl font-black text-red-900">{sel.name}</h2>
                <ServiceBadge hours={sinceSvc}/>
              </div>
              <p className="text-xs text-red-500 mb-3">{sel.type}</p>
              <div className="flex gap-1.5 mb-3">
                {STATUSES.map(function(s){
                  return (
                    <button key={s} onClick={function(){updateCondition(sel.id,s);}}
                      className={"flex-1 py-1.5 rounded-xl text-xs font-bold border transition " + (sel.condition===s?"bg-red-600 text-white border-red-600 shadow":"bg-red-50 text-red-700 border-red-200")}>
                      {s==="Work"?"🟢":s==="Idle"?"🟡":"🔴"} {s}
                    </button>
                  );
                })}
              </div>
              <div className="grid grid-cols-3 gap-2 mb-3">
                {[["Since Svc",sinceSvc.toFixed(1)],["Total Hrs",tot.toFixed(1)],["Next Svc","#"+svcNo]].map(function(item){
                  const l=item[0], v=item[1];
                  return (
                    <div key={l} className="bg-red-50 rounded-xl p-2.5 text-center border border-red-100">
                      <p className="text-lg font-bold text-red-900">{v}</p>
                      <p className="text-xs text-red-500">{l}</p>
                    </div>
                  );
                })}
              </div>
              <div className="flex justify-between text-xs text-red-500 mb-1">
                <span>{sinceSvc.toFixed(1)} / {SERVICE_INTERVAL} hrs</span>
                <span>{Math.max(0,SERVICE_INTERVAL-sinceSvc).toFixed(1)} hrs left</span>
              </div>
              <ProgressBar hours={sinceSvc}/>
              {canService && (
                <button onClick={function(){setShowSvc(true);}} className="w-full mt-3 text-white rounded-xl py-2 font-bold text-sm transition" style={{background:"linear-gradient(90deg,#dc2626,#b91c1c)"}}>
                  🔧 Mark as Serviced
                </button>
              )}
            </div>

            <div className="rounded-2xl shadow overflow-hidden bg-white border border-red-100">
              <div className="px-4 pt-4 pb-2">
                <h3 className="font-bold text-lg mb-3" style={{color:"#b91c1c"}}>🔧 Service History</h3>
              </div>
              <div className="flex justify-between items-center px-4 py-3 mb-2 font-bold text-white text-sm" style={{background:"#dc2626"}}>
                <span>Total Hours</span><span>{tot.toFixed(1)} hrs</span>
              </div>
              {sel.serviceLogs.length===0?(
                <div className="px-4 py-6 text-center text-gray-400 text-sm">No service records yet.</div>
              ):(
                <table className="w-full text-xs">
                  <thead>
                    <tr className="bg-red-50">
                      <th className="px-2 py-2 text-red-700 font-bold text-left">Svc #</th>
                      <th className="px-2 py-2 text-red-700 font-bold">Date</th>
                      <th className="px-2 py-2 text-red-700 font-bold">Site</th>
                      <th className="px-2 py-2 text-red-700 font-bold" style={{color:"#b91c1c"}}>Do No.</th>
                      <th className="px-2 py-2 text-red-700 font-bold">Hrs</th>
                      <th className="px-2 py-2 text-red-700 font-bold">Meter</th>
                      <th className="px-2 py-2 w-8"></th>
                    </tr>
                  </thead>
                  <tbody>
                    {sel.serviceLogs.map(function(l,i){
                      const svcNum=sel.serviceLogs.length-i;
                      const fmtDate=l.date.slice(8)+"/"+l.date.slice(5,7)+"/"+l.date.slice(0,4);
                      const sortedAsc=sel.serviceLogs.slice().reverse();
                      const idxAsc=sortedAsc.length-1-i;
                      const hrsMeter=sortedAsc.slice(0,idxAsc+1).reduce(function(s,x){return s+x.hoursAtService;},0);
                      return (
                        <tr key={i} className="border-b last:border-0 border-red-50">
                          <td className="px-2 py-2 text-red-900 font-black">#{svcNum}</td>
                          <td className="px-2 py-2 text-center text-gray-500 font-bold">{fmtDate}</td>
                          <td className="px-2 py-2 text-center text-gray-600">{l.site||"—"}</td>
                          <td className="px-2 py-2 text-center font-bold" style={{color:"#b91c1c"}}>{l.doNo||"—"}</td>
                          <td className="px-2 py-2 text-right font-bold text-gray-600">{l.hoursAtService.toFixed(1)}</td>
                          <td className="px-2 py-2 text-right font-black" style={{color:"#b91c1c"}}>{hrsMeter.toFixed(1)}</td>
                          <td className="px-2 py-2 text-center">
                            <button onClick={function(){delServiceLog(i);}} className="text-xs text-red-500 border border-red-400 hover:bg-red-50 rounded px-1 py-0.5">🗑</button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>
          </div>

          <div className="px-0 space-y-3">
            <div className="rounded-2xl shadow overflow-hidden bg-white border border-red-100"
              style={{position:"relative",left:"50%",right:"50%",marginLeft:"-50vw",marginRight:"-50vw",width:"100vw",maxWidth:"100vw"}}>
              <div className="px-4 pt-3 pb-2">
                <h3 className="font-black text-red-900 mb-2">📋 Daily Hour Log</h3>
                <div className="flex items-center justify-between bg-red-50 rounded-xl px-3 py-1.5 mb-2">
                  <button onClick={prevMonth} className="text-red-600 font-bold px-2 text-lg hover:text-red-900">‹</button>
                  <span className="text-sm font-black text-red-900">{MONTHS[viewMonth]} {viewYear}</span>
                  <button onClick={nextMonth} className="text-red-600 font-bold px-2 text-lg hover:text-red-900">›</button>
                </div>
                <button onClick={cleanupDuplicates} className="text-xs text-amber-600 border border-amber-400 rounded-lg px-2 py-1 mb-2 hover:bg-amber-50">🧹 Clean up duplicate entries</button>
                <p className="text-xs text-gray-400 mb-1">Tap any date column to add/edit log. You can paste a range of values from Excel into the Hrs row.</p>
              </div>
              <div className="pb-3 px-2" style={{overflowX:"auto",WebkitOverflowScrolling:"touch"}}>
                <table className="text-xs border-collapse" style={{minWidth:(dateRows.length*42+90)+"px"}}>
                  <tbody>
                    <tr>
                      <td className="px-1 py-2 font-black text-white sticky left-0 z-10" style={{background:"#7f1d1d",fontSize:"10px",minWidth:"44px"}}>Date</td>
                      {dateRows.map(function(row){
                        const isToday=row.date===today();
                        const isSvcDate=sel.serviceLogs.some(function(s){return s.date===row.date;});
                        const isSunday=new Date(row.date+"T00:00:00").getDay()===0;
                        const hasEntry=row.entries.length>0;
                        const e=row.entries[0];
                        const bg=isSvcDate?"#fde68a":isToday?"#7f1d1d":isSunday?"#fef9c3":hasEntry?"#fecaca":"#fafafa";
                        return (
                          <td key={row.date} onClick={function(){
                              setDayPopup({date:row.date,hours:e?String(e.hours):"",site:(e&&e.site)||"",doNo:(e&&e.doNo)||"",note:(e&&e.note)||"",logId:(e&&e.id)||null});
                            }}
                            className="px-1 py-2 text-center font-bold cursor-pointer hover:opacity-80"
                            style={{minWidth:"38px",background:bg,color:isToday?"white":isSunday?"#a16207":"#7f1d1d",border:"1px solid #fecaca",fontSize:"11px"}}>
                            {row.date.slice(8)}
                            {isSvcDate && <div style={{fontSize:"8px"}}>🔧</div>}
                          </td>
                        );
                      })}
                      <td className="px-1 py-2 font-black text-white text-center sticky right-0" style={{background:"#7f1d1d",fontSize:"9px",minWidth:"44px"}}>Tot</td>
                    </tr>
                    <tr>
                      <td className="px-1 py-1.5 font-black text-red-900 sticky left-0 z-10 bg-red-50" style={{fontSize:"10px"}}>Hrs</td>
                      {dateRows.map(function(row){
                        const e=row.entries[0];
                        const allDates=dateRows.map(function(r){return r.date;});
                        const isSunday=new Date(row.date+"T00:00:00").getDay()===0;
                        const currentVal=quickHrs[row.date]!==undefined?quickHrs[row.date]:(e?String(e.hours):"");
                        return (
                          <td key={row.date} className="px-0.5 py-1" style={{border:"1px solid #fecaca",background:isSunday?"#fefce8":"white"}}>
                            <input
                              ref={function(el){ if(el) hrsInputRefs.current[row.date]=el; }}
                              type="number" min="0" step="0.5" placeholder="—"
                              value={currentVal}
                              onChange={function(ev){
                                const v=ev.target.value;
                                setQuickHrs(function(q){ const n=Object.assign({},q); n[row.date]=v; return n; });
                                saveQuickHrs(row.date,v);
                              }}
                              onKeyDown={function(ev){handleQuickHrsKeyDown(ev,row.date,allDates);}}
                              onPaste={function(ev){handleHrsPaste(ev,row.date,allDates);}}
                              className="w-full text-center font-bold bg-transparent focus:outline-none focus:ring-2 focus:ring-red-400 rounded"
                              style={{color:e?"#7f1d1d":"#94a3b8",fontSize:"13px",padding:"6px 0px"}}
                            />
                          </td>
                        );
                      })}
                      <td className="px-1 py-1.5 font-black text-red-900 text-center sticky right-0 bg-red-50" style={{fontSize:"10px"}}>{summary.total.toFixed(1)}</td>
                    </tr>
                    <tr>
                      <td className="px-1 py-1.5 font-black text-red-900 sticky left-0 z-10 bg-red-50" style={{fontSize:"10px"}}>Site</td>
                      {dateRows.map(function(row){
                        const e=row.entries[0];
                        const isSunday=new Date(row.date+"T00:00:00").getDay()===0;
                        return (
                          <td key={row.date} onClick={function(){
                              setDayPopup({date:row.date,hours:e?String(e.hours):"",site:(e&&e.site)||"",doNo:(e&&e.doNo)||"",note:(e&&e.note)||"",logId:(e&&e.id)||null});
                            }}
                            className="px-1 py-1.5 text-center cursor-pointer hover:bg-red-50 text-slate-500 overflow-hidden"
                            style={{border:"1px solid #fecaca",background:isSunday?"#fefce8":"white",whiteSpace:"nowrap",fontSize:"9px",maxWidth:"38px",textOverflow:"ellipsis"}}>
                            {(e&&e.site)||"—"}
                          </td>
                        );
                      })}
                      <td className="px-1 py-1.5 bg-red-50 sticky right-0"></td>
                    </tr>
                    <tr>
                      <td className="px-1 py-1.5 font-black text-red-900 sticky left-0 z-10 bg-red-50" style={{fontSize:"10px"}}>Do</td>
                      {dateRows.map(function(row){
                        const e=row.entries[0];
                        const isSunday=new Date(row.date+"T00:00:00").getDay()===0;
                        return (
                          <td key={row.date} onClick={function(){
                              setDayPopup({date:row.date,hours:e?String(e.hours):"",site:(e&&e.site)||"",doNo:(e&&e.doNo)||"",note:(e&&e.note)||"",logId:(e&&e.id)||null});
                            }}
                            className="px-1 py-1.5 text-center cursor-pointer hover:bg-red-50 font-bold overflow-hidden"
                            style={{border:"1px solid #fecaca",background:isSunday?"#fefce8":"white",color:"#991b1b",whiteSpace:"nowrap",fontSize:"9px",maxWidth:"38px",textOverflow:"ellipsis"}}>
                            {(e&&e.doNo)||"—"}
                          </td>
                        );
                      })}
                      <td className="px-1 py-1.5 bg-red-50 sticky right-0"></td>
                    </tr>
                    <tr>
                      <td className="px-1 py-1.5 font-black text-red-900 sticky left-0 z-10 bg-red-50" style={{fontSize:"10px"}}>Note</td>
                      {dateRows.map(function(row){
                        const e=row.entries[0];
                        const isSunday=new Date(row.date+"T00:00:00").getDay()===0;
                        return (
                          <td key={row.date} onClick={function(){
                              setDayPopup({date:row.date,hours:e?String(e.hours):"",site:(e&&e.site)||"",doNo:(e&&e.doNo)||"",note:(e&&e.note)||"",logId:(e&&e.id)||null});
                            }}
                            className="px-1 py-1.5 text-center cursor-pointer hover:bg-red-50 text-slate-400 overflow-hidden"
                            style={{border:"1px solid #fecaca",background:isSunday?"#fefce8":"white",whiteSpace:"nowrap",fontSize:"9px",maxWidth:"38px",textOverflow:"ellipsis"}}>
                            {(e&&e.note)||"—"}
                          </td>
                        );
                      })}
                      <td className="px-1 py-1.5 bg-red-50 sticky right-0"></td>
                    </tr>
                    <tr>
                      <td className="px-1 py-1.5 font-black text-red-900 sticky left-0 z-10 bg-red-50" style={{fontSize:"10px"}}>Act</td>
                      {dateRows.map(function(row){
                        const e=row.entries[0];
                        const isSunday=new Date(row.date+"T00:00:00").getDay()===0;
                        return (
                          <td key={row.date} className="px-1 py-1.5 text-center" style={{border:"1px solid #fecaca",background:isSunday?"#fefce8":"white"}}>
                            <div className="flex gap-1 justify-center">
                              <button onClick={function(){
                                  setDayPopup({date:row.date,hours:e?String(e.hours):"",site:(e&&e.site)||"",doNo:(e&&e.doNo)||"",note:(e&&e.note)||"",logId:(e&&e.id)||null});
                                }}
                                className="text-red-500 hover:text-red-700" style={{fontSize:"11px"}}>✏️</button>
                              {e && (
                                <button onClick={function(){delLog(e.id);}} className="text-red-500 hover:text-red-700" style={{fontSize:"11px"}}>🗑</button>
                              )}
                            </div>
                          </td>
                        );
                      })}
                      <td className="px-1 py-1.5 bg-red-50 sticky right-0"></td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <p className="text-xs text-gray-400 px-4 pb-2">👉 Swipe left/right to see all 31 days</p>
              <div className="px-4 py-2 flex gap-3 text-xs border-t border-red-100">
                <span className="text-green-600 font-bold">🟢 {summary.work.toFixed(1)}h Work</span>
                <span className="text-yellow-600 font-bold">🟡 {summary.idle.toFixed(1)}h Idle</span>
                <span className="text-red-600 font-bold">🔴 {summary.repair.toFixed(1)}h Repair</span>
              </div>
            </div>

            <button onClick={function(){delMachine(sel.id);}} className="w-full text-red-500 text-sm py-2 hover:text-red-700 transition">🗑 Delete Machine</button>
          </div>
        </div>

        {editMachine && (
          <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
            <div className="rounded-2xl p-5 max-w-xs w-full shadow-xl bg-white">
              <h3 className="font-black text-red-900 text-lg mb-4">✏️ Edit Machine</h3>
              <div className="space-y-3">
                <div><label className="text-xs font-black text-red-600">Machine Name *</label><input className={inp} value={editMachine.name} onChange={function(e){setEditMachine(function(m){return Object.assign({},m,{name:e.target.value});});}}/></div>
                <div>
                  <label className="text-xs font-black text-red-600">Type</label>
                  <select className={inp} value={editMachine.type} onChange={function(e){setEditMachine(function(m){return Object.assign({},m,{type:e.target.value});});}}>
                    {MACHINE_TYPES.map(function(t){return <option key={t}>{t}</option>;})}
                  </select>
                </div>
                <div><label className="text-xs font-black text-red-600">Model</label><input className={inp} value={editMachine.model} onChange={function(e){setEditMachine(function(m){return Object.assign({},m,{model:e.target.value});});}}/></div>
                <div><label className="text-xs font-black text-red-600">Date Made</label><input type="date" className={inp} value={editMachine.dateMade} onChange={function(e){setEditMachine(function(m){return Object.assign({},m,{dateMade:e.target.value});});}}/></div>
                <div className="border-t border-red-100 pt-3">
                  <p className="text-xs font-black text-red-900 uppercase mb-2">Operator</p>
                  {settings.operators.length>0?(
                    <select className={inp} value={editMachine.operator} onChange={function(e){setEditMachine(function(m){return Object.assign({},m,{operator:e.target.value});});}}>
                      <option value="">— Select —</option>
                      {settings.operators.map(function(op){return <option key={op}>{op}</option>;})}
                    </select>
                  ):(
                    <input className={inp} value={editMachine.operator} onChange={function(e){setEditMachine(function(m){return Object.assign({},m,{operator:e.target.value});});}}/>
                  )}
                  <div className="mt-2"><label className="text-xs font-black text-red-600">Start Working Date</label>
                    <input type="date" className={inp} value={editMachine.operatorStart} onChange={function(e){setEditMachine(function(m){return Object.assign({},m,{operatorStart:e.target.value});});}}/>
                  </div>
                </div>
              </div>
              <div className="flex gap-2 mt-4">
                <button onClick={function(){setEditMachine(null);}} className="flex-1 border border-gray-300 rounded-xl py-2 font-bold text-gray-500 text-sm">Cancel</button>
                <button onClick={saveEditMachine} className="flex-1 text-white rounded-xl py-2 font-bold text-sm" style={{background:"linear-gradient(90deg,#dc2626,#b91c1c)"}}>Save</button>
              </div>
            </div>
          </div>
        )}

        {dayPopup && (
          <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
            <div className="rounded-2xl p-5 max-w-xs w-full shadow-xl bg-white">
              <h3 className="font-black text-red-900 text-lg mb-1">
                📅 {dayPopup.date.slice(8)}-{dayPopup.date.slice(5,7)}-{dayPopup.date.slice(0,4)}
              </h3>
              <p className="text-xs text-gray-400 mb-3">Changes save automatically</p>
              <div className="space-y-2 mb-4">
                <div>
                  <label className="text-xs text-red-600 font-black">Hours</label>
                  <input type="number" min="0" step="0.5" className={inp} placeholder="e.g. 10"
                    value={dayPopup.hours}
                    onChange={function(e){
                      const v=e.target.value;
                      setDayPopup(function(p){return Object.assign({},p,{hours:v});});
                      setQuickHrs(function(q){ const n=Object.assign({},q); n[dayPopup.date]=v; return n; });
                      saveQuickHrs(dayPopup.date,v);
                    }}/>
                </div>
                <div>
                  <label className="text-xs text-red-600 font-black">Project Site</label>
                  {settings.sites.length>0?(
                    <select className={inp} value={dayPopup.site} onChange={function(e){
                        const v=e.target.value;
                        setDayPopup(function(p){return Object.assign({},p,{site:v});});
                        upd(function(p){return p.map(function(m){
                          if(m.id!==selId) return m;
                          return Object.assign({},m,{hourLogs:m.hourLogs.map(function(l){return l.date===dayPopup.date?Object.assign({},l,{site:v}):l;})});
                        });});
                      }}>
                      <option value="">— Site —</option>
                      {settings.sites.map(function(s){return <option key={s}>{s}</option>;})}
                    </select>
                  ):(
                    <input className={inp} placeholder="Project site" value={dayPopup.site} onChange={function(e){
                        const v=e.target.value;
                        setDayPopup(function(p){return Object.assign({},p,{site:v});});
                        upd(function(p){return p.map(function(m){
                          if(m.id!==selId) return m;
                          return Object.assign({},m,{hourLogs:m.hourLogs.map(function(l){return l.date===dayPopup.date?Object.assign({},l,{site:v}):l;})});
                        });});
                      }}/>
                  )}
                </div>
                <div>
                  <label className="text-xs text-red-600 font-black">Do No.</label>
                  <input className={inp} placeholder="Do No." value={dayPopup.doNo} onChange={function(e){
                      const v=e.target.value;
                      setDayPopup(function(p){return Object.assign({},p,{doNo:v});});
                      upd(function(p){return p.map(function(m){
                        if(m.id!==selId) return m;
                        return Object.assign({},m,{hourLogs:m.hourLogs.map(function(l){return l.date===dayPopup.date?Object.assign({},l,{doNo:v}):l;})});
                      });});
                    }}/>
                </div>
                <div>
                  <label className="text-xs text-red-600 font-black">Note</label>
                  <input className={inp} placeholder="Note (optional)" value={dayPopup.note} onChange={function(e){
                      const v=e.target.value;
                      setDayPopup(function(p){return Object.assign({},p,{note:v});});
                      upd(function(p){return p.map(function(m){
                        if(m.id!==selId) return m;
                        return Object.assign({},m,{hourLogs:m.hourLogs.map(function(l){return l.date===dayPopup.date?Object.assign({},l,{note:v}):l;})});
                      });});
                    }}/>
                </div>
              </div>
              <div className="flex gap-2">
                <button onClick={function(){setDayPopup(null);}} className="flex-1 border border-gray-300 rounded-xl py-2 font-bold text-gray-500 text-sm">Close</button>
                {sel.hourLogs.some(function(l){return l.date===dayPopup.date;}) && (
                  <button onClick={deleteDayPopupLog} className="flex-1 border border-red-400 text-red-500 rounded-xl py-2 font-bold text-sm hover:bg-red-50">🗑 Delete</button>
                )}
              </div>
            </div>
          </div>
        )}

        {editLog && (
          <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
            <div className="rounded-2xl p-5 max-w-xs w-full shadow-xl bg-white">
              <h3 className="font-black text-red-900 text-lg mb-3">✏️ Edit Log Entry</h3>
              <div className="space-y-2 mb-3">
                <div><label className="text-xs text-red-600">Date</label><input type="date" className={inp} value={editLog.date} onChange={function(e){setEditLog(function(l){return Object.assign({},l,{date:e.target.value});});}}/></div>
                <div><label className="text-xs text-red-600">Hours</label><input type="number" min="0" step="0.5" className={inp} value={editLog.hours} onChange={function(e){setEditLog(function(l){return Object.assign({},l,{hours:e.target.value});});}}/></div>
                <div>
                  <label className="text-xs text-red-600">Project Site</label>
                  {settings.sites.length>0?(
                    <select className={inp} value={editLog.site||""} onChange={function(e){setEditLog(function(l){return Object.assign({},l,{site:e.target.value});});}}>
                      <option value="">— Site —</option>
                      {settings.sites.map(function(s){return <option key={s}>{s}</option>;})}
                    </select>
                  ):(
                    <input className={inp} placeholder="Project site" value={editLog.site||""} onChange={function(e){setEditLog(function(l){return Object.assign({},l,{site:e.target.value});});}}/>
                  )}
                </div>
                <div><label className="text-xs text-red-600">Do No.</label><input className={inp} placeholder="Do No." value={editLog.doNo||""} onChange={function(e){setEditLog(function(l){return Object.assign({},l,{doNo:e.target.value});});}}/></div>
                <div><label className="text-xs text-red-600">Note</label><input className={inp} placeholder="Note (optional)" value={editLog.note} onChange={function(e){setEditLog(function(l){return Object.assign({},l,{note:e.target.value});});}}/></div>
              </div>
              <div className="flex gap-2">
                <button onClick={function(){setEditLog(null);}} className="flex-1 border border-gray-300 rounded-xl py-2 font-bold text-gray-500 text-sm">Cancel</button>
                <button onClick={saveEdit} className="flex-1 text-white rounded-xl py-2 font-bold text-sm" style={{background:"linear-gradient(90deg,#dc2626,#b91c1c)"}}>Save</button>
              </div>
            </div>
          </div>
        )}

        {showSvc && (
          <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
            <div className="rounded-2xl p-5 max-w-xs w-full shadow-xl text-center bg-white">
              <div className="text-3xl mb-2 text-center">🔧</div>
              <h3 className="font-bold text-red-900 text-lg mb-1 text-center">Confirm Service #{svcNo}</h3>
              <p className="text-gray-500 text-sm mb-3 text-center">
                Hours: <strong className="text-red-900">{Math.min(sinceSvc,SERVICE_INTERVAL).toFixed(1)} hrs</strong>
                {sinceSvc>SERVICE_INTERVAL && <span className="text-red-500 text-xs"> (capped at 400)</span>}
              </p>
              <div className="space-y-3 mb-4">
                <div>
                  <label className="text-xs font-black text-red-600">📍 Project Site</label>
                  {settings.sites.length>0?(
                    <select className={inp} value={svcSite} onChange={function(e){setSvcSite(e.target.value);}}>
                      <option value="">— Select Site —</option>
                      {settings.sites.map(function(s){return <option key={s}>{s}</option>;})}
                    </select>
                  ):(
                    <input className={inp} placeholder="Project site (optional)" value={svcSite} onChange={function(e){setSvcSite(e.target.value);}}/>
                  )}
                </div>
                <div>
                  <label className="text-xs font-black text-red-600">Do No.</label>
                  <input className={inp} placeholder="Delivery Order No. (optional)" value={svcDoNo} onChange={function(e){setSvcDoNo(e.target.value);}}/>
                </div>
              </div>
              <div className="flex gap-2">
                <button onClick={function(){setShowSvc(false);setSvcSite("");setSvcDoNo("");}} className="flex-1 border border-gray-300 rounded-xl py-2 font-bold text-gray-500 text-sm">Cancel</button>
                <button onClick={markServiced} className="flex-1 text-white rounded-xl py-2 font-bold text-sm" style={{background:"linear-gradient(90deg,#dc2626,#b91c1c)"}}>Confirm</button>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  }

  // ── DASHBOARD ──
  if(!loaded){
    return (
      <div className="min-h-screen flex items-center justify-center" style={{background:"#fff5f5"}}>
        <div className="text-center">
          <div className="text-5xl mb-3 animate-pulse">🔩</div>
          <p className="text-red-700 font-bold">Loading shared data...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen" style={{background:"#fff5f5"}}>
      {syncing && (
        <div className="fixed top-0 left-0 right-0 z-50 flex items-center justify-center gap-2 py-2 text-white text-sm font-bold"
          style={{background:"linear-gradient(90deg,#16a34a,#15803d)",boxShadow:"0 2px 10px rgba(0,0,0,0.2)"}}>
          <span className="inline-block w-2 h-2 rounded-full bg-white" style={{animation:"pulse 1s infinite"}}/>
          ● Syncing data...
        </div>
      )}
      {!syncing && syncError && (
        <div className="px-4 py-2 text-white text-xs font-bold text-center" style={{background:"#b91c1c"}}>
          {syncError}
        </div>
      )}
      {usingLocalBackup && (
        <div className="px-4 py-1.5 text-white text-xs font-bold text-center" style={{background:"#ca8a04"}}>
          📴 Offline mode — showing locally saved data, not the shared database.
        </div>
      )}
      <div className="bg-white border-b-2 border-red-500" style={{boxShadow:"0 2px 20px rgba(220,38,38,0.15)"}}>
        <div className="p-4 max-w-md mx-auto">
          <div className="flex justify-between items-center">
            <div>
              <h1 className="font-black text-red-900 leading-none" style={{fontSize:"26px",letterSpacing:"-0.5px"}}>
                🔩 <span style={{color:"#dc2626"}}>Service</span> Machine
              </h1>
              <p className="text-xs font-bold mt-0.5 text-red-400">
                Service interval: <span style={{color:"#dc2626",fontWeight:900}}>{SERVICE_INTERVAL} hrs</span>
              </p>
            </div>
            <button onClick={function(){setView("add");}}
              className="font-black text-white text-sm px-5 py-2.5 rounded-full transition active:scale-95"
              style={{background:"linear-gradient(135deg,#dc2626,#b91c1c)",boxShadow:"0 0 16px rgba(220,38,38,0.4)"}}>
              + Add
            </button>
          </div>
          <div className="flex gap-2 mt-3">
            {[["⚙️","Settings","settings"],["📥","Import","import"],["🖨️","Print","print"]].map(function(item){
              const ic=item[0], lb=item[1], vw=item[2];
              return (
                <button key={vw} onClick={function(){setView(vw);}}
                  className="flex-1 font-bold text-red-700 text-xs py-2 rounded-xl transition active:scale-95 bg-red-50 border border-red-200">
                  {ic} {lb}
                </button>
              );
            })}
            <button onClick={exportFullBackup}
              className="flex-1 font-bold text-red-700 text-xs py-2 rounded-xl transition active:scale-95 bg-red-50 border border-red-200">
              💾 Backup
            </button>
            <button onClick={function(){setShowSearch(function(s){return !s;});setSearch("");}}
              className="font-bold text-xs px-3 py-2 rounded-xl transition active:scale-95"
              style={showSearch?{background:"#dc2626",color:"white",border:"1px solid #dc2626"}:{background:"#fef2f2",color:"#dc2626",border:"1px solid #fecaca"}}>
              🔍
            </button>
          </div>
          {showSearch && (
            <div className="mt-2">
              <input autoFocus
                className="w-full rounded-xl px-4 py-2 text-sm font-bold focus:outline-none border border-red-300"
                placeholder="Search machine, operator, model..."
                value={search} onChange={function(e){setSearch(e.target.value);}}/>
              {search && <p className="text-xs font-bold mt-1 ml-1 text-red-600">{filtered.length} result(s) found</p>}
            </div>
          )}
        </div>
      </div>

      <div className="p-4 max-w-md mx-auto">
        {machines.length>0 && (
          <div className="grid grid-cols-4 gap-2 mb-4 mt-1">
            {[["Total",machines.length,"#dc2626"],
              ["Work",machines.filter(function(m){return m.condition==="Work";}).length,"#16a34a"],
              ["Idle",machines.filter(function(m){return m.condition==="Idle";}).length,"#ca8a04"],
              ["Repair",machines.filter(function(m){return m.condition==="Repair";}).length,"#7f1d1d"]
            ].map(function(item){
              const l=item[0], v=item[1], col=item[2];
              return (
                <div key={l} className="rounded-2xl p-3 text-center bg-white" style={{border:"1.5px solid "+col+"40"}}>
                  <p className="text-2xl font-black" style={{color:col}}>{v}</p>
                  <p className="text-xs font-black text-gray-500">{l}</p>
                </div>
              );
            })}
          </div>
        )}
        {machines.length>0 && (
          <div>
            <div className="flex gap-2 overflow-x-auto pb-1 mb-2">
              {FILTER_STATUS.map(function(s){
                return (
                  <button key={s} onClick={function(){setFStatus(s);}}
                    className="whitespace-nowrap px-3 py-1.5 rounded-full text-xs font-black transition"
                    style={fStatus===s?{background:"#dc2626",color:"white"}:{background:"#fef2f2",color:"#64748b",border:"1px solid #fecaca"}}>
                    {s==="Work"?"🟢 ":s==="Idle"?"🟡 ":s==="Repair"?"🔴 ":"⚙️ "}{s}
                  </button>
                );
              })}
            </div>
            <div className="flex gap-2 overflow-x-auto pb-2 mb-3">
              {FILTER_TYPES.map(function(t){
                return (
                  <button key={t} onClick={function(){setFType(t);}}
                    className="whitespace-nowrap px-3 py-1.5 rounded-full text-xs font-black transition"
                    style={fType===t?{background:"#dc2626",color:"white"}:{background:"#fef2f2",color:"#64748b",border:"1px solid #fecaca"}}>
                    {t}
                  </button>
                );
              })}
            </div>
          </div>
        )}
        {machines.length===0?(
          <div className="text-center py-20">
            <div className="text-6xl mb-4">🏗️</div>
            <p className="font-black text-red-900 text-xl mb-1">No machines yet</p>
            <p className="text-sm font-bold text-red-500">Tap + Add to get started</p>
          </div>
        ):filtered.length===0?(
          <div className="text-center py-16">
            <div className="text-5xl mb-3">🔍</div>
            <p className="font-black text-red-900 text-lg">No machine found</p>
          </div>
        ):(
          <div className="space-y-3">
            {filtered.map(function(m){
              const tot=totalHrs(m), msvc=sinceServiceHrs(m);
              const s=SS[m.condition];
              const pct=Math.min(100,(msvc/SERVICE_INTERVAL)*100);
              const barCol=pct>=100?"#7f1d1d":pct>=75?"#ca8a04":"#dc2626";
              return (
                <div key={m.id} onClick={function(){setSelId(m.id);setView("detail");}}
                  className="rounded-2xl cursor-pointer transition overflow-hidden bg-white"
                  style={{border:"1px solid #fecaca",boxShadow:"0 2px 12px rgba(220,38,38,0.08)"}}>
                  <div style={{height:"3px",background:"linear-gradient(90deg,#dc2626,transparent)"}}/>
                  <div className="p-4">
                    <div className="flex justify-between items-start mb-2">
                      <div className="flex-1 min-w-0 mr-2">
                        <h3 className="font-black text-red-900 text-lg leading-tight truncate">{m.name}</h3>
                        <div className="flex items-center gap-1.5 mt-1.5 flex-wrap">
                          <span className="text-xs font-black rounded-full px-2.5 py-0.5" style={{background:"#fee2e2",color:"#dc2626",border:"1px solid #fca5a5"}}>{m.type}</span>
                          {m.model && <span className="text-xs font-bold rounded-full px-2 py-0.5" style={{background:"#fff7ed",color:"#c2410c",border:"1px solid #fed7aa"}}>🔩 {m.model}</span>}
                          <span className={"text-xs font-black px-2 py-0.5 rounded-full border "+s.pill}>
                            <span className={"inline-block w-1.5 h-1.5 rounded-full mr-1 "+s.dot} style={{verticalAlign:"middle"}}/>
                            {m.condition}
                          </span>
                        </div>
                        {m.operator && <p className="text-xs font-bold mt-1.5 text-gray-500">👷 {m.operator}</p>}
                      </div>
                      <ServiceBadge hours={msvc}/>
                    </div>
                    <div className="flex justify-between items-center text-xs mt-2 mb-2">
                      <span className="text-gray-500">Since svc: <strong className="text-red-900">{msvc.toFixed(1)}</strong><span className="text-gray-400"> / {SERVICE_INTERVAL} hrs</span></span>
                      <span className="font-black" style={{color:"#dc2626"}}>Total: {tot.toFixed(1)} hrs</span>
                    </div>
                    <div className="rounded-full overflow-hidden bg-red-100" style={{height:"6px"}}>
                      <div className="h-full rounded-full transition-all" style={{width:pct+"%",background:"linear-gradient(90deg,"+barCol+","+barCol+"cc)"}}/>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}
