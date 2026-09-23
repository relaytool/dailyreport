const $ = id => document.getElementById(id);
const defaultAssets = ["Black Pallets","Wooden Sleeves","Wood Pallets","THG Pallets","Magnum Lids","Red Sleeves","Magnum","Car Bags"];
let assets = [...defaultAssets];
let importedTransactions = [];
let importedInventory = {};
let uploadedWorkbook = false;
let discardedReasons = {};

const londonFormatter = new Intl.DateTimeFormat("en-GB", {
  timeZone: "Europe/London", year: "numeric", month: "2-digit", day: "2-digit",
  hour: "2-digit", minute: "2-digit", hourCycle: "h23"
});
const londonDateFormatter = new Intl.DateTimeFormat("en-CA", {
  timeZone: "Europe/London", year: "numeric", month: "2-digit", day: "2-digit"
});
const now = new Date();
const londonDate = londonDateFormatter.format(now);
$("reportDate").value = londonDate;
$("signoffDate").value = new Date(now.getTime() - now.getTimezoneOffset() * 60000).toISOString().slice(0,16);

function normaliseHeader(value){
  return String(value ?? "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
}
function findHeader(headers, names){
  const wanted = names.map(normaliseHeader);
  return headers.findIndex(h => wanted.includes(normaliseHeader(h)));
}
function londonParts(value){
  const d = value instanceof Date ? value : new Date(value);
  if(Number.isNaN(d.getTime())) return null;
  const parts = londonFormatter.formatToParts(d);
  const get = type => parts.find(p => p.type === type)?.value || "";
  return {date:`${get("year")}-${get("month")}-${get("day")}`, hour:Number(get("hour")), minute:Number(get("minute"))};
}
function localDateKey(value){ return londonParts(value)?.date || ""; }
function formatImportedTime(value){
  const p = londonParts(value);
  return p ? `${String(p.hour).padStart(2,"0")}:${String(p.minute).padStart(2,"0")}` : String(value ?? "");
}
function formatDate(v){
  if(!v) return "—";
  const [y,m,d] = v.split("-");
  return `${d}/${m}/${y}`;
}
function formatDateTime(v){
  if(!v) return "—";
  const [date,time] = v.split("T");
  return `${formatDate(date)} • ${time}`;
}
function text(id,fallback="—"){ return ($(id)?.value || "").trim() || fallback; }
function escapeHtml(value){
  return String(value ?? "").replace(/[&<>'"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;","\"":"&quot;"}[c]));
}
function nextDateKey(date){
  const d = new Date(`${date}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate()+1);
  return d.toISOString().slice(0,10);
}
function shiftInfo(){
  const date = $("reportDate").value;
  const shift = $("shiftSelect").value;
  if(shift === "night") return {key:shift, label:"Night Shift", window:"19:00 – 07:00", start:"19:00", end:"07:00", nextDate:nextDateKey(date)};
  return {key:shift, label:"Day Shift", window:"07:00 – 19:00", start:"07:00", end:"19:00", nextDate:date};
}
function updateShiftWindow(){
  const info = shiftInfo();
  $("shiftWindow").textContent = `Selected shift: ${info.label} • ${info.window}${info.key === "night" ? ` • ends ${formatDate(info.nextDate)}` : ""}`;
}
updateShiftWindow();

function isInSelectedShift(timestamp){
  const p = londonParts(timestamp);
  if(!p) return false;
  const selected = $("reportDate").value;
  const info = shiftInfo();
  const mins = p.hour * 60 + p.minute;
  if(info.key === "day") return p.date === selected && mins >= 420 && mins < 1140;
  return (p.date === selected && mins >= 1140) || (p.date === info.nextDate && mins < 420);
}
function shiftRows(){ return importedTransactions.filter(r => isInSelectedShift(r.timestamp)); }

function todayDiscardedRows(){
  const selectedDate=$("reportDate").value;
  return importedTransactions.filter(r =>
    localDateKey(r.timestamp)===selectedDate &&
    ["DISCARDED","DISCARD"].includes(r.movement)
  );
}

function discardedReasonKey(date,asset){ return `${date}::${asset}`; }

function allTimeClientBalances(){
  const balances=new Map();
  importedTransactions.forEach(r=>{
    const client=r.client||"Unknown Client";
    const asset=r.asset||"Unknown Asset";
    if(!["SENT","RECEIVED","RETURNED","COLLECTED"].includes(r.movement)) return;
    if(!balances.has(client)) balances.set(client,new Map());
    const m=balances.get(client);
    const current=m.get(asset)||0;
    if(r.movement==="SENT") m.set(asset,current+r.quantity);
    else m.set(asset,current-r.quantity);
  });
  return balances;
}

function buildInventoryForm(){
  $("inventoryForm").innerHTML = assets.map((a,i)=>`
    <div class="inventory-item">
      <strong>${escapeHtml(a)}</strong>
      <div class="inv-grid">
        <label><small>Opening</small><input type="number" min="0" id="open_${i}" readonly></label>
        <label><small>Dispatched</small><input type="number" min="0" id="dispatch_${i}" readonly></label>
        <label><small>Closing</small><input type="number" min="0" id="close_${i}" readonly></label>
        <label><small>Discarded</small><input type="number" min="0" id="discard_${i}" readonly></label>
      </div>
    </div>`).join("");
}
buildInventoryForm();

function parseWorkbook(file){
  if(!file) return;
  const status=$("uploadStatus");
  status.textContent="Reading Asset Transactions and Asset Inventory...";
  status.className="upload-status";
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      if(typeof XLSX === "undefined") throw new Error("Spreadsheet reader did not load. Check your internet connection and try again.");
      const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array",cellDates:true});
      const txSheetName=wb.SheetNames.find(n=>normaliseHeader(n)==="assettransactions");
      const invSheetName=wb.SheetNames.find(n=>normaliseHeader(n)==="assetinventory");
      if(!txSheetName || !invSheetName) throw new Error(`Required sheets not found. Found: ${wb.SheetNames.join(", ")}`);
      parseTransactions(wb.Sheets[txSheetName]);
      parseInventory(wb.Sheets[invSheetName]);
      uploadedWorkbook=true;
      buildInventoryForm();
      applyInventoryForShift();
      renderInventorySummaryPreview();
      renderClientMovementPreview();
      renderDiscardedAssetsTodayPreview();
      updateSummary();
      updateUploadStatus();
    }catch(err){
      console.error(err);
      uploadedWorkbook=false; importedTransactions=[]; importedInventory={}; discardedReasons={}; assets=[...defaultAssets]; buildInventoryForm();
      $("clientMovementPreview").innerHTML='<div class="upload-status">Upload the workbook to see this shift.</div>';
      $("discardedAssetsPreview").innerHTML='<div class="upload-status">Upload the workbook to see discarded assets today.</div>';
      status.textContent=err.message || "Could not read the spreadsheet.";
      status.className="upload-status error";
    }
  };
  reader.onerror=()=>{status.textContent="Could not read the selected spreadsheet.";status.className="upload-status error"};
  reader.readAsArrayBuffer(file);
}

function parseTransactions(sheet){
  const rows=XLSX.utils.sheet_to_json(sheet,{header:1,defval:"",raw:true});
  const headerRowIndex=rows.findIndex(r=>r.some(v=>normaliseHeader(v)==="timestamp"));
  if(headerRowIndex<0) throw new Error('Could not find the header row in "Asset Transactions".');
  const headers=rows[headerRowIndex].map(String);
  const idx={
    timestamp:findHeader(headers,["timestamp","date","datetime"]),
    client:findHeader(headers,["client"]),
    movement:findHeader(headers,["movement","transaction","type"]),
    asset:findHeader(headers,["asset","asset type","assettype"]),
    quantity:findHeader(headers,["quantity","qty","count"]),
    user:findHeader(headers,["user","verified by","verifiedby"])
  };
  if([idx.timestamp,idx.client,idx.movement,idx.asset,idx.quantity].some(i=>i<0)) throw new Error('"Asset Transactions" needs Timestamp, Client, Movement, Asset and Quantity columns.');
  importedTransactions=rows.slice(headerRowIndex+1).map(r=>({
    timestamp:r[idx.timestamp], client:String(r[idx.client]??"").trim(), movement:String(r[idx.movement]??"").trim().toUpperCase(),
    asset:String(r[idx.asset]??"").trim(), quantity:Number(r[idx.quantity]||0), user:idx.user>=0?String(r[idx.user]??"").trim():""
  })).filter(r=>r.timestamp!=="" && r.asset!=="" && Number.isFinite(r.quantity));
}

function parseInventory(sheet){
  const rows=XLSX.utils.sheet_to_json(sheet,{header:1,defval:"",raw:true});
  const headerIndex=rows.findIndex(r=>r.some(v=>normaliseHeader(v)==="asset"));
  if(headerIndex<0) throw new Error('Could not find the header row in "Asset Inventory".');
  const headers=rows[headerIndex].map(String);
  const assetIdx=findHeader(headers,["asset"]), balanceIdx=findHeader(headers,["balance","quantity","qty","count"]);
  if(assetIdx<0||balanceIdx<0) throw new Error('"Asset Inventory" needs Asset and Balance columns.');
  importedInventory={};
  rows.slice(headerIndex+1).forEach(r=>{
    const asset=String(r[assetIdx]??"").trim();
    const balance=Number(r[balanceIdx]??0);
    if(asset) importedInventory[asset]=Number.isFinite(balance)?balance:0;
  });
  const sourceAssets=Object.keys(importedInventory);
  if(sourceAssets.length) assets=sourceAssets;
}

function movementData(rows){
  const received=rows.filter(r=>r.movement==="RECEIVED");
  const sent=rows.filter(r=>r.movement==="SENT");
  const discarded=rows.filter(r=>["DISCARDED","DISCARD"].includes(r.movement));
  return {received,sent,discarded};
}
function aggregateByClient(rows){
  const clients=new Map();
  rows.forEach(r=>{
    const client=r.client||"Unknown Client";
    const asset=r.asset||"Unknown Asset";
    if(!clients.has(client)) clients.set(client,new Map());
    const m=clients.get(client);
    m.set(asset,(m.get(asset)||0)+r.quantity);
  });
  return clients;
}
function combinedClientData(rows){
  const {received,sent}=movementData(rows);
  const receivedMap=aggregateByClient(received), sentMap=aggregateByClient(sent);
  const allTimeMap=allTimeClientBalances();
  const names=new Set([...receivedMap.keys(),...sentMap.keys()]);
  return [...names].sort((a,b)=>a.localeCompare(b)).map(client=>{
    const r=receivedMap.get(client)||new Map(), s=sentMap.get(client)||new Map();
    const allTime=allTimeMap.get(client)||new Map();
    const typeSet=new Set([...r.keys(),...s.keys()]);
    const types=[...typeSet].sort((a,b)=>a.localeCompare(b)).map(asset=>({
      asset,
      received:r.get(asset)||0,
      sent:s.get(asset)||0,
      withClient:(s.get(asset)||0)-(r.get(asset)||0),
      allTimeWithClient:allTime.get(asset)||0
    }));
    return {client,types};
  });
}

function renderDiscardedAssetsTodayPreview(){
  const box=$("discardedAssetsPreview");
  if(!uploadedWorkbook){
    box.innerHTML='<div class="upload-status">Upload the workbook to see discarded assets today.</div>';
    return;
  }
  const data=assetTotals(todayDiscardedRows());
  if(!data.length){
    box.innerHTML='<div class="upload-status">No assets were discarded today.</div>';
    return;
  }
  const date=$("reportDate").value;
  box.innerHTML=`<div class="discarded-today-head"><span>Asset type</span><span>Count</span><span>Reason</span></div>
    <div class="discarded-today-list">${data.map(x=>{
      const key=discardedReasonKey(date,x.asset);
      return `<div class="discarded-today-row">
        <strong>${escapeHtml(x.asset)}</strong>
        <b>${x.quantity}</b>
        <input type="text" data-discard-reason data-key="${escapeHtml(key)}" value="${escapeHtml(discardedReasons[key]||"")}" placeholder="Enter reason">
      </div>`;
    }).join("")}</div>`;
  box.querySelectorAll("input[data-discard-reason]").forEach(input=>{
    input.addEventListener("input",()=>{ discardedReasons[input.dataset.key]=input.value; });
  });
}
function assetTotals(rows){
  const map=new Map();
  rows.forEach(r=>map.set(r.asset,(map.get(r.asset)||0)+r.quantity));
  return [...map.entries()].map(([asset,quantity])=>({asset,quantity})).sort((a,b)=>a.asset.localeCompare(b.asset));
}
function inventorySnapshotData(){
  const rows=shiftRows();
  const {received,sent}=movementData(rows);
  const recMap=new Map(), sentMap=new Map();
  received.forEach(r=>recMap.set(r.asset,(recMap.get(r.asset)||0)+r.quantity));
  sent.forEach(r=>sentMap.set(r.asset,(sentMap.get(r.asset)||0)+r.quantity));
  const all=new Set([...assets,...recMap.keys(),...sentMap.keys()]);
  return [...all].map(asset=>({
    asset,
    received:recMap.get(asset)||0,
    sent:sentMap.get(asset)||0,
    difference:(recMap.get(asset)||0)-(sentMap.get(asset)||0)
  })).sort((a,b)=>a.asset.localeCompare(b.asset));
}
function renderInventorySummaryPreview(){
  const box=$("inventorySummaryPreview");
  if(!uploadedWorkbook){box.innerHTML='<div class="upload-status">Upload the workbook to see this shift inventory.</div>';return;}
  const rows=shiftRows();
  const info=shiftInfo();
  const data=inventorySnapshotData();
  box.innerHTML=data.map((x,i)=>{
    const idx=assets.indexOf(x.asset);
    const opening=Number($("open_"+idx)?.value||0);
    const closing=Number($("close_"+idx)?.value||0);
    return `<div class="inventory-summary-card">
      <div class="inventory-summary-head">
        <strong>${escapeHtml(x.asset)}</strong>
        <span class="shift-difference ${x.difference>=0?"positive":"negative"}">${x.difference>=0?"+":""}${x.difference}</span>
      </div>
      <div class="inventory-summary-values">
        <div class="small-balance"><span>At opening</span><b>${opening}</b></div>
        <div class="movement-value in"><span>In</span><b>${x.received}</b></div>
        <div class="movement-value out"><span>Out</span><b>${x.sent}</b></div>
        <div class="small-balance"><span>At closing</span><b>${closing}</b></div>
      </div>
    </div>`;
  }).join("");
}
function renderClientMovementPreview(){
  const box=$("clientMovementPreview");
  if(!uploadedWorkbook){box.innerHTML='<div class="upload-status">Upload the workbook to see this shift.</div>';return;}
  const data=combinedClientData(shiftRows());
  if(!data.length){box.innerHTML='<div class="upload-status">No IN or OUT transactions found for this shift.</div>';return;}
  box.innerHTML=data.map(c=>`<div class="client-preview-card">
    <div class="client-preview-head">
      <div><span class="client-label">Client</span><strong>${escapeHtml(c.client)}</strong></div>
    </div>
    <div class="client-movement-head"><span>Asset type</span><span>In</span><span>Out</span><span>With Client Today</span><span>All-Time With Client</span></div>
    <div class="client-movement-list">
      ${c.types.map(t=>`<div class="client-movement-row"><span>${escapeHtml(t.asset)}</span><b class="movement-in">${t.received}</b><b class="movement-out">${t.sent}</b><b class="with-client ${t.withClient>=0?"positive":"negative"}">${t.withClient>=0?"+":""}${t.withClient}</b><b class="with-client ${t.allTimeWithClient>=0?"positive":"negative"}">${t.allTimeWithClient>=0?"+":""}${t.allTimeWithClient}</b></div>`).join("")}
    </div>
  </div>`).join("");
}

function applyInventoryForShift(){
  if(!uploadedWorkbook) return;
  const selectedRows=shiftRows();
  const {received,sent,discarded}=movementData(selectedRows);
  const selectedEndDate=shiftInfo().key==="day" ? $("reportDate").value : shiftInfo().nextDate;
  const endHour=shiftInfo().key==="day" ? 19*60 : 7*60;
  assets.forEach((asset,i)=>{
    const closingCurrent=Number(importedInventory[asset]??0);
    let afterShiftNet=0;
    importedTransactions.forEach(r=>{
      if(r.asset!==asset) return;
      const p=londonParts(r.timestamp); if(!p) return;
      const d=p.date, mins=p.hour*60+p.minute;
      let after=false;
      if(shiftInfo().key==="day") after=(d>selectedEndDate)||(d===selectedEndDate&&mins>=endHour);
      else after=(d>selectedEndDate)||(d===selectedEndDate&&mins>=endHour);
      if(!after) return;
      if(["RECEIVED","RETURNED","COLLECTED"].includes(r.movement)) afterShiftNet+=r.quantity;
      else if(["SENT","DISCARDED","DISCARD"].includes(r.movement)) afterShiftNet-=r.quantity;
    });
    const closingForShift=closingCurrent-afterShiftNet;
    const rec=received.filter(r=>r.asset===asset).reduce((s,r)=>s+r.quantity,0);
    const dis=sent.filter(r=>r.asset===asset).reduce((s,r)=>s+r.quantity,0);
    const disc=discarded.filter(r=>r.asset===asset).reduce((s,r)=>s+r.quantity,0);
    const opening=closingForShift-rec+dis+disc;
    $("open_"+i).value=Math.max(0,opening);
    $("dispatch_"+i).value=dis;
    $("close_"+i).value=Math.max(0,closingForShift);
    $("discard_"+i).value=disc;
  });
}

function updateSummary(){
  const rows=shiftRows();
  const {received,sent}=movementData(rows);
  const receivedTotal=received.reduce((s,r)=>s+r.quantity,0);
  const sentTotal=sent.reduce((s,r)=>s+r.quantity,0);
  return {receivedTotal,sentTotal};
}

function updateUploadStatus(){
  const rows=shiftRows(), {received,sent}=movementData(rows);
  const info=shiftInfo();
  $("uploadStatus").textContent=`Loaded Asset Transactions + Asset Inventory • ${info.label}, ${formatDate($("reportDate").value)} • ${rows.length} transactions (${received.length} received, ${sent.length} sent).`;
  $("uploadStatus").className="upload-status success";
}

function renderReportInventorySummary(){
  const box=$("rInventorySummaryCards");
  box.innerHTML="";
  inventorySnapshotData().forEach(x=>{
    const idx=assets.indexOf(x.asset);
    const opening=Number($("open_"+idx)?.value||0);
    const closing=Number($("close_"+idx)?.value||0);
    const el=document.createElement("div");
    el.className="inventory-summary-card";
    el.innerHTML=`
      <div class="inventory-summary-head">
        <strong>${escapeHtml(x.asset)}</strong>
        <span class="shift-difference ${x.difference>=0?"positive":"negative"}">${x.difference>=0?"+":""}${x.difference}</span>
      </div>
      <div class="inventory-summary-values">
        <div class="small-balance"><span>At opening</span><b>${opening}</b></div>
        <div class="movement-value in"><span>In</span><b>${x.received}</b></div>
        <div class="movement-value out"><span>Out</span><b>${x.sent}</b></div>
        <div class="small-balance"><span>At closing</span><b>${closing}</b></div>
      </div>`;
    box.appendChild(el);
  });
}
function renderReportClientMovement(rows){
  const box=$("rClientMovement"); box.innerHTML="";
  const data=combinedClientData(rows);
  if(!data.length){box.innerHTML='<div class="empty-report">No IN or OUT transactions recorded for this shift.</div>';return;}
  data.forEach(c=>{
    const card=document.createElement("div");
    card.className="client-report-card";
    card.innerHTML=`
      <div class="client-report-head">
        <div><span>Client</span><h3>${escapeHtml(c.client)}</h3></div>
      </div>
      <div class="client-movement-head"><span>Asset type</span><span>In</span><span>Out</span><span>With Client Today</span><span>All-Time With Client</span></div>
      <div class="client-movement-list">
        ${c.types.map(t=>`<div class="client-movement-row"><span>${escapeHtml(t.asset)}</span><b class="movement-in">${t.received}</b><b class="movement-out">${t.sent}</b><b class="with-client ${t.withClient>=0?"positive":"negative"}">${t.withClient>=0?"+":""}${t.withClient}</b><b class="with-client ${t.allTimeWithClient>=0?"positive":"negative"}">${t.allTimeWithClient>=0?"+":""}${t.allTimeWithClient}</b></div>`).join("")}
      </div>`;
    box.appendChild(card);
  });
}

function renderReportDiscardedAssetsToday(){
  const box=$("rDiscardedAssetsToday");
  box.innerHTML="";
  const data=assetTotals(todayDiscardedRows());
  if(!data.length){
    box.innerHTML='<div class="empty-report">No assets were discarded today.</div>';
    return;
  }
  const date=$("reportDate").value;
  const head=document.createElement("div");
  head.className="discarded-report-head";
  head.innerHTML='<span>Asset type</span><span>Count</span><span>Reason</span>';
  box.appendChild(head);
  data.forEach(x=>{
    const row=document.createElement("div");
    row.className="discarded-report-row";
    const key=discardedReasonKey(date,x.asset);
    row.innerHTML=`<strong>${escapeHtml(x.asset)}</strong><b>${x.quantity}</b><span>${escapeHtml(discardedReasons[key]||"Reason not entered")}</span>`;
    box.appendChild(row);
  });
}

function values(containerId){return [...$(containerId).children].map(row=>[...row.querySelectorAll("input,select")].map(x=>x.value.trim()));}
function fillTable(id,rows,cols){const tbody=$(id);tbody.innerHTML="";if(!rows.length){tbody.innerHTML=`<tr><td colspan="${cols}" style="text-align:center;color:#7b8791">No entries recorded</td></tr>`;return;}rows.forEach(r=>{const tr=document.createElement("tr");r.forEach(v=>{const td=document.createElement("td");td.textContent=v||"—";tr.appendChild(td)});tbody.appendChild(tr)})}

function generateReport(){
  if(!uploadedWorkbook){showToast("Upload the XLSX file first");return;}
  applyInventoryForShift();
  renderInventorySummaryPreview();
  const rows=shiftRows(), info=shiftInfo();
  $("rDate").textContent=formatDate($("reportDate").value);
  $("rShift").textContent=info.label;
  $("rShiftWindow").textContent=info.key==="night"?`${info.window} • ${formatDate($("reportDate").value)} → ${formatDate(info.nextDate)}`:info.window;
  $("rPrepared").textContent=text("preparedBy");
  const status=$("overallStatus").value;
  $("rStatus").textContent=status;
  $("rStatus").className="status-pill "+(status==="Critical"?"critical":status==="Attention Required"?"attention":"");
  renderReportInventorySummary();
  renderReportClientMovement(rows);
  renderReportDiscardedAssetsToday();
  $("rReceivedNotes").textContent=text("receivedNotes");
  $("rSentNotes").textContent=text("sentNotes");
  fillTable("rIssueRows",values("issueRows"),7);
  $("rTomorrow").textContent=text("tomorrowPlan");
  $("rSignoff").textContent=formatDateTime($("signoffDate").value);
  $("rEodStatus").textContent=$("eodStatus").value;

  const report=$("report");
  $("generateBtn").disabled=true; $("generateBtn").textContent="Generating...";
  html2canvas(report,{scale:3,backgroundColor:"#fff",useCORS:true,logging:false}).then(canvas=>canvas.toBlob(blob=>{
    const stamp=($("reportDate").value||"report").replaceAll("-","");
    const shift=info.key==="night"?"Night":"Day";
    const a=document.createElement("a"); a.download=`EOD_Asset_Report_${stamp}_${shift}.png`; a.href=URL.createObjectURL(blob); a.click();
    setTimeout(()=>URL.revokeObjectURL(a.href),1000); $("generateBtn").disabled=false; $("generateBtn").textContent="Generate Report"; showToast("High-quality end-of-shift report generated");
  },"image/png")).catch(err=>{console.error(err);$("generateBtn").disabled=false;$("generateBtn").textContent="Generate Report";showToast("Could not generate the image")});
}
function showToast(msg){const t=$("toast");t.textContent=msg;t.classList.add("show");setTimeout(()=>t.classList.remove("show"),2500)}

$("xlsxFile").addEventListener("change",e=>parseWorkbook(e.target.files[0]));
$("reportDate").addEventListener("change",()=>{updateShiftWindow();if(uploadedWorkbook){applyInventoryForShift();renderInventorySummaryPreview();renderClientMovementPreview();renderDiscardedAssetsTodayPreview();updateSummary();updateUploadStatus()}});
$("shiftSelect").addEventListener("change",()=>{updateShiftWindow();if(uploadedWorkbook){applyInventoryForShift();renderInventorySummaryPreview();renderClientMovementPreview();renderDiscardedAssetsTodayPreview();updateSummary();updateUploadStatus()}});
$("generateBtn").addEventListener("click",generateReport);
$("clearBtn").addEventListener("click",()=>{
  if(!confirm("Clear all entered data?"))return;
  document.querySelectorAll(".card input,.card textarea").forEach(x=>{if(x.type==="number")x.value=0;else x.value=""});
  $("reportDate").value=londonDate; $("signoffDate").value=new Date().toISOString().slice(0,16); $("shiftSelect").value="day";
  $("issueRows").innerHTML=""; assets=[...defaultAssets]; importedTransactions=[]; importedInventory={}; uploadedWorkbook=false; discardedReasons={}; buildInventoryForm(); updateShiftWindow(); updateSummary(); $("inventorySummaryPreview").innerHTML='<div class="upload-status">Upload the workbook to see this shift inventory.</div>';
  $("clientMovementPreview").innerHTML='<div class="upload-status">Upload the workbook to see this shift.</div>';
  $("discardedAssetsPreview").innerHTML='<div class="upload-status">Upload the workbook to see discarded assets today.</div>'; $("uploadStatus").textContent="No spreadsheet uploaded yet."; $("uploadStatus").className="upload-status";
});
