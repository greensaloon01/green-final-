// ============================================================
// GREEN SALON — BILLING SYSTEM v25
// Owner: Harsha | Developer: Shebin K Babu (one stop solutions - dlb)
// ============================================================
// HOW DATA FLOWS:
//   submitEntry   → Entries sheet (raw log, every entry)
//                 → Daily tab    (amount+tip+time per staff column, each entry new row)
//                 → Monthly tab  (TODAY's ROW: staff column += amount, Extra += tip,
//                                 Cash/Online += payment, Total/Comm/Diff recalculated)
//   submitProduct → ProductSales sheet (raw log)
//                 → Daily tab    (Product column)
//                 → Monthly tab  (TODAY's ROW: Product += amount)
//   submitExpense → Expenses sheet (raw log)
//                 → Monthly tab  (TODAY's ROW: Expenses += amount)
//
// Monthly tab structure (ONE ROW PER DAY, columns updated on every submit):
//   Date | Staff1 | Staff2 ... | Extra | Total | Product | Expenses | Commission | Online | Cash | Difference
//   Each submit ADDS to existing cell — never duplicates a row.
//   TOTAL row rebuilt after every submit.
// ============================================================
// SETUP:
//   1. Extensions → Apps Script → clear all → paste this
//   2. Replace MASTER_SHEET_ID and BRANCH_SHEETS IDs below
//   3. Run: firstTimeSetup  (approve all permissions)
//   4. Run: setupTriggers   (run ONCE only)
//   5. Deploy → New Deployment → Web App → Execute as: Me → Anyone
//   6. Copy /exec URL → paste in owner-panel Settings → API URL
// ============================================================

const MASTER_SHEET_ID = "PASTE_MASTER_SHEET_ID_HERE";
const OWNER_PASSWORD  = "harsha@greensalon2026";
const BRANCH_SHEETS   = {
  "branch1": "PASTE_BRANCH1_SHEET_ID_HERE",
  "branch2": "PASTE_BRANCH2_SHEET_ID_HERE",
  "branch3": "PASTE_BRANCH3_SHEET_ID_HERE",
};

// Multi-admin support — add more rows as needed
const ADMIN_ACCOUNTS = [
  { id:"admin1", name:"Harsha",  password:"harsha@greensalon2026", role:"owner"   },
  { id:"admin2", name:"Manager", password:"manager@green2026",      role:"manager" },
];

const C_DARK  = "#1a5c38";
const C_MED   = "#2d8653";
const C_WHITE = "#ffffff";
const C_ALT   = "#e8f5ee";
const C_RAW   = "#0f172a";
const C_WARN  = "#f59e0b";

// ── ROUTER ───────────────────────────────────────────────────
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    switch (d.action) {
      case "ownerLogin":        return R(ownerLogin(d));
      case "adminLogin":        return R(adminLogin(d));
      case "getBranches":       return R(getBranches());
      case "addBranch":         return R(addBranch(d));
      case "removeBranch":      return R(removeBranch(d));
      case "recoverBranch":     return R(recoverBranch(d));
      case "renameBranch":      return R(renameBranch(d));
      case "getStaffAdmin":     return R(getStaffAdmin(d));
      case "addStaff":          return R(addStaff(d));
      case "removeStaff":       return R(removeStaff(d));
      case "renameStaff":       return R(renameStaff(d));
      case "updateStaffComm":   return R(updateStaffComm(d));
      case "updateServices":    return R(updateServices(d));
      case "updateProducts":    return R(updateProducts(d));
      case "submitEntry":       return R(submitEntry(d));
      case "submitProduct":     return R(submitProduct(d));
      case "submitExpense":     return R(submitExpense(d));
      case "getMyEntries":      return R(getMyEntries(d));
      case "getTodayAll":       return R(getTodayAll(d));
      case "getBranchSummary":  return R(getBranchSummary(d));
      case "deleteEntry":       return R(deleteEntry(d));
      case "deleteProduct":     return R(deleteProduct(d));
      case "deleteExpense":     return R(deleteExpense(d));
      case "getMonthSummary":   return R(getMonthSummary(d));
      case "setReportEmails":   return R(setReportEmails(d));
      case "getReportEmails":   return R(getReportEmails(d));
      case "sendManualReport":  return R(sendManualReport(d));
      case "submitReport":      return R(submitReport(d));
      case "getReports":        return R(getReports(d));
      case "resolveReport":     return R(resolveReport(d));
      case "logAdminAction":    return R(logAdminAction(d));
      case "getAdminLog":       return R(getAdminLog(d));
      case "fixBranch":         return R(fixBranch(d));
      case "fixAllBranches":    return R(fixAllBranches(d));
      default: return E("Unknown action: " + d.action);
    }
  } catch (ex) { return E(ex.message + " | Stack: " + ex.stack); }
}

function doGet(e) {
  try {
    const a = e.parameter.action, bid = e.parameter.branchId;
    switch (a) {
      case "getStaff":    return R(getStaff(bid));
      case "getServices": return R(getServices(bid));
      case "getProducts": return R(getProducts(bid));
      case "getBranches": return R(getBranches());
      default: return E("Unknown GET action: " + a);
    }
  } catch (ex) { return E(ex.message); }
}

function R(d) { return ContentService.createTextOutput(JSON.stringify({success:true,...d})).setMimeType(ContentService.MimeType.JSON); }
function E(m) { return ContentService.createTextOutput(JSON.stringify({success:false,error:m})).setMimeType(ContentService.MimeType.JSON); }

// ── SHEET HELPERS ─────────────────────────────────────────────
function masterTab(name) {
  const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function branchSS(branchId) {
  const sid = BRANCH_SHEETS[branchId];
  if (sid && !sid.includes("_SHEET_ID_HERE")) {
    try { return SpreadsheetApp.openById(sid); } catch(ex) {}
  }
  const rows = masterTab("Branches").getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === branchId && rows[i][4] !== false && rows[i][4] !== "FALSE")
      return SpreadsheetApp.openById(rows[i][3]);
  }
  throw new Error("Branch not found: " + branchId);
}
function branchTab(branchId, tabName) {
  const ss = branchSS(branchId);
  return ss.getSheetByName(tabName) || ss.insertSheet(tabName);
}
function getExistingTab(branchId, tabName) {
  return branchSS(branchId).getSheetByName(tabName);
}
function todayStr() { return Utilities.formatDate(new Date(),"Asia/Kolkata","dd-MM-yyyy"); }
function nowIST()   { return Utilities.formatDate(new Date(),"Asia/Kolkata","dd-MMM-yyyy HH:mm:ss"); }
function monthName() {
  const ist = new Date(new Date().toLocaleString("en-US",{timeZone:"Asia/Kolkata"}));
  return ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][ist.getMonth()]+" "+ist.getFullYear();
}
function hdrStyle(rng,bg,fg){rng.setBackground(bg).setFontColor(fg).setFontWeight("bold").setHorizontalAlignment("center");}
function branchDisplayName(branchId){
  try{const rows=masterTab("Branches").getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===branchId)return rows[i][1];}}catch(ex){}
  return "Green Salon";
}
function activeStaffNames(ss){
  const st=ss.getSheetByName("Staff");if(!st||st.getLastRow()<2)return[];
  return st.getDataRange().getValues().slice(1).filter(r=>r[6]!==false&&r[6]!=="FALSE").map(r=>String(r[1]));
}

// ── AUTH ──────────────────────────────────────────────────────
function ownerLogin(d){if(d.password!==OWNER_PASSWORD)throw new Error("Wrong password");return{ownerName:"Harsha",adminId:"admin1",role:"owner"};}
function adminLogin(d){
  const a=ADMIN_ACCOUNTS.find(x=>x.name===d.name&&x.password===d.password);
  if(!a)throw new Error("Invalid admin credentials");
  return{adminId:a.id,adminName:a.name,role:a.role};
}

// ── ADMIN LOG ─────────────────────────────────────────────────
function logAdminAction(d){masterTab("AdminLog").appendRow([nowIST(),d.adminId||"",d.adminName||"",d.branchId||"",d.action||"",d.details||""]);return{logged:true};}
function getAdminLog(d){const rows=masterTab("AdminLog").getDataRange().getValues();const h=rows[0];return{log:rows.slice(1).map(r=>{const o={};h.forEach((k,i)=>o[k]=r[i]);return o;})};}

// ── FIRST TIME SETUP ──────────────────────────────────────────
function firstTimeSetup() {
  const bsh=masterTab("Branches");
  if(bsh.getLastRow()===0){bsh.appendRow(["BranchID","Name","Location","SheetID","Active","CreatedAt","DeletedAt"]);hdrStyle(bsh.getRange(1,1,1,7),C_RAW,C_WHITE);}
  const set=masterTab("Settings");
  if(set.getLastRow()===0){set.appendRow(["BranchID","Email1","Email2","Email3","UpdatedAt"]);hdrStyle(set.getRange(1,1,1,5),C_RAW,C_WHITE);}
  const al=masterTab("AdminLog");
  if(al.getLastRow()===0){al.appendRow(["Timestamp","AdminID","AdminName","BranchID","Action","Details"]);hdrStyle(al.getRange(1,1,1,6),C_RAW,C_WHITE);}
  const defs=[{id:"branch1",name:"Branch 1",loc:"JC Nagar"},{id:"branch2",name:"Branch 2",loc:"Koramangala"},{id:"branch3",name:"Branch 3",loc:"Indiranagar"}];
  const existing=bsh.getDataRange().getValues().map(r=>r[0]);
  defs.forEach(b=>{
    const sid=BRANCH_SHEETS[b.id];
    if(sid.includes("_SHEET_ID_HERE")){Logger.log("⚠️ "+b.id+" Sheet ID not set — skip");return;}
    if(!existing.includes(b.id))bsh.appendRow([b.id,b.name,b.loc,sid,true,nowIST(),""]);
    try{initBranch(sid,b.name);}catch(ex){Logger.log("❌ "+b.name+": "+ex.message);}
  });
  Logger.log("✅ Setup done. Run setupTriggers() then Deploy.");
}

function initBranch(sheetId,branchName){
  const ss=SpreadsheetApp.openById(sheetId);
  let st=ss.getSheetByName("Staff")||ss.insertSheet("Staff");
  if(st.getLastRow()===0){st.appendRow(["ID","Name","PIN","Commission%","HasCommission","PhotoURL","Active"]);hdrStyle(st.getRange(1,1,1,7),C_RAW,C_WHITE);st.appendRow(["S001","Staff 1","1111",40,true,"",true]);st.appendRow(["S002","Staff 2","2222",40,true,"",true]);st.appendRow(["S003","Staff 3","3333",35,true,"",true]);}
  let sv=ss.getSheetByName("Services")||ss.insertSheet("Services");
  if(sv.getLastRow()===0){sv.appendRow(["ServiceName","Price","Active"]);hdrStyle(sv.getRange(1,1,1,3),C_RAW,C_WHITE);[["Haircut",150],["Shave",80],["Facial",300],["Hair Colour",500],["Head Massage",100],["Beard Trim",60],["Threading",40],["Waxing",200]].forEach(r=>sv.appendRow([r[0],r[1],true]));}
  let pd=ss.getSheetByName("Products")||ss.insertSheet("Products");
  if(pd.getLastRow()===0){pd.appendRow(["ProductName","Price","Active"]);hdrStyle(pd.getRange(1,1,1,3),C_RAW,C_WHITE);[["Shampoo",200],["Hair Oil",150],["Conditioner",180],["Hair Serum",250]].forEach(r=>pd.appendRow([r[0],r[1],true]));}
  let en=ss.getSheetByName("Entries")||ss.insertSheet("Entries");
  if(en.getLastRow()===0){en.appendRow(["RowID","Timestamp","Date","StaffID","StaffName","Service","Amount","Tip","Payment","CommApplies","Flagged"]);hdrStyle(en.getRange(1,1,1,11),C_RAW,C_WHITE);}
  let ps=ss.getSheetByName("ProductSales")||ss.insertSheet("ProductSales");
  if(ps.getLastRow()===0){ps.appendRow(["RowID","Timestamp","Date","StaffID","StaffName","Product","Amount","Payment","Flagged"]);hdrStyle(ps.getRange(1,1,1,9),C_RAW,C_WHITE);}
  let ex=ss.getSheetByName("Expenses")||ss.insertSheet("Expenses");
  if(ex.getLastRow()===0){ex.appendRow(["RowID","Timestamp","Date","StaffID","StaffName","Description","Amount","Payment","Flagged"]);hdrStyle(ex.getRange(1,1,1,9),C_RAW,C_WHITE);}
  let rp=ss.getSheetByName("Reports")||ss.insertSheet("Reports");
  if(rp.getLastRow()===0){rp.appendRow(["ReportID","Timestamp","StaffID","StaffName","EntryRowID","EntryDetails","ReportType","Message","CorrectedValue","Status","ResolvedBy","ResolvedAt","ActionTaken"]);hdrStyle(rp.getRange(1,1,1,13),C_RAW,C_WHITE);}
  buildDailyTab(ss,branchName);
  buildMonthlyTab(ss,branchName,monthName());
  Logger.log("✅ "+branchName+" initialized");
}

// ── DAILY TAB (raw log per entry — clears at midnight) ────────
// Columns: StaffName | StaffName Tip | StaffName Time | ... | Product | Product Time
// Each entry: new row in staff's column. Values: "150C" (cash) or "300P" (online)
function buildDailyTab(ss,branchName){
  let sh=ss.getSheetByName("Daily")||ss.insertSheet("Daily");
  if(sh.getLastRow()>0)return sh;
  const names=activeStaffNames(ss);
  const headers=[];
  names.forEach(n=>{headers.push(n);headers.push(n+" Tip");headers.push(n+" Time");});
  headers.push("Product");headers.push("Product Time");
  sh.appendRow(headers);
  hdrStyle(sh.getRange(1,1,1,headers.length),C_DARK,C_WHITE);
  sh.setFrozenRows(1);
  for(let c=1;c<=headers.length;c++){const h=String(sh.getRange(1,c).getValue());sh.setColumnWidth(c,h.endsWith(" Time")?190:70);}
  return sh;
}

function dailyColMap(ss){
  const sh=ss.getSheetByName("Daily");if(!sh||sh.getLastRow()===0)return{};
  const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const m={};
  h.forEach((v,i)=>{const s=String(v);if(!s)return;
    if(s.endsWith(" Time")){const b=s.slice(0,-5);if(!m[b])m[b]={};m[b].time=i+1;}
    else if(s.endsWith(" Tip")){const b=s.slice(0,-4);if(!m[b])m[b]={};m[b].tip=i+1;}
    else{if(!m[s])m[s]={};m[s].amt=i+1;}
  });
  return m;
}

function ensureStaffInDaily(ss,name){
  const sh=ss.getSheetByName("Daily");if(!sh)return;
  const map=dailyColMap(ss);if(map[name]&&map[name].amt)return;
  const last=sh.getLastColumn();
  sh.insertColumnsAfter(last-2,3);
  sh.getRange(1,last-1).setValue(name);
  sh.getRange(1,last).setValue(name+" Tip");
  sh.getRange(1,last+1).setValue(name+" Time");
  hdrStyle(sh.getRange(1,last-1,1,3),C_DARK,C_WHITE);
  sh.setColumnWidth(last-1,70);sh.setColumnWidth(last,70);sh.setColumnWidth(last+1,190);
}

// Writes one entry into Daily tab (each entry = new row in staff column)
function writeDailyEntry(ss,staffName,amount,tip,payment,isProduct){
  const sh=ss.getSheetByName("Daily");if(!sh)return;
  if(!isProduct)ensureStaffInDaily(ss,staffName);
  const map=dailyColMap(ss);
  const key=isProduct?"Product":staffName;
  const info=map[key];if(!info||!info.amt)return;
  const amtVal=amount+(payment==="Cash"?"C":"P");
  const ts=Utilities.formatDate(new Date(),"Asia/Kolkata","dd-MMM-yyyy hh:mm:ss a");
  // Find first empty row in this column (below header)
  const lr=Math.max(sh.getLastRow(),1);
  const colVals=lr>1?sh.getRange(2,info.amt,lr-1,1).getValues():[];
  let row=2;
  for(let r=0;r<colVals.length;r++){if(!colVals[r][0]){row=r+2;break;}if(r===colVals.length-1)row=lr+1;}
  sh.getRange(row,info.amt).setValue(amtVal).setHorizontalAlignment("center");
  if(!isProduct&&info.tip&&tip>0)sh.getRange(row,info.tip).setValue(tip+(payment==="Cash"?"C":"P")).setHorizontalAlignment("center");
  if(info.time)sh.getRange(row,info.time).setValue(ts);
}

// ── MONTHLY TAB (ONE ROW PER DAY — totals accumulated) ────────
// Structure: row1=branch title merged, row2=headers, row3+=date rows + TOTAL row
// On each submit: find today's row (or create), ADD value to correct column, recalc.
// NEVER creates duplicate rows. NEVER stores individual entries here.
function buildMonthlyTab(ss,branchName,tabName){
  if(ss.getSheetByName(tabName))return ss.getSheetByName(tabName);
  const sh=ss.insertSheet(tabName);
  drawMonthlyFrame(sh,branchName,activeStaffNames(ss));
  return sh;
}
function drawMonthlyFrame(sh,branchName,staffNames){
  const cols=["Date",...staffNames,"Extra","Total","Product","Expenses","Commission","Online","Cash","Difference"];
  const nc=cols.length;
  sh.getRange(1,1,1,nc).merge().setValue(branchName).setBackground(C_DARK).setFontColor(C_WHITE).setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1,36);
  sh.getRange(2,1,1,nc).setValues([cols]);
  hdrStyle(sh.getRange(2,1,1,nc),C_MED,C_WHITE);
  sh.setRowHeight(2,26);sh.setFrozenRows(2);
  sh.setColumnWidth(1,110);for(let c=2;c<=nc;c++)sh.setColumnWidth(c,90);
}
function monthColMap(sh){if(!sh||sh.getLastRow()<2)return{};const h=sh.getRange(2,1,1,sh.getLastColumn()).getValues()[0];const m={};h.forEach((v,i)=>{if(v)m[String(v)]=i+1;});return m;}
function ensureStaffInMonthly(sh,staffName,branchName){
  const h=sh.getRange(2,1,1,sh.getLastColumn()).getValues()[0];if(h.includes(staffName))return;
  const ei=h.indexOf("Extra");if(ei<0)return;
  sh.insertColumnBefore(ei+1);sh.getRange(2,ei+1).setValue(staffName);hdrStyle(sh.getRange(2,ei+1,1,1),C_MED,C_WHITE);sh.setColumnWidth(ei+1,90);
  const nc=sh.getLastColumn();sh.getRange(1,1,1,nc).merge().setValue(branchName).setBackground(C_DARK).setFontColor(C_WHITE).setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle");
  for(let r=3;r<=sh.getLastRow();r++){if(String(sh.getRange(r,1).getValue())!=="TOTAL")sh.getRange(r,ei+1).setValue(0);}
}

// KEY FUNCTION: upsert today's row and ADD the value to correct column
// entryType: "service" | "product" | "expense"
function updateMonthly(branchId,entryType,staffName,svcAmt,tipAmt,payment,prodAmt,expAmt){
  const ss=branchSS(branchId);const bn=branchDisplayName(branchId);const tab=monthName();
  let sh=ss.getSheetByName(tab);if(!sh)sh=buildMonthlyTab(ss,bn,tab);
  // Ensure staff column exists
  if(entryType==="service"&&staffName&&svcAmt>0)ensureStaffInMonthly(sh,staffName,bn);
  const cm=monthColMap(sh);const dt=todayStr();const nc=sh.getLastColumn();
  // Find today's row and TOTAL row
  // FIX: Read date cells as display strings to avoid Date-object ISO mismatch.
  // We force the date column to plain text so it always matches todayStr().
  const lr=sh.getLastRow();
  let dr=-1,tr=-1;
  if(lr>=3){
    const dateVals=sh.getRange(3,1,lr-2,1).getDisplayValues();
    dateVals.forEach((row,idx)=>{
      const v=String(row[0]).trim();
      if(v===dt) dr=idx+3;
      if(v==="TOTAL") tr=idx+3;
    });
  }
  // Create today's row if not found
  if(dr<0){
    const zeroRow=new Array(nc).fill(0);
    zeroRow[0]=dt; // plain string dd-MM-yyyy
    if(tr>0){
      sh.insertRowBefore(tr);
      sh.getRange(tr,1,1,nc).setValues([zeroRow]);
      dr=tr; tr=tr+1;
    }else{
      sh.appendRow(zeroRow); dr=sh.getLastRow();
    }
    // Force date cell to plain text so it never becomes a Date object
    sh.getRange(dr,1).setNumberFormat("@");
    const bg=dr%2===0?C_WHITE:C_ALT;
    sh.getRange(dr,1,1,nc).setBackground(bg).setHorizontalAlignment("center");
    sh.getRange(dr,1).setHorizontalAlignment("left");
  }
  // Helper: add value to a named column in today's row
  function addVal(colKey,val){
    if(!colKey||!val||Number(val)<=0)return;
    const col=cm[colKey];if(!col)return;
    const cur=Number(sh.getRange(dr,col).getValue())||0;
    sh.getRange(dr,col).setValue(cur+Number(val));
  }
  // Route to correct columns by type
  if(entryType==="service"){
    if(svcAmt>0)addVal(staffName,svcAmt);
    if(tipAmt>0)addVal("Extra",tipAmt);
    const money=(Number(svcAmt)||0)+(Number(tipAmt)||0);
    if(money>0)addVal(payment==="Cash"?"Cash":"Online",money);
  }else if(entryType==="product"){
    if(prodAmt>0){addVal("Product",prodAmt);addVal(payment==="Cash"?"Cash":"Online",prodAmt);}
  }else if(entryType==="expense"){
    if(expAmt>0)addVal("Expenses",expAmt);
  }
  // Recalc Total, Commission, Difference for today's row
  recalcRow(sh,dr,cm,nc,branchId);
  // Rebuild TOTAL row
  rebuildTotal(sh,nc);
}

// Recalc Total = sum of all staff cols; Commission = per-staff rate; Diff = online+cash-total
function recalcRow(sh,row,cm,nc,branchId){
  const FIXED=new Set(["Date","Extra","Total","Product","Expenses","Commission","Online","Cash","Difference"]);
  const h=sh.getRange(2,1,1,nc).getValues()[0];
  // Build per-staff commission rates
  const commMap={};
  if(branchId){
    try{
      const ss=branchSS(branchId);
      const sRows=ss.getSheetByName("Staff").getDataRange().getValues().slice(1);
      sRows.forEach(r=>{if(r[6]!==false&&r[6]!=="FALSE")commMap[String(r[1])]={pct:Number(r[3])||0,has:r[4]===true||r[4]==="TRUE"};});
    }catch(ex){}
  }
  let totalSvc=0,totalComm=0;
  h.forEach((v,i)=>{
    if(!v||FIXED.has(String(v)))return;
    const amt=Number(sh.getRange(row,i+1).getValue())||0;
    totalSvc+=amt;
    const sn=String(v);
    if(commMap[sn]&&commMap[sn].has)totalComm+=amt*(commMap[sn].pct/100);
    else if(commMap[sn]&&!commMap[sn].has){}// fixed salary
    else totalComm+=amt*0.40;// fallback
  });
  if(cm["Total"])sh.getRange(row,cm["Total"]).setValue(totalSvc);
  if(cm["Commission"])sh.getRange(row,cm["Commission"]).setValue(Math.round(totalComm));
  const onl=cm["Online"]?Number(sh.getRange(row,cm["Online"]).getValue())||0:0;
  const csh=cm["Cash"]?Number(sh.getRange(row,cm["Cash"]).getValue())||0:0;
  if(cm["Difference"])sh.getRange(row,cm["Difference"]).setValue(onl+csh-totalSvc);
}

function rebuildTotal(sh,nc){
  const last=sh.getLastRow();let tr=-1;
  // Use getDisplayValue so date cells stored as text are matched correctly
  for(let r=3;r<=last;r++){if(sh.getRange(r,1).getDisplayValue().trim()==="TOTAL"){tr=r;break;}}
  const sums=new Array(nc).fill(0);
  const endR=tr>0?tr:last+1;
  for(let r=3;r<endR;r++){const v=sh.getRange(r,1,1,nc).getValues()[0];for(let c=1;c<nc;c++)sums[c]+=Number(v[c])||0;}
  const totalRow=["TOTAL",...sums.slice(1)];
  if(tr<0){sh.appendRow(totalRow);tr=sh.getLastRow();}
  else sh.getRange(tr,1,1,nc).setValues([totalRow]);
  hdrStyle(sh.getRange(tr,1,1,nc),C_DARK,C_WHITE);
  sh.getRange(tr,1).setHorizontalAlignment("left");
}

// ── TRIGGERS ──────────────────────────────────────────────────
function setupTriggers(){
  ScriptApp.getProjectTriggers().forEach(t=>{if(["midnightReset","checkMonthEnd","sendDailyReport","sendMonthlyReport"].includes(t.getHandlerFunction()))ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger("midnightReset").timeBased().everyDays(1).atHour(23).nearMinute(55).inTimezone("Asia/Kolkata").create();
  ScriptApp.newTrigger("checkMonthEnd").timeBased().everyDays(1).atHour(0).nearMinute(5).inTimezone("Asia/Kolkata").create();
  ScriptApp.newTrigger("sendDailyReport").timeBased().everyDays(1).atHour(23).nearMinute(0).inTimezone("Asia/Kolkata").create();
  ScriptApp.newTrigger("sendMonthlyReport").timeBased().everyDays(1).atHour(22).nearMinute(55).inTimezone("Asia/Kolkata").create();
  Logger.log("✅ 4 triggers created.");
}
function midnightReset(){masterTab("Branches").getDataRange().getValues().slice(1).forEach(row=>{if(row[4]===false||row[4]==="FALSE")return;try{const sh=SpreadsheetApp.openById(row[3]).getSheetByName("Daily");if(sh&&sh.getLastRow()>1)sh.deleteRows(2,sh.getLastRow()-1);}catch(ex){Logger.log("reset err "+row[1]+": "+ex.message);}});}
function checkMonthEnd(){const tab=monthName();masterTab("Branches").getDataRange().getValues().slice(1).forEach(row=>{if(row[4]===false||row[4]==="FALSE")return;try{const ss=SpreadsheetApp.openById(row[3]);if(!ss.getSheetByName(tab)){buildMonthlyTab(ss,row[1],tab);Logger.log("✅ Created: "+tab+" for "+row[1]);}}catch(ex){Logger.log("monthEnd err: "+ex.message);}});}

// ── BRANCHES ──────────────────────────────────────────────────
function getBranches(){const rows=masterTab("Branches").getDataRange().getValues();const all=rows.slice(1).map(r=>({id:r[0],name:r[1],location:r[2],sheetId:r[3],active:r[4]!==false&&r[4]!=="FALSE",deletedAt:r[6]||""}));return{branches:all.filter(b=>b.active),deleted:all.filter(b=>!b.active)};}
function addBranch(d){if(!d.name||!d.sheetId)throw new Error("Name and SheetID required");try{SpreadsheetApp.openById(d.sheetId);}catch(ex){throw new Error("Cannot access Sheet — must be shared with same Gmail");}const id="branch"+Date.now();masterTab("Branches").appendRow([id,d.name,d.location||"",d.sheetId,true,nowIST(),""]);initBranch(d.sheetId,d.name);return{branchId:id};}
function removeBranch(d){const sh=masterTab("Branches"),rows=sh.getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===d.branchId){sh.getRange(i+1,5).setValue(false);sh.getRange(i+1,7).setValue(nowIST());return{};}}throw new Error("Branch not found");}
function recoverBranch(d){const sh=masterTab("Branches"),rows=sh.getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===d.branchId){sh.getRange(i+1,5).setValue(true);sh.getRange(i+1,7).setValue("");return{name:rows[i][1]};}}throw new Error("Branch not found");}
function renameBranch(d){if(!d.newName)throw new Error("New name required");const sh=masterTab("Branches"),rows=sh.getDataRange().getValues();let old="";for(let i=1;i<rows.length;i++){if(rows[i][0]===d.branchId){old=rows[i][1];sh.getRange(i+1,2).setValue(d.newName);break;}}try{const ss=branchSS(d.branchId),msh=ss.getSheetByName(monthName());if(msh){const nc=msh.getLastColumn();msh.getRange(1,1,1,nc).merge().setValue(d.newName).setBackground(C_DARK).setFontColor(C_WHITE).setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle");}}catch(ex){}return{oldName:old,newName:d.newName};}

// ── STAFF ──────────────────────────────────────────────────────
function getStaff(bid){return{staff:_staffRows(bid)};}
function getStaffAdmin(d){return{staff:_staffRows(d.branchId)};}
function _staffRows(bid){return branchTab(bid,"Staff").getDataRange().getValues().slice(1).filter(r=>r[6]!==false&&r[6]!=="FALSE").map(r=>({id:r[0],name:r[1],pin:r[2],photoUrl:r[5]||"",hasCommission:r[4],commissionPct:Number(r[3])||0}));}
function addStaff(d){
  // BACKEND GUARD: reject empty or invalid submissions — prevents ghost rows
  const name=String(d.name||"").trim();
  const pin=String(d.pin||"").trim();
  if(!name||name.length<1) throw new Error("Staff name is required");
  if(!pin||pin.length!==4||!/^\d{4}$/.test(pin)) throw new Error("PIN must be exactly 4 digits");
  const sh=branchTab(d.branchId,"Staff");
  const id="S"+Date.now();
  sh.appendRow([id,name,pin,Number(d.commissionPct)||40,d.hasCommission!==false,d.photoUrl||"",true]);
  const ss=branchSS(d.branchId);
  const bn=branchDisplayName(d.branchId);
  ensureStaffInDaily(ss,name);
  const msh=ss.getSheetByName(monthName());
  if(msh)ensureStaffInMonthly(msh,name,bn);
  return{staffId:id};
}
function removeStaff(d){const sh=branchTab(d.branchId,"Staff"),rows=sh.getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===d.staffId){sh.getRange(i+1,7).setValue(false);return{};}}throw new Error("Staff not found");}
function renameStaff(d){
  if(!d.newName)throw new Error("New name required");
  const sh=branchTab(d.branchId,"Staff"),rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(rows[i][0]===d.staffId){
      const oldName=rows[i][1];
      sh.getRange(i+1,2).setValue(d.newName);
      // BUG1 FIX: rename column headers in Daily and Monthly tabs
      try{
        const ss=branchSS(d.branchId);
        const daily=ss.getSheetByName("Daily");
        if(daily&&daily.getLastRow()>0){
          const hdr=daily.getRange(1,1,1,daily.getLastColumn()).getValues()[0];
          hdr.forEach((v,ci)=>{
            const sv=String(v);
            if(sv===oldName)              daily.getRange(1,ci+1).setValue(d.newName);
            else if(sv===oldName+" Tip")  daily.getRange(1,ci+1).setValue(d.newName+" Tip");
            else if(sv===oldName+" Time") daily.getRange(1,ci+1).setValue(d.newName+" Time");
          });
        }
        const msh=ss.getSheetByName(monthName());
        if(msh&&msh.getLastRow()>1){
          const mhdr=msh.getRange(2,1,1,msh.getLastColumn()).getValues()[0];
          mhdr.forEach((v,ci)=>{if(String(v)===oldName)msh.getRange(2,ci+1).setValue(d.newName);});
        }
      }catch(ex){Logger.log("renameStaff header sync: "+ex.message);}
      return{oldName,newName:d.newName};
    }
  }
  throw new Error("Staff not found");
}
function updateStaffComm(d){const sh=branchTab(d.branchId,"Staff"),rows=sh.getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===d.staffId){sh.getRange(i+1,4).setValue(Number(d.commissionPct)||0);sh.getRange(i+1,5).setValue(d.hasCommission===true||d.hasCommission==="true");return{};}}throw new Error("Staff not found");}

// ── SERVICES / PRODUCTS ────────────────────────────────────────
function getServices(bid){return{services:branchTab(bid,"Services").getDataRange().getValues().slice(1).filter(r=>r[2]!==false&&r[2]!=="FALSE").map(r=>({name:r[0],price:r[1]}))}}
function updateServices(d){const sh=branchTab(d.branchId,"Services");if(sh.getLastRow()>1)sh.deleteRows(2,sh.getLastRow()-1);if(d.services&&d.services.length)d.services.forEach(s=>sh.appendRow([s.name,s.price,true]));return{};}
function getProducts(bid){return{products:branchTab(bid,"Products").getDataRange().getValues().slice(1).filter(r=>r[2]!==false&&r[2]!=="FALSE").map(r=>({name:r[0],price:r[1]}))}}
function updateProducts(d){const sh=branchTab(d.branchId,"Products");if(sh.getLastRow()>1)sh.deleteRows(2,sh.getLastRow()-1);if(d.products&&d.products.length)d.products.forEach(p=>sh.appendRow([p.name,p.price,true]));return{};}

// ── SUBMIT ENTRIES ─────────────────────────────────────────────
// Each submit: raw log row → Entries/ProductSales/Expenses sheet
//              + writeDailyEntry (new row in Daily tab)
//              + updateMonthly  (ADD to today's total row in Monthly tab)
function submitEntry(d){
  const{branchId,staffId,staffName,service,amount,tip,paymentMethod}=d;
  if(!branchId||!staffName||!service||!paymentMethod)throw new Error("Missing required fields");
  const amt=Number(amount)||0,tip2=Number(tip)||0;
  if(amt<=0)throw new Error("Amount must be > 0");
  const sRows=branchTab(branchId,"Staff").getDataRange().getValues();let comm=true;
  for(let i=1;i<sRows.length;i++){if(sRows[i][0]===staffId){comm=sRows[i][4]===true||sRows[i][4]==="TRUE";break;}}
  const rid="E"+Date.now(),ts=nowIST(),dt=todayStr();
  branchTab(branchId,"Entries").appendRow([rid,ts,dt,String(staffId||""),String(staffName),String(service),amt,tip2,String(paymentMethod),comm,false]);
  writeDailyEntry(branchSS(branchId),staffName,amt,tip2,paymentMethod,false);
  updateMonthly(branchId,"service",staffName,amt,tip2,paymentMethod,0,0);
  return{rowId:rid,timestamp:ts};
}
function submitProduct(d){
  const{branchId,staffId,staffName,product,amount,paymentMethod}=d;
  if(!branchId||!product||!paymentMethod)throw new Error("Missing required fields");
  const amt=Number(amount)||0;if(amt<=0)throw new Error("Amount must be > 0");
  const rid="P"+Date.now(),ts=nowIST(),dt=todayStr();
  branchTab(branchId,"ProductSales").appendRow([rid,ts,dt,String(staffId||"GLOBAL"),String(staffName||"Branch"),String(product),amt,String(paymentMethod),false]);
  writeDailyEntry(branchSS(branchId),"Product",amt,0,paymentMethod,true);
  updateMonthly(branchId,"product","",0,0,paymentMethod,amt,0);
  return{rowId:rid,timestamp:ts};
}
function submitExpense(d){
  const{branchId,staffId,staffName,description,amount,paymentMethod}=d;
  if(!branchId||!description||!paymentMethod)throw new Error("Missing required fields");
  const amt=Number(amount)||0;if(amt<=0)throw new Error("Amount must be > 0");
  const rid="X"+Date.now(),ts=nowIST(),dt=todayStr();
  branchTab(branchId,"Expenses").appendRow([rid,ts,dt,String(staffId||"GLOBAL"),String(staffName||"Branch"),String(description),amt,String(paymentMethod),false]);
  updateMonthly(branchId,"expense","",0,0,paymentMethod,0,amt);
  return{rowId:rid,timestamp:ts};
}

// ── GET ENTRIES ────────────────────────────────────────────────
// getMyEntries — returns TODAY's entries for THIS staff (for My Log)
function getMyEntries(d){
  const{branchId,staffId}=d;const dt=todayStr();
  const sh=branchTab(branchId,"Entries");
  if(sh.getLastRow()<2)return{entries:[],totalAmount:0,totalTip:0};
  const vals=sh.getDataRange().getValues();
  // Use getDisplayValues for date column to handle Date object vs string
  const disp=sh.getDataRange().getDisplayValues();
  const out=[];let ta=0,tt=0;
  for(let i=1;i<vals.length;i++){
    const r=vals[i];const dateCell=disp[i][2]; // col C display value
    const flagged=r[10]===true||r[10]==="TRUE";
    if(String(r[3])===String(staffId)&&(dateCell===dt||String(r[2])===dt)&&!flagged){
      out.push({rowId:r[0],timestamp:r[1],service:r[5],amount:r[6],tip:r[7],paymentMethod:r[8]});
      ta+=Number(r[6])||0;tt+=Number(r[7])||0;
    }
  }
  return{entries:out,totalAmount:ta,totalTip:tt};
}

// getBranchSummary — landing card: total entries + service revenue + product revenue + tips
function getBranchSummary(d){
  const bid=d.branchId;const dt=todayStr();
  const eR=branchTab(bid,"Entries").getDataRange().getValues();
  let totalEntries=0,totalRevenue=0,totalTips=0;
  const eDisp=branchTab(bid,"Entries").getDataRange().getDisplayValues();
  eR.slice(1).forEach((r,i)=>{const dc=eDisp[i+1]?eDisp[i+1][2]:String(r[2]);if((dc===dt||String(r[2])===dt)&&r[10]!==true&&r[10]!=="TRUE"){totalEntries++;totalRevenue+=Number(r[6])||0;totalTips+=Number(r[7])||0;}});
  // Include product sales in revenue
  const pSh=branchTab(bid,"ProductSales");const pR=pSh.getDataRange().getValues();
  const pDisp=pSh.getDataRange().getDisplayValues();
  let totalProducts=0;
  pR.slice(1).forEach((r,i)=>{const dc=pDisp[i+1]?pDisp[i+1][2]:String(r[2]);if((dc===dt||String(r[2])===dt)&&r[8]!==true&&r[8]!=="TRUE")totalProducts+=Number(r[6])||0;});
  totalRevenue+=totalProducts;
  return{totalEntries,totalRevenue,totalTips,totalProducts};
}

// getTodayAll — full today snapshot for owner panel
function getTodayAll(d){
  const bid=d.branchId;const dt=todayStr();
  const eR=branchTab(bid,"Entries").getDataRange().getValues();
  const entries=[];const sm={};
  const eDisp2=branchTab(bid,"Entries").getDataRange().getDisplayValues();
  eR.slice(1).forEach((r,ri)=>{
    const dc=eDisp2[ri+1]?eDisp2[ri+1][2]:String(r[2]);
    if(dc!==dt&&String(r[2])!==dt)return;
    const fl=r[10]===true||r[10]==="TRUE";
    entries.push({rowId:r[0],timestamp:r[1],staffId:r[3],staffName:r[4],service:r[5],amount:r[6],tip:r[7],paymentMethod:r[8],commissionApplies:r[9],flagged:fl});
    if(!fl){const sn=String(r[4]);if(!sm[sn])sm[sn]={name:sn,totalAmount:0,totalTip:0,entries:0,products:0};sm[sn].totalAmount+=Number(r[6])||0;sm[sn].totalTip+=Number(r[7])||0;sm[sn].entries++;}
  });
  const pR=branchTab(bid,"ProductSales").getDataRange().getValues();const ps=[];let totalProductRevenue=0;
  const pShT=branchTab(bid,"ProductSales");const pDisp2=pShT.getDataRange().getDisplayValues();
  pR.slice(1).forEach((r,ri)=>{
    const pdc=pDisp2[ri+1]?pDisp2[ri+1][2]:String(r[2]);
    if(pdc!==dt&&String(r[2])!==dt)return;
    const fl=r[8]===true||r[8]==="TRUE";
    ps.push({rowId:r[0],timestamp:r[1],staffName:r[4],product:r[5],amount:r[6],paymentMethod:r[7],flagged:fl});
    if(!fl){totalProductRevenue+=Number(r[6])||0;const sn=String(r[4]);if(!sm[sn])sm[sn]={name:sn,totalAmount:0,totalTip:0,entries:0,products:0};sm[sn].products+=Number(r[6])||0;}
  });
  const xR=branchTab(bid,"Expenses").getDataRange().getValues();const xs=[];let xe=0;
  const xShT=branchTab(bid,"Expenses");const xDisp2=xShT.getDataRange().getDisplayValues();
  xR.slice(1).forEach((r,ri)=>{
    const xdc=xDisp2[ri+1]?xDisp2[ri+1][2]:String(r[2]);
    if(xdc!==dt&&String(r[2])!==dt)return;
    const fl=r[8]===true||r[8]==="TRUE";
    xs.push({rowId:r[0],timestamp:r[1],staffName:r[4],description:r[5],amount:r[6],paymentMethod:r[7],flagged:fl});
    if(!fl)xe+=Number(r[6])||0;
  });
  return{entries,staffTotals:Object.values(sm),productSales:ps,expenses:xs,totalExp:xe,totalProductRevenue};
}

function getMonthSummary(d){
  const bid=d.branchId;const tab=monthName();const sh=getExistingTab(bid,tab);if(!sh)return{summary:[],month:tab};
  const all=sh.getDataRange().getValues();if(all.length<3)return{summary:[],month:tab};
  const headers=all[1];return{summary:all.slice(2).filter(r=>r[0]).map(r=>{const o={};headers.forEach((k,i)=>{o[String(k)]=r[i];});return o;}),month:tab};
}

// ── DELETE (soft flag) ─────────────────────────────────────────
function deleteEntry(d){const sh=branchTab(d.branchId,"Entries"),rows=sh.getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===d.rowId){sh.getRange(i+1,11).setValue(true);return{};}}throw new Error("Not found");}
function deleteProduct(d){const sh=branchTab(d.branchId,"ProductSales"),rows=sh.getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===d.rowId){sh.getRange(i+1,9).setValue(true);return{};}}throw new Error("Not found");}
function deleteExpense(d){const sh=branchTab(d.branchId,"Expenses"),rows=sh.getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===d.rowId){sh.getRange(i+1,9).setValue(true);return{};}}throw new Error("Not found");}

// ── REPORTS ────────────────────────────────────────────────────
function submitReport(d){
  const{branchId,staffId,staffName,entryRowId,entryDetails,reportType,message,correctedValue}=d;
  if(!branchId||!staffName||!reportType)throw new Error("Missing required fields");
  const rid="R"+Date.now();
  branchTab(branchId,"Reports").appendRow([rid,nowIST(),String(staffId||""),String(staffName),String(entryRowId||""),String(entryDetails||""),String(reportType),String(message||""),String(correctedValue||""),"Pending","","",""]);
  return{reportId:rid,status:"Pending"};
}
function getReports(d){
  const sh=getExistingTab(d.branchId,"Reports");if(!sh)return{reports:[]};
  const rows=sh.getDataRange().getValues();if(rows.length<2)return{reports:[]};
  const headers=rows[0];
  let reps=rows.slice(1).filter(r=>r[0]).map(r=>{const o={};headers.forEach((k,i)=>o[String(k)]=r[i]);return o;});
  const sf=d.status&&typeof d.status==="string"&&d.status.trim();
  if(sf){if(sf==="Resolved")reps=reps.filter(r=>String(r["Status"]||"").startsWith("Resolved"));else reps=reps.filter(r=>r["Status"]===sf);}
  return{reports:reps};
}
function resolveReport(d){
  const{branchId,reportId,action,adminId,adminName,note}=d;
  const sh=branchTab(branchId,"Reports"),rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(rows[i][0]===reportId){
      const sm={delete_entry:"Resolved - Deleted",mark_valid:"Resolved - Valid",corrected:"Resolved - Corrected",ignored:"Ignored"};
      sh.getRange(i+1,10).setValue(sm[action]||"Resolved");
      sh.getRange(i+1,11).setValue(adminName||adminId||"Admin");
      sh.getRange(i+1,12).setValue(nowIST());
      sh.getRange(i+1,13).setValue(note||action);
      if(action==="delete_entry"&&rows[i][4]){try{deleteEntry({branchId,rowId:rows[i][4]});}catch(ex){}}
      masterTab("AdminLog").appendRow([nowIST(),adminId||"",adminName||"",branchId,"resolveReport","Report "+reportId+" → "+action]);
      return{resolved:true,action};
    }
  }
  throw new Error("Report not found");
}

// ── FIX / RECOVERY ─────────────────────────────────────────────
function fixBranch(d){
  const bid=d.branchId;const ss=branchSS(bid);const bn=branchDisplayName(bid);const results=[];
  const tabs=[
    {name:"Staff",       hdr:["ID","Name","PIN","Commission%","HasCommission","PhotoURL","Active"]},
    {name:"Services",    hdr:["ServiceName","Price","Active"]},
    {name:"Products",    hdr:["ProductName","Price","Active"]},
    {name:"Entries",     hdr:["RowID","Timestamp","Date","StaffID","StaffName","Service","Amount","Tip","Payment","CommApplies","Flagged"]},
    {name:"ProductSales",hdr:["RowID","Timestamp","Date","StaffID","StaffName","Product","Amount","Payment","Flagged"]},
    {name:"Expenses",    hdr:["RowID","Timestamp","Date","StaffID","StaffName","Description","Amount","Payment","Flagged"]},
    {name:"Reports",     hdr:["ReportID","Timestamp","StaffID","StaffName","EntryRowID","EntryDetails","ReportType","Message","CorrectedValue","Status","ResolvedBy","ResolvedAt","ActionTaken"]},
  ];
  tabs.forEach(t=>{let sh=ss.getSheetByName(t.name);if(!sh){sh=ss.insertSheet(t.name);sh.appendRow(t.hdr);hdrStyle(sh.getRange(1,1,1,t.hdr.length),C_RAW,C_WHITE);results.push("Created: "+t.name);}else if(sh.getLastRow()===0){sh.appendRow(t.hdr);hdrStyle(sh.getRange(1,1,1,t.hdr.length),C_RAW,C_WHITE);results.push("Header added: "+t.name);}else results.push("OK: "+t.name);});
  const daily=ss.getSheetByName("Daily");if(!daily||daily.getLastRow()===0){buildDailyTab(ss,bn);results.push("Daily tab rebuilt");}else results.push("OK: Daily");
  const tab=monthName();if(!ss.getSheetByName(tab)){buildMonthlyTab(ss,bn,tab);results.push("Monthly created: "+tab);}else results.push("OK: "+tab);
  return{fixed:true,details:results};
}
function fixAllBranches(d){
  const rows=masterTab("Branches").getDataRange().getValues();const results=[];
  for(let i=1;i<rows.length;i++){
    if(rows[i][4]===false||rows[i][4]==="FALSE")continue;
    try{const r=fixBranch({branchId:rows[i][0]});results.push({branch:rows[i][1],fixed:r.fixed,details:r.details});}
    catch(ex){results.push({branch:rows[i][1],fixed:false,details:[ex.message]});}
  }
  return{results};
}

// ── EMAIL ──────────────────────────────────────────────────────
function setReportEmails(d){const sh=masterTab("Settings"),rows=sh.getDataRange().getValues();const em=d.emails||[];for(let i=1;i<rows.length;i++){if(rows[i][0]===d.branchId){sh.getRange(i+1,2).setValue(em[0]||"");sh.getRange(i+1,3).setValue(em[1]||"");sh.getRange(i+1,4).setValue(em[2]||"");sh.getRange(i+1,5).setValue(nowIST());return{};}}sh.appendRow([d.branchId,em[0]||"",em[1]||"",em[2]||"",nowIST()]);return{};}
function getReportEmails(d){const rows=masterTab("Settings").getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===d.branchId)return{emails:[rows[i][1]||"",rows[i][2]||"",rows[i][3]||""].filter(Boolean)};}return{emails:[]};}
function getBranchEmails(branchId){const rows=masterTab("Settings").getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===branchId)return[rows[i][1],rows[i][2],rows[i][3]].filter(Boolean);}return[];}
function getSheetAsCSV(sheetId,sheetName){try{const ss=SpreadsheetApp.openById(sheetId);const sh=ss.getSheetByName(sheetName);if(!sh||sh.getLastRow()<1)return"";return sh.getDataRange().getValues().map(row=>row.map(cell=>{const s=String(cell).replace(/"/g,'""');return s.includes(",")||s.includes('"')||s.includes('\n')?`"${s}"`:s;}).join(",")).join("\n");}catch(ex){return"";}}
function sendDailyReport(){masterTab("Branches").getDataRange().getValues().slice(1).forEach(row=>{if(row[4]===false||row[4]==="FALSE")return;const em=getBranchEmails(row[0]);if(!em.length)return;try{_sendDailyCSV(row[0],row[1],row[3],em);}catch(ex){Logger.log("daily err "+row[1]+": "+ex.message);}});}
function _sendDailyCSV(bid,bname,sheetId,emails){const dt=todayStr();const atts=[];["Entries","ProductSales","Expenses","Daily"].forEach(t=>{const csv=getSheetAsCSV(sheetId,t);if(csv)atts.push(Utilities.newBlob(csv,"text/csv",t+"_"+dt.replace(/-/g,"")+".csv"));});const body="Hello,\n\nDaily report for "+bname+".\nDate: "+dt+"\n\nRegards,\nGreen Salon";emails.forEach(e=>{try{MailApp.sendEmail({to:e,subject:"Green Salon — "+bname+" — Daily Report — "+dt,body,attachments:atts});}catch(ex){Logger.log("send err "+e);}});}
function sendMonthlyReport(){const now=new Date();const ist=new Date(now.toLocaleString("en-US",{timeZone:"Asia/Kolkata"}));const last=new Date(ist.getFullYear(),ist.getMonth()+1,0).getDate();if(ist.getDate()!==last)return;masterTab("Branches").getDataRange().getValues().slice(1).forEach(row=>{if(row[4]===false||row[4]==="FALSE")return;const em=getBranchEmails(row[0]);if(!em.length)return;try{_sendMonthlyCSV(row[0],row[1],row[3],em);}catch(ex){Logger.log("monthly err: "+ex.message);}});}
function _sendMonthlyCSV(bid,bname,sheetId,emails){const tab=monthName();const csv=getSheetAsCSV(sheetId,tab);if(!csv)return;const ist=new Date(new Date().toLocaleString("en-US",{timeZone:"Asia/Kolkata"}));const ym=ist.getFullYear()+"-"+String(ist.getMonth()+1).padStart(2,"0");const blob=Utilities.newBlob(csv,"text/csv","Monthly_"+ym+".csv");const body="Hello,\n\nMonthly report for "+bname+".\nMonth: "+tab+"\n\nRegards,\nGreen Salon";emails.forEach(e=>{try{MailApp.sendEmail({to:e,subject:"Green Salon — "+bname+" — Monthly Report — "+tab,body,attachments:[blob]});}catch(ex){Logger.log("monthly send err "+e);}});}
function sendManualReport(d){const em=getBranchEmails(d.branchId);if(!em.length)throw new Error("No emails saved. Go to Settings.");const bname=branchDisplayName(d.branchId);const rows=masterTab("Branches").getDataRange().getValues();let sid="";for(let i=1;i<rows.length;i++){if(rows[i][0]===d.branchId){sid=rows[i][3];break;}}if(!sid)throw new Error("Branch sheet not found");if(d.type==="monthly")_sendMonthlyCSV(d.branchId,bname,sid,em);else _sendDailyCSV(d.branchId,bname,sid,em);return{sent:em.length,recipients:em};}
