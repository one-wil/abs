/**
 * GESTION DE PRÉSENCE & ABSENCES — API JSONP
 * Compatible avec une interface hébergée sur GitHub Pages, sans Worker.
 */
const CONFIG = {
  SPREADSHEET_ID: "1qOPtJiDc8gBTWlqm9-pnUwoqJDC1nClZmIrdpxcLzqU",
  PERSONNEL_SHEET: "Personnel",
  ATTENDANCE_SHEET: "Pointages",
  TIMEZONE: "Africa/Algiers"
};

function doGet(e) {
  try {
    const p=(e && e.parameter) || {};
    const action=p.action || "";
    const payload=p.payload ? JSON.parse(p.payload) : {};
    let result;
    switch(action) {
      case "getPersonnelData": result=getPersonnelData(); break;
      case "getDashboard": result=getDashboard(payload.date); break;
      case "saveAttendance": result=saveAttendance(payload); break;
      case "getAttendanceHistory": result=getAttendanceHistory(payload); break;
      default: result={message:"API Gestion de Présence opérationnelle."};
    }
    return output_(p.callback,{success:true,result});
  } catch(err) {
    return output_((e && e.parameter && e.parameter.callback),{
      success:false,error:String(err && err.message ? err.message : err)
    });
  }
}

function output_(callback,data) {
  const json=JSON.stringify(data);
  // callback is accepted only when it matches a simple JavaScript identifier.
  if(callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback+"("+json+");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function getSpreadsheet_(){return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);}
function normalize_(v){
  return String(v==null?"":v).trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g,"");
}
function headerMap_(headers){
  const map={};headers.forEach((h,i)=>map[normalize_(h)]=i);return map;
}
function findColumn_(map,aliases){
  for(const a of aliases){const k=normalize_(a);if(Object.prototype.hasOwnProperty.call(map,k))return map[k];}
  return -1;
}
function getPersonnelData(){
  const sh=getSpreadsheet_().getSheetByName(CONFIG.PERSONNEL_SHEET);
  if(!sh)throw new Error('Onglet "'+CONFIG.PERSONNEL_SHEET+'" introuvable.');
  const values=sh.getDataRange().getDisplayValues();
  if(values.length<2)return {employees:[],count:0};
  const hm=headerMap_(values[0]);
  const cMat=findColumn_(hm,["Matricule","N° Matricule","Numero matricule"]);
  const cName=findColumn_(hm,["Nom et Prénom","Nom et Prenom","Nom & Prénom","Nom complet","Nom"]);
  const cFirst=findColumn_(hm,["Prénom","Prenom"]);
  const cFunction=findColumn_(hm,["Fonction","Poste","Emploi"]);
  const cAirport=findColumn_(hm,["Aéroport","Aeroport","Airport"]);
  const cDivision=findColumn_(hm,["Division","Département","Departement"]);
  if(cMat<0)throw new Error('Colonne "Matricule" introuvable dans Personnel.');
  if(cName<0)throw new Error('Colonne "Nom" ou "Nom et Prénom" introuvable dans Personnel.');
  const employees=[];
  for(let r=1;r<values.length;r++){
    const row=values[r],matricule=String(row[cMat]||"").trim();
    if(!matricule)continue;
    let fullName=String(row[cName]||"").trim();
    if(cFirst>=0 && cFirst!==cName)fullName=[fullName,String(row[cFirst]||"").trim()].filter(Boolean).join(" ");
    employees.push({
      matricule,nomPrenom:fullName,
      fonction:cFunction>=0?row[cFunction]:"",
      aeroport:cAirport>=0?row[cAirport]:"",
      division:cDivision>=0?row[cDivision]:""
    });
  }
  return {employees,count:employees.length};
}
function ensureAttendanceSheet_(){
  const ss=getSpreadsheet_();let sh=ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);
  if(!sh){
    sh=ss.insertSheet(CONFIG.ATTENDANCE_SHEET);
    sh.appendRow(["Date","Matricule","Nom et Prénom","Fonction","Aéroport","Division","Statut","Dernière modification","Utilisateur"]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,9).setFontWeight("bold").setBackground("#172554").setFontColor("#ffffff");
  }
  return sh;
}
function getToday_(){return Utilities.formatDate(new Date(),CONFIG.TIMEZONE,"yyyy-MM-dd");}
function getDashboard(date){
  const target=date||getToday_(),values=ensureAttendanceSheet_().getDataRange().getDisplayValues();
  const records=values.slice(1).filter(r=>r[0]===target);
  const stats={total:records.length,ok:0,abs:0,imp:0};
  records.forEach(r=>{const s=String(r[6]||"").toUpperCase();if(s==="OK")stats.ok++;else if(s==="ABS")stats.abs++;else if(s==="IMP")stats.imp++;});
  return {date:target,stats};
}
function saveAttendance(payload){
  if(!payload||!payload.matricule)throw new Error("Matricule obligatoire.");
  const status=String(payload.statut||"").toUpperCase();
  if(!["ABS","IMP","OK"].includes(status))throw new Error("Statut invalide.");
  const matricule=String(payload.matricule).trim(),date=String(payload.date||getToday_());
  const employee=getPersonnelData().employees.find(e=>e.matricule===matricule);
  if(!employee)throw new Error("Matricule introuvable dans la base du personnel.");
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try{
    const sh=ensureAttendanceSheet_(),values=sh.getDataRange().getDisplayValues();
    let targetRow=-1;
    for(let i=1;i<values.length;i++){
      if(values[i][0]===date&&String(values[i][1]).trim()===matricule){targetRow=i+1;break;}
    }
    const now=Utilities.formatDate(new Date(),CONFIG.TIMEZONE,"yyyy-MM-dd HH:mm:ss");
    let user="";try{user=Session.getActiveUser().getEmail()||"";}catch(e){}
    const row=[date,employee.matricule,employee.nomPrenom,employee.fonction,employee.aeroport,employee.division,status,now,user];
    if(targetRow>0){sh.getRange(targetRow,1,1,row.length).setValues([row]);return {success:true,action:"updated",message:"Le pointage existant a été mis à jour."};}
    sh.appendRow(row);return {success:true,action:"created",message:"Pointage enregistré avec succès."};
  }finally{lock.releaseLock();}
}
function getAttendanceHistory(filters){
  filters=filters||{};const values=ensureAttendanceSheet_().getDataRange().getDisplayValues();
  const rows=values.slice(1).filter(r=>{
    if(filters.date&&r[0]!==filters.date)return false;
    if(filters.statut&&r[6]!==filters.statut)return false;
    if(filters.division&&normalize_(r[5]).indexOf(normalize_(filters.division))===-1)return false;
    if(filters.search&&!normalize_(r[1]+" "+r[2]+" "+r[4]+" "+r[5]).includes(normalize_(filters.search)))return false;
    return true;
  }).reverse().slice(0,500);
  return rows.map(r=>({date:r[0],matricule:r[1],nomPrenom:r[2],fonction:r[3],aeroport:r[4],division:r[5],statut:r[6],modification:r[7],utilisateur:r[8]}));
}
