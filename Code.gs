/**
 * GESTION DE PRÉSENCE & ABSENCES — Google Apps Script
 * Configure SPREADSHEET_ID before deployment.
 */

const CONFIG = {
  SPREADSHEET_ID: "COLLER_ICI_ID_DU_GOOGLE_SHEETS",
  PERSONNEL_SHEET: "Personnel",
  ATTENDANCE_SHEET: "Pointages",
  TIMEZONE: "Africa/Algiers"
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Gestion de Présence & Absences")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSpreadsheet_() {
  if (!CONFIG.SPREADSHEET_ID || CONFIG.SPREADSHEET_ID.indexOf("COLLER_ICI") === 0) {
    throw new Error("Veuillez renseigner SPREADSHEET_ID dans Code.gs.");
  }
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function normalize_(v) {
  return String(v == null ? "" : v).trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function headerMap_(headers) {
  const map = {};
  headers.forEach((h, i) => map[normalize_(h)] = i);
  return map;
}

function findColumn_(map, aliases) {
  for (const alias of aliases) {
    const key = normalize_(alias);
    if (Object.prototype.hasOwnProperty.call(map, key)) return map[key];
  }
  return -1;
}

function getPersonnelData() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(CONFIG.PERSONNEL_SHEET);
  if (!sh) throw new Error('Onglet "' + CONFIG.PERSONNEL_SHEET + '" introuvable.');
  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return { employees: [], count: 0 };

  const hm = headerMap_(values[0]);
  const cMat = findColumn_(hm, ["Matricule", "N° Matricule", "Numero matricule"]);
  const cName = findColumn_(hm, ["Nom et Prénom", "Nom et Prenom", "Nom & Prénom", "Nom complet", "Nom"]);
  const cFirst = findColumn_(hm, ["Prénom", "Prenom"]);
  const cFunction = findColumn_(hm, ["Fonction", "Poste", "Emploi"]);
  const cAirport = findColumn_(hm, ["Aéroport", "Aeroport", "Airport"]);
  const cDivision = findColumn_(hm, ["Division", "Département", "Departement"]);

  if (cMat < 0) throw new Error('Colonne "Matricule" introuvable dans Personnel.');
  if (cName < 0) throw new Error('Colonne "Nom" ou "Nom et Prénom" introuvable dans Personnel.');

  const employees = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const matricule = String(row[cMat] || "").trim();
    if (!matricule) continue;
    let fullName = String(row[cName] || "").trim();
    if (cFirst >= 0 && cFirst !== cName) {
      const first = String(row[cFirst] || "").trim();
      fullName = [fullName, first].filter(Boolean).join(" ");
    }
    employees.push({
      matricule,
      nomPrenom: fullName,
      fonction: cFunction >= 0 ? row[cFunction] : "",
      aeroport: cAirport >= 0 ? row[cAirport] : "",
      division: cDivision >= 0 ? row[cDivision] : ""
    });
  }
  return { employees, count: employees.length };
}

function ensureAttendanceSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.ATTENDANCE_SHEET);
    sh.appendRow(["Date", "Matricule", "Nom et Prénom", "Fonction", "Aéroport", "Division", "Statut", "Dernière modification", "Utilisateur"]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#172554").setFontColor("#ffffff");
  }
  return sh;
}

function getToday_() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd");
}

function getDashboard(date) {
  const targetDate = date || getToday_();
  const sh = ensureAttendanceSheet_();
  const values = sh.getDataRange().getDisplayValues();
  const records = values.slice(1).filter(r => r[0] === targetDate);
  const stats = { total: records.length, ok: 0, abs: 0, imp: 0 };
  records.forEach(r => {
    const status = String(r[6] || "").toUpperCase();
    if (status === "OK") stats.ok++;
    else if (status === "ABS") stats.abs++;
    else if (status === "IMP") stats.imp++;
  });
  return { date: targetDate, stats };
}

function saveAttendance(payload) {
  if (!payload || !payload.matricule) throw new Error("Matricule obligatoire.");
  const status = String(payload.statut || "").toUpperCase();
  if (!["ABS", "IMP", "OK"].includes(status)) throw new Error("Statut invalide.");
  const matricule = String(payload.matricule).trim();
  const date = String(payload.date || getToday_());
  const personnel = getPersonnelData().employees;
  const employee = personnel.find(e => e.matricule === matricule);
  if (!employee) throw new Error("Matricule introuvable dans la base du personnel.");

  const sh = ensureAttendanceSheet_();
  const values = sh.getDataRange().getDisplayValues();
  let targetRow = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === date && String(values[i][1]).trim() === matricule) {
      targetRow = i + 1;
      break;
    }
  }

  const now = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss");
  let user = "";
  try { user = Session.getActiveUser().getEmail() || ""; } catch (e) {}
  const row = [date, employee.matricule, employee.nomPrenom, employee.fonction,
    employee.aeroport, employee.division, status, now, user];

  if (targetRow > 0) {
    sh.getRange(targetRow, 1, 1, row.length).setValues([row]);
    return { success: true, action: "updated", message: "Le pointage existant a été mis à jour." };
  }
  sh.appendRow(row);
  return { success: true, action: "created", message: "Pointage enregistré avec succès." };
}

function getAttendanceHistory(filters) {
  filters = filters || {};
  const sh = ensureAttendanceSheet_();
  const values = sh.getDataRange().getDisplayValues();
  const rows = values.slice(1).filter(r => {
    if (filters.date && r[0] !== filters.date) return false;
    if (filters.statut && r[6] !== filters.statut) return false;
    if (filters.division && normalize_(r[5]).indexOf(normalize_(filters.division)) === -1) return false;
    if (filters.search) {
      const q = normalize_(filters.search);
      if (!normalize_(r[1] + " " + r[2] + " " + r[4] + " " + r[5]).includes(q)) return false;
    }
    return true;
  }).reverse().slice(0, 500);
  return rows.map(r => ({
    date: r[0], matricule: r[1], nomPrenom: r[2], fonction: r[3],
    aeroport: r[4], division: r[5], statut: r[6], modification: r[7], utilisateur: r[8]
  }));
}
