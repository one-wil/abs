/**
 * GESTION DE PRÉSENCE & ABSENCES
 * Google Apps Script API — JSONP
 * Hébergement frontend : GitHub Pages / Cloudflare Pages
 * Aucun Cloudflare Worker n'est utilisé.
 */

const CONFIG = {
  SPREADSHEET_ID: "1qOPtJiDc8gBTWlqm9-pnUwoqJDC1nClZmIrdpxcLzqU",
  PERSONNEL_SHEET: "Personnel",
  ATTENDANCE_SHEET: "Pointages",
  TIMEZONE: "Africa/Algiers"
};

/** Point d'entrée appelé par index.html via JSONP. */
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const callback = params.callback || "";

  try {
    const action = String(params.action || "");
    let payload = {};
    if (params.payload) {
      payload = JSON.parse(params.payload);
      if (!payload || typeof payload !== "object" || Array.isArray(payload)) payload = {};
    }

    let result;
    switch (action) {
      case "getPersonnelData":
        result = getPersonnelData();
        break;
      case "getDashboard":
        result = getDashboard(payload.date);
        break;
      case "saveAttendance":
        result = saveAttendance(payload);
        break;
      case "getAttendanceHistory":
        result = getAttendanceHistory(payload);
        break;
      case "getReportData":
        // يجب أن تكون النتيجة مصفوفة مباشرة لأن index.html يعالجها كقائمة.
        result = getReportData(payload);
        break;
      default:
        throw new Error("Action API inconnue : " + action);
    }

    return output_(callback, { success: true, result: result });
  } catch (err) {
    return output_(callback, {
      success: false,
      error: String(err && err.message ? err.message : err)
    });
  }
}

/** إخراج JSONP مع التحقق من اسم callback. */
function output_(callback, data) {
  const json = JSON.stringify(data);
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function getSpreadsheet_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function normalize_(value) {
  return String(value == null ? "" : value).trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function headerMap_(headers) {
  const map = {};
  headers.forEach(function (header, index) {
    map[normalize_(header)] = index;
  });
  return map;
}

function findColumn_(map, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const key = normalize_(aliases[i]);
    if (Object.prototype.hasOwnProperty.call(map, key)) return map[key];
  }
  return -1;
}

/** قراءة قاعدة الموظفين من ورقة Personnel. */
function getPersonnelData() {
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.PERSONNEL_SHEET);
  if (!sheet) throw new Error('Onglet "' + CONFIG.PERSONNEL_SHEET + '" introuvable.');

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return { employees: [], count: 0 };

  const headers = headerMap_(values[0]);
  const cMat = findColumn_(headers, ["Matricule", "N° Matricule", "Numero matricule"]);
  const cName = findColumn_(headers, ["Nom et Prénom", "Nom et Prenom", "Nom & Prénom", "Nom complet", "Nom"]);
  const cFirst = findColumn_(headers, ["Prénom", "Prenom"]);
  const cFunction = findColumn_(headers, ["Fonction", "Poste", "Emploi"]);
  const cAirport = findColumn_(headers, ["Aéroport", "Aeroport", "Airport"]);
  const cDivision = findColumn_(headers, ["Division", "Département", "Departement"]);

  if (cMat < 0) throw new Error('Colonne "Matricule" introuvable dans Personnel.');
  if (cName < 0) throw new Error('Colonne "Nom" ou "Nom et Prénom" introuvable dans Personnel.');

  const employees = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const matricule = String(row[cMat] || "").trim();
    if (!matricule) continue;

    let fullName = String(row[cName] || "").trim();
    if (cFirst >= 0 && cFirst !== cName) {
      fullName = [fullName, String(row[cFirst] || "").trim()].filter(Boolean).join(" ");
    }

    employees.push({
      matricule: matricule,
      nomPrenom: fullName,
      fonction: cFunction >= 0 ? row[cFunction] : "",
      aeroport: cAirport >= 0 ? row[cAirport] : "",
      division: cDivision >= 0 ? row[cDivision] : ""
    });
  }
  return { employees: employees, count: employees.length };
}

/** إنشاء ورقة Pointages عند أول استعمال. */
function ensureAttendanceSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.ATTENDANCE_SHEET);
    sheet.appendRow([
      "Date", "Matricule", "Nom et Prénom", "Fonction", "Aéroport",
      "Division", "Statut", "Dernière modification", "Utilisateur"
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 9)
      .setFontWeight("bold")
      .setBackground("#172554")
      .setFontColor("#ffffff");
  }
  return sheet;
}

function getToday_() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd");
}

/** إحصائيات لوحة التحكم لتاريخ محدد. */
function getDashboard(date) {
  const targetDate = normalizeDate_(date || getToday_());
  const values = ensureAttendanceSheet_().getDataRange().getDisplayValues();
  const records = values.slice(1).filter(function (row) {
    return normalizeDate_(row[0]) === targetDate;
  });

  const stats = { total: records.length, ok: 0, abs: 0, imp: 0 };
  records.forEach(function (row) {
    const status = String(row[6] || "").trim().toUpperCase();
    if (status === "OK") stats.ok++;
    else if (status === "ABS") stats.abs++;
    else if (status === "IMP") stats.imp++;
  });
  return { date: targetDate, stats: stats };
}

/** تسجيل الحضور أو تحديث سجل الموظف في اليوم نفسه دون تكرار. */
function saveAttendance(payload) {
  payload = payload || {};
  const matricule = String(payload.matricule || "").trim();
  if (!matricule) throw new Error("Matricule obligatoire.");

  const status = String(payload.statut || "").trim().toUpperCase();
  if (["ABS", "IMP", "OK"].indexOf(status) === -1) {
    throw new Error("Statut invalide. Valeurs autorisées : ABS, IMP, OK.");
  }

  const date = normalizeDate_(payload.date || getToday_());
  const employee = getPersonnelData().employees.find(function (item) {
    return item.matricule === matricule;
  });
  if (!employee) throw new Error("Matricule introuvable dans la base du personnel.");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = ensureAttendanceSheet_();
    const values = sheet.getDataRange().getDisplayValues();
    let targetRow = -1;

    for (let i = 1; i < values.length; i++) {
      if (normalizeDate_(values[i][0]) === date && String(values[i][1]).trim() === matricule) {
        targetRow = i + 1;
        break;
      }
    }

    const now = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss");
    let user = "";
    try { user = Session.getActiveUser().getEmail() || ""; } catch (ignore) {}

    const row = [
      date, employee.matricule, employee.nomPrenom, employee.fonction,
      employee.aeroport, employee.division, status, now, user
    ];

    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
      return { success: true, action: "updated", message: "Le pointage existant a été mis à jour." };
    }

    sheet.appendRow(row);
    return { success: true, action: "created", message: "Pointage enregistré avec succès." };
  } finally {
    lock.releaseLock();
  }
}

/** سجل تاريخي مع المرشحات. */
function getAttendanceHistory(filters) {
  filters = filters || {};
  const values = ensureAttendanceSheet_().getDataRange().getDisplayValues();

  return values.slice(1).filter(function (row) {
    if (filters.date && normalizeDate_(row[0]) !== normalizeDate_(filters.date)) return false;
    if (filters.statut && String(row[6]).toUpperCase() !== String(filters.statut).toUpperCase()) return false;
    if (filters.division && normalize_(row[5]).indexOf(normalize_(filters.division)) === -1) return false;
    if (filters.search && normalize_(row[1] + " " + row[2] + " " + row[4] + " " + row[5]).indexOf(normalize_(filters.search)) === -1) return false;
    return true;
  }).reverse().slice(0, 500).map(function (row) {
    return {
      date: normalizeDate_(row[0]),
      matricule: row[1],
      nomPrenom: row[2],
      fonction: row[3],
      aeroport: row[4],
      division: row[5],
      statut: row[6],
      modification: row[7],
      utilisateur: row[8]
    };
  });
}

/**
 * REPORTING : retourne TOUJOURS un tableau (array) d'enregistrements.
 * Les filtres envoyés par index.html sont : dateFrom, dateTo,
 * division, aeroport, statut.
 */
function getReportData(filters) {
  filters = filters || {};
  const dateFrom = filters.dateFrom ? normalizeDate_(filters.dateFrom) : "";
  const dateTo = filters.dateTo ? normalizeDate_(filters.dateTo) : "";
  const division = normalize_(filters.division || "");
  const airport = normalize_(filters.aeroport || "");
  const status = String(filters.statut || "").trim().toUpperCase();

  const values = ensureAttendanceSheet_().getDataRange().getDisplayValues();
  if (values.length <= 1) return [];

  const reportRows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const date = normalizeDate_(row[0]);
    const rowStatus = String(row[6] || "").trim().toUpperCase();

    if (!date || !String(row[1] || "").trim()) continue;
    if (dateFrom && date < dateFrom) continue;
    if (dateTo && date > dateTo) continue;
    if (status && rowStatus !== status) continue;
    if (division && normalize_(row[5]).indexOf(division) === -1) continue;
    if (airport && normalize_(row[4]).indexOf(airport) === -1) continue;

    reportRows.push({
      date: date,
      matricule: String(row[1] || ""),
      nomPrenom: String(row[2] || ""),
      fonction: String(row[3] || ""),
      aeroport: String(row[4] || ""),
      division: String(row[5] || ""),
      statut: rowStatus
    });
  }

  reportRows.reverse();
  return reportRows;
}

/** توحيد التاريخ إلى yyyy-MM-dd سواء جاء كنص أو كتاريخ من Sheets. */
function normalizeDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, CONFIG.TIMEZONE, "yyyy-MM-dd");
  }
  const text = String(value == null ? "" : value).trim();
  if (!text) return "";

  // yyyy-MM-dd أو yyyy-MM-dd HH:mm:ss
  let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (match) return match[1] + "-" + String(match[2]).padStart(2, "0") + "-" + String(match[3]).padStart(2, "0");

  // dd/MM/yyyy
  match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (match) return match[3] + "-" + String(match[2]).padStart(2, "0") + "-" + String(match[1]).padStart(2, "0");

  return text;
}
