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
  TIMEZONE: "Africa/Algiers",
  SESSIONS_SHEET: "Sessions",
  TOKEN_TTL_SECONDS: 21600
};

// Identité de l’utilisateur authentifié pendant l’exécution courante.
var REQUEST_USER_ = "";

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
    // Seules l’authentification initiale et la définition du premier mot de passe
    // d’une session sont accessibles sans jeton.
    if (action === "login") {
      result = login_(payload);
      return output_(callback, { success: true, result: result });
    }
    if (action === "setSessionPassword") {
      result = setSessionPassword_(payload);
      return output_(callback, { success: true, result: result });
    }

    const auth = requireAuth_(payload.authToken);
    REQUEST_USER_ = auth.name;
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
      case "updateAttendance":
        result = updateAttendance(payload);
        break;
      case "deleteAttendance":
        result = deleteAttendance(payload);
        break;
      case "getAttendanceHistory":
        result = getAttendanceHistory(payload);
        break;
      case "getReportData":
        // يجب أن تكون النتيجة مصفوفة مباشرة لأن index.html يعالجها كقائمة.
        result = getReportData(payload);
        break;
      case "adminListSessions":
        requireAdmin_(auth);
        result = adminListSessions_();
        break;
      case "adminCreateSession":
        requireAdmin_(auth);
        result = adminCreateSession_(payload);
        break;
      case "adminDeleteSession":
        requireAdmin_(auth);
        result = adminDeleteSession_(payload);
        break;
      case "adminResetSessionPassword":
        requireAdmin_(auth);
        result = adminResetSessionPassword_(payload);
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

/**
 * INITIALISATION ADMIN :
 * 1) Dans les paramètres du projet Apps Script > Propriétés du script,
 *    ajouter ADMIN_INITIAL_PASSWORD avec le mot de passe choisi par le propriétaire.
 * 2) Exécuter une seule fois initializeAdminAccount() depuis l’éditeur.
 * 3) La propriété temporaire est supprimée après création du hash.
 */
function initializeAdminAccount() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("ADMIN_PASSWORD_HASH")) {
    throw new Error("Le compte ADMIN est déjà initialisé.");
  }
  const initialPassword = props.getProperty("ADMIN_INITIAL_PASSWORD");
  if (!initialPassword || initialPassword.length < 10) {
    throw new Error("Définissez ADMIN_INITIAL_PASSWORD (10 caractères minimum) dans les propriétés du script.");
  }
  const salt = Utilities.getUuid() + Utilities.getUuid();
  props.setProperties({
    ADMIN_PASSWORD_SALT: salt,
    ADMIN_PASSWORD_HASH: passwordHash_(initialPassword, salt)
  });
  props.deleteProperty("ADMIN_INITIAL_PASSWORD");
  return "Compte ADMIN initialisé. Vous pouvez déployer la nouvelle version.";
}

function passwordHash_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + ":" + String(password), Utilities.Charset.UTF_8);
  return Utilities.base64Encode(bytes);
}

function ensureSessionsSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(CONFIG.SESSIONS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SESSIONS_SHEET);
    sh.appendRow(["Nom de session", "Salt", "PasswordHash", "Statut", "Date création", "Dernière modification"]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#172554").setFontColor("#ffffff");
  }
  return sh;
}

function issueToken_(name, role) {
  const token = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
  CacheService.getScriptCache().put("AUTH_" + token, JSON.stringify({ name: name, role: role }), CONFIG.TOKEN_TTL_SECONDS);
  return token;
}

function requireAuth_(token) {
  token = String(token || "");
  if (!token) throw new Error("Session expirée. Veuillez vous reconnecter.");
  const raw = CacheService.getScriptCache().get("AUTH_" + token);
  if (!raw) throw new Error("Session expirée. Veuillez vous reconnecter.");
  const auth = JSON.parse(raw);
  // Révoque immédiatement l’accès si l’administrateur supprime ou réinitialise la session.
  if (auth.role === "SESSION") {
    const rows = ensureSessionsSheet_().getDataRange().getDisplayValues();
    const active = rows.slice(1).some(r => normalize_(r[0]) === normalize_(auth.name) && String(r[3]).toUpperCase() === "ACTIVE" && !!r[2]);
    if (!active) throw new Error("Session révoquée ou réinitialisée. Veuillez vous reconnecter.");
  }
  return auth;
}

function requireAdmin_(auth) {
  if (!auth || auth.role !== "ADMIN" || auth.name !== "ADMIN") throw new Error("Accès réservé à l’administrateur.");
}

function login_(payload) {
  const name = String(payload.sessionName || "").trim();
  const password = String(payload.password || "");
  if (!name) throw new Error("Nom de session obligatoire.");
  if (normalize_(name) === "admin") {
    if (!password) throw new Error("Mot de passe ADMIN obligatoire.");
    const props = PropertiesService.getScriptProperties();
    const salt = props.getProperty("ADMIN_PASSWORD_SALT");
    const hash = props.getProperty("ADMIN_PASSWORD_HASH");
    if (!salt || !hash) throw new Error("Le compte ADMIN n’est pas encore initialisé. Consultez les instructions de configuration.");
    if (passwordHash_(password, salt) !== hash) throw new Error("Mot de passe incorrect.");
    return { token: issueToken_("ADMIN", "ADMIN"), name: "ADMIN", role: "ADMIN" };
  }
  const sh = ensureSessionsSheet_();
  const rows = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < rows.length; i++) {
    if (normalize_(rows[i][0]) !== normalize_(name)) continue;
    if (String(rows[i][3]).toUpperCase() !== "ACTIVE") throw new Error("Cette session est désactivée.");
    if (!rows[i][2]) return { mustSetPassword: true, sessionName: rows[i][0] };
    if (passwordHash_(password, rows[i][1]) !== rows[i][2]) throw new Error("Mot de passe incorrect.");
    return { token: issueToken_(rows[i][0], "SESSION"), name: rows[i][0], role: "SESSION" };
  }
  throw new Error("Nom de session introuvable. Contactez l’administrateur.");
}

function setSessionPassword_(payload) {
  const name = String(payload.sessionName || "").trim();
  const password = String(payload.password || "");
  if (!name || normalize_(name) === "admin") throw new Error("Nom de session invalide.");
  if (password.length < 8) throw new Error("Le mot de passe doit contenir au moins 8 caractères.");
  const sh = ensureSessionsSheet_();
  const rows = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < rows.length; i++) {
    if (normalize_(rows[i][0]) !== normalize_(name)) continue;
    if (String(rows[i][3]).toUpperCase() !== "ACTIVE") throw new Error("Cette session est désactivée.");
    if (rows[i][2]) throw new Error("Le mot de passe est déjà défini. Connectez-vous normalement ou contactez l’administrateur pour le réinitialiser.");
    const salt = Utilities.getUuid() + Utilities.getUuid();
    const now = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss");
    sh.getRange(i + 1, 2, 1, 5).setValues([[salt, passwordHash_(password, salt), "ACTIVE", rows[i][4] || now, now]]);
    return { token: issueToken_(rows[i][0], "SESSION"), name: rows[i][0], role: "SESSION" };
  }
  throw new Error("Nom de session introuvable. Contactez l’administrateur.");
}

function adminListSessions_() {
  const rows = ensureSessionsSheet_().getDataRange().getDisplayValues();
  return rows.slice(1).filter(r => r[0]).map(r => ({ name: r[0], status: r[3] || "ACTIVE", passwordSet: !!r[2], createdAt: r[4] || "" }));
}

function adminCreateSession_(payload) {
  const name = String(payload.sessionName || "").trim();
  if (!name || name.length > 80 || normalize_(name) === "admin") throw new Error("Nom de session invalide (1 à 80 caractères). ADMIN est réservé.");
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = ensureSessionsSheet_();
    const rows = sh.getDataRange().getDisplayValues();
    if (rows.slice(1).some(r => normalize_(r[0]) === normalize_(name))) throw new Error("Cette session existe déjà.");
    const now = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss");
    sh.appendRow([name, "", "", "ACTIVE", now, now]);
    return { message: "Session créée. Son titulaire pourra choisir son mot de passe à la première connexion." };
  } finally { lock.releaseLock(); }
}

function adminDeleteSession_(payload) {
  const name = String(payload.sessionName || "").trim();
  if (!name || normalize_(name) === "admin") throw new Error("Le compte ADMIN ne peut pas être supprimé.");
  const sh = ensureSessionsSheet_(); const rows = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < rows.length; i++) {
    if (normalize_(rows[i][0]) === normalize_(name)) { sh.deleteRow(i + 1); return { message: "Session supprimée." }; }
  }
  throw new Error("Session introuvable.");
}

function adminResetSessionPassword_(payload) {
  const name = String(payload.sessionName || "").trim();
  if (!name || normalize_(name) === "admin") throw new Error("Nom de session invalide.");
  const sh = ensureSessionsSheet_(); const rows = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < rows.length; i++) {
    if (normalize_(rows[i][0]) === normalize_(name)) {
      const now = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss");
      sh.getRange(i + 1, 2, 1, 5).setValues([["", "", "ACTIVE", rows[i][4] || now, now]]);
      return { message: "Mot de passe réinitialisé. Le titulaire devra en choisir un nouveau." };
    }
  }
  throw new Error("Session introuvable.");
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
    let user = REQUEST_USER_ || "";
    if (!user) { try { user = Session.getActiveUser().getEmail() || ""; } catch (ignore) {} }

    const row = [
      date, employee.matricule, employee.nomPrenom, employee.fonction,
      employee.aeroport, employee.division, status, now, user
    ];

    if (targetRow > 0) {
      if (date !== getToday_()) {
        throw new Error("Modification refusée : seuls les pointages du jour peuvent être modifiés.");
      }
      sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
      return { success: true, action: "updated", message: "Le pointage du jour a été mis à jour." };
    }

    sheet.appendRow(row);
    return { success: true, action: "created", message: "Pointage enregistré avec succès." };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Modifier le statut d'un pointage : uniquement pour la date du jour.
 * La règle est vérifiée côté serveur, pas seulement dans l'interface.
 */
function updateAttendance(payload) {
  payload = payload || {};
  const matricule = String(payload.matricule || "").trim();
  const date = normalizeDate_(payload.date || "");
  const status = String(payload.statut || "").trim().toUpperCase();
  if (!matricule || !date) throw new Error("Matricule et date obligatoires.");
  if (date !== getToday_()) throw new Error("Modification refusée : seuls les pointages du jour sont modifiables.");
  if (["ABS", "IMP", "OK"].indexOf(status) === -1) throw new Error("Statut invalide.");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = ensureAttendanceSheet_();
    const values = sheet.getDataRange().getDisplayValues();
    for (let i = 1; i < values.length; i++) {
      if (normalizeDate_(values[i][0]) === date && String(values[i][1] || "").trim() === matricule) {
        const rowNumber = i + 1;
        const now = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss");
        let user = REQUEST_USER_ || "";
        if (!user) { try { user = Session.getActiveUser().getEmail() || ""; } catch (ignore) {} }
        sheet.getRange(rowNumber, 7).setValue(status);
        sheet.getRange(rowNumber, 8).setValue(now);
        sheet.getRange(rowNumber, 9).setValue(user);
        return { success: true, message: "Le pointage du jour a été modifié avec succès." };
      }
    }
    throw new Error("Pointage du jour introuvable pour ce matricule.");
  } finally {
    lock.releaseLock();
  }
}

/** Supprimer un pointage : uniquement pour la date du jour. */
function deleteAttendance(payload) {
  payload = payload || {};
  const matricule = String(payload.matricule || "").trim();
  const date = normalizeDate_(payload.date || "");
  if (!matricule || !date) throw new Error("Matricule et date obligatoires.");
  if (date !== getToday_()) throw new Error("Suppression refusée : seuls les pointages du jour sont supprimables.");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = ensureAttendanceSheet_();
    const values = sheet.getDataRange().getDisplayValues();
    for (let i = values.length - 1; i >= 1; i--) {
      if (normalizeDate_(values[i][0]) === date && String(values[i][1] || "").trim() === matricule) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Le pointage du jour a été supprimé." };
      }
    }
    throw new Error("Pointage du jour introuvable pour ce matricule.");
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
      utilisateur: row[8],
      modifiable: normalizeDate_(row[0]) === getToday_()
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
