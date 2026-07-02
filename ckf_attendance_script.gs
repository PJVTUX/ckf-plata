// CKF Attendance Tracker — Google Apps Script
// Вставьте этот код в новый Apps Script проект
// Разверните как Web App: Execute as "Me", Who can access "Anyone"

const SHEET_NAME  = "Prisutnost";
const SHEET_SUMM  = "Pregled";
const SHEET_EMP   = "Zaposleni";

// ─── MAIN ENTRY POINT ────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.event === "arrive")  return handleArrive(data);
    if (data.event === "leave")   return handleLeave(data);
    if (data.event === "manual")  return handleManual(data);
    if (data.event === "register") return handleRegister(data);
    return jsonResp({ ok: false, error: "unknown event" });
  } catch(err) {
    return jsonResp({ ok: false, error: err.toString() });
  }
}

function doGet(e) {
  const action = e.parameter.action || "status";
  if (action === "status")        return jsonResp(getCurrentStatus());
  if (action === "today")         return jsonResp(getTodayLog());
  if (action === "history")       return jsonResp(getHistory(e.parameter.days || 7));
  if (action === "employees")     return jsonResp(getEmployees());
  if (action === "empstatus")     return jsonResp(getEmpStatus(e.parameter.name));
  if (action === "hub_employees") return jsonResp(getHubEmployees());
  return jsonResp({ ok: false, error: "unknown action" });
}

// ─── HANDLERS ────────────────────────────────────────────────────────────────

function getShift(date) {
  // P = first shift (arrived before 14:00), D = second shift (14:00 or later)
  const hour = parseInt(Utilities.formatDate(date, "Europe/Belgrade", "HH"), 10);
  return hour < 14 ? "P" : "D";
}

function handleArrive(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAME);
  const now   = new Date();

  // Check if already open (no departure yet for today)
  if (isAlreadyPresent(sheet, data.employee)) {
    return jsonResp({ ok: true, status: "already_present" });
  }

  const smena = getShift(now);

  sheet.appendRow([
    now,                        // A: Timestamp
    data.employee,              // B: Employee name
    "DOLAZAK",                  // C: Event
    Utilities.formatDate(now, "Europe/Belgrade", "HH:mm"), // D: Time
    Utilities.formatDate(now, "Europe/Belgrade", "dd.MM.yyyy"), // E: Date
    "",                         // F: Departure time
    "",                         // G: Hours worked
    data.mac || "",             // H: MAC address
    "auto",                     // I: Source
    smena                       // J: Shift (P = morning, D = afternoon)
  ]);

  return jsonResp({ ok: true, status: "arrived", employee: data.employee, time: Utilities.formatDate(now, "Europe/Belgrade", "HH:mm"), smena: smena });
}

function handleLeave(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAME);
  const now   = new Date();
  const today = Utilities.formatDate(now, "Europe/Belgrade", "dd.MM.yyyy");

  // Find the most recent DOLAZAK row for this employee today without a departure
  const rows = sheet.getDataRange().getValues();
  let arriveRow = -1;
  let arriveTime = null;

  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i];
    if (row[1] === data.employee && row[2] === "DOLAZAK" && row[4] === today && !row[5]) {
      arriveRow = i + 1; // 1-indexed
      arriveTime = row[0];
      break;
    }
  }

  if (arriveRow === -1) {
    // No open arrival found — just log departure
    sheet.appendRow([now, data.employee, "ODLAZAK", Utilities.formatDate(now, "Europe/Belgrade", "HH:mm"), today, "", "", data.mac || "", "auto"]);
    return jsonResp({ ok: true, status: "left_no_arrival" });
  }

  // Calculate hours
  const diffMs    = now - arriveTime;
  const hours     = diffMs / (1000 * 60 * 60);
  const hoursStr  = formatHours(hours);

  // Update the arrival row with departure info
  sheet.getRange(arriveRow, 6).setValue(Utilities.formatDate(now, "Europe/Belgrade", "HH:mm")); // F: departure
  sheet.getRange(arriveRow, 7).setValue(hoursStr); // G: hours

  return jsonResp({ ok: true, status: "left", employee: data.employee, hours: hoursStr });
}

function handleManual(data) {
  // Admin manual entry: { employee, type, date, time }
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAME);
  const dt    = new Date(data.date + " " + data.time);
  const smena = data.type === "arrive" ? getShift(dt) : "";

  sheet.appendRow([dt, data.employee, data.type === "arrive" ? "DOLAZAK" : "ODLAZAK",
    data.time, data.date, "", "", "", "manual", smena]);

  return jsonResp({ ok: true, status: "manual_added" });
}

function handleRegister(data) {
  // Register a new employee: { name, mac, position }
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_EMP);

  // Check if exists
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] && rows[i][1].toString().toLowerCase() === data.mac.toLowerCase()) {
      // Update
      sheet.getRange(i + 1, 1).setValue(data.name);
      sheet.getRange(i + 1, 3).setValue(data.position || "");
      return jsonResp({ ok: true, status: "updated" });
    }
  }

  sheet.appendRow([data.name, data.mac.toLowerCase(), data.position || "", new Date()]);
  return jsonResp({ ok: true, status: "registered" });
}

// ─── QUERIES ─────────────────────────────────────────────────────────────────

function getCurrentStatus() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { ok: true, present: [], absent: [] };

  const now   = new Date();
  const today = Utilities.formatDate(now, "Europe/Belgrade", "dd.MM.yyyy");
  const rows  = sheet.getDataRange().getValues();

  // Who arrived today and hasn't left
  const status = {};
  for (const row of rows) {
    if (!row[0]) continue;
    const date = Utilities.formatDate(new Date(row[0]), "Europe/Belgrade", "dd.MM.yyyy");
    if (date !== today) continue;
    const emp = row[1], event = row[2];
    if (event === "DOLAZAK") status[emp] = { present: true, since: row[3] };
    if (event === "ODLAZAK" || (event === "DOLAZAK" && row[5])) status[emp] = { present: false };
  }

  const present = [], absent = [];
  for (const [emp, s] of Object.entries(status)) {
    if (s.present) present.push({ name: emp, since: s.since });
    else absent.push({ name: emp });
  }

  return { ok: true, present, absent, time: Utilities.formatDate(now, "Europe/Belgrade", "HH:mm dd.MM.yyyy") };
}

function getTodayLog() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { ok: true, log: [] };

  const today = Utilities.formatDate(new Date(), "Europe/Belgrade", "dd.MM.yyyy");
  const rows  = sheet.getDataRange().getValues();
  const log   = [];

  for (const row of rows) {
    if (!row[0]) continue;
    const date = Utilities.formatDate(new Date(row[0]), "Europe/Belgrade", "dd.MM.yyyy");
    if (date !== today) continue;
    log.push({ employee: row[1], event: row[2], time: row[3], hours: row[6] || "" });
  }

  return { ok: true, log, date: today };
}

function getHistory(days) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { ok: true, history: [] };

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - parseInt(days));

  const rows    = sheet.getDataRange().getValues();
  const summary = {};

  for (const row of rows) {
    if (!row[0] || row[2] !== "DOLAZAK" || !row[6]) continue;
    const ts = new Date(row[0]);
    if (ts < cutoff) continue;
    const emp  = row[1];
    const date = row[4];
    const hrs  = parseHours(row[6]);
    if (!summary[emp]) summary[emp] = { total: 0, days: {} };
    if (!summary[emp].days[date]) summary[emp].days[date] = 0;
    summary[emp].days[date] += hrs;
    summary[emp].total += hrs;
  }

  // Format
  const result = [];
  for (const [name, data] of Object.entries(summary)) {
    const daysList = Object.entries(data.days).map(([d, h]) => ({ date: d, hours: formatHours(h) }));
    daysList.sort((a, b) => a.date.localeCompare(b.date));
    result.push({ name, total: formatHours(data.total), days: daysList });
  }

  return { ok: true, history: result, period: days + " dana" };
}

function getEmployees() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_EMP);
  if (!sheet) return { ok: true, employees: [] };

  const rows = sheet.getDataRange().getValues();
  const emps = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) emps.push({ name: rows[i][0], mac: rows[i][1], position: rows[i][2] });
  }
  return { ok: true, employees: emps };
}

// ─── EMPSTATUS — is a specific employee currently at work? ───────────────────
// GET ?action=empstatus&name=Ime+Prezime
// Returns: { present: true/false, since: "HH:mm" }

function getEmpStatus(name) {
  if (!name) return { ok: false, error: "name required" };
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { ok: true, present: false, since: "" };

  const today = Utilities.formatDate(new Date(), "Europe/Belgrade", "dd.MM.yyyy");
  const rows  = sheet.getDataRange().getValues();

  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i];
    if (row[1] === name && row[2] === "DOLAZAK" && row[4] === today && !row[5]) {
      return { ok: true, present: true, since: row[3] || "" };
    }
  }
  return { ok: true, present: false, since: "" };
}

// ─── HUB_EMPLOYEES — return employee list with PINs for hub login ─────────────
// GET ?action=hub_employees
// Returns: [{name, position, pin}]
// Sheet "Zaposleni" must have columns: A=Ime, B=MAC, C=Pozicija, D=Registrovan, E=PIN
// (PIN column E is optional — returns "" if missing)

function getHubEmployees() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_EMP);
  if (!sheet) return [];

  const rows = sheet.getDataRange().getValues();
  const emps = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    emps.push({
      name:     rows[i][0],
      position: rows[i][2] || "",
      pin:      (rows[i][4] || "").toString()
    });
  }
  return emps;
}

// ─── HELPERS ─────────────────────────────────────────────────────────────────

function isAlreadyPresent(sheet, employee) {
  const today = Utilities.formatDate(new Date(), "Europe/Belgrade", "dd.MM.yyyy");
  const rows  = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i];
    if (row[1] === employee && row[2] === "DOLAZAK" && row[4] === today && !row[5]) return true;
  }
  return false;
}

function formatHours(h) {
  const hrs = Math.floor(h);
  const min = Math.round((h - hrs) * 60);
  return hrs + "h " + (min < 10 ? "0" : "") + min + "m";
}

function parseHours(str) {
  if (!str) return 0;
  const m = str.toString().match(/(\d+)h\s*(\d+)m/);
  if (!m) return 0;
  return parseInt(m[1]) + parseInt(m[2]) / 60;
}

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === SHEET_NAME) {
      sheet.appendRow(["Timestamp", "Zaposleni", "Dogadjaj", "Vreme", "Datum", "Odlazak", "Sati", "MAC", "Izvor", "Smena"]);
      sheet.getRange(1, 1, 1, 10).setFontWeight("bold");
    }
    if (name === SHEET_EMP) {
      sheet.appendRow(["Ime", "MAC adresa", "Pozicija", "Registrovan"]);
      sheet.getRange(1, 1, 1, 4).setFontWeight("bold");
    }
  }
  return sheet;
}

function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
