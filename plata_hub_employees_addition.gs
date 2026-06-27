// ══════════════════════════════════════════════════════════════════════
//  ДОБАВИТЬ В СУЩЕСТВУЮЩИЙ ПЛАТА СКРИПТ (Google Apps Script)
//  Инструкция:
//  1. Откройте плата Apps Script в Google
//  2. Найдите функцию doGet(e)
//  3. Добавьте строку перед "return jsonResp(...)":
//       if (action === "hub_employees") return getHubEmployeesFromPlata();
//  4. Добавьте функцию getHubEmployeesFromPlata() ниже
//  5. Сохраните и задеплойте новую версию (Deploy → New deployment)
// ══════════════════════════════════════════════════════════════════════

// ── Добавить в doGet(e) ──────────────────────────────────────────────
//
//  function doGet(e) {
//    const action = e.parameter.action || "";
//    ...
//    if (action === "hub_employees") return getHubEmployeesFromPlata();  // ← ДОБАВИТЬ
//    ...
//  }

// ── Новая функция ────────────────────────────────────────────────────

function getHubEmployeesFromPlata() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Ищем лист с сотрудниками (подбери правильное имя листа!)
    // Типичные названия: "Radnici", "Zaposleni", "Employees", "Sheet1"
    const EMPLOYEE_SHEET = "Radnici"; // ← ИЗМЕНИ если другое название

    const sheet = ss.getSheetByName(EMPLOYEE_SHEET);
    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: "Sheet not found: " + EMPLOYEE_SHEET }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(h => h.toString().toLowerCase().trim());

    // Автоопределение колонок по заголовкам
    const nameCol = findCol(headers, ["name","ime","radnik","zaposleni","naziv"]);
    const pinCol  = findCol(headers, ["pin","lozinka","pass","password"]);
    const posCol  = findCol(headers, ["position","pozicija","radno mjesto","role"]);

    if (nameCol === -1 || pinCol === -1) {
      return ContentService
        .createTextOutput(JSON.stringify({
          error: "Columns not found",
          headers: headers,
          tip: "Sheet must have columns for name (Ime/Radnik) and pin (PIN)"
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const emps = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const name = (row[nameCol] || "").toString().trim();
      const pin  = (row[pinCol]  || "").toString().trim();
      if (!name || !pin) continue;
      emps.push({
        name:     name,
        position: posCol >= 0 ? (row[posCol] || "").toString().trim() : "",
        pin:      pin
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify(emps))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Вспомогательная функция — найти колонку по возможным заголовкам
function findCol(headers, candidates) {
  for (const h of headers) {
    for (const c of candidates) {
      if (h.includes(c)) return headers.indexOf(h);
    }
  }
  return -1;
}
