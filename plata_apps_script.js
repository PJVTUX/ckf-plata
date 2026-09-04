// ═══════════════════════════════════════════════════════════════════
// CKF PLATA — Google Apps Script  v4.1
// Развернуть → Управление развёртыванием → ✏️ → Новая версия → Развернуть
// ═══════════════════════════════════════════════════════════════════

const SHEET_ID      = "1TFRG00D1MwvJYSyf7o3cPGXqLgIhx-7_uJmqtkf7x9M";
const SHEET_NAME    = "PLATA";
const RASPORED_NAME = "RASPORED";

const COL_LABEL = 2; // Колонка C (0-based) — метки строк в RASPORED
const COL_DAY1  = 4; // Колонка E (0-based) — день 1 в RASPORED

function doGet(e) {
  const pin    = (e.parameter.pin    || "").trim();
  const action = (e.parameter.action || "").trim().toLowerCase();
  const debug  = (e.parameter.debug  || "") === "1";

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    // ── Режим отладки ────────────────────────────────────────────────
    if (debug) {
      const sheet = ss.getSheetByName(SHEET_NAME);
      const data  = sheet.getDataRange().getValues();
      return jsonResponse({
        headers: data[0], row1: data[1] || [], row2: data[2] || [], totalRows: data.length
      });
    }

    // ── Календарь (RASPORED) ─────────────────────────────────────────
    if (action === "calendar") {
      const target = (e.parameter.target || "").trim().toUpperCase();
      return getCalendarData(ss, pin, target);
    }

    // ── Сохранение дня (RASPORED) ────────────────────────────────────
    if (action === "save") {
      const target = (e.parameter.target || "").trim().toUpperCase();
      const day    = parseInt(e.parameter.day   || "0");
      const smena  = (e.parameter.smena  || "").trim().toUpperCase();
      const hours  = e.parameter.hours !== undefined ? e.parameter.hours : null;
      const upl    = e.parameter.upl    !== undefined ? e.parameter.upl    : null;
      return saveCalendarDay(ss, pin, target, day, smena, hours, upl);
    }

    // ── Статистика прихода (RASPORED) ────────────────────────────────
    if (action === "stats") {
      if (!pin) return jsonResponse({ error: "PIN nije naveden" });
      const plSheet = ss.getSheetByName(SHEET_NAME);
      const plData  = plSheet.getDataRange().getValues();
      const plHdrs  = plData[0].map(h => String(h).trim().toUpperCase());
      const plCl    = {};
      plHdrs.forEach((h, i) => { if (h && plCl[h] === undefined) plCl[h] = i; });
      const pnCl  = plCl["PIN"]      !== undefined ? plCl["PIN"]      : 0;
      const posCl = plCl["POSITION"] !== undefined ? plCl["POSITION"] : null;
      const rdCl  = plCl["RADNIK"]   !== undefined ? plCl["RADNIK"]   : 1;
      const sRow  = plData.find((r, i) => i > 0 && String(r[pnCl]).trim() === pin);
      if (!sRow) return jsonResponse({ error: "PIN nije pronađen" });
      const sPos  = posCl !== null ? String(sRow[posCl] || "").trim().toLowerCase() : "worker";
      const sName = String(sRow[rdCl] || "").trim();
      if (sPos !== "administrator" && sPos !== "owner") {
        return jsonResponse({ error: "Pristup odbijen" });
      }
      return getStatsData(ss, sName, sPos);
    }

    // ── Зарплатные данные (PLATA) ────────────────────────────────────
    if (!pin) return jsonResponse({ error: "PIN nije naveden" });

    const sheet = ss.getSheetByName(SHEET_NAME);
    const data  = sheet.getDataRange().getValues();

    const headers = data[0].map(h => String(h).trim().toUpperCase());
    const col = {};
    headers.forEach((h, i) => { if (h && col[h] === undefined) col[h] = i; });

    const C = (a, b, fallback) => {
      if (col[a] !== undefined) return col[a];
      if (b && col[b] !== undefined) return col[b];
      return fallback;
    };

    const pinCol    = C("PIN",       null,       0);
    const radnikCol = C("RADNIK",    null,       1);
    const periodCol = C("PERIOD",    null,       2);
    const plataCol  = C("PLATA",     null,       3);
    const satnicaCol= C("SATNICA",   null,       4);
    const odrCol    = C("ODRADJENO", "ODRAĐENO", 5);
    const zarCol    = C("ZARADJENO", "ZARAĐENO", 6);
    const uplCol    = C("UPLACENO",  "UPLAĆENO", 7);
    const dugCol    = C("DUG",       null,       8);
    const godCol    = col["GODISNJI"] !== undefined ? col["GODISNJI"] : null;
    const posCol    = col["POSITION"] !== undefined ? col["POSITION"] : null;

    const myRow = data.find((r, i) => i > 0 && String(r[pinCol]).trim() === pin);
    if (!myRow) return jsonResponse({ error: "PIN nije pronađen" });

    const myPosition = posCol !== null ? String(myRow[posCol] || "").trim().toLowerCase() : "worker";
    const myName     = String(myRow[radnikCol] || "").trim();
    const isAdmin    = myPosition === "administrator";
    const canSeeAll  = isAdmin || myPosition === "owner";

    const records = [];
    for (let i = 1; i < data.length; i++) {
      const row    = data[i];
      const rowPin = String(row[pinCol] || "").trim();
      if (!rowPin) continue;
      if (!canSeeAll && rowPin !== pin) continue;

      const rowPos = posCol !== null ? String(row[posCol] || "").trim().toLowerCase() : "worker";
      if (rowPos === "owner") continue;

      const rec = {
        radnik:    String(row[radnikCol] || "").trim(),
        position:  rowPos,
        period:    String(row[periodCol] || "").trim(),
        PLATA:     toNum(row[plataCol]),
        satnica:   String(row[satnicaCol] || "").trim(),
        ODRADJENO: toNum(row[odrCol]),
        ZARADJENO: toNum(row[zarCol]),
        UPLACENO:  toNum(row[uplCol]),
        DUG:       toNum(row[dugCol]),
      };
      if (canSeeAll && godCol !== null) rec.GODISNJI = toNum(row[godCol]);
      records.push(rec);
    }

    if (records.length === 0) return jsonResponse({ error: "PIN nije pronađen" });

    records.sort((a, b) => {
      const p = String(b.period).localeCompare(String(a.period));
      if (p !== 0) return p;
      const aA = a.position === "administrator" ? 0 : 1;
      const bA = b.position === "administrator" ? 0 : 1;
      if (aA !== bA) return aA - bA;
      return String(a.radnik).localeCompare(String(b.radnik));
    });

    return jsonResponse({ radnik: myName, position: myPosition, records });

  } catch(err) {
    return jsonResponse({ error: "Greška servera: " + err.message });
  }
}

// ── Чтение расписания из RASPORED ───────────────────────────────────
function getCalendarData(ss, pin, targetName) {
  if (!pin) return jsonResponse({ error: "PIN nije naveden" });

  // Аутентификация через PLATA
  const plataSheet = ss.getSheetByName(SHEET_NAME);
  const plataData  = plataSheet.getDataRange().getValues();

  const headers = plataData[0].map(h => String(h).trim().toUpperCase());
  const col = {};
  headers.forEach((h, i) => { if (h && col[h] === undefined) col[h] = i; });

  const pinCol    = col["PIN"]      !== undefined ? col["PIN"]      : 0;
  const radnikCol = col["RADNIK"]   !== undefined ? col["RADNIK"]   : 1;
  const posCol    = col["POSITION"] !== undefined ? col["POSITION"] : null;

  const myRow = plataData.find((r, i) => i > 0 && String(r[pinCol]).trim() === pin);
  if (!myRow) return jsonResponse({ error: "PIN nije pronađen" });

  const myName     = String(myRow[radnikCol] || "").trim().toUpperCase();
  const myPosition = posCol !== null ? String(myRow[posCol] || "").trim().toLowerCase() : "worker";
  const canSeeAll  = myPosition === "administrator" || myPosition === "owner";

  // Список сотрудников для admin/owner (без owner)
  const employeeList = [];
  if (canSeeAll) {
    const seen = {};
    for (let i = 1; i < plataData.length; i++) {
      const r   = plataData[i];
      const rPn = String(r[pinCol] || "").trim();
      if (!rPn) continue;
      const rPs = posCol !== null ? String(r[posCol] || "").trim().toLowerCase() : "worker";
      if (rPs === "owner") continue;
      const rNm = String(r[radnikCol] || "").trim().toUpperCase();
      if (rNm && !seen[rNm]) { seen[rNm] = true; employeeList.push(rNm); }
    }
  }

  // Какого сотрудника показывать
  const showName = (canSeeAll && targetName) ? targetName : myName;

  // Читаем RASPORED
  const rasporedSheet = ss.getSheetByName(RASPORED_NAME);
  if (!rasporedSheet) return jsonResponse({ error: "List RASPORED nije pronađen" });

  const data = rasporedSheet.getDataRange().getValues();

  // Находим блок сотрудника (строка где C = имя)
  let empRow = -1;
  for (let i = 0; i < data.length; i++) {
    const lbl = String(data[i][COL_LABEL] || "").trim().toUpperCase();
    if (lbl === showName) { empRow = i; break; }
  }

  if (empRow === -1) {
    return jsonResponse({ error: "Zaposleni nije pronađen u rasporedu", employees: employeeList });
  }

  // Ищем строки MAJ/месяц, SMENA, ODRADJENO, UPLAĆENO в блоке сотрудника
  const MONTHS = ["JAN","FEB","MAR","APR","MAJ","JUN","JUL","AVG","SEP","OKT","NOV","DEC",
                  "JANUAR","FEBRUAR","MART","APRIL","JUNI","JULI","AVGUST","SEPTEMBAR","OKTOBAR","NOVEMBAR","DECEMBAR"];
  let majRow = -1, smenaRow = -1, odrRow = -1, uplRow = -1, godRow = -1;

  const KNOWN_LABELS = ["SATNICA","ZARADJENO","ZARAĐENO","UPLACENO","UPLAĆENO",
                        "DUG","ZVANICNO","GODISNJI","GODISNJE","GODIŠNJI","GODIŠNJE",
                        "SMENA","ODRADJENO","ODRAĐENO"];

  for (let i = empRow + 1; i < Math.min(empRow + 25, data.length); i++) {
    const lbl = String(data[i][COL_LABEL] || "").trim().toUpperCase();
    if (MONTHS.some(m => lbl === m || lbl.startsWith(m + " "))) { majRow   = i; continue; }
    if (lbl === "SMENA")                                          { smenaRow = i; continue; }
    if (lbl === "ODRADJENO" || lbl === "ODRAĐENO")               { odrRow   = i; continue; }
    if (lbl === "UPLACENO"  || lbl === "UPLAĆENO"  ||
        lbl === "UPLAČENO"  || lbl === "PLACENO"   ||
        lbl === "PLAĆENO"   || lbl === "PLAČENO")                { uplRow   = i; continue; }
    if (lbl === "GODISNJI"  || lbl === "GODISNJE"  ||
        lbl === "GODIŠNJI"  || lbl === "GODIŠNJE")               { godRow   = i; continue; }
    // Если встретили следующего сотрудника — стоп
    if (lbl !== "" && !KNOWN_LABELS.includes(lbl) &&
        MONTHS.every(m => !lbl.startsWith(m)) &&
        majRow >= 0 && smenaRow >= 0 && odrRow >= 0) break;
  }

  // Собираем данные по дням
  const days = [];
  for (let d = 0; d < 31; d++) {
    const colIdx  = COL_DAY1 + d;
    const dayCell = majRow >= 0 ? String(data[majRow][colIdx] || "").trim() : "";
    if (!dayCell) break; // Конец месяца

    // Имя дня — только буквы (убираем цифры типа "PET 1" → "PET")
    const dayName = dayCell.replace(/\d+/g, "").trim().toUpperCase();
    let smena     = smenaRow >= 0 ? String(data[smenaRow][colIdx] || "").trim().toUpperCase() : "";
    // Нормализация: старые значения 1/2 → P/D
    if (smena === "1") smena = "P";
    else if (smena === "2") smena = "D";

    const hours   = odrRow >= 0 ? toNum(data[odrRow][colIdx]) : 0;
    const upl     = uplRow >= 0 ? toNum(data[uplRow][colIdx]) : 0;
    const god     = godRow >= 0 ? toNum(data[godRow][colIdx]) : 0;

    days.push({ num: d + 1, name: dayName, smena: smena, hours: hours, upl: upl, god: god });
  }

  const monthLabel = majRow >= 0 ? String(data[majRow][COL_LABEL] || "").trim().toUpperCase() : "";

  return jsonResponse({
    month:     monthLabel,
    days:      days,
    employee:  showName,
    employees: employeeList
  });
}

// ── Сохранение одного дня в RASPORED ────────────────────────────────
function saveCalendarDay(ss, pin, targetName, day, smena, hoursRaw, uplRaw) {
  if (!pin) return jsonResponse({ error: "PIN nije naveden" });
  if (!day || day < 1 || day > 31) return jsonResponse({ error: "Neispravan dan" });

  // Аутентификация — только administrator
  const plataSheet = ss.getSheetByName(SHEET_NAME);
  const plataData  = plataSheet.getDataRange().getValues();
  const headers    = plataData[0].map(h => String(h).trim().toUpperCase());
  const col = {};
  headers.forEach((h, i) => { if (h && col[h] === undefined) col[h] = i; });

  const pinCol    = col["PIN"]      !== undefined ? col["PIN"]      : 0;
  const radnikCol = col["RADNIK"]   !== undefined ? col["RADNIK"]   : 1;
  const posCol    = col["POSITION"] !== undefined ? col["POSITION"] : null;

  const myRow = plataData.find((r, i) => i > 0 && String(r[pinCol]).trim() === pin);
  if (!myRow) return jsonResponse({ error: "PIN nije pronađen" });

  const myPosition = posCol !== null ? String(myRow[posCol] || "").trim().toLowerCase() : "worker";
  const myName2    = String(myRow[radnikCol] || "").trim().toUpperCase();
  const isSelf     = !targetName || targetName.toUpperCase() === myName2;
  // Администратор может писать за любого; обычный сотрудник — только за себя
  if (myPosition !== "administrator" && !isSelf) {
    return jsonResponse({ error: "Pristup odbijen" });
  }

  // Имя сотрудника для записи
  const showName = targetName || myName2;

  // Читаем RASPORED
  const rasporedSheet = ss.getSheetByName(RASPORED_NAME);
  if (!rasporedSheet) return jsonResponse({ error: "List RASPORED nije pronađen" });
  const data = rasporedSheet.getDataRange().getValues();

  // Находим блок сотрудника
  let empRow = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][COL_LABEL] || "").trim().toUpperCase() === showName) {
      empRow = i; break;
    }
  }
  if (empRow === -1) return jsonResponse({ error: "Zaposleni nije pronađen u rasporedu" });

  // Ищем строки SMENA, ODRADJENO, UPLACENO
  const MONTHS = ["JAN","FEB","MAR","APR","MAJ","JUN","JUL","AVG","SEP","OKT","NOV","DEC",
                  "JANUAR","FEBRUAR","MART","APRIL","JUNI","JULI","AVGUST","SEPTEMBAR","OKTOBAR","NOVEMBAR","DECEMBAR"];
  const KNOWN  = ["SATNICA","ZARADJENO","ZARAĐENO","UPLACENO","UPLAĆENO","UPLAČENO",
                  "PLACENO","PLAĆENO","PLAČENO","DUG","ZVANICNO","GODISNJI","SMENA",
                  "ODRADJENO","ODRAĐENO"];
  let smenaRow = -1, odrRow = -1, uplRow = -1;

  for (let i = empRow + 1; i < Math.min(empRow + 15, data.length); i++) {
    const lbl = String(data[i][COL_LABEL] || "").trim().toUpperCase();
    if (lbl === "SMENA")                                         { smenaRow = i; continue; }
    if (lbl === "ODRADJENO" || lbl === "ODRAĐENO")              { odrRow   = i; continue; }
    if (lbl === "UPLACENO"  || lbl === "UPLAĆENO" ||
        lbl === "UPLAČENO"  || lbl === "PLACENO"  ||
        lbl === "PLAĆENO"   || lbl === "PLAČENO")               { uplRow   = i; continue; }
    if (lbl !== "" && !KNOWN.includes(lbl) && MONTHS.every(m => !lbl.startsWith(m))) break;
  }

  // Колонка дня (0-based → 1-based для Sheets)
  const sheetCol = COL_DAY1 + day - 1 + 1; // +1 за 1-based индекс Sheets

  if (smenaRow >= 0) rasporedSheet.getRange(smenaRow + 1, sheetCol).setValue(smena);
  if (odrRow   >= 0) {
    const h = hoursRaw !== null ? (parseFloat(String(hoursRaw).replace(",",".")) || "") : "";
    rasporedSheet.getRange(odrRow + 1, sheetCol).setValue(h);
  }
  if (uplRow   >= 0) {
    const u = uplRaw !== null ? (parseFloat(String(uplRaw).replace(",",".")) || "") : "";
    rasporedSheet.getRange(uplRow + 1, sheetCol).setValue(u);
  }

  return jsonResponse({ ok: true, employee: showName, day: day });
}

// ── Вспомогательные функции ──────────────────────────────────────────
function toNum(v) {
  if (v === "" || v === null || v === undefined) return 0;
  const n = parseFloat(String(v).replace(",", "."));
  return isNaN(n) ? 0 : Math.round(n * 100) / 100;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── Чтение статистики выручки из RASPORED ───────────────────────────
function getStatsData(ss, myName, myPos) {
  const sheet = ss.getSheetByName(RASPORED_NAME);
  if (!sheet) return jsonResponse({ error: "List RASPORED nije pronađen" });
  const data = sheet.getDataRange().getValues();

  const MONTHS_LIST = ["JAN","FEB","MAR","APR","MAJ","JUN","JUL","AVG","SEP","OKT","NOV","DEC",
                       "JANUAR","FEBRUAR","MART","APRIL","JUNI","JULI","AVGUST","SEPTEMBAR",
                       "OKTOBAR","NOVEMBAR","DECEMBAR"];
  const isMonth = s => MONTHS_LIST.some(m => s === m || s.startsWith(m + " "));

  // Маппинг меток выручки → ключ (нормализация вариантов написания)
  const REV_MAP = {
    "KUCAN KES":"kucan_kes","KUĆAN KES":"kucan_kes","KUCAN":"kucan_kes",
    "GLOVO":"glovo","WOLT":"wolt","KARTICA":"kartica","KES":"kes",
    "UKUPNO":"ukupno","NABAVKA":"nabavka",
    "PORUDZBINA":"porudzbina","PORUDZB":"porudzbina","NARUDZBINA":"porudzbina"
  };
  const isRevLbl = s => s in REV_MAP;

  // Значения дней из строки RASPORED (31 значение)
  const getDayVals = function(rowIdx) {
    if (rowIdx < 0 || rowIdx >= data.length) return Array(31).fill(0);
    const row = data[rowIdx];
    var vals = [];
    for (var d = 0; d < 31; d++) { vals.push(toNum(row[COL_DAY1 + d])); }
    return vals;
  };

  var monthLabel = "";
  var revRows    = {}; // ключ → индекс строки

  // Сканируем первые 30 строк (блок выручки в начале RASPORED)
  // Проверяем несколько колонок: C (COL_LABEL), D (COL_LABEL+1), а также A и B
  // т.к. неизвестно точно в какой колонке стоят метки выручки
  var SCAN_COLS = [COL_LABEL, COL_LABEL + 1, COL_LABEL - 1, COL_LABEL - 2];
  for (var i = 0; i < Math.min(30, data.length); i++) {
    for (var ci = 0; ci < SCAN_COLS.length; ci++) {
      var col = SCAN_COLS[ci];
      if (col < 0) continue;
      var lbl = String(data[i][col] || "").trim().toUpperCase();
      if (!lbl) continue;

      if (!monthLabel && isMonth(lbl)) {
        monthLabel = lbl;
      } else if (isRevLbl(lbl) && !(REV_MAP[lbl] in revRows)) {
        revRows[REV_MAP[lbl]] = i;
      } else if (lbl.startsWith("UKUPNO") && (lbl.indexOf("PORUD") >= 0 || lbl.indexOf("NARUDZ") >= 0)) {
        if (!("porudzbina" in revRows)) revRows["porudzbina"] = i;
      }
    }
  }

  // Отладка: если ничего не нашли — вернуть первые строки RASPORED для диагностики
  if (!("ukupno" in revRows)) {
    var dbgRows = [];
    for (var di = 0; di < Math.min(15, data.length); di++) {
      dbgRows.push(data[di].slice(0, 8).map(function(v){ return String(v || ""); }));
    }
    return jsonResponse({ error: "Ukupno row not found", debug_rows: dbgRows, revRows: JSON.stringify(revRows), month: monthLabel });
  }

  var ukupno = getDayVals("ukupno" in revRows ? revRows["ukupno"] : -1);
  var total  = ukupno.reduce(function(s, v) { return s + v; }, 0);

  return jsonResponse({
    me:         { name: myName, position: myPos },
    month:      monthLabel,
    days:       ukupno,
    total:      total,
    channels: {
      glovo:     getDayVals("glovo"     in revRows ? revRows["glovo"]     : -1),
      wolt:      getDayVals("wolt"      in revRows ? revRows["wolt"]      : -1),
      kartica:   getDayVals("kartica"   in revRows ? revRows["kartica"]   : -1),
      kes:       getDayVals("kes"       in revRows ? revRows["kes"]       : -1),
      kucan_kes: getDayVals("kucan_kes" in revRows ? revRows["kucan_kes"] : -1)
    },
    nabavka:    getDayVals("nabavka"    in revRows ? revRows["nabavka"]    : -1),
    porudzbina: getDayVals("porudzbina" in revRows ? revRows["porudzbina"] : -1)
  });
}

// ══════════════════════════════════════════════════════════════════════
// PRISUTNOST — обработка POST-запросов (приход / уход сотрудников)
// Данные пишутся в лист "Prisutnost" этой же таблицы
// ══════════════════════════════════════════════════════════════════════

const PRIS_SHEET = "Prisutnost";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.event === "arrive") return handleArrive(data);
    if (data.event === "leave")  return handleLeave(data);
    return jsonResponse({ ok: false, error: "unknown event" });
  } catch(err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function getShift(date) {
  // P = первая смена (приход до 14:00), D = вторая смена (14:00 и позже)
  const hour = parseInt(Utilities.formatDate(date, "Europe/Belgrade", "HH"), 10);
  return hour < 14 ? "P" : "D";
}

function getPrisSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(PRIS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PRIS_SHEET);
    sheet.appendRow(["Datum", "Zaposleni", "Dolazak", "Odlazak", "Sati", "Smena", "Pozicija"]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function isAlreadyPresent(sheet, employee) {
  const today = Utilities.formatDate(new Date(), "Europe/Belgrade", "dd.MM.yyyy");
  const rows  = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][0] === today && rows[i][1] === employee && !rows[i][3]) return true;
  }
  return false;
}

function handleArrive(data) {
  const sheet = getPrisSheet();
  const now   = new Date();
  const today = Utilities.formatDate(now, "Europe/Belgrade", "dd.MM.yyyy");
  const time  = Utilities.formatDate(now, "Europe/Belgrade", "HH:mm");

  if (isAlreadyPresent(sheet, data.employee)) {
    return jsonResponse({ ok: true, status: "already_present" });
  }

  sheet.appendRow([
    today,           // A: Datum
    data.employee,   // B: Zaposleni
    time,            // C: Dolazak
    "",              // D: Odlazak
    "",              // E: Sati
    getShift(now),   // F: Smena (P / D)
    data.position || ""  // G: Pozicija
  ]);

  return jsonResponse({ ok: true, status: "arrived", time: time });
}

function handleLeave(data) {
  const sheet = getPrisSheet();
  const now   = new Date();
  const today = Utilities.formatDate(now, "Europe/Belgrade", "dd.MM.yyyy");
  const time  = Utilities.formatDate(now, "Europe/Belgrade", "HH:mm");

  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][0] === today && rows[i][1] === data.employee && !rows[i][3]) {
      // Нашли открытую строку прихода — заполняем уход и часы
      const arriveStr = rows[i][2]; // "HH:mm"
      let hoursStr = "";
      if (arriveStr) {
        const [ah, am] = arriveStr.split(":").map(Number);
        const [lh, lm] = time.split(":").map(Number);
        const diff = (lh * 60 + lm) - (ah * 60 + am);
        if (diff > 0) {
          const h = Math.floor(diff / 60), m = diff % 60;
          hoursStr = h + "h " + (m < 10 ? "0" : "") + m + "m";
        }
      }
      const row = i + 1; // 1-indexed
      sheet.getRange(row, 4).setValue(time);     // D: Odlazak
      sheet.getRange(row, 5).setValue(hoursStr); // E: Sati
      return jsonResponse({ ok: true, status: "left", hours: hoursStr });
    }
  }

  // Нет открытого прихода — пишем только уход
  sheet.appendRow([today, data.employee, "", time, "", "", data.position || ""]);
  return jsonResponse({ ok: true, status: "left_no_arrival" });
}
