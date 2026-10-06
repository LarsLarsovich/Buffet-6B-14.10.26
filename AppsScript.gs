/**
 * Anmeldung — Backend als Google Apps Script.
 *
 * Tabellenblatt "Anmeldungen 2" (wird beim ersten Aufruf automatisch
 * angelegt, falls es noch nicht existiert):
 *
 *        A                B                C          ...  F           G                  H       ...
 *   1   Aktivitäten    Palatschinken    Kassa(5)            Grill
 *   2   (leer)
 *   3   Namen/Zeiten   Samstag 10:00    Samstag 11:00       Samstag 13:00[A]  Fotografieren(4)*  Lesen(10)
 *   4   Max Mustermann
 *   5   Erika Musterfrau
 *   6   ...
 *
 * - Zeile 1, ab Spalte B: die Aktivitäten (Liste endet an der ersten
 *   leeren Zelle). Schreibst du direkt dahinter (OHNE Leerzeichen) eine
 *   Zahl in Klammern, z. B. "Kassa(5)", dürfen sich pro Zeitslot bis zu 5
 *   Personen für diese Aktivität eintragen. Ohne Klammer gilt automatisch
 *   1 Person.
 *
 * - Zeile 3, Spalte B bis F (FEST, max. 5 Zeitslots): die Zeitslots/
 *   Uhrzeiten — werden 1:1 so auf der Webseite angezeigt (ohne eine
 *   eventuelle Buchstaben-Markierung, siehe unten).
 *
 * - Zeile 3, ab Spalte G (FEST, unabhängig davon wie viele Zeitslots
 *   benutzt werden): die Ankreuz-Aktivitäten ohne Zeitbezug (Liste endet
 *   an der ersten leeren Zelle). Format: "Name(Max)*" — "(Max)" (optional,
 *   Standard 1) legt die maximale Anzahl Personen fest, ein "*" am Ende
 *   (optional) markiert sie als Teil der Pflicht-Gruppe: mindestens eine
 *   "*"-Aktivität muss gewählt werden, SOLANGE NICHT jede einzelne
 *   "*"-Aktivität bereits ihr eigenes Max. erreicht hat — ist jede "*"
 *   voll, entfällt die Pflicht für alle, die sich danach anmelden.
 *
 * - Buchstaben-Verknüpfung: Trägst du bei einer Aktivität (Zeile 1) UND
 *   einem Zeitslot (Zeile 3, Spalte B–F) direkt dahinter (OHNE
 *   Leerzeichen) denselben Großbuchstaben in eckigen Klammern ein, z. B.
 *   "Mittagessen(6)[A]" und "13:00[A]", wird bei diesem einen Zeitslot
 *   NUR NOCH diese eine Aktivität angezeigt — alle anderen (auch
 *   unmarkierte) sind für diesen Zeitslot gesperrt. Reihenfolge immer
 *   erst runde, dann eckige Klammer. Hat ein Zeitslot/eine Aktivität
 *   einen Buchstaben ohne Gegenstück, wird er wie unverknüpft behandelt.
 *
 * - Spalte A, ab Zeile 4: die Namensliste. Jede Liste endet bei der
 *   ersten leeren Zelle — keine Lücken lassen.
 *
 * - In den Zellen ab B4 (Spalte B bis F) steht die Aktivität, die diese
 *   Person in diesem Zeitslot macht (ohne Klammer-Zusätze, leer = nichts
 *   gebucht). In den Zellen ab G4 steht bei den Ankreuz-Aktivitäten
 *   einfach "X" (leer = nicht angekreuzt).
 *
 * Löschen/Entfernen ist über die Webseite NICHT möglich — nur wer dieses
 * Sheet bearbeiten kann, kann eine Zelle leeren und den Platz wieder
 * freigeben.
 *
 * Einrichtung:
 * 1. Im bestehenden Google Sheet: Erweiterungen -> Apps Script.
 * 2. Kompletten Inhalt dieser Datei einfügen (vorhandenen Code ersetzen).
 * 3. Bereitstellen -> Verwalten -> neue Version wählen -> bereitstellen
 *    (die URL bleibt dabei gleich).
 */

const SHEET_NAME = "Anmeldungen 2";
const ACTIVITY_ROW = 1;
const SLOT_ROW = 3;
const NAME_START_ROW = 4;
const START_COL = 2; // Spalte B
const TIME_SLOT_MAX_COL = START_COL + 4; // Spalte F — max. 5 Zeitslots (B–F)
const CHECK_START_COL = 7; // Spalte G — fix, Start der Ankreuz-Aktivitäten

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(ACTIVITY_ROW, 1).setValue("Aktivitäten");
    sheet.getRange(SLOT_ROW, 1).setValue("Namen/Zeiten");
  }
  return sheet;
}

// Liest die ANGEZEIGTEN Werte (nicht die internen, z. B. bei Uhrzeiten),
// damit auf der Webseite exakt das steht, was im Sheet zu sehen ist.
// Liest ab startCol bis zur ersten leeren Zelle (kein festes Limit).
function readRowUntilBlank_(sheet, row, startCol) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < startCol) return [];
  const vals = sheet.getRange(row, startCol, 1, lastCol - startCol + 1).getDisplayValues()[0];
  const out = [];
  for (const v of vals) {
    const s = String(v || "").trim();
    if (!s) break;
    out.push(s);
  }
  return out;
}

// Wie oben, aber zusätzlich auf eine feste Endspalte begrenzt (für den
// Zeitslot-Bereich, der nicht in den fixen Ankreuz-Bereich hineinlaufen darf).
function readRowRange_(sheet, row, startCol, endCol) {
  if (endCol < startCol) return [];
  const lastCol = Math.min(endCol, sheet.getLastColumn());
  if (lastCol < startCol) return [];
  const vals = sheet.getRange(row, startCol, 1, lastCol - startCol + 1).getDisplayValues()[0];
  const out = [];
  for (const v of vals) {
    const s = String(v || "").trim();
    if (!s) break;
    out.push(s);
  }
  return out;
}

function readColUntilBlank_(sheet, col, startRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];
  const vals = sheet.getRange(startRow, col, lastRow - startRow + 1, 1).getDisplayValues();
  const out = [];
  for (const r of vals) {
    const s = String(r[0] || "").trim();
    if (!s) break;
    out.push(s);
  }
  return out;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// "Mittagessen(6)[A]" -> { name: "Mittagessen", capacity: 6, letter: "A" }
// "Kassa(5)" -> { name: "Kassa", capacity: 5, letter: null }
// "Grill[B]" -> { name: "Grill", capacity: 1, letter: "B" }
// "Grill" -> { name: "Grill", capacity: 1, letter: null }
// Klammern müssen direkt (ohne Leerzeichen davor) angehängt sein, sonst
// zählen sie als normaler Teil des Namens.
function parseActivity_(raw) {
  let s = String(raw || "").trim();
  let letter = null;

  const letterMatch = s.match(/\[([A-Z])\]$/);
  if (letterMatch) {
    const before = s.slice(0, s.length - letterMatch[0].length);
    if (!/\s$/.test(before)) {
      letter = letterMatch[1];
      s = before.trim();
    }
  }

  let capacity = 1;
  const capMatch = s.match(/^(.+)\((\d+)\)$/);
  if (capMatch && !/\s$/.test(capMatch[1])) {
    capacity = Math.max(1, parseInt(capMatch[2], 10) || 1);
    s = capMatch[1].trim();
  }

  return { name: s, capacity: capacity, letter: letter };
}

// "13:00[A]" -> { name: "13:00", letter: "A" }
// "13:00" -> { name: "13:00", letter: null }
function parseSlot_(raw) {
  let s = String(raw || "").trim();
  let letter = null;

  const letterMatch = s.match(/\[([A-Z])\]$/);
  if (letterMatch) {
    const before = s.slice(0, s.length - letterMatch[0].length);
    if (!/\s$/.test(before)) {
      letter = letterMatch[1];
      s = before.trim();
    }
  }

  return { name: s, letter: letter };
}

// "Fotografieren(4)*" -> { name: "Fotografieren", capacity: 4, required: true }
// "Lesen(10)" -> { name: "Lesen", capacity: 10, required: false }
// "Spenden*" -> { name: "Spenden", capacity: 1, required: true }
function parseCheckboxItem_(raw) {
  let s = String(raw || "").trim();
  let required = false;

  if (s.endsWith("*")) {
    required = true;
    s = s.slice(0, -1).trim();
  }

  let capacity = 1;
  const capMatch = s.match(/^(.+)\((\d+)\)$/);
  if (capMatch && !/\s$/.test(capMatch[1])) {
    capacity = Math.max(1, parseInt(capMatch[2], 10) || 1);
    s = capMatch[1].trim();
  }

  return { name: s, capacity: capacity, required: required };
}

function readState_(sheet) {
  const rawActivities = readRowUntilBlank_(sheet, ACTIVITY_ROW, START_COL);
  const parsedActivities = rawActivities.map(parseActivity_);
  const activities = parsedActivities.map(function (a) { return a.name; });
  const activityCapacity = {};
  const activityLetter = {};
  parsedActivities.forEach(function (a) {
    activityCapacity[a.name] = a.capacity;
    if (a.letter) activityLetter[a.name] = a.letter;
  });

  const rawSlots = readRowRange_(sheet, SLOT_ROW, START_COL, TIME_SLOT_MAX_COL);
  const parsedSlots = rawSlots.map(parseSlot_);
  const slots = parsedSlots.map(function (s) { return s.name; });
  const slotLetter = {};
  parsedSlots.forEach(function (s) { if (s.letter) slotLetter[s.name] = s.letter; });

  const rawCheck = readRowUntilBlank_(sheet, SLOT_ROW, CHECK_START_COL);
  const parsedCheck = rawCheck.map(parseCheckboxItem_);
  const checkItems = parsedCheck.map(function (c) { return c.name; });
  const checkCapacity = {};
  const checkRequired = {};
  parsedCheck.forEach(function (c) {
    checkCapacity[c.name] = c.capacity;
    checkRequired[c.name] = c.required;
  });

  const names = readColUntilBlank_(sheet, 1, NAME_START_ROW);

  const roster = {};
  if (names.length && slots.length) {
    const range = sheet.getRange(NAME_START_ROW, START_COL, names.length, slots.length).getDisplayValues();
    names.forEach(function (name, r) {
      roster[name] = {};
      slots.forEach(function (slot, c) { roster[name][slot] = String(range[r][c] || "").trim(); });
    });
  }

  const checkRoster = {};
  if (names.length && checkItems.length) {
    const range = sheet.getRange(NAME_START_ROW, CHECK_START_COL, names.length, checkItems.length).getDisplayValues();
    names.forEach(function (name, r) {
      checkRoster[name] = {};
      checkItems.forEach(function (item, c) {
        checkRoster[name][item] = String(range[r][c] || "").trim().toUpperCase() === "X";
      });
    });
  }

  // Buchstaben-Verknüpfung auflösen: pro Zeitslot, welche Aktivitäten erlaubt sind.
  const activitiesByLetter = {};
  Object.keys(activityLetter).forEach(function (name) {
    const L = activityLetter[name];
    if (!activitiesByLetter[L]) activitiesByLetter[L] = [];
    activitiesByLetter[L].push(name);
  });
  const slotsByLetter = {};
  Object.keys(slotLetter).forEach(function (name) {
    const L = slotLetter[name];
    if (!slotsByLetter[L]) slotsByLetter[L] = [];
    slotsByLetter[L].push(name);
  });

  // Aktivitäten, die NICHT exklusiv an einen Zeitslot gebunden sind (kein
  // Buchstabe, oder Buchstabe ohne passenden Zeitslot) — das sind die
  // normalen Spalten der Tabelle.
  const normalActivities = activities.filter(function (name) {
    const al = activityLetter[name];
    if (!al) return true;
    return !(slotsByLetter[al] && slotsByLetter[al].length);
  });

  // Pro Zeitslot: ist er exklusiv verknüpft (slotLinked), und falls ja,
  // welche Aktivität(en) sind dann die einzige(n) erlaubte(n) (slotAllowedActivities).
  const slotLinked = {};
  const slotAllowedActivities = {};
  slots.forEach(function (slot) {
    const L = slotLetter[slot];
    if (L && activitiesByLetter[L] && activitiesByLetter[L].length) {
      slotLinked[slot] = true;
      slotAllowedActivities[slot] = activitiesByLetter[L].slice();
      return;
    }
    slotLinked[slot] = false;
    slotAllowedActivities[slot] = normalActivities.slice();
  });

  return {
    activities: activities,
    activityCapacity: activityCapacity,
    normalActivities: normalActivities,
    slots: slots,
    slotLinked: slotLinked,
    slotAllowedActivities: slotAllowedActivities,
    checkItems: checkItems,
    checkCapacity: checkCapacity,
    checkRequired: checkRequired,
    names: names,
    roster: roster,
    checkRoster: checkRoster
  };
}

function doGet(e) {
  const sheet = getSheet_();
  return jsonOut_(readState_(sheet));
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    if (body.action !== "submitBatch") {
      return jsonOut_({ ok: false, message: "Unbekannte Aktion." });
    }

    const name = String(body.name || "").trim();
    const items = Array.isArray(body.items) ? body.items : [];
    const checkItemsRequested = Array.isArray(body.checkItems)
      ? body.checkItems.map(function (s) { return String(s || "").trim(); }).filter(Boolean)
      : [];
    const maxTotal = Number(body.maxTotalPerPerson) || 999;

    if (!name) return jsonOut_({ ok: false, message: "Kein Name angegeben." });
    if (!items.length && !checkItemsRequested.length) {
      return jsonOut_({ ok: false, message: "Keine Auswahl übermittelt." });
    }

    const state = readState_(sheet);
    const nameRowIdx = state.names.indexOf(name);
    if (nameRowIdx === -1) {
      return jsonOut_({ ok: false, message: "Unbekannter Name (nicht in Spalte A gefunden)." });
    }
    const row = NAME_START_ROW + nameRowIdx;

    // Pflicht-Gruppe (*) vorab prüfen — bevor irgendetwas geschrieben wird.
    const requiredItems = state.checkItems.filter(function (it) { return state.checkRequired[it]; });
    if (requiredItems.length) {
      const alreadyHasRequired = requiredItems.some(function (it) { return (state.checkRoster[name] || {})[it]; });
      const batchHasRequired = checkItemsRequested.some(function (it) { return requiredItems.indexOf(it) !== -1; });
      const groupFull = requiredItems.every(function (it) {
        const cap = state.checkCapacity[it] || 1;
        const count = state.names.filter(function (n) { return (state.checkRoster[n] || {})[it]; }).length;
        return count >= cap;
      });
      if (!alreadyHasRequired && !batchHasRequired && !groupFull) {
        return jsonOut_({ ok: false, message: "Bitte mindestens eine Pflicht-Aktivität (mit *) auswählen." });
      }
    }

    let currentTotal = Object.values(state.roster[name] || {}).filter(function (v) { return v; }).length;
    const usedSlotsThisBatch = {};
    const applied = [];
    const failed = [];

    items.forEach(function (item) {
      const slot = String(item.slot || "").trim();
      const activity = String(item.activity || "").trim();

      if (state.slots.indexOf(slot) === -1) { failed.push({ type: "slot", slot: slot, activity: activity, message: "Unbekannter Zeitslot." }); return; }
      if (state.activities.indexOf(activity) === -1) { failed.push({ type: "slot", slot: slot, activity: activity, message: "Unbekannte Aktivität." }); return; }

      const allowed = state.slotAllowedActivities[slot] || state.activities;
      if (allowed.indexOf(activity) === -1) {
        failed.push({ type: "slot", slot: slot, activity: activity, message: "Aktivität an diesem Zeitslot nicht verfügbar." });
        return;
      }

      const col = START_COL + state.slots.indexOf(slot);

      if ((state.roster[name] || {})[slot] || usedSlotsThisBatch[slot]) {
        failed.push({ type: "slot", slot: slot, activity: activity, message: "In diesem Zeitslot schon etwas anderes gebucht." });
        return;
      }

      const capacity = state.activityCapacity[activity] || 1;
      const occupantColumn = sheet.getRange(NAME_START_ROW, col, state.names.length, 1).getDisplayValues();
      let countOthers = 0;
      for (let i = 0; i < occupantColumn.length; i++) {
        if (i === nameRowIdx) continue;
        if (String(occupantColumn[i][0] || "").trim() === activity) countOthers++;
      }
      if (countOthers >= capacity) {
        failed.push({ type: "slot", slot: slot, activity: activity, message: "Platz ist voll (" + countOthers + "/" + capacity + ")." });
        return;
      }

      if (currentTotal >= maxTotal) {
        failed.push({ type: "slot", slot: slot, activity: activity, message: "Maximale Anzahl erreicht." });
        return;
      }

      sheet.getRange(row, col).setValue(activity);
      usedSlotsThisBatch[slot] = true;
      currentTotal++;
      applied.push({ type: "slot", slot: slot, activity: activity });
    });

    const usedCheckThisBatch = {};
    checkItemsRequested.forEach(function (item) {
      if (state.checkItems.indexOf(item) === -1) {
        failed.push({ type: "check", item: item, message: "Unbekannte Aktivität." });
        return;
      }
      if ((state.checkRoster[name] || {})[item] || usedCheckThisBatch[item]) {
        failed.push({ type: "check", item: item, message: "Bereits angekreuzt." });
        return;
      }

      const col = CHECK_START_COL + state.checkItems.indexOf(item);
      const capacity = state.checkCapacity[item] || 1;
      const occupantColumn = sheet.getRange(NAME_START_ROW, col, state.names.length, 1).getDisplayValues();
      let countOthers = 0;
      for (let i = 0; i < occupantColumn.length; i++) {
        if (i === nameRowIdx) continue;
        if (String(occupantColumn[i][0] || "").trim().toUpperCase() === "X") countOthers++;
      }
      if (countOthers >= capacity) {
        failed.push({ type: "check", item: item, message: "Platz ist voll (" + countOthers + "/" + capacity + ")." });
        return;
      }

      sheet.getRange(row, col).setValue("X");
      usedCheckThisBatch[item] = true;
      applied.push({ type: "check", item: item });
    });

    return jsonOut_({ ok: true, applied: applied, failed: failed });
  } catch (err) {
    return jsonOut_({ ok: false, message: "Fehler: " + err.message });
  } finally {
    lock.releaseLock();
  }
}
