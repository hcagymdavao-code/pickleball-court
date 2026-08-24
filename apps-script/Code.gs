var SHEET_NAME = "Schedules";
var GUIDE_NAME = "Guide";
var HEADERS = [
  "Date",
  "Time",
  "Slot",
  "Court",
  "Court Label",
  "Notes",
  "Saved At",
  "Id",
];
var ADMIN_PASSWORD = "hca12345";
var SCHEMA_VERSION = 2;
var DEFAULT_COURT_LABELS = {
  1: "Court 1 (Left Wing)",
  2: "Court 2 (Right Wing)",
};

function doGet(e) {
  e = e || { parameter: {} };
  setupIfNeeded();

  var action = e.parameter.action || "list";
  var result;

  try {
    if (action === "list" || action === "setup") {
      result = okResult(readSchedules());
    } else if (action === "save" || action === "delete") {
      if (e.parameter.password !== ADMIN_PASSWORD) {
        result = { ok: false, error: "Incorrect password." };
      } else if (!e.parameter.date) {
        result = { ok: false, error: "Date is required." };
      } else {
        var slots = [];
        if (action === "save" && e.parameter.slots) {
          slots = JSON.parse(e.parameter.slots);
        }
        replaceDate(e.parameter.date, slots);
        result = okResult(readSchedules());
      }
    } else {
      result = { ok: false, error: "Unknown action." };
    }
  } catch (err) {
    result = { ok: false, error: String(err) };
  }

  return jsonp(e, result);
}

function doPost(e) {
  setupIfNeeded();
  var data = {};
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ ok: false, error: "Invalid JSON" });
  }

  if (data.password !== ADMIN_PASSWORD) {
    return json({ ok: false, error: "Incorrect password." });
  }

  if (data.action === "save") {
    replaceDate(data.date, data.slots || []);
  } else if (data.action === "delete") {
    replaceDate(data.date, []);
  } else {
    return json({ ok: false, error: "Unknown action." });
  }

  return json(okResult(readSchedules()));
}

function setupIfNeeded() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (sheet.getName() !== SHEET_NAME) {
    sheet.setName(SHEET_NAME);
  }

  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var firstRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (firstRow[0] !== "Date" || firstRow[1] !== "Time") {
    if (sheet.getLastRow() > 0) {
      sheet.insertRowBefore(1);
    }
    writeHeaders(sheet);
  } else if (firstRow.indexOf("Court") === -1) {
    sheet.insertColumnAfter(3);
    sheet.insertColumnAfter(4);
    writeHeaders(sheet);
    fillMissingCourts(sheet);
  } else {
    writeHeaders(sheet);
  }

  repairShiftedRows(sheet);
  try {
    ensureGuide(ss);
  } catch (err) {
    // Guide is optional. A write failure here must not block saving.
  }
}

function repairShiftedRows(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  var range = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
  var values = range.getValues();
  var changed = false;
  var i;

  for (i = 0; i < values.length; i += 1) {
    var row = values[i];
    var court = String(row[3] || "");
    var label = String(row[4] || "");
    var notes = String(row[5] || "");
    var saved = String(row[6] || "");
    var courtIsCode = court === "1" || court === "2";
    var labelIsIso = /^\d{4}-\d{2}-\d{2}T/.test(label);
    var notesAreIso = /^\d{4}-\d{2}-\d{2}T/.test(notes);

    if (courtIsCode) {
      continue;
    }

    if (labelIsIso) {
      row[7] = notes;
      row[6] = label;
      row[5] = court;
      row[3] = "1";
      row[4] = DEFAULT_COURT_LABELS["1"];
      changed = true;
    } else if (notesAreIso) {
      row[7] = saved;
      row[6] = notes;
      row[5] = court;
      row[3] = "1";
      row[4] = DEFAULT_COURT_LABELS["1"];
      changed = true;
    }
  }

  if (changed) {
    range.setValues(values);
  }
}

function fillMissingCourts(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }
  var values = [];
  var i;
  for (i = 2; i <= lastRow; i += 1) {
    values.push(["1", DEFAULT_COURT_LABELS["1"]]);
  }
  sheet.getRange(2, 4, lastRow - 1, 2).setValues(values);
}

function writeHeaders(sheet) {
  var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setValues([HEADERS]);
  headerRange
    .setFontWeight("bold")
    .setBackground("#1b4332")
    .setFontColor("#ffffff");
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 190);
  sheet.setColumnWidth(4, 70);
  sheet.setColumnWidth(5, 180);
  sheet.setColumnWidth(6, 280);
  sheet.setColumnWidth(7, 200);
  sheet.setColumnWidth(8, 180);
}

function ensureGuide(ss) {
  var guide = ss.getSheetByName(GUIDE_NAME);
  if (guide) {
    return;
  }

  guide = ss.insertSheet(GUIDE_NAME);
  var rows = padRows(
    [
      ["HCA Pickleball Gym — how this sheet works"],
      [""],
      ["One booked hour on one court = one row on the Schedules tab."],
      [
        "The website reads and writes these rows after the admin password is confirmed.",
      ],
      [""],
      ["Column", "Meaning", "Example"],
      ["Date", "Court date as YYYY-MM-DD", "2026-08-25"],
      ["Time", "Hour start in 24-hour time", "06:00"],
      ["Slot", "Display label shown in the app", "6:00 AM – 6:59 AM"],
      ["Court", "1 = Left Wing, 2 = Right Wing", "1"],
      ["Court Label", "Readable court name", "Court 1 (Left Wing)"],
      ["Notes", "Optional session note", "Open play"],
      ["Saved At", "When the row was last written", "2026-08-23T07:18:51.255Z"],
      ["Id", "Unique row id used by the app", "1787469531255-z2j9am"],
      [""],
      ["Do not rename the Schedules tab or these header names."],
    ],
    3,
  );
  guide.getRange(1, 1, rows.length, 3).setValues(rows);
  guide.getRange("A1").setFontWeight("bold").setFontSize(14);
  guide.getRange("A6:C6").setFontWeight("bold").setBackground("#d4e157");
  guide.setColumnWidth(1, 120);
  guide.setColumnWidth(2, 420);
  guide.setColumnWidth(3, 280);
}

function readSchedules() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  var h;
  for (h = 0; h < headers.length; h += 1) {
    map[String(headers[h])] = h;
  }

  var values = sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .getValues();
  var items = [];
  var i;
  for (i = 0; i < values.length; i += 1) {
    var row = values[i];
    var date = normalizeDate(cell(row, map, "Date", 0));
    var time = normalizeTime(cell(row, map, "Time", 1));
    if (
      !date ||
      date === "Date" ||
      !time ||
      time === "Time" ||
      String(cell(row, map, "Slot", 2) || "") === "Slot"
    ) {
      continue;
    }
    var court = String(cell(row, map, "Court", 3) || "1");
    if (court !== "2") {
      court = "1";
    }
    var hour = Number(String(time).split(":")[0]);
    items.push({
      date: date,
      time: time,
      timeLabel: String(cell(row, map, "Slot", 2) || ""),
      court: court,
      courtLabel:
        String(cell(row, map, "Court Label", 4) || "") ||
        DEFAULT_COURT_LABELS[court],
      notes: String(
        cell(row, map, "Notes", map.Notes != null ? map.Notes : 5) || "",
      ),
      savedAt: String(cell(row, map, "Saved At", 6) || ""),
      id: String(cell(row, map, "Id", 7) || ""),
      hour: isNaN(hour) ? 0 : hour,
    });
  }
  return items;
}

function cell(row, map, name, fallback) {
  if (map[name] != null) {
    return row[map[name]];
  }
  return row[fallback];
}

function replaceDate(dateKey, slots) {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = dates.length - 1; i >= 0; i -= 1) {
      if (normalizeDate(dates[i][0]) === dateKey) {
        sheet.deleteRow(i + 2);
      }
    }
  }

  if (!slots || !slots.length) {
    return;
  }

  var rows = slots.map(function (slot) {
    var court = String(slot.court || "1");
    if (court !== "2") {
      court = "1";
    }
    return [
      dateKey,
      slot.time,
      slot.timeLabel || slot.label || "",
      court,
      slot.courtLabel || DEFAULT_COURT_LABELS[court],
      slot.notes || "",
      slot.savedAt || new Date().toISOString(),
      slot.id || Utilities.getUuid(),
    ];
  });
  sheet
    .getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length)
    .setValues(rows);
}

function padRows(rows, columnCount) {
  return rows.map(function (row) {
    var padded = row.slice();
    while (padded.length < columnCount) {
      padded.push("");
    }
    return padded.slice(0, columnCount);
  });
}

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function normalizeDate(value) {
  if (
    Object.prototype.toString.call(value) === "[object Date]" &&
    !isNaN(value)
  ) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
  }
  return String(value || "").slice(0, 10);
}

function normalizeTime(value) {
  if (
    Object.prototype.toString.call(value) === "[object Date]" &&
    !isNaN(value)
  ) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
  }
  return String(value || "");
}

function okResult(schedules) {
  return {
    ok: true,
    schemaVersion: SCHEMA_VERSION,
    schedules: schedules,
  };
}

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
    ContentService.MimeType.JSON,
  );
}

function jsonp(e, data) {
  var payload = JSON.stringify(data);
  var callback = e && e.parameter && e.parameter.callback;
  if (callback) {
    return ContentService.createTextOutput(
      callback + "(" + payload + ")",
    ).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json(data);
}
