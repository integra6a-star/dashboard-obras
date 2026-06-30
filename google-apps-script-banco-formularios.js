const SPREADSHEET_ID = "1ogj1kp_x7YB430IUcV_W5uanOXKOl1Bti7J6Xj6FLnY";
const TOKEN = "";

const SHEETS = {
  meioambiente_checklist: "meioambiente_checklists",
  seguranca_inspecao: "seguranca_inspecoes",
  qualidade_checklist: "qualidade_checklists"
};

function doPost(e) {
  const row = JSON.parse((e && e.postData && e.postData.contents) || "{}");
  if (TOKEN && row.token !== TOKEN) {
    return textOutput({ ok:false, error:"token_invalido" });
  }

  const form = row.form || "formularios";
  const sheet = getSheet(form);
  ensureHeader(sheet);
  sheet.appendRow([
    row.id || Utilities.getUuid(),
    form,
    row.createdAt || new Date().toISOString(),
    row.syncedAt || "",
    JSON.stringify(row.payload || {})
  ]);

  return textOutput({ ok:true, id:row.id });
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  const callback = params.callback || "callback";
  const form = params.form || "";
  const records = listRecords(form);
  return ContentService
    .createTextOutput(`${callback}(${JSON.stringify({ ok:true, records })});`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function listRecords(form) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const names = form ? [SHEETS[form] || form] : Object.values(SHEETS);
  const records = [];

  names.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    values.forEach(row => {
      if (!row[0]) return;
      let payload = {};
      try {
        payload = JSON.parse(row[4] || "{}");
      } catch (err) {
        payload = {};
      }
      records.push({
        id:String(row[0]),
        form:String(row[1] || form),
        createdAt:toIso(row[2]),
        syncedAt:toIso(row[3]),
        payload
      });
    });
  });

  return records;
}

function getSheet(form) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const name = SHEETS[form] || form;
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureHeader(sheet) {
  if (sheet.getLastRow() > 0) return;
  sheet.appendRow(["id", "form", "createdAt", "syncedAt", "payload"]);
}

function textOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function toIso(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return value.toISOString();
  }
  return String(value);
}
