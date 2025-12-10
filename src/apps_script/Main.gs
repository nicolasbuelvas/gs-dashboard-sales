/** 
 * SALES DASHBOARD (Sheets API)
 * Google Sheets + Apps Script + Google Sheets Api
*/

const SHEET_NAME = "sales_data_sample";
const SPREADSHEET_ID = SpreadsheetApp.getActive().getId();

const EXPECTED_HEADERS = [
  "ORDERNUMBER","QUANTITYORDERED","PRICEEACH","ORDERLINENUMBER","SALES","ORDERDATE",
  "STATUS","QTR_ID","MONTH_ID","YEAR_ID","PRODUCTLINE","MSRP","PRODUCTCODE",
  "CUSTOMERNAME","PHONE","ADDRESSLINE1","ADDRESSLINE2","CITY","STATE","POSTALCODE",
  "COUNTRY","TERRITORY","CONTACTLASTNAME","CONTACTFIRSTNAME","DEALSIZE"
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Sales Dashboard")
    .addItem("Open Dashboard", "showDashboard")
    .addToUi();
}

function showDashboard() {
  const html = HtmlService.createTemplateFromFile("Ui")
    .evaluate()
    .setWidth(380)
    .setTitle("Sales Dashboard");
  SpreadsheetApp.getUi().showSidebar(html);
}

function getFullscreenHTML() {
  return HtmlService.createTemplateFromFile("Ui").evaluate().getContent();
}

function getSalesData_vC() {
  try {
    const range = SHEET_NAME;
    const resp = Sheets.Spreadsheets.Values.get(SPREADSHEET_ID, range);
    const rows = resp.values || [];
    return { raw: rows };
  } catch (e) {
    return { error: String(e), raw: [] };
  }
}

/**
 * DEBUG PIPELINE
*/

function debugSalesPipeline() {
  const log = msg => Logger.log("🔍 " + msg);

  log("=== START DEBUG PIPELINE ===");

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    log("❌ Sheet not found: " + SHEET_NAME);
    return;
  }
  log("✔ Sheet found");

  let rows;
  try {
    rows = sheet.getDataRange().getValues();
    log("✔ Rows loaded: " + rows.length);
  } catch (e) {
    log("❌ Error reading sheet: " + e);
    return;
  }

  const rawHeaders = rows[0].map(h => String(h ?? "").trim());
  log("Headers: " + JSON.stringify(rawHeaders));

  EXPECTED_HEADERS.forEach(h => {
    if (!rawHeaders.includes(h)) log("⚠ Missing header: " + h);
  });

  log("=== Normalizing ===");
  try {
    const normalized = normalizeRowsInMemory_debug(rows);
    log("✔ Normalized rows: " + normalized.length);
  } catch (e) {
    log("❌ Error normalizing: " + e);
  }

  log("=== END DEBUG PIPELINE ===");
}

function normalizeRowsInMemory_debug(rows) {
  const log = msg => Logger.log("   ↳ " + msg);

  const rawHeaders = rows[0].map(h => String(h ?? "").trim());
  const idxMap = {};
  rawHeaders.forEach((h, i) => idxMap[h] = i);

  log("Index map: " + JSON.stringify(idxMap));

  const out = [];

  for (let r = 1; r < rows.length; r++) {
    log("Row " + r);

    const row = rows[r];
    const temp = {};
    let validSales = 0;
    let validOrder = 0;

    EXPECTED_HEADERS.forEach(h => {
      let val = "";

      if (idxMap.hasOwnProperty(h)) {
        val = row[idxMap[h]];
      } else {
        for (const k in idxMap) {
          if (k.toUpperCase().replace(/\s+/g, "") === h) {
            val = row[idxMap[k]];
            break;
          }
        }
      }

      if (["PRICEEACH","SALES","MSRP"].includes(h)) {
        let s = String(val).replace(/\./g, "").replace(",", ".");
        const num = parseFloat(s);
        temp[h] = isNaN(num) ? 0 : num;
        if (h === "SALES") validSales = temp[h];
      }
      else if (["ORDERNUMBER","QUANTITYORDERED","ORDERLINENUMBER","QTR_ID","MONTH_ID","YEAR_ID"].includes(h)) {
        const num = parseInt(String(val).replace(/\D/g, ""), 10);
        temp[h] = isNaN(num) ? 0 : num;
        if (h === "ORDERNUMBER") validOrder = temp[h];
      }
      else if (h === "ORDERDATE") {
        if (val instanceof Date) temp[h] = val;
        else {
          const p = String(val).split(" ")[0]?.split("/") || [];
          if (p.length === 3) {
            const d = new Date(+p[2], p[1]-1, +p[0]);
            temp[h] = isNaN(d) ? "" : d;
          } else temp[h] = "";
        }
      }
      else {
        temp[h] = String(val).trim();
      }
    });

    if (validSales > 0 && validOrder > 0) {
      out.push(temp);
      log("✔ Added");
    } else {
      log("✖ Discarded");
    }
  }

  return out;
}