/**
 * SALES DASHBOARD (Sheets API)
 * Google Sheets + Apps Script
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
    .setWidth(900)
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

function exportDataSheetToPDF() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("Sheet not found: " + SHEET_NAME);

    const url = ss.getUrl();
    const exportUrl = url.replace(/\/edit.*$/, '') +
      'export?format=pdf' +
      '&gid=' + sheet.getSheetId() +
      '&size=letter' +
      '&portrait=true' +
      '&fitw=true' +
      '&sheetnames=false&printtitle=false&pagenumbers=false' +
      '&gridlines=true&fzr=false';

    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + token }
    });

    const blob = response.getBlob().setName(ss.getName() + ' - ' + SHEET_NAME + '.pdf');

    try {
      const file = DriveApp.getFileById(ss.getId());
      const parents = file.getParents();
      if (parents.hasNext()) {
        const folder = parents.next();
        folder.createFile(blob);
      } else {
        DriveApp.createFile(blob);
      }
    } catch (e) {
      DriveApp.createFile(blob);
    }

    return { success: true };
  } catch (e) {
    return { error: String(e) };
  }
}

/* Debug pipeline - server-side checks and normalization trace */
function debugSalesPipeline() {
  const log = msg => Logger.log(msg);

  log("START DEBUG PIPELINE");

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    log("Sheet not found: " + SHEET_NAME);
    return { error: "Sheet not found: " + SHEET_NAME };
  }
  log("Sheet found");

  let rows;
  try {
    rows = sheet.getDataRange().getValues();
    log("Rows loaded: " + rows.length);
  } catch (e) {
    log("Error reading sheet: " + e);
    return { error: String(e) };
  }

  const rawHeaders = rows[0].map(h => String(h ?? "").trim());
  log("Headers: " + JSON.stringify(rawHeaders));

  const missing = EXPECTED_HEADERS.filter(h => !rawHeaders.includes(h));
  if (missing.length) {
    log("Missing headers: " + missing.join(", "));
  } else {
    log("All expected headers present.");
  }

  log("Normalizing sample rows (first 50 rows) for inspection");
  const sampleRows = rows.slice(0, 51); // header + up to 50 rows
  const normalizedSample = normalizeRowsInMemory_debug(sampleRows);
  log("Normalized sample count: " + normalizedSample.length);

  log("END DEBUG PIPELINE");
  return { rowsLoaded: rows.length, normalizedSampleCount: normalizedSample.length, missingHeaders: missing };
}

function normalizeRowsInMemory_debug(rows) {
  const log = msg => Logger.log(msg);

  const rawHeaders = rows[0].map(h => String(h ?? "").trim());
  const idxMap = {};
  rawHeaders.forEach((h, i) => idxMap[h] = i);

  log("Index map: " + JSON.stringify(idxMap));

  const out = [];

  for (let r = 1; r < rows.length; r++) {
    log("Processing row " + r);

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
      log("Added");
    } else {
      log("Discarded");
    }
  }

  return out;
}