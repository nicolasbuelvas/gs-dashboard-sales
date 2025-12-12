/**
 * SALES DASHBOARD – Web App version
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
    .addItem("Open Dashboard", "openDashboardWebApp")
    .addToUi();
}

function openDashboardWebApp() {
  const url = getWebAppUrl();
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(
      `<html>
        <body style="font-family:Arial;padding:20px;">
          <p>Opening dashboard...</p>
          <script>window.open("${url}", "_blank");google.script.host.close();</script>
        </body>
      </html>`
    ).setWidth(200).setHeight(80),
    "Opening"
  );
}

function getWebAppUrl() {
  return "https://script.google.com/macros/s/AKfycbwsfCBO2tuBYCZT01_bYhiZCliewV98M4nw4pkA-uQDu7_4DfZE52fUiJ9f9LBTeWvg/exec";
}

function doGet() {
  const dataResp = getSalesData_vC();

  const t = HtmlService.createTemplateFromFile('Ui');
  t.INITIAL_RAW = JSON.stringify(dataResp.raw || []);

  return t.evaluate()
    .setTitle("Sales Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSalesData_vC() {
  try {
    const resp = Sheets.Spreadsheets.Values.get(SPREADSHEET_ID, SHEET_NAME);
    return { raw: resp.values || [] };
  } catch (e) {
    return { error: String(e), raw: [] };
  }
}

function exportDataSheetToPDF() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("Sheet not found: " + SHEET_NAME);

    const fileId = ss.getId();
    const sheetId = sheet.getSheetId();

    const exportUrl =
      "https://docs.google.com/spreadsheets/d/" +
      fileId +
      "/export" +
      "?format=pdf" +
      "&portrait=true" +
      "&size=letter" +
      "&fitw=true" +
      "&sheetnames=false" +
      "&printtitle=false" +
      "&pagenumbers=false" +
      "&gridlines=true" +
      "&fzr=false" +
      "&gid=" + sheetId;

    const token = ScriptApp.getOAuthToken();

    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true
    });

    const blob = response.getBlob().setName(
      ss.getName() + " - " + SHEET_NAME + ".pdf"
    );

    // ⬇⬇⬇ Instead of saving to Drive, return blob as Base64
    return {
      success: true,
      fileName: blob.getName(),
      base64: Utilities.base64Encode(blob.getBytes())
    };

  } catch (e) {
    return { error: String(e) };
  }
}

function debugSalesPipeline() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { error: "Sheet not found: " + SHEET_NAME };

  const rows = sheet.getDataRange().getValues();
  const headers = rows[0].map(h => String(h ?? "").trim());

  const missing = EXPECTED_HEADERS.filter(h => !headers.includes(h));
  const sample = normalizeRowsInMemory_debug(rows.slice(0, 51));

  return {
    rowsLoaded: rows.length,
    normalizedSampleCount: sample.length,
    missingHeaders: missing
  };
}

function normalizeRowsInMemory_debug(rows) {
  const headers = rows[0].map(h => String(h ?? "").trim());
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const out = [];

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const temp = {};
    let valSales = 0;
    let valOrder = 0;

    EXPECTED_HEADERS.forEach(h => {
      let v = idx.hasOwnProperty(h) ? row[idx[h]] : "";

      if (["PRICEEACH","SALES","MSRP"].includes(h)) {
        const num = parseFloat(String(v).replace(/\./g, "").replace(",", "."));
        temp[h] = isNaN(num) ? 0 : num;
        if (h === "SALES") valSales = temp[h];
      }
      else if (["ORDERNUMBER","QUANTITYORDERED","ORDERLINENUMBER","QTR_ID","MONTH_ID","YEAR_ID"].includes(h)) {
        const num = parseInt(String(v).replace(/\D/g, ""), 10);
        temp[h] = isNaN(num) ? 0 : num;
        if (h === "ORDERNUMBER") valOrder = temp[h];
      }
      else if (h === "ORDERDATE") {
        if (v instanceof Date) temp[h] = v;
        else {
          const p = String(v).split(" ")[0].split("/");
          if (p.length === 3) {
            const d = new Date(+p[2], p[1] - 1, +p[0]);
            temp[h] = isNaN(d) ? "" : d;
          } else temp[h] = "";
        }
      }
      else temp[h] = String(v).trim();
    });

    if (valSales > 0 && valOrder > 0) out.push(temp);
  }

  return out;
}