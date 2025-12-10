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
  var html = HtmlService.createTemplateFromFile("ui").evaluate()
    .setTitle("Sales Dashboard")
    .setWidth(1200);
  SpreadsheetApp.getUi().showSidebar(html);
}

function normalizeSheet() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("SalesData");
  if (!sheet) throw new Error("Hoja 'SalesData' no encontrada.");

  const raw = sheet.getDataRange().getValues();
  if (!raw || raw.length <= 1) throw new Error("Hoja vacía o sin datos.");

  const rawHeaders = raw[0].map(h => String(h || "").trim());
  // Mapeo de índice por header (sin limpiar)
  const idxMap = {};
  rawHeaders.forEach((h,i)=> idxMap[h] = i);

  const out = [];
  out.push(EXPECTED_HEADERS.slice());

  for (let r = 1; r < raw.length; r++) {
    const row = raw[r];
    const newRow = [];

    EXPECTED_HEADERS.forEach(h => {
      let val = "";

      if (idxMap.hasOwnProperty(h)) {
        val = row[idxMap[h]];
      } else {
        // Intentar con variantes: mayúsculas/minúsculas, espacios, etc.
        for (const k in idxMap) {
          if (k.toUpperCase().replace(/\s+/g,"") === h) {
            val = row[idxMap[k]];
            break;
          }
        }
      }

      if (val === null || val === undefined || val === "") {
        newRow.push("");
      } else {
        // Número decimal: reconocer coma
        if (["PRICEEACH","SALES","MSRP"].includes(h)) {
          let s = String(val).replace(/\./g, "").replace(",", ".");
          const num = parseFloat(s);
          newRow.push(isNaN(num) ? 0 : num);
        }
        // Enteros
        else if (["ORDERNUMBER","QUANTITYORDERED","ORDERLINENUMBER","QTR_ID","MONTH_ID","YEAR_ID"].includes(h)) {
          const num = parseInt(String(val).replace(/\D/g, ""), 10);
          newRow.push(isNaN(num) ? 0 : num);
        }
        // Fecha
        else if (h === "ORDERDATE") {
          let dd = String(val).trim();
          // esperar formato dd/mm/yyyy HH:MM o mm/dd/yyyy HH:MM
          const parts = dd.split(" ");
          const datePart = parts[0];
          const timePart = parts[1] || "";
          const sp = datePart.split("/");
          if (sp.length === 3) {
            let day = parseInt(sp[0],10);
            let month = parseInt(sp[1],10) - 1;
            let year = parseInt(sp[2],10);
            let hour = 0, min = 0;
            if (timePart) {
              const tp = timePart.split(":");
              hour = parseInt(tp[0],10) || 0;
              min = parseInt(tp[1],10) || 0;
            }
            const dt = new Date(year, month, day, hour, min);
            newRow.push(dt);
          } else {
            newRow.push("");
          }
        }
        else {
          newRow.push(String(val).trim());
        }
      }
    });

    // Filtrar sólo filas con ventas > 0
    const sales = newRow[EXPECTED_HEADERS.indexOf("SALES")];
    const orderNum = newRow[EXPECTED_HEADERS.indexOf("ORDERNUMBER")];
    if (sales > 0 && orderNum > 0) {
      out.push(newRow);
    }
  }

  sheet.clearContents();
  sheet.getRange(1,1,out.length, EXPECTED_HEADERS.length).setValues(out);
  // Puedes agregar formatos si quieres
}

function getSalesData() {
  try { normalizeSheet(); }
  catch(e) { Logger.log("normalizeSheet error: " + e); }

  const sheet = SpreadsheetApp.getActive().getSheetByName("SalesData");
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  const headers = rows.shift();
  return rows.map(r => {
    const obj = {};
    headers.forEach((h,i) => obj[h] = r[i]);
    return obj;
  });
}