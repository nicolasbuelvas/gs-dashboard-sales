# gs-dashboard-sales

**Interactive Sales Dashboard built with Google Sheets + Apps Script**

An interactive sales dashboard that visualizes your Google Sheets data in real time. Provides KPIs, charts, dynamic filters, auto-refresh, and PDF export functionality.

## Features

* **Dynamic Filters:** Filter data by year and product line.
* **KPIs:** Display Total Sales, Total Orders, and Average Order Value.
* **Charts:**

  * Sales by Product Line (Pie Chart)
  * Sales by Month (Line Chart)
  * Top 10 Customers (Bar Chart)
* **Auto-Refresh:** Refresh data automatically at configurable intervals.
* **PDF Export:** Export the current sheet to PDF directly from the dashboard.
* **Debugging:** Quickly check data consistency and missing headers.

## How It Works

1. **Setup:**

   * Open your Google Sheet containing sales data.
   * Ensure the sheet has the following headers (case-sensitive):
     `ORDERNUMBER, QUANTITYORDERED, PRICEEACH, ORDERLINENUMBER, SALES, ORDERDATE, STATUS, QTR_ID, MONTH_ID, YEAR_ID, PRODUCTLINE, MSRP, PRODUCTCODE, CUSTOMERNAME, PHONE, ADDRESSLINE1, ADDRESSLINE2, CITY, STATE, POSTALCODE, COUNTRY, TERRITORY, CONTACTLASTNAME, CONTACTFIRSTNAME, DEALSIZE`

2. **Deploy Web App:**

   * Go to **Extensions → Apps Script → Deploy → New Deployment → Web App**.
   * Set **Execute as:** Me (your account).
   * Set **Who has access:** Anyone (or Anyone with link).
   * Copy the deployment URL and use it to open the dashboard.

3. **Open Dashboard:**

   * From Google Sheets, use the **Sales Dashboard → Open Dashboard** menu.
   * The dashboard will open in a new tab.

4. **Interact:**

   * Use filters to select specific years or products.
   * View KPIs and charts update in real time.
   * Click “Export PDF” to save a snapshot of your data to Google Drive.
   * Optionally enable auto-refresh for live updates.

## Tech Stack

* Google Sheets
* Google Apps Script (Server-side logic, PDF export, data normalization)
* Google Charts (Visualization)
* HTML/CSS/JS (Client-side UI)

## Notes

* Ensure popups are allowed in your browser for the dashboard to open correctly.
* All calculations are done client-side based on normalized data from the sheet.
* The dashboard is designed to be fully interactive and responsive.
* When using Google Sheets, it is recommended to set your region to United States (US). This ensures numeric and date formats (like decimal points and commas) work correctly, avoiding issues with data parsing and calculations.