/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      console.log("Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.");
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = createTable;
    document.getElementById("create-tableCustom").onclick = createCustomTable;
    document.getElementById("filter-table").onclick = filterTable;
    document.getElementById("open-dialog").onclick = openDialog;
    document.getElementById("sort-table").onclick = sortTable;
    document.getElementById("create-chart").onclick = createChart;
    document.getElementById("freeze-header").onclick = freezeHeader;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

async function createTable() {
  await Excel.run(async (context) => {
    // TODO1: Queue table creation logic here.
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "Resumen";
    // TODO2: Queue commands to populate the table with data.
    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add at the end*/, [
      ["1/1/2017", "The Phone Company", "Communications", "120"],
      ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
      ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
      ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
      ["1/11/2017", "Bellows College", "Education", "350.1"],
      ["1/15/2017", "Trey Research", "Other", "135"],
      ["1/15/2017", "Best For You Organics ", "Groceries", "97.88"],
    ]);
    // TODO3: Queue commands to format the table.
    expensesTable.columns.getItemAt(3).getRange().numberFormat = [["\u20AC#,##0.00"]];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();

    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    // eslint-disable-next-line no-undef
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}
async function createCustomTable() {
  await Excel.run(async (context) => {
    //RESETEA LA TABLA
    context.workbook.worksheets.getItemOrNullObject("CPPO").delete();
    const sheet = context.workbook.worksheets.add("CPPO");

    sheet.getRange("A1").values = "Volume in Month of contracted sale/Volume in Monat Abschluss.";
    sheet.getRange("A2").values = "Sales Contracts Development CPPO.";
    sheet.getRange("A3:C3").merge();

    //encabezado
    CrearEncabezado(sheet);
    sheet.getRange("P1").values = [[new Date().toLocaleDateString("en-US")]];
    //Titulo
    let x = 4;
    let anio = 2014;
    for (let index = 0; index < 9; index++) {
      CrearReporteLinea(sheet, anio, x);
      x += 4;
      anio++;
    }

    sheet.activate();
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    // eslint-disable-next-line no-undef
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function CrearReporteLinea(sheet, anio, x) {
  const yearTitle = sheet.getRange(`A${x}:A${x + 3}`);
  yearTitle.merge();
  yearTitle.values = anio;
  //valores
  const Revenue = sheet.getRange(`B${x}`);
  Revenue.values = "Rev";
  const Vol = sheet.getRange(`B${x + 2}`);
  Vol.values = "Vol";
  // "c4"
  const Dolar = sheet.getRange(`C${x}`);
  Dolar.values = "$";
  const promedio = sheet.getRange(`C${x + 1}`);
  promedio.values = "$/MT";
  const Dollar2 = sheet.getRange(`C${x + 2}`);
  Dollar2.values = "$";
  // P4=SUM(D4:O4)
  // =SUM(D5:O5)
  // SUM(D6:O6)
  Sumar(sheet, `=SUM(D${x}:O${x})`, `P${x}`);
  Sumar(sheet, `=SUM(D${x + 1}:O${x + 1})`, `P${x + 1}`);
  Sumar(sheet, `=SUM(D${x + 2}:O${x + 2})`, `P${x + 2}`);
}

function CrearEncabezado(sheet) {
  const data = [["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dec", "Total"]];
  sheet.getRange("D3:p3").values = data;
}

function Sumar(sheet, rango, celda) {
  const sumRangeP6 = sheet.getRange(celda);
  sumRangeP6.formulas = [[rango]];
  sumRangeP6.format.fill.color = "LightBlue";
  sumRangeP6.format.font.bold = true;
}

async function filterTable() {
  await Excel.run(async (context) => {
    // TODO1: Queue commands to filter out all expense categories except
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    // currentWorksheet.load("tables");
    const expensesTable = currentWorksheet.tables.getItem("Resumen");
    expensesTable.load("filter");
    const categoryFilter = expensesTable.columns.getItem("Category").filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    // eslint-disable-next-line no-undef
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}
async function sortTable() {
  await Excel.run(async (context) => {
    // TODO1: Queue commands to sort the table by Merchant name.
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem("Resumen");
    const sortFields = [
      {
        key: 1, // Merchant column
        ascending: false,
      },
    ];
    expensesTable.sort.apply(sortFields);
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    // eslint-disable-next-line no-undef
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}
async function createChart() {
  await Excel.run(async (context) => {
    // TODO1: Queue commands to get the range of data to be charted.
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem("Resumen");
    const dataRange = expensesTable.getDataBodyRange();
    // TODO2: Queue command to create the chart and define its type.
    const chart = currentWorksheet.charts.add("ColumnClustered", dataRange, "Auto");
    // TODO3: Queue commands to position and format the chart.
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "Right";
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = "Value in \u20AC";
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    // eslint-disable-next-line no-undef
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}
async function freezeHeader() {
  await Excel.run(async (context) => {
    // TODO1: Queue commands to keep the header visible when the user scrolls.
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    // eslint-disable-next-line no-undef
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}
function openDialog() {
  // TODO1: Call the Office Common API that opens a dialog
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/popup.html",
    { height: 45, width: 55 },

    // TODO2: Add callback parameter.
    function (result) {
      dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
  );
}
function processMessage(arg) {
  document.getElementById("user-name").innerHTML = arg.message;
  dialog.close();
}
let dialog = null;
