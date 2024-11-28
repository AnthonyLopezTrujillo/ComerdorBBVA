// Data 
const SS_MENUS_ID = "";
const SS_DATOS_MENUS = SpreadsheetApp.openById(SS_MENUS_ID);
const SHEET_CALENDARIOS_MENU_NOMBRE = "BASE_MENU";
const SHEET_CALENDARIOS_MENUS = SS_DATOS_MENUS.getSheetByName(SHEET_CALENDARIOS_MENU_NOMBRE);

function obtenerCalendariosMenus() {
  var query = `SELECT *`;
  var resultado = executeQuery(SS_MENUS_ID,SHEET_CALENDARIOS_MENU_NOMBRE , query, "A2:K", true);

  var itemResultado = [];
  if (resultado != null && resultado.length > 0) {
    itemResultado = resultado;
  }
  return itemResultado;
}

function abrirMenuFunciones() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Administrar Menus')
    .addItem('🧹 Limpiar Base', 'borrarTextoBaseMenu')
    .addItem('❌ Eliminar Eventos', 'mostrarFormEliminar')
    .addItem('✅ Generar Eventos', 'mostrarFormCrearEventos')
    .addToUi();
}

function mostrarFormEliminar() {
  var html = HtmlService.createHtmlOutputFromFile('FormEliminarEventos.html');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function mostrarFormCrearEventos() {
  var html = HtmlService.createHtmlOutputFromFile('FormCrearEventos.html');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/* ************************** */
/* **** QUERYS GENERALES **** */
/* ************************** */

// Query generico para busquedas en archivos
function executeQuery(spreadSheetId, sheetName, queryFormula, range, showHeader) {
  var lastRow = SpreadsheetApp.openById(spreadSheetId).getSheetByName(sheetName).getLastRow();
  var qvizURL = 'https://docs.google.com/spreadsheets/d/' + spreadSheetId
    + '/gviz/tq?tqx=out:json&headers=1&sheet=' + sheetName
    + '&range=' + range + lastRow
    + '&tq=' + encodeURIComponent(queryFormula);
  console.log("qvizURL=" + qvizURL);
  var ret = UrlFetchApp.fetch(qvizURL, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getContentText();
  var resp = JSON.parse(ret.replace("/*O_o*/", "").replace("google.visualization.Query.setResponse(", "").slice(0, -2));
  var data = resp.table.rows.map(row => {
    return row.c.map(cols => {
      return cols === null ? '' : cols.f !== undefined ? cols.f : (cols.v === null ? '' : cols.v);
    });
  });
  if (showHeader) {
    var header = resp.table.cols.map(col => {
      return col.label;
    });
    return [header].concat(data);
  } else {
    return data;
  }
}
