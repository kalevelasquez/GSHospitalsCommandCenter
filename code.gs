const ADMIN_EMAILS = ["tobarcristina0616@gmail.com"];

/**
 * Función que se ejecuta al abrir el Google Sheets.
 * Crea el menú personalizado en la barra superior.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('HOSPITAL MANAGER')
      .addItem('🔍 Open Explorer', 'abrirExplorador')
      .addItem('➕ Add New Hospital', 'abrirFormulario')
      .addSeparator()
      .addItem('📊 Run Data Audit', 'ejecutarAuditoria')
      .addItem('🔍 Check Orphan Records', 'detectarRegistrosHuerfanos')
      .addItem('🧹 Clear Search Tool', 'limpiarBuscador')
      .addSeparator()
      .addItem('🔄 Refresh Menu/Dashboard', 'onOpen')
      .addToUi();
}

/**
 * Verifica si el usuario actual es administrador.
 */
function esAdmin() {
  var email = Session.getActiveUser().getEmail();
  return ADMIN_EMAILS.indexOf(email) > -1;
}

/**
 * Borra el contenido de la celda de búsqueda en el Dashboard.
 */
function limpiarBuscador() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName("Dashboard");
  if (dashboard) {
    dashboard.getRange("I8").clearContent(); 
  }
}

/**
 * Obtiene todos los hospitales y los "limpia" para evitar errores de fecha.
 */
function obtenerTodosLosHospitales() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hospitals");
  var datos = sheet.getDataRange().getValues();
  datos.shift(); // Quitar cabecera
  return JSON.parse(JSON.stringify(datos));
}

/**
 * Busca un registro por ID en ambas hojas.
 */
function obtenerDatosPorId(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetHosp = ss.getSheetByName("Hospitals");
  var sheetBill = ss.getSheetByName("Records Billing");
  var idBuscar = String(id).trim();
  
  var datosHosp = sheetHosp.getDataRange().getValues();
  var datosBill = sheetBill.getDataRange().getValues();
  var resultado = { hosp: null, bill: null };
  
  for(var i=0; i<datosHosp.length; i++) {
    if(String(datosHosp[i][0]).trim() === idBuscar) {
      resultado.hosp = JSON.parse(JSON.stringify(datosHosp[i]));
      break;
    }
  }
  for(var j=0; j<datosBill.length; j++) {
    if(String(datosBill[j][0]).trim() === idBuscar) {
      resultado.bill = JSON.parse(JSON.stringify(datosBill[j]));
      break;
    }
  }
  return resultado;
}

/**
 * Procesa la creación o edición de un registro.
 */
function procesarRegistro(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetHosp = ss.getSheetByName("Hospitals");
    var sheetBill = ss.getSheetByName("Records Billing");
    
    var idHosp = data.idHosp || "HOSP-" + new Date().getTime().toString().slice(-6);
    var fechaActual = new Date();
    var usuario = Session.getActiveUser().getEmail() || "System";
    
    var filaHosp = encontrarFilaPorId(sheetHosp, idHosp);
    var filaBill = encontrarFilaPorId(sheetBill, idHosp);
    
    var statusActual = "Active";
    if(filaHosp > 0) {
      statusActual = sheetHosp.getRange(filaHosp, 12).getValue() || "Active";
    }

    var valoresHosp = [
      idHosp, data.hospName, data.state, data.county, data.city, 
      data.healthSystem, data.address, data.zip, data.website, 
      data.phone, data.piVolume, statusActual
    ];

    // Array mapeado con los NUEVOS CAMPOS insertados según tu requerimiento
    var valoresBill = [
      idHosp, data.hospName, data.mrPhone, data.mrFax, data.mrEmail, data.mrPortal, 
      data.copyService, data.reqMethod, data.procTime, data.feeReq, data.prepayReq, 
      data.hipaaForm, data.clientId, data.hospBillPhone, data.hospBillFax, 
      data.hospBillEmail, data.hospBillAdd, data.physBillComp, data.physBillPhone, 
      data.physBillFax, data.physBillEmail, data.physBillAdd, data.radBillComp, 
      data.radBillPhone, data.radBillFax, data.radBillEmail, data.radBillAdd, 
      data.imgBillComp, data.imgBillPhone, data.imgBillFax, data.imgBillEmail, data.imgBillAdd,
      // --- NUEVOS CAMPOS INSERTADOS AQUÍ ---
      data.irPhone, data.irFax, data.irEmail, data.irPortal, 
      data.copyServiceUsed, data.prefReqMethod, data.stdProcTime, 
      data.recFeeReq, data.prepayReqBill, data.hospHipaa, data.reqClientId,
      // -------------------------------------
      data.lienPhone, data.lienEmail, data.lienVendor, data.lienContact, 
      data.erDeptName, data.erBillComp, data.erBillContact, data.specInstr, 
      data.delays, data.intNotes, fechaActual, usuario
    ];

    if(filaHosp > 0 && filaBill > 0) {
      sheetHosp.getRange(filaHosp, 1, 1, valoresHosp.length).setValues([valoresHosp]);
      sheetBill.getRange(filaBill, 1, 1, valoresBill.length).setValues([valoresBill]);
      return "✅ Record updated successfully: " + idHosp;
    } else {
      sheetHosp.appendRow(valoresHosp);
      sheetBill.appendRow(valoresBill);
      return "✅ New record created successfully: " + idHosp;
    }
  } catch (e) { return "❌ Error: " + e.toString(); }
}

/**
 * Mueve un hospital a la papelera (Deleted).
 */
function eliminarHospital(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetHosp = ss.getSheetByName("Hospitals");
  var fila = encontrarFilaPorId(sheetHosp, id);
  if (fila > 0) {
    var fecha = new Date();
    var usuario = Session.getActiveUser().getEmail() || "System";
    sheetHosp.getRange(fila, 12, 1, 3).setValues([["Deleted", fecha, usuario]]);
    return "✅ Record moved to Recycle Bin";
  }
  throw "Record not found.";
}

/**
 * Restaura un hospital de la papelera a Active.
 */
function restaurarHospital(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetHosp = ss.getSheetByName("Hospitals");
  var fila = encontrarFilaPorId(sheetHosp, id);
  if (fila > 0) {
    sheetHosp.getRange(fila, 12, 1, 3).setValues([["Active", "", ""]]);
    return "✅ Hospital restored successfully.";
  }
  throw "Error finding record.";
}

/**
 * Elimina físicamente el registro de todas las hojas (Solo Admin).
 */
function eliminarDefinitivamente(id) {
  if (!esAdmin()) throw "Unauthorized: Admin access required";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetHosp = ss.getSheetByName("Hospitals");
  var sheetBill = ss.getSheetByName("Records Billing");
  var filaHosp = encontrarFilaPorId(sheetHosp, id);
  var filaBill = encontrarFilaPorId(sheetBill, id);
  if (filaHosp > 0) sheetHosp.deleteRow(filaHosp);
  if (filaBill > 0) sheetBill.deleteRow(filaBill);
  return "🔥 Record permanently deleted from all sheets";
}

/**
 * Helper para encontrar el número de fila mediante un ID.
 */
function encontrarFilaPorId(hoja, id) {
  if (hoja.getLastRow() < 1) return 0;
  var datos = hoja.getRange(1, 1, hoja.getLastRow()).getValues().flat();
  var idx = datos.indexOf(String(id).trim());
  return idx > -1 ? idx + 1 : 0;
}

// --- FUNCIONES DE APERTURA DE INTERFACES ---

function abrirFormulario() {
  var template = HtmlService.createTemplateFromFile('Formulario');
  template.initialData = null;
  var html = template.evaluate().setWidth(1500).setHeight(980).setTitle('Hospital Master Registry');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function abrirExplorador() {
  var html = HtmlService.createHtmlOutputFromFile('Explorador')
      .setWidth(1500).setHeight(980)
      .setTitle('Hospital Explorer');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function abrirFormularioEdicion(datos) {
  var template = HtmlService.createTemplateFromFile('Formulario');
  template.initialData = datos;
  var html = template.evaluate().setWidth(1500).setHeight(980).setTitle('Edit Record');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * Carga la configuración de Estados y Condados.
 */
function getGeographicConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Geographic Config");
  const values = sheet.getDataRange().getValues();
  const config = {};
  const states = values[0];
  states.forEach((state, colIndex) => {
    if (state) {
      config[state] = [];
      for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
        const county = values[rowIndex][colIndex];
        if (county) config[state].push(county);
      }
    }
  });
  return config;
}

/**
 * Analiza registros activos en busca de campos críticos vacíos.
 */
function ejecutarAuditoria() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Hospitals");
  var datos = sheet.getDataRange().getValues();
  var incompletos = 0;
  var listaIncompletos = [];
  for (var i = 1; i < datos.length; i++) {
    var status = String(datos[i][11]).trim();
    if (status === "Active" || status === "") { 
      if (!datos[i][1] || !datos[i][2] || !datos[i][9]) {
        incompletos++;
        listaIncompletos.push(datos[i][1] || "ID: " + datos[i][0]);
      }
    }
  }
  if (incompletos > 0) {
    SpreadsheetApp.getUi().alert("⚠️ Found " + incompletos + " incomplete records.");
  } else {
    SpreadsheetApp.getUi().alert("✅ Database Healthy.");
  }
}





function repararIdsYRelaciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetHosp = ss.getSheetByName("Hospitals");
  const sheetBill = ss.getSheetByName("Records Billing");
  
  const dataHosp = sheetHosp.getDataRange().getValues();
  const dataBill = sheetBill.getDataRange().getValues();
  
  let ultimoIdNum = 0;

  // 1. Identificar el número más alto de ID actual para no duplicar
  dataHosp.forEach(fila => {
    if (fila[0] && fila[0].toString().startsWith("HOSP-")) {
      let num = parseInt(fila[0].toString().split("-")[1]);
      if (num > ultimoIdNum) ultimoIdNum = num;
    }
  });

  // 2. Reparar Hoja de Hospitales y crear un mapa de nombres
  let nombreAIdMap = {};
  for (let i = 1; i < dataHosp.length; i++) {
    let nombreHosp = dataHosp[i][1];
    if (!dataHosp[i][0]) { // Si no tiene ID
      ultimoIdNum++;
      let nuevoId = "HOSP-" + ultimoIdNum.toString().padStart(6, '0');
      sheetHosp.getRange(i + 1, 1).setValue(nuevoId);
      nombreAIdMap[nombreHosp] = nuevoId;
    } else {
      nombreAIdMap[nombreHosp] = dataHosp[i][0];
    }
  }

  // 3. Reparar Hoja de Records Billing usando el mapa de nombres
  // Esto asume que el nombre del hospital coincide en ambas hojas
  for (let j = 1; j < dataBill.length; j++) {
    let nombreHospBill = dataBill[j][1]; // Columna B: Hospital Name
    
    if (!dataBill[j][0]) { // Si la columna A (ID) está vacía
      if (nombreAIdMap[nombreHospBill]) {
        sheetBill.getRange(j + 1, 1).setValue(nombreAIdMap[nombreHospBill]);
      } else {
        // Si el hospital no existe en la primera hoja, creamos un ID nuevo
        ultimoIdNum++;
        let nuevoIdExtra = "HOSP-" + ultimoIdNum.toString().padStart(6, '0');
        sheetBill.getRange(j + 1, 1).setValue(nuevoIdExtra);
        // OJO: Aquí deberías crear también la fila en Hospitals para que no quede huérfano
      }
    }
  }
  
  SpreadsheetApp.getUi().alert("✅ Reparación completada. Se han asignado IDs y vinculado registros por nombre.");
}


/**
 * Busca IDs que existen en una hoja pero faltan en la otra (Registros Huérfanos).
 */
function detectarRegistrosHuerfanos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetHosp = ss.getSheetByName("Hospitals");
  const sheetBill = ss.getSheetByName("Records Billing");
  
  // Obtener todos los IDs (columna A)
  const idsHosp = sheetHosp.getRange(2, 1, sheetHosp.getLastRow() - 1).getValues().flat().map(id => String(id).trim());
  const idsBill = sheetBill.getRange(2, 1, sheetBill.getLastRow() - 1).getValues().flat().map(id => String(id).trim());
  
  let huerfanosEnHosp = []; // ID está en Hospitals pero NO en Billing
  let huerfanosEnBill = []; // ID está en Billing pero NO en Hospitals
  
  // 1. Buscar en Hospitals -> Billing
  idsHosp.forEach((id, index) => {
    if (id && idsBill.indexOf(id) === -1) {
      // Obtenemos el nombre del hospital para que sea fácil identificarlo
      let nombre = sheetHosp.getRange(index + 2, 2).getValue();
      huerfanosEnHosp.push("- " + id + " (" + nombre + ")");
    }
  });
  
  // 2. Buscar en Billing -> Hospitals
  idsBill.forEach((id, index) => {
    if (id && idsHosp.indexOf(id) === -1) {
      let nombre = sheetBill.getRange(index + 2, 2).getValue();
      huerfanosEnBill.push("- " + id + " (" + nombre + ")");
    }
  });
  
  // 3. Mostrar resultados
  let mensaje = "";
  
  if (huerfanosEnHosp.length > 0) {
    mensaje += "⚠️ IDS EN 'HOSPITALS' SIN FACTURACIÓN:\n" + huerfanosEnHosp.join("\n") + "\n\n";
  }
  
  if (huerfanosEnBill.length > 0) {
    mensaje += "⚠️ IDS EN 'RECORDS BILLING' SIN REGISTRO GENERAL:\n" + huerfanosEnBill.join("\n") + "\n\n";
  }
  
  if (mensaje === "") {
    SpreadsheetApp.getUi().alert("✅ ¡Perfecto! No se encontraron registros huérfanos. Ambas hojas están sincronizadas.");
  } else {
    // Mostrar el reporte en una ventana grande
    var htmlOutput = HtmlService
      .createHtmlOutput('<pre style="font-family: sans-serif;">' + mensaje + '</pre>')
      .setWidth(600)
      .setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Reporte de Registros Huérfanos');
  }
}
