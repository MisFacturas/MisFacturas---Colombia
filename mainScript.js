function onOpen() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();


  ui.createAddonMenu()
    .addItem('Inicio', 'showSidebar')
    .addToUi();

  console.log("setActiveSheet Inicio");

  var sheet = ss.getSheetByName("Inicio");
  SpreadsheetApp.setActiveSheet(sheet);

  console.log("onOpenReturning");
  return;

}

function onOpen(e) {

  let ui = SpreadsheetApp.getUi();

  Logger.log("ScriptApp.AuthMode.NONE")
  ui.createAddonMenu()
    .addItem('Inicio', 'showSidebar2')
    .addItem('Instalar', 'IniciarMisfacturas')
    .addItem("Desinstalar", "eliminarHojasFactura").addToUi();

  return;
}

function OnOpenSheetInicio() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Inicio");
  SpreadsheetApp.setActiveSheet(sheet);
}

function showSidebar2() {
  console.log("showSidebar2 Enters");
  let ui = SpreadsheetApp.getUi();
  console.log("setActiveSheet2 Inicio");
  let requiredSheets = ["Inicio", "Productos", "Datos de emisor", "Clientes", "Factura", "ListadoEstado", "ClientesInvalidos", "Copia de Factura", "Datos", "Listado Facturas", "Listado Facturas Data"];
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  for (let sheetName of requiredSheets) {
    let sheet = ss.getSheetByName(sheetName);
    if (sheet == null) {
      let respuesta = ui.alert(`Faltan hojas de calculo requeridas para el funcionamiento de misfacturas. Primero debes de instalar las hojas necesarias ¿Deseas instalarlas ya?`, ui.ButtonSet.YES_NO);
      if (respuesta == ui.Button.YES) {
        iniciarHojasFactura();
        OnOpenSheetInicio();
        agregarDataValidations();
        let htmlOutput = HtmlService.createHtmlOutput(plantillaVincularMF()).setWidth(500).setHeight(250);
        ui.showModalDialog(htmlOutput, 'Vinculación requerida');
      } else {
        return;
      }
    }
  }

  var html = HtmlService.createHtmlOutputFromFile('main')
    .setTitle('Menú');
  SpreadsheetApp.getUi()
    .showSidebar(html);

  console.log("showSidebar Exits");
}

function grantAccessToTemplate() {
  const plantillaID = "1FgLge7RWvu3R-51Se6ekjN3pOTZE40HLwafRIwoIMUg";
  const plantilla = SpreadsheetApp.openById(plantillaID);
  const userEmail = Session.getEffectiveUser().getEmail();
  plantilla.addEditor(userEmail);
}

function revokeAccessToTemplate() {
  const plantillaID = "1FgLge7RWvu3R-51Se6ekjN3pOTZE40HLwafRIwoIMUg";
  const plantilla = SpreadsheetApp.openById(plantillaID);
  const userEmail = Session.getEffectiveUser().getEmail();
  plantilla.removeEditor(userEmail);
}

function iniciarHojasFactura() {
  Logger.log("Inicio instalación de hojas");

  grantAccessToTemplate();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const plantillaID = "1FgLge7RWvu3R-51Se6ekjN3pOTZE40HLwafRIwoIMUg";
  const plantilla = SpreadsheetApp.openById(plantillaID);

  const nombresHojas = ["Inicio", "Productos", "Datos de emisor", "Clientes", "Factura", "ListadoEstado", "ClientesInvalidos", "Copia de Factura", "Datos", "Listado Facturas", "Listado Facturas Data"];
  const hojasInvisibles = ["ListadoEstado", "Datos", "ClientesInvalidos", "Copia de Factura", "Listado Facturas Data"];
  const hojaBloqueada = ["Inicio"]

  // Instalar hojas desde la plantilla si no existen
  nombresHojas.forEach(nombreHoja => {
    if (nombreHoja === "Datos") return; // Saltar "Datos" para instalarla al final

    let hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) {
      const hojaPlantilla = plantilla.getSheetByName(nombreHoja);
      if (hojaPlantilla) {
        // Copiar hoja
        const hojaCopia = hojaPlantilla.copyTo(ss).setName(nombreHoja);

        // Bloquear la hoja completa si está en la lista de bloqueadas e invisibles
        if (hojasInvisibles.includes(nombreHoja)) {
          hojaCopia.hideSheet(); // Hacer la hoja invisible
        }

        if (hojaBloqueada.includes(nombreHoja)) {
          const protection = hojaCopia.protect();
          protection.removeEditors(protection.getEditors()); // Bloquear completamente
          protection.addEditor(Session.getEffectiveUser()); // Solo el propietario tiene acceso
        }

      } else {
        SpreadsheetApp.getUi().alert('La hoja "' + nombreHoja + '" no existe en la plantilla.');
      }
    }
  });

  // Siempre instalar o reinstalar la hoja "Datos" al final
  reinstalarHojaDatos(ss, plantilla);

  // Eliminar hojas que no pertenezcan a la lista de hojas instaladas
  ss.getSheets().forEach(hoja => {
    const nombreHoja = hoja.getName();
    if (!nombresHojas.includes(nombreHoja)) {
      Logger.log(hoja.getName())
      Logger.log("hojaname")
      ss.deleteSheet(hoja);
    }
  });


  revokeAccessToTemplate();

  SpreadsheetApp.getUi().alert("Hojas instaladas satisfactoriamente.");
  SpreadsheetApp.getUi().alert("Recuerda que tu configuracion regional del sheet debe de estar en Estados Unidos para su correcto funcionamiento.");
  //SpreadsheetApp.getUi().alert("Recuerda que antes de utilizar misfacturas debes de crear la carpeta donde se guardarán las facturas. Dirígete a la hoja Datos de emisor y dale clic en el botón crear carpeta.");
}

function agregarDataValidations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDatos = ss.getSheetByName("Datos");
  const hojaFacturas = ss.getSheetByName("Factura");
  const hojaValoresC = ss.getSheetByName("Clientes");
  const hojaValoresP = ss.getSheetByName("Productos");
  const hojaValoresCInvalidos = ss.getSheetByName("ClientesInvalidos");
  const HojaValorescopiaFactura = ss.getSheetByName("Copia de Factura");

  // Rango donde aplicar los dropdowns
  const rangoDropdownCliente = hojaDatos.getRange("H2");
  const rangoDropdownClienteInvalido = hojaDatos.getRange("H6");
  const rangoDropdownProductos = hojaDatos.getRange("I11");
  const rangoDropdownClienteF = hojaFacturas.getRange("B2:C2");
  const rangoDropdownProductoF = hojaFacturas.getRange("B15");
  const rangoDropdownCopiaFacturaCliente = HojaValorescopiaFactura.getRange("B2:C2");
  const rangoDropdownCopiaFacturaProducto = HojaValorescopiaFactura.getRange("B15");
  const rangoDropdownFacturaMedioDePago = hojaFacturas.getRange("J3");
  const rangoDropdownFacturaMoneda = hojaFacturas.getRange("J4");
  const rangoDropdownCopiaFacturaMedioDePago = HojaValorescopiaFactura.getRange("J3");
  const rangoDropdownCopiaFacturaMoneda = HojaValorescopiaFactura.getRange("J4");
  const rangoDropdownProductosColumnaD = hojaValoresP.getRange("D2:D");
  const rangoDropdownProductosColumnaG = hojaValoresP.getRange("G2:G");
  const rangoDropdownProductosColumnaI = hojaValoresP.getRange("I2:I");
  const rangoDropdownProductosColumnaK = hojaValoresP.getRange("K2:K");
  const rangoDropdownProductosColumnaM = hojaValoresP.getRange("M2:M");
  const rangoDropdownClientesColumnaI = hojaValoresC.getRange("I2:I");
  const rangoDropdownClientesColumnaM = hojaValoresC.getRange("M2:M");
  const rangoDropdownClientesColumnaN = hojaValoresC.getRange("N2:N");
  const rangoDropdownClientesColumnaU = hojaValoresC.getRange("U2:U");



  // Rango de valores para los dropdowns
  const rangoValoresClienteInvalido = hojaValoresCInvalidos.getRange("W2:W");
  const rangoValoresClienteDatos = hojaValoresC.getRange("W2:W");
  const rangoValoresProductosDatos = hojaValoresP.getRange("P2:P");
  const rangoValoresClienteFactura = hojaValoresC.getRange("$W$2:$W");
  const rangoValoresProductosFactura = hojaValoresP.getRange("$P$2:$P");
  const rangoValoresFacturaMedioDePago = hojaDatos.getRange("R18:R37");
  const rangoValoresFacturaMoneda = hojaDatos.getRange("X18:X196");
  const rangoValoresCopiaFacturaMedioDePago = hojaDatos.getRange("R18:R937");
  const rangoValoresCopiaFacturaMoneda = hojaDatos.getRange("X18:X196");
  const rangoValoresProductosColumnaD = hojaDatos.getRange("F36:F43");
  const rangoValoresProductosColumnaG = hojaDatos.getRange("B35:B399");
  const rangoValoresProductosColumnaI = hojaDatos.getRange("K26:K29");
  const rangoValoresProductosColumnaK = hojaDatos.getRange("K31:K35");
  const rangoValoresProductosColumnaM = hojaDatos.getRange("M26:M68");
  const rangoValoresClientesColumnaI = hojaDatos.getRange("D3:D14");
  const rangoValoresClientesColumnaM = hojaDatos.getRange("A24:A195");
  const rangoValoresClientesColumnaN = hojaDatos.getRange("F57:F90");
  const rangoValoresClientesColumnaU = hojaDatos.getRange("D18:D21");


  // Crear y aplicar validaciones
  const reglas = [
    {
      rango: rangoDropdownCliente,
      valores: rangoValoresClienteDatos
    },
    {
      rango: rangoDropdownClienteInvalido,
      valores: rangoValoresClienteInvalido
    },
    {
      rango: rangoDropdownProductos,
      valores: rangoValoresProductosDatos
    },
    {
      rango: rangoDropdownClienteF,
      valores: rangoValoresClienteFactura
    },
    {
      rango: rangoDropdownProductoF,
      valores: rangoValoresProductosFactura
    },
    {
      rango: rangoDropdownCopiaFacturaCliente,
      valores: rangoValoresClienteFactura
    },
    {
      rango: rangoDropdownCopiaFacturaProducto,
      valores: rangoValoresProductosFactura
    },
    {
      rango: rangoDropdownProductosColumnaD,
      valores: rangoValoresProductosColumnaD
    },
    {
      rango: rangoDropdownProductosColumnaG,
      valores: rangoValoresProductosColumnaG
    },
    {
      rango: rangoDropdownProductosColumnaI,
      valores: rangoValoresProductosColumnaI
    },
    {
      rango: rangoDropdownProductosColumnaK,
      valores: rangoValoresProductosColumnaK
    },
    {
      rango: rangoDropdownProductosColumnaM,
      valores: rangoValoresProductosColumnaM
    },
    {
      rango: rangoDropdownClientesColumnaI,
      valores: rangoValoresClientesColumnaI
    },
    {
      rango: rangoDropdownClientesColumnaM,
      valores: rangoValoresClientesColumnaM
    },
    {
      rango: rangoDropdownClientesColumnaN,
      valores: rangoValoresClientesColumnaN
    },
    {
      rango: rangoDropdownClientesColumnaU,
      valores: rangoValoresClientesColumnaU
    },
    {
      rango: rangoDropdownCopiaFacturaMedioDePago,
      valores: rangoValoresCopiaFacturaMedioDePago
    },
    {
      rango: rangoDropdownCopiaFacturaMoneda,
      valores: rangoValoresCopiaFacturaMoneda
    },
    {
      rango: rangoDropdownFacturaMedioDePago,
      valores: rangoValoresFacturaMedioDePago
    },
    {
      rango: rangoDropdownFacturaMoneda,
      valores: rangoValoresFacturaMoneda
    }
  ];

  reglas.forEach(({ rango, valores }) => {
    const regla = SpreadsheetApp.newDataValidation()
      .requireValueInRange(valores, true) // Usar valores del rango especificado
      .setAllowInvalid(false) // No permitir valores fuera del rango
      .build();
    rango.setDataValidation(regla); // Aplicar la regla
  });
}

function IniciarMisfacturas() {
  let ui = SpreadsheetApp.getUi();

  let hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Datos de emisor");
  if (hoja == null) {
    iniciarHojasFactura()
    OnOpenSheetInicio()
    agregarDataValidations()
    let htmlOutput = HtmlService.createHtmlOutput(plantillaVincularMF()).setWidth(500).setHeight(250);
    ui.showModalDialog(htmlOutput, 'Vinculación requerida');


  } else {
    let respuesta = ui.alert('Si vuelves a instalar, solo se instalaran las hojas no existan o que hayan sido eliminadas?', ui.ButtonSet.YES_NO);
    if (respuesta == ui.Button.YES) {
      iniciarHojasFactura()
      OnOpenSheetInicio()
      agregarDataValidations()
      let htmlOutput = HtmlService.createHtmlOutput(plantillaVincularMF()).setWidth(500).setHeight(250);
      ui.showModalDialog(htmlOutput, 'Vinculación requerida');
      
    } else {
      return
    }
    
  }

}

function abrirMenuVinculacion(inHoja) {
  // Código para abrir el menú de vinculación en el sidebar
  if (inHoja) {
    let htmlOutput = HtmlService.createHtmlOutputFromFile('menuVincularFactura')
      .setTitle('Vincular misfacturas');
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
  } else {
    let htmlOutput = HtmlService.createHtmlOutputFromFile('menuVincular')
      .setTitle('Vincular misfacturas');
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Datos de emisor");
    SpreadsheetApp.setActiveSheet(sheet);
  }
}

function eliminarHojasFactura() {
  let ui = SpreadsheetApp.getUi();
  Logger.log("Inicio de eliminación de hojas");
  let respuesta = ui.alert('Recuerda que al desinstalar las hojas se eliminará toda la información de las mismas. Esta función solo debe ejecutarse si tienes un problema irreparable con las hojas. ¿Estás seguro de continuar?', ui.ButtonSet.YES_NO);
  if (respuesta == ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const nombresHojas = ["Inicio", "Productos", "Datos de emisor", "Clientes", "Factura", "ListadoEstado", "ClientesInvalidos", "Copia de Factura", "Datos", "Listado Facturas", "Listado Facturas Data"];

    // Crear una hoja nueva en blanco
    let nuevaHoja = ss.getSheetByName("Hoja en blanco");
    if (!nuevaHoja) {
      nuevaHoja = ss.insertSheet("Hoja en blanco");
      Logger.log("Se creó una nueva hoja en blanco");
    }

    // Recorrer todas las hojas del archivo
    ss.getSheets().forEach(hoja => {
      const nombreHoja = hoja.getName();
      if (nombresHojas.includes(nombreHoja)) {
        ss.deleteSheet(hoja);
        Logger.log(`Hoja eliminada: ${nombreHoja}`);
      }
    });

    SpreadsheetApp.getUi().alert("Hojas eliminadas satisfactoriamente.");
  } else {
    return;
  }
}


function reinstalarHojaDatos(ss, plantilla) {
  Logger.log("Reinstalando hoja Datos...");

  const nombreHoja = "Datos";
  let hojaDatos = ss.getSheetByName(nombreHoja);
  Logger.log("After getting hojadatos")
  // Eliminar la hoja "Datos" si ya existe
  if (hojaDatos) {
    ss.deleteSheet(hojaDatos);
    Logger.log("AIFF ")
  }

  // Copiar la hoja "Datos" desde la plantilla
  const hojaPlantilla = plantilla.getSheetByName(nombreHoja);
  if (hojaPlantilla) {
    const hojaCopia = hojaPlantilla.copyTo(ss).setName(nombreHoja);
    Logger.log("dentro if")
    // Hacer la hoja invisible
    hojaCopia.hideSheet();

    Logger.log("hoja aca")
  } else {
    SpreadsheetApp.getUi().alert('La hoja "Datos" no existe en la plantilla.');
  }
}


function pruebaLogo() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Datos de emisor");
  var celdaLogo = hoja.getRange("B12").getValue();
  hoja.getRange("B20").setValue(celdaLogo);
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('main')
    .setTitle('Menú');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}
function showPreProductos() {
  console.log("Attempting to show Productos");
  var html = HtmlService.createHtmlOutputFromFile('preProductos')
    .setTitle('Productos');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function openDatosEmisorSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Datos de emisor");
  SpreadsheetApp.setActiveSheet(sheet);
}

function showConfiguracion() {
  var html = HtmlService.createHtmlOutputFromFile('menuConfiguracion')
    .setTitle('Datos emisor');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function showDesvincular() {
  var html = HtmlService.createHtmlOutputFromFile('menuDesvincular')
    .setTitle('Desvincular cuenta');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function showAggProductos() {
  var html = HtmlService.createHtmlOutputFromFile('agregarProducto')
    .setTitle('Agregar Productos');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function openFacturaSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Factura");
  SpreadsheetApp.setActiveSheet(sheet);
}

function openHistorialSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Listado Facturas");
  SpreadsheetApp.setActiveSheet(sheet);
}

function showMenuFactura() {
  var html = HtmlService.createHtmlOutputFromFile('menuFactura')
    .setTitle('Menú Factura');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function showNuevaFactura() {
  var html = HtmlService.createHtmlOutputFromFile('nuevaFactura').setTitle("Nueva factura")
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function showAgregarProdcuto() {
  var html = HtmlService.createHtmlOutputFromFile('menuAgregarProducto').setTitle("Agregar Producto")
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function openClientesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Clientes");
  SpreadsheetApp.setActiveSheet(sheet);
}

function showClientes() {
  var html = HtmlService.createHtmlOutputFromFile('menuCliente')
    .setTitle('Menu cliente');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function showVincularCuenta() {
  var html = HtmlService.createHtmlOutputFromFile('menuVincular')
    .setTitle('Vincular cuenta');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}


function showEliminarInfo() {
  var html = HtmlService.createHtmlOutputFromFile('menuEliminarInfo')
    .setTitle('Eliminar informacion');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}


function eliminarTotalidadInformacion() {
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatos = spreadsheet.getSheetByName('Datos');
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let hojaProductos = spreadsheet.getSheetByName('Productos');
  let hojaClientes = spreadsheet.getSheetByName("Clientes");
  let hojaListadoEstado = spreadsheet.getSheetByName('ListadoEstado');
  let ClientesInvalidos = spreadsheet.getSheetByName('ClientesInvalidos');
  let hojaFacturasData = spreadsheet.getSheetByName('Listado Facturas Data');

  borrarInfoHoja(hojaProductos)
  borrarInfoHoja(hojaClientes)
  borrarInfoHoja(hojaListadoEstado)
  borrarInfoHoja(ClientesInvalidos)
  borrarInfoHoja(hojaDatosEmisor)
  borrarInfoHoja(hojaFacturasData)
  hojaDatos.getRange("Q11").setValue(0)
  hojaDatos.getRange("F47").setValue("")
  hojaDatosEmisor.getRange("B13").setBackground('#FFC7C7')
  hojaDatosEmisor.getRange("B13").setValue("Desvinculado")
  SpreadsheetApp.getUi().alert('Informacion eliminada correctamente');


}

function borrarInfoHoja(hoja) {
  let lastrow = Number(hoja.getLastRow())
  let nombreHoja = hoja.getSheetName()
  Logger.log("borrarInfoHoja")
  Logger.log("nombreHoja " + nombreHoja)
  Logger.log("lastrow " + lastrow)
  if (nombreHoja === "Datos de emisor") {
    Logger.log("Hoja es datos emisor ")
    hoja.getRange(1, 2, 12).setValue("")
    for (let i = 18; i < 50; i++) {
      hoja.getRange(i, 1, 1, 6).setValue("")
      hoja.getRange(i, 1, 1, 6).setBackground(null)
    }
  } else if (nombreHoja === "Clientes") {
    Logger.log("Hoja es clientes")
    hoja.deleteRows(3, lastrow)
    let maxRows = hoja.getMaxRows()
    Logger.log("maxRows " + maxRows)
    let dif = 1000 - maxRows
    Logger.log("dif " + dif)
    hoja.insertRows(maxRows, dif)
  }
  else {
    Logger.log("else")
    hoja.deleteRows(2, lastrow)
    let maxRows = hoja.getMaxRows()
    Logger.log("maxRows " + maxRows)
    let dif = 1000 - maxRows
    Logger.log("dif " + dif)
    hoja.insertRows(maxRows, dif)
  }
}

function mensajeBorrarInfoError() {
  Logger.log("Error borrar info")
  SpreadsheetApp.getUi().alert('Si deseas eliminar toda la informacion de misfacturas asegurate de escribir ELIMINAR en el campo');
}

function DesvincularMisfacturas(cambioAmb) {
  Logger.log("Desvincular")
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let hojaDatos = spreadsheet.getSheetByName('Datos');
  hojaDatosEmisor.getRange("B13").setBackground('#FFC7C7')
  hojaDatosEmisor.getRange("B13").setValue("Desvinculado")
  hojaDatos.getRange("F47").setValue("")
  if (!cambioAmb) {
    SpreadsheetApp.getUi().alert('Haz desvinculado exitosamente misfacturas');
  }
  for (let i = 18; i < 30; i++) {
    hojaDatosEmisor.getRange(i, 1, 1, 6).setValue("")
    hojaDatosEmisor.getRange(i, 1, 1, 6).setBackground(null)
  }
}

function mensajeErrorDesvincularMisfacturas() {
  Logger.log("Error Desvincular")
  SpreadsheetApp.getUi().alert('Si deseas desvincular misfacturas asegurate de escribir DESVINCULAR en el campo');
}

function openProductosSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Productos");
  SpreadsheetApp.setActiveSheet(sheet);
}

function processForm(data) {
  let existe = verificarIdentificacionUnica(data.codigoReferencia, "Productos")
  if (existe) {
    SpreadsheetApp.getUi().alert("El codigo de referencia ya existe, por favor poner un codigo de referencia unico");
    throw new Error('por favor poner un Numero de Identificacion unico');
  }
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Productos");
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    //Crea las variables para guardar los datos del producto
    const tipo = data.tipoProducto;
    const codigoReferencia = data.codigoReferencia;
    const nombre = data.nombre;
    const precioUnitario = parseFloat(data.precioUnitario);
    const unidadDeMedida = data.unidadDeMedida;
    const referenciaAdicional = data.referenciaAdicional;
    const numeroReferenciaAdicional = referenciaAdicionalCodigos[referenciaAdicional];
    let iva = "";
    let tarifaIva = String(data.tarifaIva) + "%";
    if (data.tarifaIva !== "0") {
      iva = "IVA";
    }
    let inc = "";
    let tarifaInc = String(data.tarifaInc) + "%";
    if (data.tarifaInc !== "0") {
      inc = "INC";
    }
    const retencionConcepto = data.tarifaReteRenta;
    let tarifaRetencion = String(validarTarifaRetencion(retencionConcepto) + "%");

    //Asigna los valores a los campos en el sheet
    //Tipo
    sheet.getRange(newRow, 1).setValue(tipo);
    sheet.getRange(newRow, 1).setHorizontalAlignment('center');
    //Codigo Referencia
    sheet.getRange(newRow, 2).setValue(codigoReferencia);
    sheet.getRange(newRow, 2).setHorizontalAlignment('center');
    //Nombre
    sheet.getRange(newRow, 3).setValue(nombre);
    sheet.getRange(newRow, 3).setHorizontalAlignment('center');
    //Referencia Adicional
    sheet.getRange(newRow, 4).setValue(referenciaAdicional);
    //Codigo Referencia Adicional
    sheet.getRange(newRow, 5).setValue(numeroReferenciaAdicional);

    //Precio Unitario
    sheet.getRange(newRow, 6).setValue(precioUnitario);
    sheet.getRange(newRow, 6).setHorizontalAlignment('normal');
    sheet.getRange(newRow, 6).setNumberFormat('$#,##0');
    //Unidad de Medida
    sheet.getRange(newRow, 7).setValue(unidadDeMedida);
    //Columna IVA
    sheet.getRange(newRow, 8).setValue(iva);
    //Tarifa IVA (formatea la celda como porcentaje)
    const tarifaIVA = sheet.getRange(newRow, 9);
    tarifaIVA.setHorizontalAlignment('center');
    tarifaIVA.setValue(tarifaIva); // Establece el valor del IVA como decimal
    //Columna INC
    sheet.getRange(newRow, 10).setValue(inc);
    //Tarifa INC (formatea la celda como porcentaje)  
    const tarifaINC = sheet.getRange(newRow, 11);
    tarifaINC.setHorizontalAlignment('center');
    tarifaINC.setValue(tarifaInc); // Establece el valor del IVA como decimal
    //Precio impuesto
    precioImpuesto = precioUnitario * (parseFloat(data.tarifaIva) / 100) + precioUnitario * (parseFloat(data.tarifaInc) / 100);
    sheet.getRange(newRow, 12).setValue(precioImpuesto);
    //Retencion concepto
    sheet.getRange(newRow, 13).setValue(retencionConcepto);
    //Tarifa Retencion (formatea la celda como porcentaje)
    const tarifaRetencionCell = sheet.getRange(newRow, 14);
    tarifaRetencionCell.setValue(tarifaRetencion); // Establece el valor del IVA como decimal
    tarifaRetencionCell.setNumberFormat('0%'); // Formatea la celda como porcentaje con dos decimales
    //Valor Retencion
    valorRetencion = precioUnitario * (parseFloat(tarifaRetencion) / 100);
    sheet.getRange(newRow, 15).setValue(valorRetencion);

    let referenciaUnica = nombre + "-" + codigoReferencia
    sheet.getRange(newRow, 16).setValue(referenciaUnica);
    sheet.getRange(newRow, 16).setHorizontalAlignment('normal');


    return SpreadsheetApp.getUi().alert("Nuevo producto creado satisfactoriamente");
  } catch (error) {
    return SpreadsheetApp.getUi().alert("Error al guardar los datos: " + error.message);
  }
}

function convertToPercentage(value) {
  return (value * 100).toFixed(2);
}

function onEdit(e) {
  // Trigger global para todas las hojas
  onEditFacturaActualizarNumeroFactura(e); // Actualiza número de factura en J2 al cambiar H2 en Factura
  const lock = LockService.getScriptLock();
  let hojaActual = e.source.getActiveSheet();
  let nombreHoja = hojaActual.getName();

  if (nombreHoja === "Datos" || nombreHoja === "ClientesInvalidos" || nombreHoja === "ListadoEstado" || nombreHoja === "Copia de Factura" || nombreHoja === "Listado Facturas Data") {
    showWarningAndHideSheet();
  }

  if (nombreHoja === "Factura") {

    let factura_sheet = hojaActual;
    let celdaEditada = e.range;
    let rowEditada = celdaEditada.getRow();
    let colEditada = celdaEditada.getColumn();
    let columnaClientes = 2; // Ajusta según sea necesario
    let rowClientes = 2;

    const productStartRow = 15; // productos empiezan aca
    let cargosDescuentosStartRow = getcargosDescuentosStartRow(hojaActual); // Assuming products end at column L
    let posRowTotalProductos = cargosDescuentosStartRow - 2; // posición (row) de Total productos

    Logger.log(`Editando celda en fila ${rowEditada}, columna ${colEditada}`);

    if (colEditada === columnaClientes && rowEditada === rowClientes) {
      verificarYCopiarCliente(e);
      ponerFechaYHoraActual();
    }
    if (rowEditada == 3 && colEditada == 8) {
      let hojaDatosEmisor = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos de emisor');
      let numeroAutorizacion = celdaEditada.getValue();
      let consecutivoFactura = 0;
      for (i = 18; i <= 20; i++) {
        if (hojaDatosEmisor.getRange(i, 1).getValue() == numeroAutorizacion) {
          consecutivoFactura = hojaDatosEmisor.getRange(i, 5).getValue();
          break;
        }
      }
    } else if (rowEditada >= productStartRow && (colEditada == 2 || colEditada == 3) && rowEditada < posRowTotalProductos) {
      let i = rowEditada;
      let productoFilaI = factura_sheet.getRange("B" + String(i)).getValue();
      let dictInformacionProducto = obtenerInformacionProducto(productoFilaI);
      let cantidadProducto = factura_sheet.getRange("C" + String(i)).getValue();

      if (cantidadProducto === "") {
        cantidadProducto = 0;
        factura_sheet.getRange("A" + String(i)).setValue(dictInformacionProducto["codigo Producto"]);
        factura_sheet.getRange("D" + String(i)).setValue(dictInformacionProducto["precio Unitario"]); // precio unitario
        factura_sheet.getRange("G" + String(i)).setValue(dictInformacionProducto["tarifa IVA"]); // %IVA
        factura_sheet.getRange("H" + String(i)).setValue(dictInformacionProducto["tarifa INC"]); // %INC
      } else {
        factura_sheet.getRange("A" + String(i)).setValue(dictInformacionProducto["codigo Producto"]);
        factura_sheet.getRange("D" + String(i)).setValue(dictInformacionProducto["precio Unitario"]); // precio unitario
        factura_sheet.getRange("E" + String(i)).setValue("=D" + String(i) + "*C" + String(i) + "-(D" + String(i) + "*C" + String(i) + ")*" + "I" + String(i) + "+J" + String(i)); // Subtotal teniendo en cuenta descuentos y cargos
        factura_sheet.getRange("F" + String(i)).setValue("=E" + String(i) + "*" + dictInformacionProducto["tarifa IVA"] + "+E" + String(i) + "*" + String(dictInformacionProducto["tarifa INC"])); // Impuestos
        factura_sheet.getRange("G" + String(i)).setValue(dictInformacionProducto["tarifa IVA"]); // %IVA
        factura_sheet.getRange("H" + String(i)).setValue(dictInformacionProducto["tarifa INC"]); // %INC
        cargos = Number(factura_sheet.getRange("J" + String(i)).getValue()); // Cargos
        Logger.log("retencion: " + dictInformacionProducto["valor Retencion"]);
        factura_sheet.getRange("K" + String(i)).setValue(dictInformacionProducto["tarifa Retencion"] * factura_sheet.getRange("E" + String(i)).getValue()); // Retencion
        factura_sheet.getRange("L" + String(i)).setValue("=E" + String(i) + "+F" + String(i));
      }

      let lastRowProducto = cargosDescuentosStartRow - 3;
      if (lastRowProducto === productStartRow) {
        let rowParaTotales = getTotalesLinea(hojaActual);
        hojaActual.getRange("K" + String(rowParaTotales)).setValue("=L15");
        hojaActual.getRange("L" + String(rowParaTotales)).setValue("=K" + String(rowParaTotales) + "-J" + String(rowParaTotales));

        let productoFilaI = factura_sheet.getRange("B15").getValue();
        let dictInformacionProducto = obtenerInformacionProducto(productoFilaI);
        let tarifaINC = dictInformacionProducto["tarifa INC"];
        let tarifaIVA = dictInformacionProducto["tarifa IVA"];
        if (tarifaINC !== 0 || tarifaIVA !== 0) {
          hojaActual.getRange("B" + String(rowParaTotales)).setValue("=E" + String(lastRowProducto));
        }
        calcularDescuentosCargosYTotales(lastRowProducto, cargosDescuentosStartRow, hojaActual);
      } else {
        calcularDescuentosCargosYTotales(lastRowProducto, cargosDescuentosStartRow, hojaActual);
      }

      updateTotalProductCounter(lastRowProducto, productStartRow, hojaActual, cargosDescuentosStartRow);
    } else if ((colEditada == 9 || colEditada == 10) && rowEditada >= productStartRow && rowEditada < posRowTotalProductos) {
      // verificar descuentos
      let i = rowEditada;
      let descuento = factura_sheet.getRange("I" + String(i)).getValue();
      if (descuento > 1 || descuento < 0) {
        SpreadsheetApp.getUi().alert("El descuento no puede ser mayor al 100% ni menor a 0%");
        factura_sheet.getRange("I" + String(i)).setValue(0);
      }

      let cargosDescuentosStartRow = getcargosDescuentosStartRow(hojaActual);
      let lastRowProducto = cargosDescuentosStartRow - 3;
      calcularDescuentosCargosYTotales(lastRowProducto, cargosDescuentosStartRow, hojaActual);
    } else if ((colEditada == 2 || colEditada == 3 || colEditada == 4) && rowEditada > posRowTotalProductos) {
      let cargosDescuentosStartRow = getcargosDescuentosStartRow(hojaActual);
      let lastRowProducto = cargosDescuentosStartRow - 3;
      calcularDescuentosCargosYTotales(lastRowProducto, cargosDescuentosStartRow, hojaActual);
    } else if (colEditada == 7 && rowEditada == 6) {
      // Entra a verificar días de vencimiento
      let valorDiasVencimiento = celdaEditada.getValue();

      // Verifica si es un entero positivo
      if (!Number.isInteger(valorDiasVencimiento) || valorDiasVencimiento <= 0) {
        // Muestra una alerta
        SpreadsheetApp.getUi().alert('El valor de días de vencimiento debe ser un entero positivo.');

        // Restablece el valor a 0
        celdaEditada.setValue(0);
      }
    } else if (colEditada == 10 && rowEditada == 4) {
      // Verifica la moneda
      let moneda = celdaEditada.getValue();
      if (moneda != "COP-Peso colombiano") {
        hojaActual.getRange(rowEditada + 1, colEditada).setBackground('#e0ecd4');
        hojaActual.getRange(rowEditada + 1, colEditada).setValue("");
        ponerFechaTasaDeCambio();
      } else {
        hojaActual.getRange(rowEditada + 1, colEditada).setBackground('#e0dcdc');
        hojaActual.getRange(rowEditada + 1, colEditada).setValue("N/A");
        hojaActual.getRange(rowEditada + 2, colEditada).setValue("N/A");
      }
    }
  } else if (nombreHoja === "Clientes") {
    let celdaEditada = e.range;
    let hojaCliente = e.source.getActiveSheet();

    let rowEditada = celdaEditada.getRow();
    let colEditada = celdaEditada.getColumn();
    let colTipoDePersona = 2
    let tipoPersona = obtenerTipoDePersona(e);

    if (colEditada == 10 || colEditada == 11 && rowEditada > 1) {
      let numeroIdentificacion = hojaCliente.getRange(rowEditada, colEditada).getValue()
      let numeroIngresadoyColumna = numeroIdentificacion + "-" + colEditada
      Logger.log("num i" + numeroIngresadoyColumna)
      let existe = verificarIdentificacionUnica(numeroIngresadoyColumna, "Clientes", true, rowEditada)
      if (existe === 1) {
        SpreadsheetApp.getUi().alert("El numero de identificacion ya existe, por favor elegir otro numero unico");
        hojaCliente.getRange(rowEditada, 10).setValue("")

        verificarDatosObligatorios(e, tipoPersona)
        throw new Error('por favor poner un Numero de Identificacion unico');

      } else if (existe === 2) {
        SpreadsheetApp.getUi().alert("El codigo del cliente ya existe, por favor elegir otro numero unico");
        hojaCliente.getRange(rowEditada, 11).setValue("")

        verificarDatosObligatorios(e, tipoPersona)
        throw new Error('por favor poner un Numero de Identificacion unico');
      }

    }

    var hojaDatos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos');
    hojaDatos.getRange("L101").setValue("=Clientes!N" + rowEditada);
    // Agregar regla de validación de datos
    let rangoValidacion = hojaCliente.getRange("O" + rowEditada);
    let regla = SpreadsheetApp.newDataValidation()
      .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRange("Datos!$M$101:$M$367"))
      .setAllowInvalid(false)
      .build();
    rangoValidacion.setDataValidation(regla);

    verificarDatosObligatorios(e, tipoPersona)
    agregarCodigoIdentificador(e, tipoPersona)
    SpreadsheetApp.flush();


  } else if (nombreHoja === "Productos") {

    let celdaEditada = e.range;
    let hojaProductos = e.source.getActiveSheet();

    let rowEditada = celdaEditada.getRow();
    let colEditada = celdaEditada.getColumn();

    if (rowEditada !== 1) {
      let codigoReferencia = hojaProductos.getRange(rowEditada, colEditada).getValue()
      let existe = verificarIdentificacionUnica(codigoReferencia, "Productos", true, rowEditada)
      if (existe) {
        SpreadsheetApp.getUi().alert("El codigo de referencia ya existe, por favor elegir otro numero unico");
        hojaProductos.getRange(rowEditada, 2).setValue("");
        throw new Error('por favor poner un codigo de referencia unico');
      }
      let tipoPersona = '';
      let valido = verificarDatosObligatoriosProductos(e);
      Logger.log("valido " + valido)
      if (valido === true) {
        agregarCodigoIdentificador(e, tipoPersona);
      } else {
        hojaProductos.getRange(rowEditada, 16).setValue("");
      }
    }

  } else if (nombreHoja === "Listado Facturas") {
    let celdaEditada = e.range;
    let rowEditada = celdaEditada.getRow();
    let colEditada = celdaEditada.getColumn();
    if (rowEditada == 5 && colEditada == 3) {
      Logger.log("dentto de selccionar filtor")
      let valor = celdaEditada.getValue()
      filtroHistorialFacturas(valor)
    }

  }

}

function calcularDescuentosCargosYTotales(lastRowProducto, cargosDescuentosStartRow, hojaActual) {

  let rowParaTotales = getTotalesLinea(hojaActual);

  //subtotal 
  hojaActual.getRange("A" + String(rowParaTotales)).setValue("=SUM(E15:E" + String(lastRowProducto) + ")")

  //base grabable
  const rangoDatos = hojaActual.getRange("E15:F" + String(lastRowProducto)).getValues();
  const condicion = 0;
  let suma = 0;
  rangoDatos.forEach(fila => {
    if (fila[1] !== condicion) {
      suma += fila[0];
    }
  });
  hojaActual.getRange("B" + String(rowParaTotales)).setValue(suma)

  //Seccion Cargos y Descuentos
  let lastCargoDescuentoRow = getLastCargoDescuentoRow(hojaActual);
  let rowSeccionCargosYDescuentos = cargosDescuentosStartRow;
  let totalDescuentosSeccionCargosyDescuentos = calcularCargYDescu(hojaActual, rowSeccionCargosYDescuentos, lastCargoDescuentoRow);


  //impuestos
  hojaActual.getRange("C" + String(rowParaTotales)).setValue("=SUM(F15:F" + String(lastRowProducto) + ")")
  //subtotal mas impuestos
  hojaActual.getRange("D" + String(rowParaTotales)).setValue("=A" + String(rowParaTotales) + "+C" + String(rowParaTotales))

  //retenciones
  hojaActual.getRange("E" + String(rowParaTotales)).setValue("=SUM(K15:K" + String(lastRowProducto) + ")")

  //descuentos
  hojaActual.getRange("F" + String(rowParaTotales)).setValue(Number(totalDescuentosSeccionCargosyDescuentos.descuentos))

  //cargos
  hojaActual.getRange("H" + String(rowParaTotales)).setValue(totalDescuentosSeccionCargosyDescuentos.cargos)


  //total
  hojaActual.getRange("K" + String(rowParaTotales)).setValue("=D" + String(rowParaTotales) + "-F" + String(rowParaTotales) + "+H" + String(rowParaTotales))
  //neto a pagar
  hojaActual.getRange("L" + String(rowParaTotales)).setValue("=K" + String(rowParaTotales) + "-J" + String(rowParaTotales))
}

function calcularDescuentos(hojaActual, lastRowProducto) {
  let rangoDatos = hojaActual.getRange("E15:K" + String(lastRowProducto)).getValues();
  let resultado = 0;

  rangoDatos.forEach(fila => {
    //suma = subtotal + impuestos + cargos + retencion
    let suma = Number(fila[0]) + Number(fila[1]) + Number(fila[5]) + Number(fila[6]);
    //producto = suma * descuento
    let producto = suma * Number(fila[4]);
    resultado += producto;

  });

  Logger.log("resultado " + resultado)
  return resultado;

}

function getLastProductRow(sheet, productStartRow, cargosDescuentosStartRow) {

  //retorna el numero de fila exacta donde esta el ulitmo producto agregado
  // si no encuntra producto agg si solo tiene un producto retorna el mismo productStartRow 
  let lastProductRow = productStartRow;

  for (let row = productStartRow; row < cargosDescuentosStartRow; row++) {

    let valorCeldaActual = sheet.getRange(row, 1).getValue()
    if (valorCeldaActual === "Total items") {
      return lastProductRow - 1;
    } else {
      lastProductRow = row;
    }
  }
  return lastProductRow;
}

function getLastCargoDescuentoRow(sheet) {
  //obtiene la row donde esta el final de la seccion de cargos y descuentos
  const lastRow = sheet.getLastRow();
  let row = 20

  for (row; row < lastRow; row++) {
    if (sheet.getRange(row, 1).getValue() === 'Subtotal') {
      return row - 3;
    }
  }
}

function getTotalesLinea(sheet) {
  //obtiene la row donde esta la linea de totales
  const lastRow = sheet.getLastRow();
  let row = 20

  for (row; row < lastRow; row++) {
    if (sheet.getRange(row, 1).getValue() === 'Subtotal') {
      return row + 1;

    }
  }
  Logger.log(" last row row " + row)
}

function getcargosDescuentosStartRow(sheet) {
  const lastRow = sheet.getLastRow();
  let row = 14

  for (row; row < lastRow; row++) {
    if (sheet.getRange(row, 1).getValue() === 'Cargos y/o Descuentos') {
      return row;
    }
  }
}

function calcularCargYDescu(hojaActual, rowSeccionCargosYDescuentos, lastCargoDescuentoRow) {
  let totalDescuentosSeccionCargosyDescuentos = { cargos: 0, descuentos: 0 };
  let totalesRow = getTotalesLinea(hojaActual);
  for (let i = rowSeccionCargosYDescuentos + 2; i < lastCargoDescuentoRow + 1; i++) {
    let celdaValorPorcentaje = hojaActual.getRange("C" + String(i)).getValue()
    if (hojaActual.getRange("A" + String(i)).getValue() === "Cargo") {
      if (String(celdaValorPorcentaje).includes("%")) {
        porcentaje = celdaValorPorcentaje.replace("%", "") / 100
        let subtotal = hojaActual.getRange("A" + String(totalesRow)).getValue();
        let base = hojaActual.getRange("D" + String(i)).setValue(subtotal);
        hojaActual.getRange("E" + String(i)).setValue("=" + subtotal + "*" + porcentaje)

      } else {
        hojaActual.getRange("D" + String(i)).setValue("N/A")
        hojaActual.getRange("E" + String(i)).setValue(celdaValorPorcentaje)
      }
      totalDescuentosSeccionCargosyDescuentos.cargos += hojaActual.getRange("E" + String(i)).getValue();
    }
    else {
      if (String(celdaValorPorcentaje).includes("%")) {
        porcentaje = celdaValorPorcentaje.replace("%", "") / 100
        let subtotal = hojaActual.getRange("A" + String(totalesRow)).getValue();
        let base = hojaActual.getRange("D" + String(i)).setValue(subtotal);
        hojaActual.getRange("E" + String(i)).setValue("=" + subtotal + "*" + porcentaje)

      } else {
        hojaActual.getRange("D" + String(i)).setValue("N/A")
        hojaActual.getRange("E" + String(i)).setValue(celdaValorPorcentaje)
      }
      totalDescuentosSeccionCargosyDescuentos.descuentos += hojaActual.getRange("E" + String(i)).getValue();
    }

  }
  return totalDescuentosSeccionCargosyDescuentos;
}

function updateTotalProductCounter(lastRowProducto, productStartRow, hojaActual, cargosDescuentosStartRow) {
  let totalProducts = 0;

  for (let i = productStartRow; i <= lastRowProducto; i++) {
    let cellValue = hojaActual.getRange("B" + String(i)).getValue();
    Logger.log("Checking cell B" + String(i) + ": " + cellValue);
    if (cellValue != "") {
      totalProducts++;
    }
  }

  let rowTotalProductos = cargosDescuentosStartRow - 2;
  Logger.log("Updating total products in row B" + String(rowTotalProductos));
  hojaActual.getRange("B" + String(rowTotalProductos)).setValue(totalProducts);
  Logger.log("Total products: " + totalProducts);
}

function verificarIdentificacionUnica(codigo, nombreHoja, inHoja, numRow) {
  // Clientes o Productos
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  if (codigo == "") {
    return false
  } else if (nombreHoja === "Clientes") {
    try {
      let columnaNumIdentificacionC = 10;
      let columnaCodigoCliente = 11;
      let lastActiveRow = sheet.getLastRow();
      let rangeNumeroIdentificaciones;
      let rangeCodigosCliente;
      let numidentificacion;
      let codigoCliente;
      let codigos = codigo.split("-")
      if (inHoja) {
        if (codigos[1] === "10") {
          Logger.log("entro a verificar si el numero es un de identificacion")
          rangeNumeroIdentificaciones = sheet.getRange(2, columnaNumIdentificacionC, lastActiveRow);
          let NumerosIdentificacion = String(rangeNumeroIdentificaciones.getValues());
          NumerosIdentificacion = NumerosIdentificacion.split(",")
          numidentificacion = codigos[0]

          if (NumerosIdentificacion.includes(String(numidentificacion))) {
            var posicionArreglo = NumerosIdentificacion.indexOf(String(numidentificacion)) + 2;
            Logger.log("posicionArreglo " + posicionArreglo + " numRow " + numRow)
            if (posicionArreglo !== numRow) {
              Logger.log(`posicion en el arreglo ${NumerosIdentificacion.indexOf(String(numidentificacion))} y en numero de fila ${numRow}`)
              Logger.log("El Num identificion ya existe.");
              return 1
            }
          } else {
            Logger.log("El Num identificion no existe.");
            return false;
          }
        } else if (codigos[1] === "11") {
          Logger.log("entro a verificar si el numero es un codigo")
          rangeCodigosCliente = sheet.getRange(2, columnaCodigoCliente, lastActiveRow);
          let CodigosCliente = String(rangeCodigosCliente.getValues());
          CodigosCliente = CodigosCliente.split(",")
          codigoCliente = codigos[0]
          if (CodigosCliente.includes(String(codigoCliente))) {
            var posicionArreglo = CodigosCliente.indexOf(String(codigoCliente)) + 2;
            Logger.log("posicionArreglo " + posicionArreglo + " numRow " + numRow)
            if (posicionArreglo !== numRow) {
              Logger.log(`posicion en el arreglo ${CodigosCliente.indexOf(String(codigoCliente))} y en numero de fila ${numRow}`)
              Logger.log("El Num identificion ya existe.");
              return 2
            }
          } else {
            Logger.log("El Num identificion no existe.");
            return false;
          }

        }

      } else {

        rangeNumeroIdentificaciones = sheet.getRange(2, columnaNumIdentificacionC, lastActiveRow);
        rangeCodigosCliente = sheet.getRange(2, columnaCodigoCliente, lastActiveRow);
        //Separar identificacion y codigo
        numidentificacion = codigos[0]
        codigoCliente = codigos[1]
      }

      let NumerosIdentificacion = String(rangeNumeroIdentificaciones.getValues());
      NumerosIdentificacion = NumerosIdentificacion.split(",")
      let CodigosCliente = String(rangeCodigosCliente.getValues());
      CodigosCliente = CodigosCliente.split(",")



      // Verificar si el código ya existe
      if (NumerosIdentificacion.includes(String(numidentificacion))) {
        var posicionArreglo = NumerosIdentificacion.indexOf(String(numidentificacion)) + 2;
        if (posicionArreglo !== numRow) {
          Logger.log(`posicion en el arreglo ${NumerosIdentificacion.indexOf(String(numidentificacion))} y en numero de fila ${numRow}`)
          Logger.log("El Num identificion ya existe.");
          return 1
        }

      } else if (CodigosCliente.includes(String(codigoCliente))) {
        var posicionArreglo = CodigosCliente.indexOf(String(codigoCliente)) + 2;
        if (posicionArreglo !== numRow) {
          Logger.log(`posicion en el arreglo ${CodigosCliente.indexOf(String(codigoCliente))} y en numero de fila ${numRow}`)
          Logger.log("El Num identificion ya existe.");
          return 2
        }
      } else {
        Logger.log("El Num identificion no existe.");
        return false;
      }
    } catch (error) {
      Logger.log("Error al verificar el Num identificion: " + error.message);
    }
  } else if (nombreHoja === "Productos") {
    try {
      let columnaNumIdentificacionP = 2;
      let lastActiveRow = sheet.getLastRow();
      let codigoNumero = Number(codigo);
      let rangeCodigoReferencia = sheet.getRange(2, columnaNumIdentificacionP, lastActiveRow);
      let datos = rangeCodigoReferencia.getValues().flat().map(Number);
      for (let i = 0; i < datos.length; i++) {
        if (datos[i] === codigoNumero && i + 2 !== numRow) {
          Logger.log(`El código "${codigoNumero}" ya existe en la hoja "${nombreHoja}".`);
          return true;
        }
      }
      return false;
    } catch (error) {
      Logger.log("Error al verificar el codigo: " + error.message);
      return false;
    }
  }
}

function agregarCodigoIdentificador(e, tipoPersona) {
  let hoja = e.source.getActiveSheet();
  let range = e.range;
  let rowEditada = range.getRow();


  if (hoja.getName() == "Clientes") {
    let estadoActual = hoja.getRange(rowEditada, 1).getValue()
    if (estadoActual == "Valido") {
      if (tipoPersona === "Natural") {
        let nombre = hoja.getRange(rowEditada, 5).getValue()
        let apellido = hoja.getRange(rowEditada, 7).getValue()
        let numeroIdentificacion = hoja.getRange(rowEditada, 10).getValue()
        let identificadorUnico = nombre + " " + apellido + "-" + numeroIdentificacion
        hoja.getRange(rowEditada, 23).setValue(identificadorUnico)
        let rangoValidacion = hoja.getRange("O" + rowEditada);
        rangoValidacion.clearDataValidations();

      } else if (tipoPersona === "Juridica") {
        let nombre = hoja.getRange(rowEditada, 4).getValue()
        let numeroIdentificacion = hoja.getRange(rowEditada, 10).getValue()
        let identificadorUnico = nombre + "-" + numeroIdentificacion
        hoja.getRange(rowEditada, 23).setValue(identificadorUnico)
        let rangoValidacion = hoja.getRange("O" + rowEditada);
        rangoValidacion.clearDataValidations();
      }
    } else {
      hoja.getRange(rowEditada, 23).setValue("")
    }
  } else if (hoja.getName() == "Productos") {
    let nombre = hoja.getRange(rowEditada, 3).getValue()
    let numeroIdentificacion = hoja.getRange(rowEditada, 2).getValue()
    let identificadorUnico = nombre + "-" + numeroIdentificacion
    hoja.getRange(rowEditada, 16).setValue(identificadorUnico)
  }
}

function slugifyF(str) {
  var map = {
    '-': ' ',
    '-': '_',
    'a': 'á|à|ã|â|À|Á|Ã|Â',
    'e': 'é|è|ê|É|È|Ê',
    'i': 'í|ì|î|Í|Ì|Î',
    'o': 'ó|ò|ô|õ|Ó|Ò|Ô|Õ',
    'u': 'ú|ù|û|ü|Ú|Ù|Û|Ü',
    'c': 'ç|Ç',
    'n': 'ñ|Ñ'
  };

  str = String(str)
  str = str.toLowerCase();

  for (var pattern in map) {
    str = str.replace(new RegExp(map[pattern], 'g'), pattern);
  };

  return str;
};

function getAdditionalDocuments() {
  var AdditionalDocuments = {
    "OrderReference": "",
    "DespatchDocumentReference": "",
    "ReceiptDocumentReference": "",
    "AdditionalDocument": []
  }
  return AdditionalDocuments;
}

function getAdditionalProperty() {
  var AdditionalProperty = [];
  return AdditionalProperty;
}

function getDelivery() {
  var Delivery = {
    "AddressLine": "",
    "CountryCode": "",
    "CountryName": "",
    "SubdivisionCode": "",
    "SubdivisionName": "",
    "CityCode": "",
    "CityName": "",
    "ContactPerson": "",
    "DeliveryDate": "",
    "DeliveryCompany": ""
  };
  return Delivery;

}
function showWarningAndHideSheet() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('No tienes permiso para editar esta hoja.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojasInvisibles = ["Datos", "ClientesInvalidos", "ListadoEstado", "Copia de Factura", "Listado Facturas Data"];
  hojasInvisibles.forEach(nombreHoja => {
    const hoja = ss.getSheetByName(nombreHoja);
    if (hoja) {
      hoja.hideSheet();
    }
  });
}

function onChange(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojasInvisibles = ["Datos", "ClientesInvalidos", "ListadoEstado", "Copia de Factura", "Listado Facturas Data"];
  hojasInvisibles.forEach(nombreHoja => {
    const hoja = ss.getSheetByName(nombreHoja);
    if (hoja && !hoja.isSheetHidden()) {
      hoja.hideSheet();
      SpreadsheetApp.getUi().alert(`La hoja "${nombreHoja}" debe permanecer oculta.`);
    }
  });
}


function cambiarAmbiente() {
  let ui = SpreadsheetApp.getUi();

  // Preguntar si el usuario está seguro
  let respuesta = ui.alert(
    '¿Estás seguro de que quieres cambiar el ambiente? Tendrás que volver a iniciar sesión.',
    ui.ButtonSet.YES_NO
  );

  if (respuesta == ui.Button.YES) {
    // Mostrar cuadro de diálogo personalizado con los ambientes
    let htmlOutput = HtmlService.createHtmlOutput(plantillaCambiarAmbiente())
      .setWidth(400)
      .setHeight(320);

    ui.showModalDialog(htmlOutput, 'Cambiar Ambiente');
  } else {
    ui.alert('No se ha cambiado el ambiente.');
  }
}

function aplicarCambioAmbiente(nuevoAmbiente) {
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');

  Logger.log("Nuevo ambiente seleccionado: " + nuevoAmbiente);

  // Actualizar el valor en la hoja y en las propiedades del documento
  const scriptProps = PropertiesService.getDocumentProperties();
  scriptProps.setProperties({
    'Ambiente': nuevoAmbiente
  });
  Logger.log("Si se creo el doc prop para el ambiente: " + nuevoAmbiente);

  abrirMenuVinculacion();
  let cambioAmb = true;
  DesvincularMisfacturas(cambioAmb);
  hojaDatosEmisor.getRange("C1002").setValue(nuevoAmbiente)
}

/**
 * Muestra un popup explicando el estado "En revisión".
 */
function mostrarPopupEstadoEnRevision() {
  let htmlOutput = HtmlService.createHtmlOutput(plantillaEstadoEnRevision()).setWidth(600).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '¿Qué significa "En revisión"?');
}

/**
 * Dummy para la opción "No mostrar nuevamente" (puedes implementar lógica de usuario si lo deseas)
 */
function noMostrarPopupEstadoEnRevision() {
  // Aquí podrías guardar en PropertiesService o similar si quieres persistencia por usuario
  return true;
}

function onEditFacturaActualizarNumeroFactura(e) {
  let hojaActual = e.source.getActiveSheet();
  let rowEditada = e.range.getRow();
  let colEditada = e.range.getColumn();

  if (hojaActual.getName() === "Factura" && colEditada === 8 && rowEditada === 2) {
    let numeroAutorizacion = hojaActual.getRange(rowEditada, colEditada).getValue();
    let hojaDatosEmisor = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos de emisor');
    let data = hojaDatosEmisor.getRange('A18:E67').getValues(); // Máximo 50 resoluciones
    let consecutivoFactura = '';
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] == numeroAutorizacion) {
        consecutivoFactura = data[i][4]; // Columna E (índice 4)
        break;
      }
    }
    hojaActual.getRange(2, 10).setValue(consecutivoFactura); // J2
  }
}