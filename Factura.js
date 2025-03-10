function descargarFacturaHtml() {
  var html = HtmlService.createHtmlOutputFromFile('descargaFacturaHistorial')
    .setTitle('Historial facutras');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function linkDescargaFactura(idFactura) {
  logearUsuario();
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let idNumber = hojaDatosEmisor.getRange("B3").getValue();
  let schemaID = 31;
  let documentNumber = idFactura;
  let documentType = 1;
  let url = `https://misfacturas.cenet.ws/integrationAPI_2/api/GetDownloadRgDocumentByNumber?SchemaID=${schemaID}&IDnumber=${idNumber}&DocumentType=${documentType}&DocumentNumber=${documentNumber}`;
  let hojaDatos = spreadsheet.getSheetByName('Datos');
  let token = hojaDatos.getRange("F47").getValue();

  let opciones = {
    "method": "get",
    "headers": { "Authorization": "misfacturas " + token },
    'muteHttpExceptions': true
  };
  Logger.log("Descargar factura antes del try");

  try {
    var respuesta = UrlFetchApp.fetch(url, opciones);
    var estatusRespuesta = respuesta.getResponseCode();
    Logger.log("Estatus de la respuesta: " + estatusRespuesta);

    if (estatusRespuesta == 200) {
      Logger.log("Si acepta el try, Descargando factura");
      const contenidoPDF = respuesta.getBlob().setName(`Factura_${documentNumber}.pdf`);
      const base64Data = Utilities.base64Encode(contenidoPDF.getBytes());
      return { documentNumber: documentNumber, base64Data: base64Data };

    } else {
      var contenidoRespuesta = JSON.parse(respuesta.getContentText());

      if (estatusRespuesta == 404 && contenidoRespuesta["Message"] === "No se encontraron resultados de facturas válidas que tengan representación gráfica") {
        SpreadsheetApp.getUi().alert("Error: No se encontraron resultados de facturas válidas que tengan representación gráfica.");
      } else {
        SpreadsheetApp.getUi().alert(
          "Error al intentar descargar la factura: " +
          (contenidoRespuesta["Message"] || "Respuesta desconocida")
        );
        Logger.log("Error al intentar descargar la factura: " + contenidoRespuesta);
      }
    }
  } catch (error) {
    Logger.log("Error al enviar el JSON a la API: " + error.message);
    SpreadsheetApp.getUi().alert("Error al intentar descargar la factura. Intente de nuevo si el error persiste comuníquese con soporte");
  }
}



function getDownloadLink(idFactura) {
  var data = linkDescargaFactura(idFactura);
  Logger.log("sale de linkdescargar")
  return data;
}


function verificarEstadoValidoFactura(estadoFactura) {
  var spreadsheet = SpreadsheetApp.getActive();
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  let estaValido = true;
  estadoFactura.push(estaValido);

  //verificar nit
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let nit = hojaDatosEmisor.getRange("B3").getValue();
  if (nit === "") {
    estaValido = false;
    estadoFactura.push("Por favor registre su NIT en la hoja Datos de Emisor");

  }

  // Verificar datos de la factura 
  let clienteActual = hojaFactura.getRange("B2").getValue();
  let informacionFactura1 = hojaFactura.getRange(3, 6, 4, 3).getValues();
  let informacionFactura2 = hojaFactura.getRange(2, 9, 5, 2).getValues();
  let moneda = hojaFactura.getRange("J4").getValue();

  if (clienteActual === "") {
    estaValido = false;
    estadoFactura.push("Cliente");
  }
  for (let i = 0; i < informacionFactura1.length; i++) {
    if (informacionFactura1[i][2] === "") {
      estaValido = false;
      estadoFactura.push(informacionFactura1[i][0]);
    }
  }
  if (moneda === "") {
    estaValido = false;
    estadoFactura.push("Moneda");
  }
  if (moneda === "COP-Peso colombiano") {
    for (let j = 0; j < 3; j++) {
      if (informacionFactura2[j][1] === "") {
        estaValido = false;
        estadoFactura.push(informacionFactura2[j][0]);
      }
    }
  } else {
    for (let j = 3; j < 5; j++) {
      if (informacionFactura2[j][1] === "") {
        estaValido = false;
        estadoFactura.push(informacionFactura2[j][0])
      };
    }
  }

  // Verificar si se agregaron productos
  let totalProductos = hojaFactura.getRange("A16").getValue();

  if (totalProductos === "Total items") {
    // no hay necesidad de encontrar TOTAL PRODUCTOS si no esta, porque eso implica que si anadio asi sea 1 prodcuto
    let valorTotalProductos = hojaFactura.getRange("B16").getValue();
    if (valorTotalProductos === 0 || valorTotalProductos === "") {
      // no agrego producto
      estaValido = false
      estadoFactura.push("No agrego producto")
    }
  }
  estadoFactura[0] = estaValido;
}

function guardarFactura() {

  SpreadsheetApp.getUi().alert("Revisando validez de la factura. Aguarde unos segundos");
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let estadoVinculacion = hojaDatosEmisor.getRange("B13").getValue();
  let estadoFactura = [];
  if (estadoVinculacion == "Desvinculado") {
    let inHoja = true;
    let htmlOutput = HtmlService.createHtmlOutput(plantillaVincularMF(inHoja)).setWidth(500).setHeight(200);
    let ui = SpreadsheetApp.getUi();
    ui.showModalDialog(htmlOutput, 'Vinculación requerida');
    return;
  }

  Logger.log("Se va a verificar la validez de la factura");
  verificarEstadoValidoFactura(estadoFactura);
  Logger.log("Estado de la factura: " + estadoFactura[0]);

  if (estadoFactura[0] === true) {

    if (logearUsuario()) {
      guardarYGenerarInvoice();
      mostrarResumenFactura();
    }


  } else {
    let mensajeError = "La factura no es válida. Por favor rellene los campos obligatorios:\n" + estadoFactura.join("\n- ");
    SpreadsheetApp.getUi().alert(mensajeError);
  }
}

function agregarFilaNueva() {

  const lock = LockService.getScriptLock();

  try {
    lock.tryLock(5000);
    var spreadsheet = SpreadsheetApp.getActive();

    let hojaFactura = spreadsheet.getSheetByName('Factura');
    let numeroFilasParaAgregar = hojaFactura.getRange("B13").getValue();

    // Verificar si numeroFilasParaAgregar es nulo, vacío o no es un número
    if (numeroFilasParaAgregar == 0 || numeroFilasParaAgregar == "" || isNaN(numeroFilasParaAgregar)) {
      SpreadsheetApp.getUi().alert("Error: Por favor ingresa un número válido de filas para agregar.");
      return; // Detener la ejecución si hay error
    }

    let cargosDescuentosStartRow = getcargosDescuentosStartRow(hojaFactura);
    const productStartRow = 15;
    const lastProductRow = getLastProductRow(hojaFactura, productStartRow, cargosDescuentosStartRow) + 1;

    Logger.log("Agregar fila nueva");
    hojaFactura.insertRows(lastProductRow, numeroFilasParaAgregar);

  } catch (error) {
    Logger.log("Error al agregar filas: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

function agregarFilaCargoDescuento() {

  const lock = LockService.getScriptLock();
  try {
    lock.tryLock(5000);
    let spreadsheet = SpreadsheetApp.getActive();
    let hojaFactura = spreadsheet.getSheetByName('Factura');
    const lastCargoDescuentoRow = getLastCargoDescuentoRow(hojaFactura);
    hojaFactura.insertRowAfter(lastCargoDescuentoRow);
  } catch (error) {
    Logger.log("Error al agregar fila de cargo/descuento: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

function agregarProductoDesdeFactura(cantidad, producto) {
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  let cargosDescuentosStartRow = getcargosDescuentosStartRow(hojaFactura);
  const productStartRow = 15;
  const lastProductRow = getLastProductRow(hojaFactura, productStartRow, cargosDescuentosStartRow);

  let dictInformacionProducto = {}
  if (producto === "" || cantidad === "" || cantidad === 0) {
    throw new Error('Porfavor elige un producto y un cantidad adecuado');
  } else {
    dictInformacionProducto = obtenerInformacionProducto(producto);
  }

  let rowParaDatos = lastProductRow
  let cantidadProductos = hojaFactura.getRange("B16").getValue()//estado defaul de total productos
  if (cantidadProductos === 0 || cantidadProductos === "") {
    hojaFactura.getRange("A15").setValue(dictInformacionProducto["codigo Producto"])
    hojaFactura.getRange("B15").setValue(producto)
    hojaFactura.getRange("C15").setValue(cantidad)
    hojaFactura.getRange("D15").setValue(dictInformacionProducto["precio Unitario"])
    hojaFactura.getRange("E15").setValue("=D15*C15")
    hojaFactura.getRange("F15").setValue(dictInformacionProducto["precio Impuesto"])
    hojaFactura.getRange("G15").setValue(dictInformacionProducto["tarifa INC"])
    hojaFactura.getRange("H15").setValue(dictInformacionProducto["tarifa IVA"])
    hojaFactura.getRange("K15").setValue(dictInformacionProducto["valor Retencion"])
    hojaFactura.getRange("L15").setValue("=(E15+F15+J15)-((E15+F15+J15)*I15)")//Total

  } else {
    hojaFactura.insertRowAfter(lastProductRow)
    rowParaDatos = lastProductRow + 1
    hojaFactura.getRange("A" + String(rowParaDatos)).setValue(dictInformacionProducto["codigo Producto"])
    hojaFactura.getRange("B" + String(rowParaDatos)).setValue(producto)
    hojaFactura.getRange("C" + String(rowParaDatos)).setValue(cantidad)
    hojaFactura.getRange("D" + String(rowParaDatos)).setValue(dictInformacionProducto["precio Unitario"])//precio unitario
    hojaFactura.getRange("E" + String(rowParaDatos)).setValue("=D" + String(rowParaDatos) + "*C" + String(rowParaDatos))//Subtotal
    hojaFactura.getRange("F" + String(rowParaDatos)).setValue("=E" + String(rowParaDatos) + "*" + dictInformacionProducto["tarifa IVA"] + "+E" + String(rowParaDatos) + "*" + String(dictInformacionProducto["tarifa INC"]))//Impuestos
    hojaFactura.getRange("G" + String(rowParaDatos)).setValue(dictInformacionProducto["tarifa INC"])//%INC
    hojaFactura.getRange("H" + String(rowParaDatos)).setValue(dictInformacionProducto["tarifa IVA"])//%IVA
    hojaFactura.getRange("K" + String(rowParaDatos)).setValue(dictInformacionProducto["valor Retencion"])
    hojaFactura.getRange("L" + String(rowParaDatos)).setValue("=(E" + String(rowParaDatos) + "+F" + String(rowParaDatos) + "+J" + String(rowParaDatos) + "-(E" + String(rowParaDatos) + "*I" + String(rowParaDatos) + ")")//Total
  }
  updateTotalProductCounter(cargosDescuentosStartRow - 3, productStartRow, hojaFactura, cargosDescuentosStartRow);
  calcularDescuentosCargosYTotales(cargosDescuentosStartRow - 3, cargosDescuentosStartRow, hojaFactura);
}

function recuperarJson() {
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaFactura = spreadsheet.getSheetByName('ListadoEstado');
  let lastRow = hojaFactura.getLastRow();
  let json = hojaFactura.getRange(lastRow, 5).getValues();
  return json[0][0];
}

function logearUsuario() {
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatos = spreadsheet.getSheetByName('Datos');
  let usuario = hojaDatos.getRange("F49").getValue();
  let contrasena = hojaDatos.getRange("F50").getValue();
  if (usuario === "" || contrasena === "") {
    let inHoja = true;
    let htmlOutput = HtmlService.createHtmlOutput(plantillaVincularMF(inHoja)).setWidth(500).setHeight(200);
    let ui = SpreadsheetApp.getUi();
    ui.showModalDialog(htmlOutput, 'Vinculación requerida');
    return false;
  }


  let url = `https://misfacturas.cenet.ws/IntegrationAPI_2/api/login?username=${usuario}&password=${contrasena}`;
  let opciones = {
    "method": "post",
    "contentType": "application/json",
    'muteHttpExceptions': true
  };
  let respuesta = UrlFetchApp.fetch(url, opciones);
  let contenidoRespuesta = respuesta.getContentText();
  let token = JSON.parse(contenidoRespuesta);
  hojaDatos.getRange("F47").setValue(token);
  Logger.log("Usuario loggeado para generar resoluciones o mandar la factura");
  return true;
}

function enviarFactura() {
  let retorno = [];
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let schemaID = 31;
  let idNumber = hojaDatosEmisor.getRange("B3").getValue();
  let templateID = 73;
  let url = `https://misfacturas.cenet.ws/integrationAPI_2/api/insertinvoice?SchemaID=${schemaID}&IDNumber=${idNumber}&TemplateID=${templateID}`;
  let json = recuperarJson();
  let hojaDatos = spreadsheet.getSheetByName('Datos');
  let token = hojaDatos.getRange("F47").getValue();

  let opciones = {
    "method": "post",
    "contentType": "application/json",
    "payload": json,
    "headers": { "Authorization": "misfacturas " + token },
    'muteHttpExceptions': true
  };
  Logger.log("Enviar factura antes del try");
  try {
    var respuesta = UrlFetchApp.fetch(url, opciones);
    var estatusRespuesta = respuesta.getResponseCode();
    Logger.log("Estatus de la respuesta: " + estatusRespuesta);

    var contenidoRespuesta = respuesta.getContentText();
    contenidoRespuesta = JSON.parse(contenidoRespuesta);
    Logger.log(contenidoRespuesta);

    if (contenidoRespuesta["DocumentId"] && contenidoRespuesta["MessageValidation"] === "Factura insertada existosamente") {
      Logger.log(contenidoRespuesta["DocumentId"], contenidoRespuesta["MessageValidation"]);
      retorno.push(true);
      retorno.push(contenidoRespuesta["DocumentId"]);
      return retorno;
    } else if (contenidoRespuesta["Message"] === "E002: El documento que intenta ingresar ya existe en el sistema") {
      SpreadsheetApp.getUi().alert("Error: La factura ya existe en el sistema de misfacturas. Por favor verifique que el número de factura sea único.");
      hojaFactura.getRange("H2").setBackground("#FFC7C7");
      return false;
    }
    else {
      SpreadsheetApp.getUi().alert(
        "Error al enviar la factura: " +
        (contenidoRespuesta["Message"] || "Respuesta desconocida")
      );
    }

  } catch (error) {
    Logger.log("Error al enviar el JSON a la API: " + error.message);
    SpreadsheetApp.getUi().alert("Error al enviar la factura a misfacturas. Intente de nuevo si el error presiste comuniquese con soporte");
  }

}

function registarEstadoFactura(idFactura, numRow) {
  let schemaID = 31;
  let documentType = 1;
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let IDNumber = hojaDatosEmisor.getRange("B3").getValue();

  let documentId = idFactura;
  let url = `https://misfacturas.cenet.ws/integrationAPI_2/api/GetDocumentStatus?SchemaID=${schemaID}&DocumentType=${documentType}&IDNumber=${IDNumber}&DocumentID=${documentId}`;

  let hojaDatos = spreadsheet.getSheetByName('Datos');
  let token = hojaDatos.getRange("F47").getValue();
  let opciones = {
    "method": "post",
    "contentType": "application/json",
    "headers": { "Authorization": "misfacturas " + token },
    'muteHttpExceptions': true
  };

  try {
    var respuesta = UrlFetchApp.fetch(url, opciones);
    var estatusRespuesta = respuesta.getResponseCode();

    if (estatusRespuesta == 200) {
      var contenidoRespuesta = respuesta.getContentText();
      contenidoRespuesta = JSON.parse(contenidoRespuesta);
      var status = contenidoRespuesta.DocumentStatus
      if (status === 74) {
        status = "Enviada";
      } else if (status === 70) {
        status = "Invalida";
      } else {
        status = "En revisión";
      }
      hojaFacturaHistorialData = spreadsheet.getSheetByName('Historial Facturas Data');
      if (numRow !== undefined) {
        hojaFacturaHistorialData.getRange(numRow, 6).setValue(status);
      } else {
        return status;
      }
    } else {
      var contenidoRespuesta = JSON.parse(respuesta.getContentText());
      if (estatusRespuesta == 404 && contenidoRespuesta["Message"] === "No se encontraron resultados de facturas válidas que tengan representación gráfica") {
        SpreadsheetApp.getUi().alert("Error: No se encontraron resultados de facturas válidas que tengan representación gráfica.");
      } else {
        SpreadsheetApp.getUi().alert(
          "Error al intentar descargar la factura: " +
          (contenidoRespuesta["Message"] || "Respuesta desconocida")
        );
        Logger.log("Error al intentar descargar la factura: " + contenidoRespuesta);
      }
    }

  } catch (error) {
    Logger.log("Error al enviar el JSON a la API: " + error.message);
  }
}


function obtenerTokenMF(usuario, contra) {
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let hojaDatos = spreadsheet.getSheetByName("Datos")

  let url = `https://misfacturas.cenet.ws/IntegrationAPI_2/api/login?username=${usuario}&password=${contra}`;
  let opciones = {
    "method": "post",
    "contentType": "application/json",
    'muteHttpExceptions': true
  };

  try {
    let respuesta = UrlFetchApp.fetch(url, opciones);
    let contenidoRespuesta = respuesta.getContentText();
    Logger.log("Respuesta de la API: " + contenidoRespuesta);
    // Intentamos parsear la respuesta como JSON
    let respuestaJson;

    try {
      respuestaJson = JSON.parse(contenidoRespuesta);
    } catch (e) {
      throw new Error("Respuesta inesperada de la API. No es JSON válido.");
    }

    // Verificar si la respuesta contiene un token en el formato esperado
    if (respuestaJson.length > 0 && typeof respuestaJson === 'string') {
      let token = respuestaJson; // Extrae el API Key
      Logger.log("API Key obtenida: " + token);
      SpreadsheetApp.getUi().alert("Se ha vinculado tu cuenta exitosamente");
      hojaDatos.getRange("F49").setValue(usuario)
      hojaDatos.getRange("F50").setValue(contra)
      hojaDatosEmisor.getRange("B13").setBackground('#ccffc7')  // Almacena el API Key en la celda
      hojaDatosEmisor.getRange("B13").setValue("Vinculado")
      hojaDatos.getRange("F47").setValue(token)
      var resdian = obtenerResolucionesDian(token, usuario);
      if (!resdian) {
        return;
      }
    } else {
      hojaDatosEmisor.getRange("B13").setBackground('#FFC7C7')
      hojaDatosEmisor.getRange("B13").setValue("Desvinculado")
      hojaDatos.getRange("F47").setValue("")
      throw new Error("Error de la API: " + contenidoRespuesta); // Muestra el error de la API

    }
  } catch (error) {
    Logger.log("Error al enviar el JSON a la API: " + error.message);
    hojaDatosEmisor.getRange("B13").setBackground('#FFC7C7')
    hojaDatosEmisor.getRange("B13").setValue("Desvinculado")
    hojaDatos.getRange("F47").setValue("")
    SpreadsheetApp.getUi().alert("Error al vincular tu cuenta. Verifica que el usuario y la contraseña estén correctos e intenta de nuevo. Si el error persiste, comunícate con soporte.");
  }
}

function obtenerResolucionesDian(token, usuario) {
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let hojaDatos = spreadsheet.getSheetByName("Datos")
  let nit = hojaDatosEmisor.getRange("B3").getValue();
  if (nit === "") {

    hojaDatosEmisor.getRange("B3").setBackground("#FFC7C7");
    SpreadsheetApp.getUi().alert("Por favor ingrese el NIT del emisor en la hoja 'Datos de emisor' y vuelva a intentar obtener las resoluciones.");
    spreadsheet.setActiveSheet(hojaDatosEmisor);
    return;

  }
  else {
    hojaDatosEmisor.getRange("B3").setBackground(null);
    if (!token || !usuario) {
      usuario = hojaDatos.getRange("F49").getValue();
      logearUsuario();
      token = hojaDatos.getRange("F47").getValue();
    }
    let url = `https://misfacturas.cenet.ws/integrationAPI_2/api/GetDianResolutions?SchemaID=31&IDNumber=${nit}`;
    let opciones = {
      "method": "get",
      "headers": { "Authorization": "misfacturas " + token },
      "contentType": "application/json",
      'muteHttpExceptions': true
    };

    try {
      const respuesta = UrlFetchApp.fetch(url, opciones);
      const contenidoTexto = respuesta.getContentText(); // Obtiene el cuerpo de la respuesta como texto
      const datos = JSON.parse(contenidoTexto); // Convierte el texto a un objeto JSON

      if (datos.InvoiceAuthorizationList && datos.InvoiceAuthorizationList.length > 0) {
        const encabezados = [
          "InvoiceAuthorizationNumber",
          //"ResolutionDateTime",
          //"StartDate",
          //"EndDate",
          "Prefix",
          "From",
          "To",
          //"TechnicalKey",
          "CurrentSecuence",
          "Estado",
          //"Observaciones"
        ];


        //hojaDatosEmisor.getRange(17, 1, 1, encabezados.length).setValues([encabezados]);

        // Preparar los datos para escribirlos en el sheet
        const filas = datos.InvoiceAuthorizationList.map(item => [
          item.InvoiceAuthorizationNumber,
          //item.ResolutionDateTime,
          //item.StartDate,
          //item.EndDate,
          item.Prefix,
          item.From,
          item.To,
          //item.TechnicalKey,
          item.CurrentSecuence,
          item.Estado,
          //item.Observaciones
        ]);

        // Limpiar la hoja desde la fila 18 hasta la 30
        hojaDatosEmisor.getRange(18, 1, 13, hojaDatosEmisor.getLastColumn()).clearContent();


        // Escribir los datos en la hoja, debajo de los encabezados
        hojaDatosEmisor.getRange(18, 1, filas.length, encabezados.length).setValues(filas);


        // Set background colors
        hojaDatosEmisor.getRange(18, 1, filas.length, 1).setBackground('#d9d9d9'); // Column A
        hojaDatosEmisor.getRange(18, 2, filas.length, 1).setBackground('#d9d9d9'); // Column B
        hojaDatosEmisor.getRange(18, 3, filas.length, 1).setBackground('#d9d9d9'); // Column C
        hojaDatosEmisor.getRange(18, 4, filas.length, 1).setBackground('#d9d9d9'); // Column D
        hojaDatosEmisor.getRange(18, 5, filas.length, 1).setBackground('#edffeb'); // Column E
        hojaDatosEmisor.getRange(18, 6, filas.length, 1).setBackground('#d9d9d9'); // Column F

        return true;
      } else {
        throw new Error("Error de la API: " + contenidoRespuesta); // Muestra el error de la API
      }
    } catch (error) {

      SpreadsheetApp.getUi().alert("Error al obtener las resoluciones dian. Verifica que el NIT sea correcto e intenta de nuevo. Si el error persiste, comunícate con soporte.");
      return false;
    }
  }



}

function limpiarHojaFactura() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const hojaFactura = spreadsheet.getSheetByName('Factura');
  const copiaFactura = spreadsheet.getSheetByName('Copia de Factura');
  const hojaInicio = spreadsheet.getSheetByName('Inicio');

  if (!copiaFactura) {
    Logger.log("No se encontró la hoja 'Copia facturas'.");
    return;
  }
  spreadsheet.setActiveSheet(hojaInicio)
  // Si existe la hoja Factura, elimínala
  if (hojaFactura) {
    spreadsheet.deleteSheet(hojaFactura);
  }

  // Copiar la hoja "Copia facturas" como nueva hoja llamada "Factura"
  const nuevaHojaFactura = copiaFactura.copyTo(spreadsheet);
  nuevaHojaFactura.setName('Factura');
  const hojaFacturaPost = spreadsheet.getSheetByName('Factura');
  spreadsheet.setActiveSheet(hojaFacturaPost)
  Logger.log("La hoja 'Factura' ha sido reemplazada correctamente.");
  grabarRangoResolucionesDian(hojaFacturaPost);
}

function inicarFacturaNueva() {
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  let hojaDatos = spreadsheet.getSheetByName('Datos');
  let consecutivoFactura = hojaDatos.getRange("Q11").getValue();
  hojaFactura.getRange("H2").setValue(consecutivoFactura);
  ponerFechaYHoraActual();
}

function verificarYCopiarCliente(e) {
  let hojaFacturas = e.source.getSheetByName('Factura');
  let hojaClientes = e.source.getSheetByName('Clientes');
  let celdaEditada = e.range;

  let nombreCliente = celdaEditada.getValue();
  let ultimaColumnaPermitida = 20; // Columna del estado en la hoja de clientes
  let datosARetornar = ["B", "O", "M", "L", "N", "Q"]; // Columnas que quiero de la hoja de clientes


  if (nombreCliente === "Cliente") {
    Logger.log("Estado default")
  } else {
    let listaConInformacion = obtenerInformacionCliente(nombreCliente);
    if (listaConInformacion["Estado"] === "No Valido") {
      SpreadsheetApp.getUi().alert("Error: El cliente seleccionado no es válido.");
    } else {
      //asigna el valor del coldigo solamente porque ese fue lo que me pidieron no mas
      hojaFacturas.getRange("B3").setValue(listaConInformacion["Código cliente"]);
    }
  }


}

function ponerFechaYHoraActual() {
  let spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName('Factura');

  let fecha = new Date();
  let fechaFormateada = Utilities.formatDate(fecha, "America/Bogota", "yyyy-MM-dd");
  let horaFormateada = Utilities.formatDate(fecha, "America/Bogota", "HH:mm:ss");

  sheet.getRange("H4").setNumberFormat("@");
  sheet.getRange("H4").setValue(fechaFormateada);

  sheet.getRange("H6").setNumberFormat("@");
  sheet.getRange("H6").setValue(horaFormateada);
}

function ponerFechaTasaDeCambio() {
  let spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName('Factura');
  let fecha = new Date();
  let fechaFormateada = Utilities.formatDate(fecha, "America/Bogota", "yyyy-MM-dd");
  sheet.getRange("J6").setNumberFormat("@");
  sheet.getRange("J6").setValue(fechaFormateada);
}

function obtenerFecha() {
  let fechaFormateada
  let spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName('Factura');
  let valorFecha = sheet.getRange("H4").getValue();
  fechaFormateada = Utilities.formatDate(new Date(valorFecha), "America/Bogota", "yyyy-MM-dd");
  return fechaFormateada
}

function obtenerDatosProductos(sheet, range, e) {
  if (range.getA1Notation() === "A14" || range.getA1Notation() === "A15" || range.getA1Notation() === "A16" || range.getA1Notation() === "A17" || range.getA1Notation() === "A18") {
    Logger.log("entro a obtenerdatos")
    var selectedProduct = range.getValue();


    // Referencia a la hoja de productos
    var productSheet = e.source.getSheetByName("Productos");
    var data = productSheet.getDataRange().getValues();

    // Encuentra el producto en la hoja de productos
    for (var i = 1; i < data.length; i++) {
      Logger.log(data[i][1])
      Logger.log(selectedProduct)
      if (data[i][1] == selectedProduct) {
        sheet.getRange("B14").setValue(data[i][0]);  // Código de referencia
        sheet.getRange("D14").setValue(data[i][2]);  // Precio unitario
        break;
      }
    }
  }

}

function grabarRangoResolucionesDian(hojaFactura) {
  var hojaDatosEmisor = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Datos de emisor");

  var columnaBase = hojaDatosEmisor.getRange("A18:A").getValues(); // Ajustar la columna según tu necesidad
  var valoresValidos = [];

  // Filtrar valores no vacíos
  for (var i = 0; i < columnaBase.length; i++) {
    if (columnaBase[i][0] !== "") {
      valoresValidos.push(columnaBase[i][0]);
    }
  }

  if (valoresValidos.length === 0) {
    Logger.log("No hay valores válidos en el rango.");
    return;
  }

  // Crear el rango de validación dinámico
  var reglaDeValidacion = SpreadsheetApp.newDataValidation()
    .requireValueInList(valoresValidos, true) // Lista desplegable basada en los valores encontrados
    .setAllowInvalid(false) // No permitir valores fuera de la lista
    .build();

  // Aplicar la validación al rango donde quieres el dropdown
  var rangoDropdown = hojaFactura.getRange("H3"); // Ajusta el rango donde irá el dropdown
  rangoDropdown.setDataValidation(reglaDeValidacion);
  Logger.log("Dropdown creado exitosamente.");

}

function getInvoiceGeneralInformation() {
  let spreadsheet = SpreadsheetApp.getActive();
  let prefactura_sheet = spreadsheet.getSheetByName('Factura');

  //Recuperar los datos de la factura del sheets
  var numeroAutorizacion = prefactura_sheet.getRange("H3").getValue();//Resolución DIAN
  var numeroFactura = prefactura_sheet.getRange("H2").getValue();
  var fechaEmision = prefactura_sheet.getRange("H4").getValue() + "T" + prefactura_sheet.getRange("H6").getValue();//fecha de emision
  var diasVencimiento = prefactura_sheet.getRange("H5").getValue();//dias de vencimiento
  var moneda = prefactura_sheet.getRange("J4").getValue();//moneda
  moneda = moneda.split("-")[0];
  var exchangeRate = prefactura_sheet.getRange("J5").getValue();//tasa de cambio
  var exchangeRateDate = prefactura_sheet.getRange("J6").getValue();//fecha de tasa de cambio
  var observaciones = prefactura_sheet.getRange("B10").getValue();//observaciones
  var fechaVencimiento = SumarDiasAFecha(diasVencimiento, prefactura_sheet.getRange("H4").getValue());


  //Agregar para el json
  var InvoiceGeneralInformation = {
    "InvoiceAuthorizationNumber": String(numeroAutorizacion),
    "PreinvoiceNumber": String(numeroFactura),
    "InvoiceNumber": String(numeroFactura),
    "IssueDate": fechaEmision,
    "Prefix": buscarPrefijo(numeroAutorizacion),
    "DaysOff": String(diasVencimiento),
    "Currency": moneda,
    "ExchangeRate": exchangeRate,
    "ExchangeRateDate": exchangeRateDate,
    "CustomizationID": "10",
    "SalesPerson": "",
    "Note": observaciones,
    "ExternalGR": false,
    "StartDateTime": "0001-01-01T00:00:00",
    "EndDateTime": "0001-01-01T00:00:00",
    "InvoiceDueDate": fechaVencimiento
  }
  return InvoiceGeneralInformation;
}

function buscarPrefijo(numeroAutorizacion) {
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let fila = 18;
  let valorCelda;

  while (true) {
    valorCelda = hojaDatosEmisor.getRange(fila, 1).getValue();
    if (valorCelda === "") {
      return null; // No existe el número de autorización
    }
    if (valorCelda == numeroAutorizacion) {
      return hojaDatosEmisor.getRange(fila, 2).getValue(); // Retorna el valor de la columna 2
    }
    fila++;
  }
}

function getPaymentSummary() {
  let spreadsheet = SpreadsheetApp.getActive();
  let prefactura_sheet = spreadsheet.getSheetByName('Factura');
  var PaymentTypeTxt = prefactura_sheet.getRange("J2").getValue();
  var PaymentMeansTxt = prefactura_sheet.getRange("J3").getValue();
  PaymentMeansTxt = PaymentMeansTxt.split("-")[1];
  var PaymentSummary = {
    "PaymentType": metodosPago[PaymentTypeTxt],
    "PaymentMeans": paymentMeansCode[PaymentMeansTxt],
    "PaymentNote": ""
  }
  return PaymentSummary;
}

function guardarYGenerarInvoice() {
  Logger.log("Guardando y generando invoice");
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaProductos = spreadsheet.getSheetByName('Productos');
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let listadoestado_sheet = spreadsheet.getSheetByName('ListadoEstado');
  let consecutivoFactura = hojaFactura.getRange("H2").getValue();
  let consecutivoFacturaActualizado = consecutivoFactura + 1;
  let numeroAutorizacion = hojaFactura.getRange("H3").getValue();
  for (let i = 18; i <= 20; i++) {
    if (hojaDatosEmisor.getRange(i, 1).getValue() == numeroAutorizacion) {
      consecutivoFactura = hojaDatosEmisor.getRange(i, 5).setValue(consecutivoFacturaActualizado);
      break;
    }
  }

  //obtener el total de prodcutos
  let posicionTotalProductos = hojaFactura.getRange("A16").getValue(); // para verificar donde esta el TOTAL
  if (posicionTotalProductos === "Total items") {
    var cantidadProductos = hojaFactura.getRange("B16").getValue();// cantidad total de productos 
    var cantidadFilasProductos = 1;
  } else {
    let startingRowTax = getcargosDescuentosStartRow(hojaFactura)
    let posicionTotalProductos = startingRowTax - 2
    var cantidadProductos = hojaFactura.getRange("B" + String(posicionTotalProductos)).getValue();// cantidad total de productos
    var cantidadFilasProductos = posicionTotalProductos - 15
    Logger.log("Cantidad de productos: " + cantidadProductos);
    Logger.log("Cantidad de filas de productos: " + cantidadFilasProductos);
  }

  let llavesParaLinea = hojaFactura.getRange("A14:L14");//llamo los headers 
  llavesParaLinea = slugifyF(llavesParaLinea.getValues()).replace(/\s/g, ''); // Todo en una sola linea
  const llavesFinales = llavesParaLinea.split(",");
  /* Creo que esto se puede cambiar a una manera mas simple, ya que los headers de la fila H7 hatsa N7 nunca van a cambiar */

  var productoInformation = [];

  let i = 15 // es 15 debido a que aqui empieza los productos elegidos por el cliente
  do {
    let filaActual = "A" + String(i) + ":L" + String(i);
    let rangoProductoActual = hojaFactura.getRange(filaActual);
    let productoFilaActual = String(rangoProductoActual.getValues());
    productoFilaActual = productoFilaActual.split(",");// cojo el producto de la linea actual y se le hace split a toda la info
    let LineaFactura = {};
    Logger.log(i);
    if (productoFilaActual[0] === "") {
      Logger.log("No hay productos en la fila " + i);
      i++;
      continue;
    }


    for (let j = 0; j < 12; j++) {
      LineaFactura[llavesFinales[j]] = productoFilaActual[j]
    }

    let filaProducto = obtenerFilaPorReferencia(Number(LineaFactura['referencia']));
    let unidadDeMedida = hojaProductos.getRange("G" + String(filaProducto)).getValue();
    Logger.log("Unidad de medida: " + unidadDeMedida);
    unidadDeMedida = unidadDeMedida.split("-")[1];
    Logger.log("Unidad de medida: " + unidadDeMedida);
    let ItemReference = String(LineaFactura['referencia']);
    let Name = String(LineaFactura['producto']);
    let Quantity = String(LineaFactura['cantidad']);
    let Price = Number(LineaFactura['preciounitario']);
    let chargeIndicator = false;
    let LineTotalTaxes = Number(LineaFactura['impuestos']);
    let LineTotal = parseFloat(LineaFactura['totaldelinea']);
    let LineExtensionAmount = parseFloat(LineaFactura['subtotal']);
    let TotalCargosLinea = Number(LineaFactura['cargos']);
    let TotalDescuentoLinea = Number(LineaFactura['descuento%']) * 100;
    let MeasureUnitCode = String(unidadDeMedida);

    let ItemTaxesInformation = [];



    function obtenerFilaPorReferencia(referencia) {
      const hojaProductos = spreadsheet.getSheetByName('Productos');
      const ultimaFila = hojaProductos.getLastRow();
      const rangoCodigosReferencia = hojaProductos.getRange(2, 2, ultimaFila - 1, 1).getValues();

      for (let i = 0; i < rangoCodigosReferencia.length; i++) {
        if (rangoCodigosReferencia[i][0] === referencia) {
          return i + 2; // +2 porque el rango empieza en la fila 2
        }
      }

      return -1; // Retorna -1 si no se encuentra la referencia
    }

    function agregarCargosDescuentos() {
      let AllowanceCharge = [];
      if (TotalDescuentoLinea > 0) {
        let Allowance = {
          "Id": 9,
          "ChargeIndicator": chargeIndicator,
          "AllowanceChargeReason": "",
          "MultiplierFactorNumeric": Number(TotalDescuentoLinea),
          "Amount": Number(Price) * Number(Quantity) * (TotalDescuentoLinea / 100),
          "BaseAmount": Number(Price) * Number(Quantity)
        }
        AllowanceCharge.push(Allowance);

      }
      if (TotalCargosLinea > 0) {
        let Charge = {
          "Id": 20,
          "ChargeIndicator": !chargeIndicator,
          "AllowanceChargeReason": "",
          "Amount": TotalCargosLinea,
          "BaseAmount": Number(Price) * Number(Quantity)
        }
        AllowanceCharge.push(Charge);

      }

      return AllowanceCharge;
    }

    function agregarImpuestos() {
      //taxes del producto en si
      if (LineaFactura["iva%"] > 0) {
        let percentIva = convertToPercentage(LineaFactura["iva%"]);
        let ivaTaxInformation = {
          Id: "01",//Id
          TaxEvidenceIndicator: false,
          TaxableAmount: LineExtensionAmount,
          TaxAmount: LineExtensionAmount * LineaFactura["iva%"],
          Percent: Number(percentIva),
          BaseUnitMeasure: 0,
          PerUnitAmount: 0,
        }
        ItemTaxesInformation.push(ivaTaxInformation);
      }


      if (LineaFactura["inc%"] > 0) {
        let percentInc = convertToPercentage(LineaFactura["inc%"]);
        let incTaxInformation = {
          Id: "04",//Id
          TaxEvidenceIndicator: false,
          TaxableAmount: LineExtensionAmount,
          TaxAmount: LineExtensionAmount * LineaFactura["inc%"],
          Percent: Number(percentInc),
          BaseUnitMeasure: 0,
          PerUnitAmount: 0,
        };
        ItemTaxesInformation.push(incTaxInformation);
      }

      if (LineaFactura["retencion"] > 0) {
        let retencionTaxInformation = {
          Id: "06",
          TaxEvidenceIndicator: true,
          TaxableAmount: Number(LineExtensionAmount),
          TaxAmount: 0,
          Percent: 0,
          BaseUnitMeasure: 0,
          PerUnitAmount: 0,
        };
        let nombreYporcentajeRetencion = buscarRetencion(LineaFactura["producto"]);
        let porcentajeRetencion = Number(nombreYporcentajeRetencion[1]) * 100;
        retencionTaxInformation.Percent = Number(porcentajeRetencion.toFixed(3));
        retencionTaxInformation.TaxAmount = Number(LineExtensionAmount) * Number(porcentajeRetencion) / 100;
        ItemTaxesInformation.push(retencionTaxInformation);
      }

      return ItemTaxesInformation;
    }


    let productoI = {//aqui organizamos todos los parametros necesarios para los productos
      ItemReference: ItemReference,
      Name: Name,
      Quatity: new Number(Quantity),
      Price: new Number(Price),

      LineAllowanceTotal: TotalDescuentoLinea,
      LineChargeTotal: TotalCargosLinea,
      LineTotalTaxes: LineTotalTaxes,
      LineTotal: LineTotal,
      LineExtensionAmount: LineExtensionAmount,
      MeasureUnitCode: MeasureUnitCode,
      FreeOFChargeIndicator: chargeIndicator,
      AdditionalReference: [],
      Nota: "",
      AdditionalProperty: [],
      TaxesInformation: agregarImpuestos(),
      AllowanceCharge: agregarCargosDescuentos()
    };


    productoInformation.push(productoI);//agregamos el producto actual a la lista total 
    i++;
  } while (i < (15 + cantidadFilasProductos));


  // Función para obtener todos los impuestos de productoInformation
  function obtenerTodosLosImpuestos(productoInformation) {
    let todosLosImpuestos = [];

    for (let i = 0; i < productoInformation.length; i++) {
      let impuestosProducto = productoInformation[i].TaxesInformation;
      if (impuestosProducto && impuestosProducto.length > 0) {
        for (let j = 0; j < impuestosProducto.length; j++) {
          todosLosImpuestos.push(impuestosProducto[j]);
        }
      }
    }

    return todosLosImpuestos;
  }

  // Función para agrupar impuestos
  function agruparImpuestos(todosLosImpuestos) {
    let impuestosAgrupados = [];

    for (let i = 0; i < todosLosImpuestos.length; i++) {
      let impuestoActual = todosLosImpuestos[i];
      let encontrado = false;

      for (let j = 0; j < impuestosAgrupados.length; j++) {
        let impuestoAgrupado = impuestosAgrupados[j];

        if (impuestoAgrupado.Id === impuestoActual.Id && impuestoAgrupado.Percent === impuestoActual.Percent) {
          impuestoAgrupado.TaxableAmount += impuestoActual.TaxableAmount;
          impuestoAgrupado.TaxAmount += impuestoActual.TaxAmount;
          encontrado = true;
          break;
        }
      }

      if (!encontrado) {
        impuestosAgrupados.push({ ...impuestoActual });
      }
    }

    return impuestosAgrupados;
  }
  function agregarCargosDescuentosTotales(subtotal) {
    let hojaFactura = spreadsheet.getSheetByName('Factura');

    let rowSeccionCargosyDescuentos = getcargosDescuentosStartRow(hojaFactura) + 2;
    let lastCargoDescuentoRow = getLastCargoDescuentoRow(hojaFactura);
    let CargosyDescuentos = [];
    let chargeIndicator = false;
    for (let i = rowSeccionCargosyDescuentos; i <= lastCargoDescuentoRow + 1; i++) {
      let celdaValorPorcentaje = hojaFactura.getRange("C" + String(i)).getValue()
      if (hojaFactura.getRange("A" + String(i)).getValue() === "Cargo") {
        let Charge = {
          "Id": 20,
          "ChargeIndicator": !chargeIndicator,
          "AllowanceChargeReason": hojaFactura.getRange("B" + String(i)).getValue(),
          "MultiplierFactorNumeric": 0,
          "Amount": hojaFactura.getRange("E" + String(i)).getValue(),
          "BaseAmount": subtotal
        }
        if (String(celdaValorPorcentaje).includes("%")) {
          Charge.MultiplierFactorNumeric = Number(celdaValorPorcentaje.replace("%", ""))
        }
        CargosyDescuentos.push(Charge);
      }
      else if (hojaFactura.getRange("A" + String(i)).getValue() === "Descuento") {
        let Allowance = {
          "Id": 9,
          "ChargeIndicator": chargeIndicator,
          "AllowanceChargeReason": hojaFactura.getRange("B" + String(i)).getValue(),
          "MultiplierFactorNumeric": 0,
          "Amount": hojaFactura.getRange("E" + String(i)).getValue(),
          "BaseAmount": subtotal
        }
        if (String(celdaValorPorcentaje).includes("%")) {
          Allowance.MultiplierFactorNumeric = Number(celdaValorPorcentaje.replace("%", ""))
        }
        CargosyDescuentos.push(Allowance);
      }
    }

    return CargosyDescuentos;
  }





  //estos es dinamico, verificar donde va el total cargo y descuento
  const posicionOriginalTotalFactura = hojaFactura.getRange("A20").getValue(); // para verificar donde esta el TOTAL
  let rangeTotales = ""


  if (posicionOriginalTotalFactura === "Subtotal") {
    rangeTotales = hojaFactura.getRange(posicionOriginalTotalFactura + 1, 1, 1, 12);//va a cambiar

  } else {
    let rowTotales = getTotalesLinea(hojaFactura)
    rangeTotales = hojaFactura.getRange(rowTotales, 1, 1, 12);
  }

  let totalesValores = String(rangeTotales.getValues())
  totalesValores = totalesValores.split(",")


  //Definir los valores para el json
  Logger.log("Valores totales: " + totalesValores);
  let pfSubTotal = parseFloat(totalesValores[0]);
  let pfBaseGrabable = parseFloat(totalesValores[1]);
  let pfSubTotalMasImpuestos = parseFloat(totalesValores[3]);
  let pfRetenciones = parseFloat(totalesValores[4]);
  let pfDescuentos = parseFloat(totalesValores[5]);
  let pfCargos = parseFloat(totalesValores[7]);
  let pfAnticipo = Number(totalesValores[9]);
  let pfNetoAPagar = Number(totalesValores[11]);


  let invoice_total = {
    "lineExtensionamount": pfSubTotal,
    "TaxExclusiveAmount": pfBaseGrabable,
    "TaxInclusiveAmount": pfSubTotalMasImpuestos,
    "AllowanceTotalAmount": pfDescuentos,
    "ChargeTotalAmount": pfCargos,
    "PrePaidAmount": pfAnticipo,
    "PayableAmount": pfNetoAPagar
  }


  let cliente = hojaFactura.getRange("B2").getValue();
  let InvoiceGeneralInformation = getInvoiceGeneralInformation();
  let CustomerInformation = getCustomerInformation(cliente);

  let sheetDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let userId = String(sheetDatosEmisor.getRange("B11").getValue());
  let companyId = String(sheetDatosEmisor.getRange("B3").getValue());
  let PaymentSummary = getPaymentSummary();

  let nuevoInvoiceResumido = JSON.stringify({
    "file": "base64",
    "Document": {
      "fileName": "nombre documento",
      "userId": userId,
      "companyId": companyId,
      "invoice": {
        "invoiceType": false,
        "contactName": "",
        "numeroIdentificacion": "",
        "invoiceDate": "",
        "numberInvoice": "",
        "taxableAmount": "",
        "Percent": "0",
        "taxAmount": '',
        "surchargeAmount": "el valor no se debe de reportar",
        "surchargeValue": "el valor no se debe de reportar",
        "PercentSurchargeEquivalence": "0",
        "PercentageRetention": "0",
        "IRPFValue": "el valor no se debe de reportar",
        "invoiceTotal": '',
        "payDate": "",
        "PaymentType": "",
        "Observations": ""
      }
    }
  }
  );

  let invoice = JSON.stringify({
    CustomerInformation: CustomerInformation,
    InvoiceGeneralInformation: InvoiceGeneralInformation,
    Delivery: getDelivery(),
    AdditionalDocuments: [],
    PaymentSummary: PaymentSummary, //por ahora esto leugo se cambia la funcion getPaymentSummary para que cumpla los parametros
    ItemInformation: productoInformation,
    InvoiceTaxTotal: agruparImpuestos(obtenerTodosLosImpuestos(productoInformation)),
    InvoiceTaxOthersTotal: null,
    InvoiceAllowanceCharge: agregarCargosDescuentosTotales(pfSubTotal),
    InvoiceTotal: invoice_total,
    Documents: []
  });

  let nombreCliente = hojaFactura.getRange("B2").getValue();
  let numeroFactura = InvoiceGeneralInformation.InvoiceNumber;
  let fecha = obtenerFecha();
  let codigoCliente = hojaFactura.getRange("B3").getValue();
  listadoestado_sheet.appendRow([fecha, numeroFactura, nombreCliente, codigoCliente, invoice, nuevoInvoiceResumido]);
}

function guardarFacturaHistorial(documentId) {
  var jsonString = recuperarJson();
  var json = JSON.parse(jsonString); // Convertir el JSON string a un objeto

  Logger.log(typeof json);
  Logger.log(json["InvoiceGeneralInformation"]);

  var prefijo = json.InvoiceGeneralInformation.Prefix || "";

  var hojaFactura = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Factura');
  var hojaListado = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Historial Facturas Data');
  var numeroFactura = hojaFactura.getRange("H2").getValue();
  var cliente = hojaFactura.getRange("B2").getValue();
  var fechaEmision = hojaFactura.getRange("H4").getValue();
  var informacionCliente = getCustomerInformation(cliente);
  var identificacion = informacionCliente.Identification;
  var lastRow = hojaListado.getLastRow();
  var newRow = lastRow + 1;

  var celdaPrefijo = hojaListado.getRange("A" + newRow);
  celdaPrefijo.setValue(prefijo);
  celdaPrefijo.setHorizontalAlignment('center');
  celdaPrefijo.setBorder(true, true, true, true, null, null, null, null);

  var celdaNumFactura = hojaListado.getRange("B" + newRow);
  celdaNumFactura.setValue(numeroFactura);
  celdaNumFactura.setHorizontalAlignment('center');
  celdaNumFactura.setBorder(true, true, true, true, null, null, null, null);

  var celdaCliente = hojaListado.getRange("C" + newRow);
  celdaCliente.setValue(cliente);
  celdaCliente.setHorizontalAlignment('center');
  celdaCliente.setBorder(true, true, true, true, null, null, null, null);

  var celdaIdentificacion = hojaListado.getRange("D" + newRow);
  celdaIdentificacion.setValue(identificacion);
  celdaIdentificacion.setHorizontalAlignment('center');
  celdaIdentificacion.setBorder(true, true, true, true, null, null, null, null);

  var celdaFecha = hojaListado.getRange("E" + newRow);
  celdaFecha.setValue(fechaEmision);
  celdaFecha.setHorizontalAlignment('center');
  celdaFecha.setBorder(true, true, true, true, null, null, null, null);

  let hojaListadoEstado = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ListadoEstado');
  let lineaJson = hojaListadoEstado.getLastRow();
  hojaListadoEstado.getRange(lineaJson, 7).setValue(documentId);
  registarEstadoFactura(documentId, newRow);
  limpiarHojaFactura();
}

function filtroHistorialFacturas(tipoFiltro) {

  Logger.log("debtro de fitrlo historial")
  Logger.log("tipoFiltro " + tipoFiltro)
  let Formula = ''
  let hojahistorial = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Historial Facturas');

  if (tipoFiltro == "Numero Idenificacion") {
    Formula = "=FILTER('Historial Facturas Data'!A2:F1000;ISNUMBER(SEARCH(F5;'Historial Facturas Data'!D2:D1000)))"
  } else if (tipoFiltro == "Nombre Cliente") {
    Formula = "=FILTER('Historial Facturas Data'!A2:F1000;ISNUMBER(SEARCH(F5;'Historial Facturas Data'!C2:C1000)))"
  } else if (tipoFiltro == "Numero Factura") {
    Formula = "=FILTER('Historial Facturas Data'!A2:F1000;ISNUMBER(SEARCH(F5;'Historial Facturas Data'!B2:B1000)))"
  } else if (tipoFiltro == "Fecha Emision") {
    Formula = "=FILTER('Historial Facturas Data'!A2:F1000;ISNUMBER(SEARCH(F5;'Historial Facturas Data'!E2:E1000)))"
  } else if (tipoFiltro == "Prefijo") {
    Formula = "=FILTER('Historial Facturas Data'!A2:F1000;ISNUMBER(SEARCH(F5;'Historial Facturas Data'!A2:A1000)))"
  } else if (tipoFiltro == "") {
    Formula = `=FILTER('Historial Facturas Data'!A2:F, 'Historial Facturas Data'!A2:A <> "")`
  }

  hojahistorial.getRange("B8").setValue(Formula)

}

function actualizarEstadoUltimasFacturas() {
  Logger.log("Actualizando estado de las últimas facturas...");
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let hojaListadoEstado = spreadsheet.getSheetByName('ListadoEstado');
  let hojaHistorialFacturasData = spreadsheet.getSheetByName('Historial Facturas Data');
  let lastRowListadoEstado = hojaListadoEstado.getLastRow();
  let lastRowHistorialFacturasData = hojaHistorialFacturasData.getLastRow();
  let numFacturas = Math.min(10, lastRowListadoEstado, lastRowHistorialFacturasData); // Obtener el número de facturas a procesar (máximo 10)

  for (let i = 0; i < numFacturas; i++) {
    let rowListadoEstado = lastRowListadoEstado - i;
    let rowHistorialFacturasData = lastRowHistorialFacturasData - i;
    let documentId = hojaListadoEstado.getRange(rowListadoEstado, 7).getValue(); // Obtener el documentId de la columna 7

    if (documentId !== "" && documentId !== null && documentId !== undefined) {
      let estado = registarEstadoFactura(documentId);
      if (estado) {
        hojaHistorialFacturasData.getRange(rowHistorialFacturasData, 6).setValue(estado); // Actualizar el estado en la columna 6
      }
    }
  }
}

function SumarDiasAFecha(dias, fecha) {

  // Convertir la fecha de string a objeto Date
  var partes = fecha.split("-");
  var fechaObj = new Date(partes[0], partes[1] - 1, partes[2]); // Año, Mes (0-indexado), Día

  // Sumar los días
  fechaObj.setDate(fechaObj.getDate() + dias);

  // Formatear la nueva fecha a yyyy-MM-dd
  var nuevaFecha = Utilities.formatDate(fechaObj, "GMT", "yyyy-MM-dd");

  return nuevaFecha;
}

function mostrarResumenFactura() {
  var jsonString = recuperarJson();
  var json = JSON.parse(jsonString); // Convertir el JSON string a un objeto

  // Extraer la información importante
  var nombreCliente = json.CustomerInformation.RegistrationName;
  var numeroFactura = json.InvoiceGeneralInformation.InvoiceNumber;

  // Extraer la información de impuestos
  var impuestos = json.InvoiceTaxTotal.map(function (tax) {
    var tipoImpuesto = tax.Id === "01" ? "IVA" : tax.Id === "04" ? "INC" : "ReteRenta";
    return {
      tipo: tipoImpuesto,
      percent: tax.Percent,
      amount: tax.TaxAmount
    };
  });

  // Extraer la información de InvoiceTotal
  var invoiceTotal = json.InvoiceTotal;

  // Crear el contenido HTML para el cuadro de diálogo
  var htmlContent = plantillaResumenFactura(nombreCliente, numeroFactura, impuestos, invoiceTotal);

  // Mostrar el cuadro de diálogo
  var htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(600)
    .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Resumen de la Factura');
}

function modificarFactura() {
  // Cerrar el cuadro de diálogo y retornar un valor para evitar que se ejecuten los siguientes UI
  return false;
}

function enviarFacturaHtml() {
  let respuesta = enviarFactura();
  if (respuesta[0] == true) {
    guardarFacturaHistorial(respuesta[1]);
    ui = SpreadsheetApp.getUi();
    ui.alert("Factura enviada correctamente, puede descargarla desde la hoja 'Historial Facturas'.");
  }
}