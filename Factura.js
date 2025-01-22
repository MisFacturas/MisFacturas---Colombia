

var spreadsheet = SpreadsheetApp.getActive();
var prefactura_sheet = spreadsheet.getSheetByName('Factura');
var listadoestado_sheet = spreadsheet.getSheetByName('ListadoEstado');
var hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
var folderId = hojaDatosEmisor.getRange("B13").getValue();


function verificarEstadoValidoFactura(estadoFactura) {
  let hojaFactura = spreadsheet.getSheetByName('Factura');

  // Verificar datos de la factura 

  let estaValido = true;
  estadoFactura.push(estaValido);

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
  if (moneda === "Pesos Colombianos") {
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

  if (totalProductos === "Total filas") {
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
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("Revisando validez de la factura. Aguarde unos segundos");
  let estadoFactura = [];
  verificarEstadoValidoFactura(estadoFactura);
  if (estadoFactura[0] === true) {
    guardarYGenerarInvoice();
    logearUsuario();
    enviarFactura();
  } else {
    let mensajeError = "La factura no es válida. Por favor rellene los campos obligatorios:\n" + estadoFactura.join("\n- ");
    SpreadsheetApp.getUi().alert(mensajeError);
  }
}
function agregarFilaNueva() {
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  let numeroFilasParaAgregar = hojaFactura.getRange("B13").getValue();

  // Verificar si numeroFilasParaAgregar es nulo, vacío o no es un número
  if (numeroFilasParaAgregar == 0 || numeroFilasParaAgregar == "" || isNaN(numeroFilasParaAgregar)) {
    SpreadsheetApp.getUi().alert("Error: Por favor, ingresa un número válido de filas para agregar.");
    return; // Detener la ejecución si hay error
  }

  let cargosDescuentosStartRow = getcargosDescuentosStartRow(hojaFactura);
  const productStartRow = 15;
  const lastProductRow = getLastProductRow(hojaFactura, productStartRow, cargosDescuentosStartRow);

  Logger.log("Agregar fila nueva");
  hojaFactura.insertRows(lastProductRow, numeroFilasParaAgregar);
}

function agregarFilaCargoDescuento() {
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  const lastCargoDescuentoRow = getLastCargoDescuentoRow(hojaFactura);
  hojaFactura.insertRowAfter(lastCargoDescuentoRow)
}

function agregarProductoDesdeFactura(cantidad, producto) {
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
  let hojaFactura = spreadsheet.getSheetByName('ListadoEstado');
  let lastRow = hojaFactura.getLastRow();
  let infoFactura = hojaFactura.getRange(lastRow, 1, 1, 20).getValues();
  let json = infoFactura[0][9];
  Logger.log(json);
  return json;
}

function logearUsuario() {
  let hojaDatos = spreadsheet.getSheetByName('Datos');
  let usuario = hojaDatos.getRange("F49").getValue();
  let contrasena = hojaDatos.getRange("F50").getValue();

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
}

function enviarFactura() {
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
      SpreadsheetApp.getUi().alert("Factura enviada correctamente a misfacturas. Si desea verla ingrese a https://misfacturas-qa.cenet.ws/Aplicacion/");
      limpiarHojaFactura();
    } else {
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


function obtenerTokenMF(usuario, contra) {
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
      obtenerResolucionesDian(token, usuario);
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
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let hojaDatos = spreadsheet.getSheetByName("Datos")
  if (!token || !usuario) {
    usuario = hojaDatos.getRange("F49").getValue();
    logearUsuario();
    token = hojaDatos.getRange("F47").getValue();
  }

  let url = `https://misfacturas.cenet.ws/integrationAPI_2/api/GetDianResolutions?SchemaID=31&IDNumber=${usuario}`;
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

      // Escribir los datos en la hoja, debajo de los encabezados
      hojaDatosEmisor.getRange(18, 1, filas.length, encabezados.length).setValues(filas);

    } else {
      throw new Error("Error de la API: " + contenidoRespuesta); // Muestra el error de la API

    }
  } catch (error) {

    SpreadsheetApp.getUi().alert("Error al obtener las resoluciones dian. Verifica que el usuario y la contraseña estén correctos e intenta de nuevo. Si el error persiste, comunícate con soporte.");
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
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  let hojaDatos = spreadsheet.getSheetByName('Datos');
  let consecutivoFactura = hojaDatos.getRange("Q11").getValue();
  hojaFactura.getRange("H2").setValue(consecutivoFactura);
  ponerFechaYHoraActual();
}

function limpiarYEliminarFila(numeroFila, hoja, hojaTax) {
  //funcion para el boton que se va a agregar al final del producto
  if (numeroFila > 20 && numeroFila < hojaTax) {
    hoja.deleteRow(numeroFila)
  } else {
    hoja.getRange("A" + String(numeroFila)).setValue("");//referencia
    hoja.getRange("B" + String(numeroFila)).setValue("");//producto
    hoja.getRange("C" + String(numeroFila)).setValue("");//cantidad
    hoja.getRange("D" + String(numeroFila)).setValue(0);//precio unitario

  }
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
  let sheet = spreadsheet.getSheetByName('Factura');
  let fecha = new Date();
  let fechaFormateada = Utilities.formatDate(fecha, "America/Bogota", "yyyy-MM-dd");
  sheet.getRange("J6").setNumberFormat("@");
  sheet.getRange("J6").setValue(fechaFormateada);
}

function obtenerFecha() {
  let fechaFormateada
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

  //Recuperar los datos de la factura del sheets
  var numeroAutorizacion = prefactura_sheet.getRange("H3").getValue();//Resolución DIAN
  var numeroFactura = prefactura_sheet.getRange("H2").getValue();
  var fechaEmision = prefactura_sheet.getRange("H4").getValue() + "T" + prefactura_sheet.getRange("H6").getValue();//fecha de emision
  var diasVencimiento = prefactura_sheet.getRange("H5").getValue();//dias de vencimiento
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
    "Prefix": "SETT",
    "DaysOff": String(diasVencimiento),
    "Currency": "COP",
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

function getPaymentSummary() {
  var PaymentTypeTxt = prefactura_sheet.getRange("J2").getValue();
  var PaymentMeansTxt = prefactura_sheet.getRange("J3").getValue();
  var PaymentSummary = {
    "PaymentType": metodosPago[PaymentTypeTxt],
    "PaymentMeans": paymentMeansCode[PaymentMeansTxt],
    "PaymentNote": ""
  }
  return PaymentSummary;
}

function guardarYGenerarInvoice() {
  let hojaProductos = spreadsheet.getSheetByName('Productos');
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
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
  if (posicionTotalProductos === "Total filas") {
    var cantidadProductos = hojaFactura.getRange("B16").getValue();// cantidad total de productos 
  } else {
    let startingRowTax = getcargosDescuentosStartRow(hojaFactura)
    let posicionTotalProductos = startingRowTax - 2
    var cantidadProductos = hojaFactura.getRange("B" + String(posicionTotalProductos)).getValue();// cantidad total de productos
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

    for (let j = 0; j < 12; j++) {
      LineaFactura[llavesFinales[j]] = productoFilaActual[j]
    }
    let filaProducto = obtenerFilaPorReferencia(Number(LineaFactura['referencia']));
    let unidadDeMedida = hojaProductos.getRange("G" + String(filaProducto)).getValue();

    let ItemReference = String(LineaFactura['referencia']);
    let Name = String(LineaFactura['producto']);
    let Quantity = String(LineaFactura['cantidad']);
    let Price = Number(LineaFactura['preciounitario']);
    let LineChargeTotal = Number(LineaFactura['cargos']);
    let chargeIndicator = false;
    let LineTotalTaxes = Number(LineaFactura['impuestos']);
    let LineTotal = parseFloat(LineaFactura['totaldelinea']);
    let LineExtensionAmount = parseFloat(LineaFactura['subtotal']);
    let LineAllowanceTotal = Number(LineaFactura['descuento%']) * 100;
    let MeasureUnitCode = String(codigosUnidadDeMedida[unidadDeMedida]);

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
      if (LineAllowanceTotal > 0) {
        let Allowance = {
          "Id": 9,
          "ChargeIndicator": chargeIndicator,
          "AllowanceChargeReason": "",
          "MultiplierFactorNumeric": Number(LineAllowanceTotal),
          "Amount": Number(Price) * Number(Quantity) * (LineAllowanceTotal / 100),
          "BaseAmount": Number(Price) * Number(Quantity)
        }
        AllowanceCharge.push(Allowance);
      }
      if (LineChargeTotal > 0) {
        let Charge = {
          "Id": 20,
          "ChargeIndicator": !chargeIndicator,
          "AllowanceChargeReason": "",
          "Amount": LineChargeTotal,
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
        retencionTaxInformation.TaxAmount = Number(LineExtensionAmount) * Number(porcentajeRetencion)/100;
        ItemTaxesInformation.push(retencionTaxInformation);
      }

      return ItemTaxesInformation;
    }


    let productoI = {//aqui organizamos todos los parametros necesarios para los productos
      ItemReference: ItemReference,
      Name: Name,
      Quatity: new Number(Quantity),
      Price: new Number(Price),

      LineAllowanceTotal: 0,
      LineChargeTotal: 0,
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
  } while (i < (15 + cantidadProductos));


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

    let rowSeccionCargosyDescuentos = getcargosDescuentosStartRow(hojaFactura)+2;
    let lastCargoDescuentoRow = getLastCargoDescuentoRow(hojaFactura);
    let CargosyDescuentos = [];
    let chargeIndicator = false;
    for (let i = rowSeccionCargosyDescuentos; i <= lastCargoDescuentoRow+1; i++) {
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
  const posicionOriginalTotalFactura = hojaFactura.getRange("A26").getValue(); // para verificar donde esta el TOTAL
  let rangeTotales = ""


  if (posicionOriginalTotalFactura === "Subtotal") {
    rangeTotales = hojaFactura.getRange(posicionOriginalTotalFactura+1, 1, 1, 12);//va a cambiar

  } else {
    let rowTotales = getTotalesLinea(hojaFactura)
    rangeTotales = hojaFactura.getRange(rowTotales, 1, 1, 12);
  }

  let totalesValores = String(rangeTotales.getValues())
  totalesValores = totalesValores.split(",")


  //Definir los valores para el json
  let pfSubTotal = parseFloat(totalesValores[0]);
  let pfBaseGrabable = parseFloat(totalesValores[1]);
  let pfSubTotalMasImpuestos = parseFloat(totalesValores[3]);
  let pfRetenciones = parseFloat(totalesValores[4]);
  let pfDescuentos = parseFloat(totalesValores[5]);
  let pfCargos = parseFloat(totalesValores[7]);
  let pfAnticipo = parseFloat(totalesValores[9]);
  let pfNetoAPagar = Number(totalesValores[10]);
  if (pfAnticipo = null) {
    pfAnticipo = 0;
  }

  let invoice_total = {
    "lineExtensionamount": pfSubTotal,
    "TaxExclusiveAmount": pfBaseGrabable,
    "TaxInclusiveAmount": pfSubTotalMasImpuestos,
    "AllowanceTotalAmount": pfDescuentos,
    "ChargeTotalAmount": pfCargos,
    "PrePaidAmount": Number(pfAnticipo),
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
  listadoestado_sheet.appendRow(["vacio", fecha, "vacio", numeroFactura, nombreCliente, codigoCliente, "vacio", "vacio", "Vacio", invoice, nuevoInvoiceResumido]);

  SpreadsheetApp.getUi().alert("Factura generada y guardada satisfactoriamente, aguarde unos segundos");

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

