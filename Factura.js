PREFACTURA_ROW = 3;
PREFACTURA_COLUMN = 2;
COL_TOTALES_PREFACTURA = 11;// K
FILA_INICIAL_PREFACTURA = 8;
COLUMNA_FINAL = 50;
ADDITIONAL_ROWS = 3 + 3; //(Personalizacion)

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

  if (totalProductos === "Total productos") {
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
  let estadoFactura = [];
  verificarEstadoValidoFactura(estadoFactura);
  if (estadoFactura[0] === true) {
    guardarYGenerarInvoice();
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
    hojaFactura.getRange("L15").setValue("=(E15+F15+J15+(K15*E15))-((E15+F15+J15+(K15*E15))*I15)")//Total

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
    hojaFactura.getRange("L" + String(rowParaDatos)).setValue("=(E" + String(rowParaDatos) + "+F" + String(rowParaDatos) + "+J" + String(rowParaDatos) + "+(K" + String(rowParaDatos) + "*E" + String(rowParaDatos) + "))-(E" + String(rowParaDatos) + "*I" + String(rowParaDatos) + ")")//Total
  }
  updateTotalProductCounter(cargosDescuentosStartRow - 3, productStartRow, hojaFactura, cargosDescuentosStartRow);
  calcularDescuentosCargosYTotales(cargosDescuentosStartRow - 3, productStartRow, hojaFactura, cargosDescuentosStartRow);
}

function recuperarJson() {
  let hojaFactura = spreadsheet.getSheetByName('ListadoEstado');
  let lastRow = hojaFactura.getLastRow();
  let infoFactura = hojaFactura.getRange(lastRow, 1, 1, 20).getValues();
  let json = infoFactura[0][9];
  Logger.log(json);
  return json;
}
function enviarFactura() {
  let schemaID = 31;
  let idNumber = 900091496;
  let templateID = 73;
  let url = `https://misfacturas.cenet.ws/integrationAPI_2/api/insertinvoice?SchemaID=${schemaID}&IDNumber=${idNumber}&TemplateID=${templateID}`;
  let json = recuperarJson();
  let hojaDatos = spreadsheet.getSheetByName('Datos');
  let APIkey = hojaDatos.getRange("F47").getValue();
  let opciones = {
    "method": "post",
    "contentType": "application/json",
    "payload": json,
    "headers": {"Authorization": "misfacturas " + APIkey},
    'muteHttpExceptions': true
  };
  Logger.log("Enviar factura antes del try");
  try {
    var respuesta = UrlFetchApp.fetch(url, opciones);
    var contenidoRespuesta = respuesta.getContentText();
    Logger.log(contenidoRespuesta); // Muestra la respuesta de la API en los logs
    SpreadsheetApp.getUi().alert("Factura enviada correctamente a misfacturas. Si desea verla ingrese a https://misfacturas-qa.cenet.ws/Aplicacion/");
  } catch (error) {
    Logger.log("Error al enviar el JSON a la API: " + error.message);
    SpreadsheetApp.getUi().alert("Error al enviar la factura a misfacturas. Intente de nuevo si el error presiste comuniquese con soporte");
  }
  limpiarHojaFactura();
}

function obtenerAPIkey(usuario, contra) {
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

    // Verificar si la respuesta contiene un API Key en el formato esperado
    if (respuestaJson.length > 0 && typeof respuestaJson === 'string') {
      let apiKey = respuestaJson; // Extrae el API Key
      Logger.log("API Key obtenida: " + apiKey);
      SpreadsheetApp.getUi().alert("Se ha vinculado tu cuenta exitosamente");
      hojaDatosEmisor.getRange("B13").setBackground('#ccffc7')  // Almacena el API Key en la celda
      hojaDatosEmisor.getRange("B13").setValue("Vinculado")
      hojaDatos.getRange("F47").setValue(apiKey)
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
  let datosSheet = spreadsheet.getSheetByName('Datos');
  let hojaProductos = spreadsheet.getSheetByName('Productos');
  let consecutivoFactura = datosSheet.getRange("Q11").getValue();
  let consecutivoFacturaActualizado = consecutivoFactura + 1;
  datosSheet.getRange("Q11").setValue(consecutivoFacturaActualizado);

  //obtener el total de prodcutos
  let posicionTotalProductos = prefactura_sheet.getRange("A16").getValue(); // para verificar donde esta el TOTAL
  if (posicionTotalProductos === "Total productos") {
    var cantidadProductos = prefactura_sheet.getRange("B16").getValue();// cantidad total de productos 
  } else {
    let startingRowTax = getcargosDescuentosStartRow(prefactura_sheet)
    let posicionTotalProductos = startingRowTax - 2
    var cantidadProductos = prefactura_sheet.getRange("B" + String(posicionTotalProductos)).getValue();// cantidad total de productos
  }

  let llavesParaLinea = prefactura_sheet.getRange("A14:L14");//llamo los headers 
  llavesParaLinea = slugifyF(llavesParaLinea.getValues()).replace(/\s/g, ''); // Todo en una sola linea
  const llavesFinales = llavesParaLinea.split(",");
  /* Creo que esto se puede cambiar a una manera mas simple, ya que los headers de la fila H7 hatsa N7 nunca van a cambiar */

  var productoInformation = [];

  let i = 15 // es 15 debido a que aqui empieza los productos elegidos por el cliente
  do {
    let filaActual = "A" + String(i) + ":L" + String(i);
    let rangoProductoActual = prefactura_sheet.getRange(filaActual);
    let productoFilaActual = String(rangoProductoActual.getValues());
    productoFilaActual = productoFilaActual.split(",");// cojo el producto de la linea actual y se le hace split a toda la info
    let LineaFactura = {};

    for (let j = 0; j < 12; j++) {// original dice que son 11=COL_TOTALES_PREFACTURA deberian ser 10 creo
      LineaFactura[llavesFinales[j]] = productoFilaActual[j]
    }
    let filaProducto = obtenerFilaPorReferencia(Number(LineaFactura['referencia']));
    let unidadDeMedida = hojaProductos.getRange("G" + String(filaProducto)).getValue();

    let ItemReference = String(LineaFactura['referencia']);
    let Name = String(LineaFactura['producto']);
    let Quantity = String(LineaFactura['cantidad']);
    let Price = Number(LineaFactura['preciounitario']);
    let LineAllowanceTotal = parseFloat(LineaFactura['descuento%']);
    let LineChargeTotal = Number(LineaFactura['cargos']);
    let LineTotalTaxes = Number(LineaFactura['impuestos']);
    let LineTotal = parseFloat(LineaFactura['totaldelinea']);
    let LineExtensionAmount = parseFloat(LineaFactura['subtotal']);
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

    function agregarImpuestos() {
      //taxes del producto en si
      if (LineaFactura["iva%"] > 0) {
        let percentIva = convertToPercentage(LineaFactura["iva%"]);
        let ivaTaxInformation = {
          Id: "01",//Id
          TaxEvidenceIndicator: false,
          TaxableAmount: LineExtensionAmount,
          TaxAmount: LineExtensionAmount * LineaFactura["iva%"],
          Percent: percentIva,
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
          Percent: percentInc,
          BaseUnitMeasure: 0,
          PerUnitAmount: 0,
        };
        ItemTaxesInformation.push(incTaxInformation);
      }
      return ItemTaxesInformation;
    }

    let productoI = {//aqui organizamos todos los parametros necesarios para 
      ItemReference: ItemReference,
      Name: Name,
      Quatity: new Number(Quantity),
      Price: new Number(Price),

      LineAllowanceTotal: LineAllowanceTotal,
      LineChargeTotal: LineChargeTotal,
      LineTotalTaxes: LineTotalTaxes,
      LineTotal: LineTotal,
      LineExtensionAmount: LineExtensionAmount,
      MeasureUnitCode: MeasureUnitCode,
      FreeOFChargeIndicator: false,
      AdditionalReference: [],
      Nota: "",
      AdditionalProperty: [],
      TaxesInformation: agregarImpuestos(),
      AllowanceCharge: []
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




  //estos es dinamico, verificar donde va el total cargo y descuento
  const posicionOriginalTotalFactura = prefactura_sheet.getRange("A29").getValue(); // para verificar donde esta el TOTAL
  let rangeTotales = ""


  if (posicionOriginalTotalFactura === "Subtotal") {
    rangeTotales = prefactura_sheet.getRange(29, 1, 1, 12);//va a cambiar

  } else {
    let rowTotales = getTotalesLinea(prefactura_sheet)
    rangeTotales = prefactura_sheet.getRange(rowTotales, 1, 1, 12);
  }

  let totalesValores = String(rangeTotales.getValues())
  totalesValores = totalesValores.split(",")


  //Definir los valores para el json
  let pfSubTotal = parseFloat(totalesValores[0]);
  let pfBaseGrabable = parseFloat(totalesValores[1]);
  let pfSubTotalMasImpuestos = parseFloat(totalesValores[2]);
  let pfRetenciones = parseFloat(totalesValores[4]);
  let pfCargos = parseFloat(totalesValores[7]);
  let pfTotal = parseFloat(totalesValores[8]);
  let pfAnticipo = parseFloat(totalesValores[9]);
  let pfNetoAPagar = parseFloat(totalesValores[10]);
  if (pfAnticipo = null) {
    pfAnticipo = 0;
  }

  let invoice_total = {
    "lineExtensionamount": pfBaseGrabable,
    "TaxExclusiveAmount": pfSubTotal,
    "TaxInclusiveAmount": pfSubTotalMasImpuestos,
    "AllowanceTotalAmount": pfRetenciones,
    "ChargeTotalAmount": pfCargos,
    "PrePaidAmount": Number(pfAnticipo),
    "PayableAmount": pfNetoAPagar,
  }


  let cliente = prefactura_sheet.getRange("B2").getValue();
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
    InvoiceAllowanceCharge: [],
    InvoiceTotal: invoice_total,
    Documents: []
  });

  let nombreCliente = prefactura_sheet.getRange("B2").getValue();
  let numeroFactura = InvoiceGeneralInformation.InvoiceNumber;
  let fecha = obtenerFecha();
  let codigoCliente = prefactura_sheet.getRange("B3").getValue();
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

