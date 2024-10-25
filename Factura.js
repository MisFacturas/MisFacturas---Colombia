PREFACTURA_ROW = 3;
PREFACTURA_COLUMN = 2;
COL_TOTALES_PREFACTURA = 11;// K
FILA_INICIAL_PREFACTURA = 8;
COLUMNA_FINAL = 50;
ADDITIONAL_ROWS = 3 + 3; //(Personalizacion)

var spreadsheet = SpreadsheetApp.getActive();
var prefactura_sheet = spreadsheet.getSheetByName('Factura');
var unidades_sheet = spreadsheet.getSheetByName('Unidades');
var listadoestado_sheet = spreadsheet.getSheetByName('ListadoEstado');
var hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
var folderId = hojaDatosEmisor.getRange("B13").getValue();


function verificarEstadoValidoFactura() {
  // en esta funcion se debe de verificar si el numero de factura ya fue utiliazado en alguna otra factura
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  
  // función que verifica si una factura cumple con los requisitos mínimos para guardar
  let estaValido = true;

  let clienteActual = hojaFactura.getRange("B2").getValue();
  let informacionFactura1 = hojaFactura.getRange(2, 6, 5, 3).getValues();
  let informacionFactura2 = hojaFactura.getRange(2, 9, 5, 2).getValues();


  // Crear una lista combinada
  let listaCombinada = [clienteActual];  // Añadir clienteActual al array
  for (let i = 0; i < informacionFactura1.length; i++) {
    listaCombinada.push(informacionFactura1[i][2]); // Añadir cada valor de informacionFactura1
  }
  for (let j = 0; j < informacionFactura2.length; j++) {
    listaCombinada.push(informacionFactura2[j][1]); // Añadir cada valor de informacionFactura2
  }

  // Recorrer 
  for (let k = 0; k < listaCombinada.length; k++) {
    if(listaCombinada[k]===""){
      estaValido=false
    }
  }

  let totalProductos=hojaFactura.getRange("A16").getValue();

  if (totalProductos==="Total productos"){
    // no hay necesidad de encontrar TOTAL PRODUCTOS si no esta, porque eso implica que si anadio asi sea 1 prodcuto
    let valorTotalProductos=hojaFactura.getRange("B16").getValue();
    if(valorTotalProductos===0 ||valorTotalProductos===""){
      // no agrego producto
      estaValido=false
    }
  }


  return estaValido;  
}

function guardarFactura(){
  let estadoFactura=verificarEstadoValidoFactura();
  if(estadoFactura){
    //factura valida
    // generar json
    guardarYGenerarInvoice()
    guardarFacturaHistorial()
    limpiarHojaFactura()
    
  }else{
    SpreadsheetApp.getUi().alert("La factura no es valida")
  }
  

}
function agregarFilaNueva(){
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  let cargosDescuentosStartRow = getcargosDescuentosStartRow(hojaFactura);
  const productStartRow = 15;
  const lastProductRow = getLastProductRow(hojaFactura, productStartRow, cargosDescuentosStartRow);
  hojaFactura.insertRowAfter(lastProductRow)
}
function agregarFilaCargoDescuento(){
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  const lastCargoDescuentoRow = getLastCargoDescuentoRow(hojaFactura);
  hojaFactura.insertRowAfter(lastCargoDescuentoRow)
}

function agregarProductoDesdeFactura(cantidad,producto){
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  let cargosDescuentosStartRow = getcargosDescuentosStartRow(hojaFactura);
  const productStartRow = 15;
  const lastProductRow = getLastProductRow(hojaFactura, productStartRow, cargosDescuentosStartRow);

  let dictInformacionProducto ={}
  if(producto==="" || cantidad==="" || cantidad===0){
    throw new Error('Porfavor elige un producto y un cantidad adecuado');
  }else{
    dictInformacionProducto = obtenerInformacionProducto(producto);
  }

  let rowParaDatos=lastProductRow
  let rowParaTotalTaxes=cargosDescuentosStartRow
  let cantidadProductos=hojaFactura.getRange("B16").getValue()//estado defaul de total productos
  if(cantidadProductos===0 || cantidadProductos===""){
    hojaFactura.getRange("A15").setValue(dictInformacionProducto["codigo Producto"])
    hojaFactura.getRange("B15").setValue(producto)
    hojaFactura.getRange("C15").setValue(cantidad)
    hojaFactura.getRange("D15").setValue(dictInformacionProducto["precio Unitario"])
    hojaFactura.getRange("E15").setValue("=D15*C15")
    hojaFactura.getRange("F15").setValue(dictInformacionProducto["precio Impuesto"])
    hojaFactura.getRange("G15").setValue(dictInformacionProducto["tarifa INC"])
    hojaFactura.getRange("H15").setValue(dictInformacionProducto["tarifa IVA"])
    hojaFactura.getRange("K15").setValue("=D15*("+dictInformacionProducto["tarifa Retencion"]+"*"+cantidad+")")

  }else{
    hojaFactura.insertRowAfter(lastProductRow)
    rowParaTotalTaxes=cargosDescuentosStartRow+1
    rowParaDatos=lastProductRow+1
    hojaFactura.getRange("A"+String(rowParaDatos)).setValue(dictInformacionProducto["codigo Producto"])
    hojaFactura.getRange("B"+String(rowParaDatos)).setValue(producto)
    hojaFactura.getRange("C"+String(rowParaDatos)).setValue(cantidad)
    hojaFactura.getRange("D"+String(rowParaDatos)).setValue(dictInformacionProducto["precio Unitario"])//precio unitario
    hojaFactura.getRange("E"+String(rowParaDatos)).setValue("=D"+String(rowParaDatos)+"*C"+String(rowParaDatos))//Subtotal
    hojaFactura.getRange("F"+String(rowParaDatos)).setValue(dictInformacionProducto["precio Impuesto"])//precio de los Impuestos
    hojaFactura.getRange("G"+String(rowParaDatos)).setValue(dictInformacionProducto["tarifa INC"])//%INC
    hojaFactura.getRange("H"+String(rowParaDatos)).setValue(dictInformacionProducto["tarifa IVA"])//%IVA
    hojaFactura.getRange("K"+String(rowParaDatos)).setValue("=D"+String(rowParaDatos)+"*("+dictInformacionProducto["tarifa Retencion"]+"*"+cantidad+")")//Valor de los Impuestos
  } 
  updateTotalProductCounter(rowParaDatos, productStartRow,hojaFactura, rowParaTotalTaxes);
  calcularDescuentosCargosYTotales(rowParaDatos,productStartRow,rowParaTotalTaxes,hojaFactura)
}

function guardarFacturaHistorial() {
  var hojaFactura = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Factura');
  var hojaListado = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Historial Facturas');
  var numeroFactura = hojaFactura.getRange("H2").getValue();
  var cliente = hojaFactura.getRange("B2").getValue();
  var fechaEmision = hojaFactura.getRange("H4").getValue();
  var estado = "Creada";
  var informacionCliente = getCustomerInformation(cliente);
  var numeroIdentificacion = informacionCliente.Identification;

  var lastRow = hojaListado.getLastRow();
  var newRow = lastRow + 1;
  var celdaNumFactura = hojaListado.getRange("A" + newRow);
  celdaNumFactura.setValue(numeroFactura);
  celdaNumFactura.setHorizontalAlignment('center');
  celdaNumFactura.setBorder(true, true, true, true, null, null, null, null);

  var celdaCliente = hojaListado.getRange("B" + newRow);
  celdaCliente.setValue(cliente);
  celdaCliente.setHorizontalAlignment('center');
  celdaCliente.setBorder(true, true, true, true, null, null, null, null);

  var celdaNumeroIdentificacion = hojaListado.getRange("C" + newRow);
  celdaNumeroIdentificacion.setValue(numeroIdentificacion);
  celdaNumeroIdentificacion.setHorizontalAlignment('center');
  celdaNumeroIdentificacion.setBorder(true, true, true, true, null, null, null, null);

  var celdaFecha = hojaListado.getRange("D" + newRow);
  celdaFecha.setValue(fechaEmision);
  celdaFecha.setHorizontalAlignment('center');
  celdaFecha.setBorder(true, true, true, true, null, null, null, null);

  var celdaEstado = hojaListado.getRange("E" + newRow);
  celdaEstado.setValue(estado);
  celdaEstado.setHorizontalAlignment('center');
  celdaEstado.setBorder(true, true, true, true, null, null, null, null);

  var idArchivo = obtenerDatosFactura(numeroFactura);
  Logger.log("idarchivo"+idArchivo)
  guardarIdArchivo(idArchivo, numeroFactura);

  showCustomDialog()
}

function guardarIdArchivo(idArchivo, numeroFactura) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Facturas ID');
  var lastRow = hoja.getLastRow();
  var newRow = lastRow + 1;
  hoja.getRange("A" + newRow).setValue(numeroFactura).setBorder(true, true, true, true, null, null, null, null);
  hoja.getRange("B" + newRow).setValue(idArchivo).setBorder(true, true, true, true, null, null, null, null);

}
function convertPdfToBase64() {
  let hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Facturas ID');
  let hojaListadoEstado=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ListadoEstado');
  let dataRange=hojaListadoEstado.getDataRange()
  let data=dataRange.getValues()

  let jsonNuevoCol=13;
  let lastRow = hojaListadoEstado.getLastRow();
  let jsonData=data[lastRow-1][jsonNuevoCol]
  Logger.log("jsonData: " + jsonData)
  let invoiceData=JSON.parse(jsonData)

  let lastRowFacturasId=hoja.getLastRow()
  var idArchivo = hoja.getRange("B" + lastRowFacturasId).getValue();
  const file = DriveApp.getFileById(idArchivo);
  const base64String = Utilities.base64Encode(file.getBlob().getBytes());
  invoiceData.Document.fileName=String(file.getName())
  Logger.log(JSON.stringify(invoiceData))
  invoiceData.file=  base64String;
  
  Logger.log("Nuevo valor de invoiceData.file: " + invoiceData.fileName);
  let nuevoJsonData = JSON.stringify(invoiceData);

  return nuevoJsonData

}
function enviarFactura(){
  Logger.log("enviarFactura")
  let url ="" //url de la api
  let json = convertPdfToBase64()
  let opciones={
    "method" : "post",
    "contentType": "application/json",
    "payload" : json,
    'muteHttpExceptions': true
  };

  try {
    var respuesta = UrlFetchApp.fetch(url, opciones);
    Logger.log("aqui quiero borrar ese log re grande"    ); // Muestra la respuesta de la API en los logs
    Logger.log(respuesta.getContentText()); // Muestra la respuesta de la API en los logs
    SpreadsheetApp.getUi().alert("Factura enviada correctamente a MisFacturas. Si desea verla ingrese a https://misfacturas-qa.cenet.ws/Aplicacion/");
  } catch (error) {
    Logger.log("Error al enviar el JSON a la API: " + error.message);
    SpreadsheetApp.getUi().alert("Error al enviar la factura a MisFacturas. Intente de nuevo si el error presiste comuniquese con soporte");
  }
}

function ProcesarFormularioFactura(data) {
  var numFactura = data.numFactura
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Facturas ID');

  var range = hoja.getRange('A2:A'); // Rango desde A2 hasta el final de la columna A
  var textFinder = range.createTextFinder(numFactura);
  var cell = textFinder.findNext();

  if (cell) {
    var fila = cell.getRow();
    var idAsociado = hoja.getRange('B' + fila).getValue();
  } else {
    return 'Factura no encontrada';
  }

  var pdf = DriveApp.getFileById(idAsociado);
  var link = pdf.getDownloadUrl();
  return link;
}

function limpiarHojaFactura(){
  let hojaFactura = spreadsheet.getSheetByName('Factura');

  //Limpiar informacion general de la factura
  hojaFactura.getRange("B2").setValue("")//Cliente
  hojaFactura.getRange("B3").setValue("")//Codigo

  hojaFactura.getRange("H4").setValue("")//fecha
  hojaFactura.getRange("H5").setValue(0)//dias vencimiento
  hojaFactura.getRange("H6").setValue("")//hora

  hojaFactura.getRange("J2").setValue("")//forma pago
  hojaFactura.getRange("J3").setValue("")//tipo de pago
  hojaFactura.getRange("J4").setValue("")//moneda
  hojaFactura.getRange("J5").setValue("")//tasa de cambio
  hojaFactura.getRange("J6").setValue("")//fecha tasa de cambio

  hojaFactura.getRange("B10").setValue("")//Osbervaciones
  hojaFactura.getRange("B11").setValue("")//Nota de pago 
  hojaFactura.getRange("L15").setValue("")//Total primer producto
  
  //Limpiar informacion productos
  let productStartRow = 15;
  let cargosDescuentosStartRow = getcargosDescuentosStartRow(hojaFactura);
  let lastProductRow = getLastProductRow(hojaFactura, productStartRow, cargosDescuentosStartRow);
  for (let j = lastProductRow; j >= Number(productStartRow)+1; j--) {
    hojaFactura.deleteRow(j);
  }
  hojaFactura.getRange("A15").setValue("")//referncia
  hojaFactura.getRange("B15").setValue("")//producto
  hojaFactura.getRange("C15").setValue("")//cantidad
  hojaFactura.getRange("D15").setValue("")//precio unitario
  hojaFactura.getRange("E15").setValue("")//Subtotal
  hojaFactura.getRange("F15").setValue("")//impuestos
  hojaFactura.getRange("G15").setValue("")//%inc
  hojaFactura.getRange("H15").setValue("")//%iva
  hojaFactura.getRange("I15").setValue("")//descuento producto
  hojaFactura.getRange("J15").setValue("")//cargos
  hojaFactura.getRange("K15").setValue("")//retencion
  
  //Limpiar informacion de los cargos y descuentos
  let cargosStartRow = 20;
  let lastCargoDescuentoRow = getLastCargoDescuentoRow(hojaFactura);
  for (let i = lastCargoDescuentoRow; i >= Number(cargosStartRow)+1; i--) {
    hojaFactura.deleteRow(i);
  }
  hojaFactura.getRange("A20").setValue("")//cargo/descuento
  hojaFactura.getRange("B20").setValue("")//concepto
  hojaFactura.getRange("C20").setValue("")//valor o porcentaje
  hojaFactura.getRange("D20").setValue("")//base
  hojaFactura.getRange("E20").setValue("")//total

  hojaFactura.getRange("B16").setValue("0")//total producto
  hojaFactura.getRange("J29").setValue("0")//anticipos
  hojaFactura.getRange("F29").setValue("0")//total descuentos
  hojaFactura.getRange("H29").setValue("0")//total cargos

}


function inicarFacturaNueva(){
  generarNumeroFactura(); 
  ponerFechaYHoraActual();
}

function limpiarYEliminarFila(numeroFila,hoja,hojaTax){
  //funcion para el boton que se va a agregar al final del producto
  if (numeroFila>20 && numeroFila<hojaTax){
    hoja.deleteRow(numeroFila)
  }else{
    hoja.getRange("A"+String(numeroFila)).setValue("");//referencia
    hoja.getRange("B"+String(numeroFila)).setValue("");//producto
    hoja.getRange("C"+String(numeroFila)).setValue("");//cantidad
    hoja.getRange("D"+String(numeroFila)).setValue(0);//precio unitario

  }
}

function verificarYCopiarCliente(e) {
  let hojaFacturas = e.source.getSheetByName('Factura');
  let hojaClientes = e.source.getSheetByName('Clientes');
  let celdaEditada = e.range;



  let nombreCliente = celdaEditada.getValue();
  let ultimaColumnaPermitida = 20; // Columna del estado en la hoja de clientes
  let datosARetornar = ["B", "O","M","L","N","Q"]; // Columnas que quiero de la hoja de clientes


  if (nombreCliente==="Cliente"){
    Logger.log("Estado default")
  }else{
    let listaConInformacion = obtenerInformacionCliente(nombreCliente);
    if (listaConInformacion["Estado"]==="No Valido"){
      SpreadsheetApp.getUi().alert("Error: El cliente seleccionado no es válido.");
    }else{
      //asigna el valor del coldigo solamente porque ese fue lo que me pidieron no mas
      hojaFacturas.getRange("B3").setValue(listaConInformacion["Código cliente"]);
    }
  }


}


function generarNumeroFactura(){
  let sheet = spreadsheet.getSheetByName('Factura');

  let numeroActual= sheet.getRange("G2").getValue();
  numeroActual=Number(numeroActual);
  numeroActual++
  sheet.getRange("G2").setValue(numeroActual);
}

function ponerFechaYHoraActual(){ 
  let sheet = spreadsheet.getSheetByName('Factura');

  let fecha = new Date();
  let fechaFormateada = Utilities.formatDate(fecha, "America/Bogota", "yyyy-MM-dd");
  let horaFormateada = Utilities.formatDate(fecha, "America/Bogota", "HH:mm:ss");

  sheet.getRange("H4").setNumberFormat("@");
  sheet.getRange("H4").setValue(fechaFormateada);

  sheet.getRange("H6").setNumberFormat("@");
  sheet.getRange("H6").setValue(horaFormateada);

  sheet.getRange("J6").setNumberFormat("@");
  sheet.getRange("J6").setValue(fechaFormateada);
}
function obtenerFecha(){
  let fechaFormateada
  let sheet = spreadsheet.getSheetByName('Factura');
  let valorFecha=sheet.getRange("H4").getValue();
  fechaFormateada = Utilities.formatDate(new Date(valorFecha), "America/Bogota", "yyyy-MM-dd");
  return fechaFormateada
}

function obtenerDatosProductos(sheet,range,e){
    if ( range.getA1Notation() === "A14" || range.getA1Notation()=== "A15" || range.getA1Notation() === "A16" || range.getA1Notation()=== "A17" || range.getA1Notation()=== "A18") {
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

function getPaymentSummary(pfTotal, pfNetoAPagar) {
  var PaymentTypeTxt = prefactura_sheet.getRange("J2").getValue();
  var PaymentMeansTxt = prefactura_sheet.getRange("J3").getValue();
  var PaymentSummary = {
    "PaymentType": PaymentTypeTxt,
    "PaymentMeans": getPaymentMeans(PaymentMeansTxt),
    "PaymentNote": "NA"
  }
  return PaymentSummary;
}

function guardarYGenerarInvoice(){

  //obtener el total de prodcutos
  let posicionTotalProductos = prefactura_sheet.getRange("A16").getValue(); // para verificar donde esta el TOTAL
  if (posicionTotalProductos==="Total productos"){
    var cantidadProductos=prefactura_sheet.getRange("B16").getValue();// cantidad total de productos 
  }else{
    let startingRowTax=getcargosDescuentosStartRow(prefactura_sheet)
    let posicionTotalProductos=startingRowTax-2
    var cantidadProductos=prefactura_sheet.getRange("B"+String(posicionTotalProductos)).getValue();// cantidad total de productos
  }

  let llavesParaLinea=prefactura_sheet.getRange("A14:L14");//llamo los headers 
  llavesParaLinea = slugifyF(llavesParaLinea.getValues()).replace(/\s/g, ''); // Todo en una sola linea
  const llavesFinales =llavesParaLinea.split(",");
  /* Creo que esto se puede cambiar a una manera mas simple, ya que los headers de la fila H7 hatsa N7 nunca van a cambiar */

  let invoiceTaxTotal=[];
  var productoInformation = [];

  let i = 15 // es 15 debido a que aqui empieza los productos elegidos por el cliente
  do{
    let filaActual = "A" + String(i) + ":L" + String(i);
    let rangoProductoActual=prefactura_sheet.getRange(filaActual);
    let productoFilaActual= String(rangoProductoActual.getValues());
    productoFilaActual=productoFilaActual.split(",");// cojo el producto de la linea actual y se le hace split a toda la info
    let LineaFactura={};

    for (let j=0;j<12;j++){// original dice que son 11=COL_TOTALES_PREFACTURA deberian ser 10 creo
      LineaFactura[llavesFinales[j]]=productoFilaActual[j]
    }

    
    let ItemReference = String(LineaFactura['referencia']);
    let Name = String(LineaFactura['producto']);
    let Quantity = String(LineaFactura['cantidad']);
    let Price = Number(LineaFactura['preciounitario']);
    let LineAllowanceTotal = parseFloat(LineaFactura['descuento%']);
    let LineChargeTotal = Number(LineaFactura['cargos']);
    let LineTotalTaxes = Number(LineaFactura['impuestos']);
    let LineTotal = parseFloat(LineaFactura['totaldelinea']);
    let LineExtensionAmount = parseFloat(LineaFactura['subtotal']);
    let MeasureUnitCode = "Sin unidad"
    let ItemTaxesInformation = [];

    function agregarImpuestos(){
    //taxes del producto en si
      if (LineaFactura["iva%"]>0){
        let percentIva = convertToPercentage(LineaFactura["iva%"]);
        let ivaTaxInformation = {
          Id: "01",//Id
          TaxEvidenceIndicator: false,
          TaxableAmount: LineExtensionAmount,
          TaxAmount: LineExtensionAmount*LineaFactura["iva%"],
          Percent: percentIva,
          BaseUnitMeasure: 0,
          PerUnitAmount: 0,
        };
        ItemTaxesInformation.push(ivaTaxInformation);
      }
      if  (LineaFactura["inc%"]>0){
        let percentInc = convertToPercentage(LineaFactura["inc%"]);
        let incTaxInformation = {
          Id: "02",//Id
          TaxEvidenceIndicator: false,
          TaxableAmount: LineExtensionAmount,
          TaxAmount: LineExtensionAmount*LineaFactura["inc%"],
          Percent: percentInc,
          BaseUnitMeasure: 0,
          PerUnitAmount: 0,      
        };
        ItemTaxesInformation.push(incTaxInformation);
      }
      invoiceTaxTotal.push(ItemTaxesInformation);
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
  }while(i<(15+cantidadProductos));

  //estos es dinamico, verificar donde va el total cargo y descuento
  const posicionOriginalTotalFactura = prefactura_sheet.getRange("A29").getValue(); // para verificar donde esta el TOTAL
  let rangeTotales=""


  if (posicionOriginalTotalFactura==="Subtotal"){
    rangeTotales=prefactura_sheet.getRange(29,1,1,12);//va a cambiar
    
  }else{
    let rowTotales = getTotalesLinea(prefactura_sheet)
    rangeTotales=prefactura_sheet.getRange(rowTotales+1,1,1,12);
  }
  
  let totalesValores=String(rangeTotales.getValues())
  totalesValores=totalesValores.split(",")
 
  //Definir los valores para el json
  let pfSubTotal = parseFloat(totalesValores[0]);
  let pfBaseGrabable = parseFloat(totalesValores[1]);
  let pfSubTotalMasImpuestos = parseFloat(totalesValores[2]);
  let pfRetenciones = parseFloat(totalesValores[4]);
  let pfCargos = parseFloat(totalesValores[7]);
  let pfTotal = parseFloat(totalesValores[8]);
  let pfAnticipo = parseFloat(totalesValores[9]);
  let pfNetoAPagar = parseFloat(totalesValores[10]);
  if (pfAnticipo = null){
    pfAnticipo=0;
  }

  let invoice_total = {
    "lineExtensionamount": pfBaseGrabable,
    "TaxExclusiveAmount": pfSubTotal,
    "TaxInclusiveAmount": pfSubTotalMasImpuestos,
    "AllowanceTotalAmount": pfRetenciones,
    "ChargeTotalAmount": pfCargos,
    "PrePaidAmount": Number(pfAnticipo),
    "PayableAmount": pfNetoAPagar , 
  }


  let cliente = prefactura_sheet.getRange("B2").getValue();
  let InvoiceGeneralInformation = getInvoiceGeneralInformation();
  let CustomerInformation = getCustomerInformation(cliente);
  
  let sheetDatosEmisor=spreadsheet.getSheetByName('Datos de emisor');
  let userId = String(sheetDatosEmisor.getRange("B11").getValue());
  let companyId = String(sheetDatosEmisor.getRange("B3").getValue());
  let PaymentSummary=getPaymentSummary(pfTotal, pfNetoAPagar);

  let nuevoInvoiceResumido=JSON.stringify({
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
    AdditionalDocuments: getAdditionalDocuments(),
    PaymentSummary: PaymentSummary, //por ahora esto leugo se cambia la funcion getPaymentSummary para que cumpla los parametros
    ItemInformation: productoInformation,
    InvoiceTaxTotal: invoiceTaxTotal,
    InvoiceTaxOthersTotal: null,
    InvoiceAllowanceCharge: [],
    InvoiceTotal: invoice_total,
    Documents: []
  });

  let nombreCliente = prefactura_sheet.getRange("B2").getValue();
  let numeroFactura = InvoiceGeneralInformation.InvoiceNumber;
  let fecha = obtenerFecha();
  let codigoCliente = prefactura_sheet.getRange("B3").getValue();
  listadoestado_sheet.appendRow(["vacio", fecha,"vacio", numeroFactura, nombreCliente, codigoCliente,"vacio" ,"vacio","Vacio", invoice, nuevoInvoiceResumido]);
  
  SpreadsheetApp.getUi().alert("Factura generada y guardada satisfactoriamente, aguarde unos segundos");
  
}

function showCustomDialog() {
  var html = HtmlService.createHtmlOutputFromFile('postFactura')
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Elige una opcion');
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



//--------------------------------------------------------------------------------------------//
function obtenerDatosFactura(factura){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ListadoEstado');
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
  var invoiceColIndex = 5; // Columna F (indexada desde 0)
  var jsonColIndex = 12; // Columna M (indexada desde 0)

  for (var i = 1; i < data.length; i++) { // Comienza en 1 para saltar la fila de encabezado
    if (data[i][invoiceColIndex] == factura) {
      var jsonData = data[i][jsonColIndex];
      if (jsonData) {
        try {
          var invoiceData = JSON.parse(jsonData);
          var facturaNumero = invoiceData.InvoiceGeneralInformation.InvoiceNumber;
          var cliente = invoiceData.CustomerInformation.RegistrationName;
          var numeroIdentificacion = invoiceData.CustomerInformation.Identification;
          var codigo = invoiceData.CustomerInformation.AdditionalAccountID;
          var direccion = invoiceData.CustomerInformation.AddressLine;
          var telefono = invoiceData.CustomerInformation.Telephone;
          var municipio = invoiceData.CustomerInformation.CityName;
          var departamento = invoiceData.CustomerInformation.SubdivisionName;
          var pais = invoiceData.CustomerInformation.CountryName;
          var fechaEmision = invoiceData.CustomerInformation.DV;
          var formaPago = invoiceData.PaymentSummary.PaymentType;
          var listaProductos = invoiceData.ItemInformation;
          var numeroProductos = 0;
          var descuentosFactura = parseFloat(invoiceData.InvoiceTotal.PrePaidAmount);
          let descuentoGeneralesFactura=parseFloat(invoiceData.InvoiceTotal.GeneralPrePaidAmount);
          var cargosFactura = parseFloat(invoiceData.InvoiceTotal.ChargeTotalAmount);
          var totalFacturaJSON = parseFloat(invoiceData.InvoiceTotal.PayableAmount);
          var valorPagar = int2word(totalFacturaJSON) //arreglar
          var notaPago = invoiceData.PaymentSummary.PaymentNote;
          var observaciones = invoiceData.InvoiceGeneralInformation.Note;

          var filasInsertadas = 0;
          var filasInsertadasPorProductos = 0;
          var grupoIva = {};

          var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copia de Plantilla'); // Hoja donde quieres insertar el NumeroIdentificacion
          if (!targetSheet) {
            targetSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Copia de Plantilla');
          }

          var hojaCeldas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Celdas Plantilla');
          
          for (var j = 0; j < listaProductos.length; j++) {
            numeroProductos += 1;
            var numeroCelda = 19 + j;
            if (numeroProductos > 1) {
              targetSheet.insertRowAfter(numeroCelda);
              targetSheet.getRange('C'+(numeroCelda+1)+':E'+(numeroCelda+1)).merge();
              filasInsertadas += 1;
              filasInsertadasPorProductos += 1;
            }
            var celdaItem = targetSheet.getRange('A'+numeroCelda);
            celdaItem.setBorder(true,true,true,true,null,null,null,null);
            celdaItem.setValue(numeroProductos);
            celdaItem.setHorizontalAlignment('center');

            var celdaReferencia = targetSheet.getRange('B'+numeroCelda);
            celdaReferencia.setBorder(true,true,true,true,null,null,null,null);
            celdaReferencia.setValue(listaProductos[j].ItemReference);
            celdaReferencia.setHorizontalAlignment('center');

            var celdaDespricion = targetSheet.getRange('C'+numeroCelda);
            celdaDespricion.setBorder(true,true,true,true,null,null,null,null);
            celdaDespricion.setValue(listaProductos[j].Name);
            celdaDespricion.setHorizontalAlignment('center');
            
            var celdaCantidad = targetSheet.getRange('F'+numeroCelda);
            celdaCantidad.setBorder(true,true,true,true,null,null,null,null);
            celdaCantidad.setValue(listaProductos[j].Quatity);
            celdaCantidad.setHorizontalAlignment('center');
            
            var celdaPrecioUnitario = targetSheet.getRange('G'+numeroCelda);
            celdaPrecioUnitario.setBorder(true,true,true,true,null,null,null,null);
            celdaPrecioUnitario.setValue(listaProductos[j].Price);
            celdaPrecioUnitario.setHorizontalAlignment('normal');
            celdaPrecioUnitario.setNumberFormat('$#,##0')

            var celdaSubtotal = targetSheet.getRange('H'+numeroCelda);
            celdaSubtotal.setBorder(true,true,true,true,null,null,null,null);
            celdaSubtotal.setFormula('=F'+numeroCelda+'*(G'+numeroCelda+'-(G'+numeroCelda+'*J'+numeroCelda+'))');
            celdaSubtotal.setHorizontalAlignment('normal');
            celdaSubtotal.setNumberFormat('$#,##0')
            
            var celdaIva = targetSheet.getRange('I'+numeroCelda);
            celdaIva.setBorder(true,true,true,true,null,null,null,null);
            var percent = listaProductos[j].TaxesInformation[0].Percent;
            percent = percent.slice(0, -1);
            percent = parseFloat(percent);
            celdaIva.setValue(percent/100);
            celdaIva.setNumberFormat('0.0%');
            celdaIva.setHorizontalAlignment('center');

            var celdaDescuento = targetSheet.getRange('J'+numeroCelda);
            celdaDescuento.setBorder(true,true,true,true,null,null,null,null);
            celdaDescuento.setValue(parseFloat(listaProductos[j].TaxesInformation[0].Descuento));
            celdaDescuento.setNumberFormat('0.0%')
            celdaDescuento.setHorizontalAlignment('center');

            var celdaRetencion = targetSheet.getRange('K'+numeroCelda);
            celdaRetencion.setBorder(true,true,true,true,null,null,null,null);
            celdaRetencion.setValue(parseFloat(listaProductos[j].TaxesInformation[0].Retencion));
            celdaRetencion.setNumberFormat('0%')
            celdaRetencion.setHorizontalAlignment('center');

            var celdaRecargoEquivalencia = targetSheet.getRange('L'+numeroCelda);
            celdaRecargoEquivalencia.setBorder(true,true,true,true,null,null,null,null);
            celdaRecargoEquivalencia.setValue(parseFloat(listaProductos[j].TaxesInformation[0].RecgEquivalencia));
            celdaRecargoEquivalencia.setNumberFormat('0.00%')
            celdaRecargoEquivalencia.setHorizontalAlignment('center');

            
            var celdaTotalLinea = targetSheet.getRange('M'+numeroCelda);
            celdaTotalLinea.setBorder(true,true,true,true,null,null,null,null);
            //subtotal+(subtotal*iva)+(subtotal*recargo)-(subtotal*retencion)
            celdaTotalLinea.setFormula('=H'+numeroCelda+'+(H'+numeroCelda+'*I'+numeroCelda+')+(H'+numeroCelda+'*L'+numeroCelda+')-(H'+numeroCelda+'*K'+numeroCelda+')');
            celdaTotalLinea.setNumberFormat('$#,##0');
            celdaTotalLinea.setHorizontalAlignment('normal');
            

            var producto = listaProductos[j]
            //crea un diccionario que la llave sea el % de iva y el valor sea el total de la linea
            
            if (grupoIva.hasOwnProperty(percent)) {
              grupoIva[percent] += producto.TaxesInformation[0].TaxableAmount;
            } else {
              grupoIva[percent] = producto.TaxesInformation[0].TaxableAmount;
            }
          }
          var contador = 0;
          var auxiliarFilasInsertadas = filasInsertadas;
          for (var key in grupoIva) {
            if (grupoIva.hasOwnProperty(key)) {
              var numeroCelda = 30 + auxiliarFilasInsertadas;
              if (contador > 0) {
                targetSheet.insertRowAfter(numeroCelda);
                targetSheet.getRange('A'+(numeroCelda+1)+':D'+(numeroCelda+1)).merge();
                targetSheet.getRange('F'+(numeroCelda+1)+':H'+(numeroCelda+1)).merge();
                targetSheet.getRange('I'+(numeroCelda+1)+':M'+(numeroCelda+1)).merge();
                filasInsertadas += 1;
                auxiliarFilasInsertadas += 1;
              } else {
                auxiliarFilasInsertadas += 1;
              }
              var celdaBaseImponible = targetSheet.getRange('A'+numeroCelda);
              celdaBaseImponible.setBorder(true,true,true,true,null,null,null,null);
              celdaBaseImponible.setValue(grupoIva[key]);
              celdaBaseImponible.setNumberFormat('$#,##0');
              celdaBaseImponible.setHorizontalAlignment('normal');
              
              var celdaPorcentajeIva = targetSheet.getRange('E'+numeroCelda);
              celdaPorcentajeIva.setBorder(true,true,true,true,null,null,null,null);
              celdaPorcentajeIva.setValue(key/100);
              celdaPorcentajeIva.setNumberFormat('0.0%');
              celdaPorcentajeIva.setHorizontalAlignment('center');
              
              var celdaIVA = targetSheet.getRange('F'+numeroCelda);
              celdaIVA.setBorder(true,true,true,true,null,null,null,null);
              celdaIVA.setFormula('=A'+numeroCelda+'*E'+numeroCelda);
              celdaIVA.setNumberFormat('$#,##0');
              celdaIVA.setHorizontalAlignment('normal');
              
              var celdaTotal = targetSheet.getRange('I'+numeroCelda);
              celdaTotal.setBorder(true,true,true,true,null,null,null,null);
              celdaTotal.setFormula('=A'+numeroCelda+'+F'+numeroCelda);
              celdaTotal.setNumberFormat('$#,##0');
              celdaTotal.setHorizontalAlignment('normal');

              contador += 1;
              
            }
          }

          //Extaccion celdas de datos cliente
          var clienteCeldaHoja = hojaCeldas.getRange('E3').getValue();
          var numeroIdentificacionCeldaHoja = hojaCeldas.getRange('E4').getValue();
          var codigoCeldaHoja = hojaCeldas.getRange('E8').getValue();
          var direccionCeldaHoja = hojaCeldas.getRange('E5').getValue();
          var telefonoCeldaHoja = hojaCeldas.getRange('E7').getValue();
          var municipioCeldaHoja = hojaCeldas.getRange('E6').getValue();
          var fechaEmisionCeldaHoja = hojaCeldas.getRange('E9').getValue();
          var formaPagoCeldaHoja = hojaCeldas.getRange('E10').getValue();

          //factura
          var celdaNumFactura = targetSheet.getRange('A9');
          //Datos Cliente
          var clienteCell = targetSheet.getRange(clienteCeldaHoja);
          var numeroIdentificacionCell = targetSheet.getRange(numeroIdentificacionCeldaHoja);
          var codigoCell = targetSheet.getRange(codigoCeldaHoja);
          var direccionCell = targetSheet.getRange(direccionCeldaHoja);
          var telefonoCell = targetSheet.getRange(telefonoCeldaHoja);
          var municipioCell = targetSheet.getRange(municipioCeldaHoja);
          var fechaEmisionCell = targetSheet.getRange(fechaEmisionCeldaHoja);
          var formaPagoCell = targetSheet.getRange(formaPagoCeldaHoja);
          var valorPagarCell = targetSheet.getRange('B'+(41+filasInsertadas));
          var notaPagoCell = targetSheet.getRange('A'+(45+filasInsertadas));
          var observacionesCell = targetSheet.getRange('A'+(50+filasInsertadas));
          var totalItemsCell = targetSheet.getRange('B'+(21+filasInsertadasPorProductos));
          var descuentosCell = targetSheet.getRange('A'+(24+filasInsertadasPorProductos));
          var cargosCell = targetSheet.getRange('D'+(24+filasInsertadasPorProductos));
          var sumaBaseImponible = targetSheet.getRange('A'+(32+filasInsertadas));
          var sumaImpIva = targetSheet.getRange('F'+(32+filasInsertadas));
          var sumaTotal = targetSheet.getRange('I'+(32+filasInsertadas));

          var totalRetenciones = targetSheet.getRange('A'+(36+filasInsertadas));
          var totalCrgEquivalencia = targetSheet.getRange('D'+(36+filasInsertadas));
          var totalCargos = targetSheet.getRange('G'+(36+filasInsertadas));
          var totalDescuentos = targetSheet.getRange('K'+(36+filasInsertadas));

          var totalDeFactura = targetSheet.getRange('H'+(38+filasInsertadas));

          celdaNumFactura.setValue("FACTURA DE VENTA NO. "+facturaNumero);
          clienteCell.setValue(cliente);
          numeroIdentificacionCell.setValue(numeroIdentificacion);
          codigoCell.setValue(codigo);
          direccionCell.setValue(direccion);
          telefonoCell.setValue(telefono);
          // Ajustar la forma en que se ve el pais - IMPORTANTE
          if (municipio == "" || departamento == "" || pais == "") {
            var columnaMunicipio = municipioCell.getColumn();
            var filaMunicipio = municipioCell.getRow();
            targetSheet.getRange(filaMunicipio, columnaMunicipio-1).clearContent();
          } else {
            municipioCell.setValue(municipio+', '+departamento+', '+pais);
          }
          fechaEmisionCell.setValue(fechaEmision);
          formaPagoCell.setValue(formaPago);
          valorPagarCell.setValue(valorPagar);
          notaPagoCell.setValue(notaPago);
          observacionesCell.setValue(observaciones);
          totalItemsCell.setValue(numeroProductos);
          descuentosCell.setValue(descuentoGeneralesFactura);
          cargosCell.setValue(cargosFactura);
          sumaBaseImponible.setFormula('=SUM(A'+(30+numeroProductos-1)+':A'+(31+filasInsertadas-1)+')');
          sumaImpIva.setFormula('=SUM(F'+(30+numeroProductos-1)+':F'+(31+filasInsertadas-1)+')');
          sumaTotal.setFormula('=SUM(I'+(30+numeroProductos-1)+':I'+(31+filasInsertadas-1)+')');
          totalRetenciones.setFormula('=SUMPRODUCT(H19:H'+(19+numeroProductos-1)+';K19:K'+(19+numeroProductos-1)+')');
          totalCrgEquivalencia.setFormula('=SUMPRODUCT(H19:H'+(19+numeroProductos-1)+';L19:L'+(19+numeroProductos-1)+')');
          totalCargos.setValue(cargosFactura);
          totalDescuentos.setFormula(descuentosFactura);
  
          totalDeFactura.setFormula('=SUM(M19:M'+(19+numeroProductos-1)+')+G'+(36+filasInsertadas)+'-A'+(24+filasInsertadasPorProductos));
          

          
          
          var itemCellPrueba = targetSheet.getRange('A19')
          var hojaEnBlanco = clienteCell.isBlank() || formaPagoCell.isBlank() || itemCellPrueba.isBlank() || celdaBaseImponible.isBlank();
          while (hojaEnBlanco) {
            sleep(1000);
            hojaEnBlanco = clienteCell.isBlank() || formaPagoCell.isBlank() || itemCellPrueba.isBlank() || celdaBaseImponible.isBlank();
          }

          if (!hojaEnBlanco){
            Logger.log("entrar hoja en blanco")
            var pdfFactura = generatePdfFromPlantilla();
            var id = subirFactura(facturaNumero, pdfFactura);
            resetPlantilla();
            return id;
          }
          

        } catch (e) {
          Logger.log('Error parsing JSON for row ' + (i + 1) + ': ' + e.message);
        }
      }
      break//ojo esto debo de quitarlo
    }
  }
  Logger.log('Invoice number ' + factura + ' not found.');
}
