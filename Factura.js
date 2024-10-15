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
  let informacionFactura1 = hojaFactura.getRange(2, 6, 5, 2).getValues();
  let informacionFactura2 = hojaFactura.getRange(2, 9, 6, 1).getValues();


  // Crear una lista combinada
  let listaCombinada = [clienteActual];  // Añadir clienteActual al array
  for (let i = 0; i < informacionFactura1.length; i++) {
    listaCombinada.push(informacionFactura1[i][0]); // Añadir cada valor de informacionFactura1
    
  }
  for (let j = 0; j < informacionFactura2.length; j++) {
    listaCombinada.push(informacionFactura2[j][0]); // Añadir cada valor de informacionFactura2
  
  }

  // Recorrer 
  for (let k = 0; k < listaCombinada.length; k++) {
    Logger.log("listaCombinada"+listaCombinada[k])
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
    SpreadsheetApp.getUi().alert("Factura no es valida")
  }
  

}
function agregarFilaNueva(){
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  let taxSectionStartRow = getTaxSectionStartRow(hojaFactura);//recordar este devuelve el lugar en donde deberian estar base imponible, toca restar -1
  const productStartRow = 15;
  const lastProductRow = getLastProductRow(hojaFactura, productStartRow, taxSectionStartRow);
  hojaFactura.insertRowAfter(lastProductRow)
}
function agregarFilaCargoDescuento(){
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  let taxSectionStartRow = getTaxSectionStartRow(hojaFactura);//recordar este devuelve el lugar en donde deberian estar base imponible, toca restar -1
  const lastCargoDescuentoRow = getLastCargoDescuentoRow(hojaFactura, taxSectionStartRow);
  hojaFactura.insertRowAfter(lastCargoDescuentoRow)
}

function agregarProductoDesdeFactura(cantidad,producto){
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  let taxSectionStartRow = getTaxSectionStartRow(hojaFactura);//recordar este devuelve el lugar en donde deberian estar base imponible, toca restar -1
  const productStartRow = 15;
  const lastProductRow = getLastProductRow(hojaFactura, productStartRow, taxSectionStartRow);

  let dictInformacionProducto ={}
  if(producto==="" || cantidad==="" || cantidad===0){
    throw new Error('Porfavor elige un producto y un cantidad adecuado');
  }else{
    dictInformacionProducto = obtenerInformacionProducto(producto);
  }

  let rowParaDatos=lastProductRow
  let rowParaTotalTaxes=taxSectionStartRow
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
    rowParaTotalTaxes=taxSectionStartRow+1
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

function onImageClick() {
  // Obtén el rango activo (última celda seleccionada)
  var range = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();

  // Obtén la dirección de la celda
  var cellAddress = range.getA1Notation();

  // Muestra la celda en un diálogo
  SpreadsheetApp.getUi().alert('La celda activa es: ' + cellAddress);
}
function probarInsertarImagen(){
  insertarImagenBorrarFila(15)
}
function insertarImagenBorrarFila(fila){
  let hojaFcatura=spreadsheet.getSheetByName('Factura');
  let imagenURL="https://i.postimg.cc/RFZ45sgp/basura3.png"
  var cell = hojaFcatura.getRange('H'+fila);
  cell.setHorizontalAlignment('center');
  var imageBlob = UrlFetchApp.fetch(imagenURL).getBlob();
  var image = hojaFcatura.insertImage(imageBlob, cell.getColumn(), cell.getRow());
  var numFactura = hojaFcatura.getRange('A'+fila).getValue();
  image.assignScript("onImageClick");
  image.setHeight(20);
  image.setWidth(20);
  image.setAnchorCellXOffset(40);
}

function guardarFacturaHistorial() {
  var hojaFactura = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Factura');
  var hojaListado = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Historial Facturas');
  var numeroFactura = hojaFactura.getRange("G2").getValue();
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

  var celdaImagen = hojaListado.getRange("F" + newRow);
  insertarImagen(newRow);
  celdaImagen.setHorizontalAlignment('center');
  celdaImagen.setBorder(true, true, true, true, null, null, null, null);

  var idArchivo = obtenerDatosFactura(numeroFactura);
  Logger.log("idarchivo"+idArchivo)
  guardarIdArchivo(idArchivo, numeroFactura);

  // var html = HtmlService.createHtmlOutputFromFile('postFactura')
  //   .setTitle('Menú');
  // SpreadsheetApp.getUi()
  //   .showSidebar(html);
  
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
  let hojaListadoEstao=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ListadoEstado');
  let dataRange=hojaListadoEstao.getDataRange()
  let data=dataRange.getValues()

  let jsonNuevoCol=13;
  let lastRow = hojaListadoEstao.getLastRow();
  let jsonData=data[lastRow-1][jsonNuevoCol]
  Logger.log("json"+jsonData)
  let invoiceData=JSON.parse(jsonData)
  let infoACambiar=invoiceData.file;
  Logger.log("infoACambiar "+infoACambiar)


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
  let url ="https://facturasapp-qa.cenet.ws/ApiGateway/InvoiceSync/LoadInvoice/LoadDocument"
  let json =convertPdfToBase64()
  let opciones={
    "method" : "post",
    "contentType": "application/json",
    "payload" : json,
    'muteHttpExceptions': true
  };

  try {
    var respuesta = UrlFetchApp.fetch(url, opciones);
    Logger.log(respuesta.getContentText()); // Muestra la respuesta de la API en los logs
    SpreadsheetApp.getUi().alert("Factura enviada correctamente a MisFacturas. Si desea verla ingrese a https://facturasapp-qa.cenet.ws/Aplicacion/");
  } catch (error) {
    Logger.log("Error al enviar el JSON a la API: " + error.message);
    SpreadsheetApp.getUi().alert("Error al enviar la factura a MisFacturas. Intente de nuevo si el error presiste comuniquese con soporte");
  }
}
function convertPdfToBase64Prueba() {
  let hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Facturas ID');
  let hojaListadoEstao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ListadoEstado');
  let dataRange = hojaListadoEstao.getDataRange();
  let data = dataRange.getValues();

  let jsonNuevoCol = 13;
  let lastRow = hojaListadoEstao.getLastRow();
  let jsonData = data[lastRow - 1][jsonNuevoCol];
  Logger.log("json" + jsonData);

  let invoiceData = JSON.parse(jsonData);
  let infoACambiar = invoiceData.file;
  Logger.log("infoACambiar " + infoACambiar);

  let lastRowFacturasId = hoja.getLastRow();
  let idArchivo = hoja.getRange("B" + lastRowFacturasId).getValue();
  const file = DriveApp.getFileById(idArchivo);
  const base64String = Utilities.base64Encode(file.getBlob().getBytes());

  invoiceData.file = base64String;
  Logger.log("Nuevo valor de invoiceData.file: " + invoiceData.file);
  
  let nuevoJsonData = JSON.stringify(invoiceData);
  Logger.log("Nuevo JSON Data: " + nuevoJsonData);

  // Crear o actualizar el archivo 'prueba.json' en Google Drive
  let folder = DriveApp.getRootFolder(); // Aquí puedes especificar una carpeta en particular
  let files = folder.getFilesByName('prueba.json');
  let jsonFile = folder.createFile('prueba.json', nuevoJsonData, "application/json");

  if (files.hasNext()) {
    jsonFile = files.next();
    jsonFile.setContent(nuevoJsonData);
    Logger.log('Archivo "prueba.json" actualizado.');
  } else {
    jsonFile = folder.createFile('prueba.json', nuevoJsonData, MimeType.JSON);
    Logger.log('Archivo "prueba.json" creado.');
  }
}



function linkDescargaFactura() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Facturas ID');
  var lastRow = hoja.getLastRow();
  var idArchivo = hoja.getRange("B" + lastRow).getValue();
  var numFactura = hoja.getRange("A" + lastRow).getValue();
  var pdf = DriveApp.getFileById(idArchivo);
  var url = pdf.getDownloadUrl();
  return {
    numFactura: numFactura,
    url: url
  };
}

function getDownloadLink() {
  var data = linkDescargaFactura();
  return data;
}

function enviarEmailPostFactura(email) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Facturas ID');
  var lastRow = hoja.getLastRow();
  var idArchivo = hoja.getRange("B" + lastRow).getValue();
  var numFactura = hoja.getRange("A" + lastRow).getValue();
  var pdfFile = DriveApp.getFileById(idArchivo).getBlob();
  var subject = `Factura ${numFactura}`
  var body = 'Adjunto encontrará la factura en formato PDF.';

  if (!email) {
    return "Por favor ingrese una dirección de correo válida.";
  }

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: body,
    attachments: [pdfFile.getAs(MimeType.PDF)]
  });

  return "PDF generado y enviado por correo electrónico a " + email;
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

function insertarImagen(fila) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Historial Facturas');
  var imageUrl = 'https://cdn.icon-icons.com/icons2/1674/PNG/512/download_111133.png'; // Reemplaza con la URL de tu imagen
  var cell = sheet.getRange('F' + fila);
  cell.setHorizontalAlignment('center');
  var imageBlob = UrlFetchApp.fetch(imageUrl).getBlob();
  var image = sheet.insertImage(imageBlob, cell.getColumn(), cell.getRow());
  image.assignScript("descargarFactura");
  image.setHeight(20);
  image.setWidth(20);
  image.setAnchorCellXOffset(40);
}

function descargarFactura() {
  var html = HtmlService.createHtmlOutputFromFile('descargaFacturaHistorial')
    .setTitle('Menú');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function guardarFilaFactura() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Historial Facturas');
  var cell = sheet.getActiveCell();
  var fila = cell.getRow();
  sheet.getRange('Z1').setValue(fila); // Guardar la fila en una celda oculta (Z1)
  generarPDFfactura();
}


function generarPDFfactura() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Historial Facturas');
  var fila = sheet.getRange('Z1').getValue(); // Leer el número de fila de la celda oculta
  var numeroFactura = sheet.getRange('A' + fila).getValue(); // Obtener el número de factura

  var resultado = obtenerDatosFactura(numeroFactura);
  if (resultado) {
    var pdfBlob = generarPDF();
  } else {
    Utilities.sleep(5000);
    var pdfBlob = generarPDF();
  }
  var url = generarPdfUrl(pdfBlob);

  // Crear un archivo temporal en el Drive para proporcionar un enlace de descarga
  var tempFile = DriveApp.createFile(pdfBlob);
  var tempFileUrl = tempFile.getDownloadUrl();

  // Enviar un enlace de descarga al usuario
  var html = '<html><body><a href="' + tempFileUrl + '">Descargar PDF de la Factura ' + numeroFactura + '</a></body></html>';
  var ui = HtmlService.createHtmlOutput(html)
    .setWidth(300)
    .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Descargar PDF');
}


function generarPDF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Copia de Plantilla');

  if (!sheet) {
    throw new Error('La hoja Plantilla no existe.');
  }

  var sheetId = sheet.getSheetId();
  var url = ss.getUrl().replace(/edit$/, '') + 'export?exportFormat=pdf&format=pdf' +
    '&gid=' + sheetId +
    '&size=A4' +  // Tamaño del papel
    '&portrait=true' +  // Orientación vertical
    '&fitw=true' +  // Ajustar a ancho de la página
    '&sheetnames=false&printtitle=false' +  // Opciones de impresión
    '&pagenumbers=false&gridlines=false' +  // Más opciones de impresión
    '&fzr=false' +  // Aislar filas congeladas
    '&top_margin=0.8' +  // Margen superior
    '&bottom_margin=0.00' +  // Margen inferior
    '&left_margin=0.50' +  // Margen izquierdo
    '&right_margin=0.50' +  // Margen derecho
    '&horizontal_alignment=CENTER' +  // Alineación horizontal
    '&vertical_alignment=TOP';  // Alineación vertical

  var token = ScriptApp.getOAuthToken();

  try {
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      var pdfBlob = response.getBlob().setName('Factura.pdf');
      return pdfBlob;
    } else {
      Logger.log('Error ' + response.getResponseCode() + ': ' + response.getContentText());
      throw new Error('Error ' + response.getResponseCode() + ': ' + response.getContentText());
    }
  } catch (e) {
    Logger.log('Exception: ' + e.message);
    throw new Error('Exception: ' + e.message);
  }
}

function generarPdfUrl(pdfBlob) {
  var base64Data = Utilities.base64Encode(pdfBlob.getBytes());
  var contentType = pdfBlob.getContentType();
  var name = pdfBlob.getName();
  return `data:${contentType};base64,${base64Data}`;
}


function limpiarHojaFactura(){
  let hojaFactura = spreadsheet.getSheetByName('Factura');

  //total productos
  hojaFactura.getRange("B2").setValue("")//Cliente
  hojaFactura.getRange("B3").setValue("")//Codigo

  hojaFactura.getRange("H6").setValue("")//hora
  hojaFactura.getRange("H4").setValue("")//fecha
  hojaFactura.getRange("J2").setValue("")//forma pago
  hojaFactura.getRange("H5").setValue(0)//dias vencimiento
  hojaFactura.getRange("J3").setValue(0)//tipo de pago
  hojaFactura.getRange("J4").setValue("")//moneda
  hojaFactura.getRange("J5").setValue("")//tasa de cambio
  hojaFactura.getRange("J6").setValue("")//fecha tasa de cambio

  hojaFactura.getRange("B10").setValue("")//Osbervaciones
  hojaFactura.getRange("B11").setValue("")//Nota de pago 
  


  //productos
  let productStartRow = 15;
  let taxSectionStartRow = getTaxSectionStartRow(hojaFactura);
  let lastProductRow = getLastProductRow(hojaFactura, productStartRow, taxSectionStartRow);
  Logger.log("limpiarHojaFactura")
  Logger.log("lastProductRow "+lastProductRow)
  Logger.log("productStartRow+1 "+(Number(productStartRow)+1))
  for (let j = lastProductRow; j >= Number(productStartRow)+1; j--) {
    hojaFactura.deleteRow(j);
    Logger.log("J" + j);
  }
  Logger.log("Salta if")

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
  

  hojaFactura.getRange("B16").setValue("0")//total producto

}


function inicarFacturaNueva(){
  generarNumeroFactura(); 
  obtenerFechaYHoraActual();
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

function obtenerFechaYHoraActual(){ 
  let sheet = spreadsheet.getSheetByName('Factura');

  let fecha = Utilities.formatDate(new Date(), "America/Bogota", "dd/MM/yyyy");
  let hora= Utilities.formatDate(new Date(), "America/Bogota", "HH:mm:ss");

  sheet.getRange("H4").setNumberFormat("dd/MM/yyyy");
  sheet.getRange("H4").setValue(String(fecha))
  
  sheet.getRange("H6").setValue(hora)

  
  let valorFecha=sheet.getRange("H4").getValue();

  let fechaFormateada = Utilities.formatDate(new Date(valorFecha), "America/Bogota", "dd/MM/yyyy");
  Logger.log("valorFecha "+valorFecha)
  Logger.log("fecha "+fecha)
  Logger.log("fechaFormateada "+fechaFormateada)

}

function ObtenerFechaFormatedada(opcion){
    let sheet = spreadsheet.getSheetByName('Factura');
    let valorFecha=sheet.getRange("H4").getValue();
    let fechaFormateada = Utilities.formatDate(new Date(valorFecha), "America/Bogota", "dd/MM/yyyy");
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

function getprefacturaValueA1(column, row) {
  return getsheetValueA1(prefactura_sheet, column, row);
}

function getprefacturaValue(column, row) {
  return getsheetValue(prefactura_sheet, column, row);
}

function updateprefacturaValue(column, row, value) {
  updatesheetValue(prefactura_sheet, column, row, value);
  return;
}

function getInvoiceGeneralInformation() {
  //Browser.msgBox('getInvoiceGeneralInformation()');

  var InvoiceAuthorizationNumber = "nulo"//Resolución Autorización
  //
  range = prefactura_sheet.getRange("G6");//dias de vencimiento
  var DaysOff = range.getValue();

  var invoice_number = getprefacturaValue(2, 7);//cambiamos los valores para llamar el numero de factura
  var InvoiceGeneralInformation = {
    "InvoiceAuthorizationNumber": InvoiceAuthorizationNumber,
    "PreinvoiceNumber": invoice_number,
    "InvoiceNumber": invoice_number,
    "DaysOff": DaysOff,
    "Currency": "COP",
    "ExchangeRate": "",
    "ExchangeRateDate": "",
    "SalesPerson": "",
    //"InvoiceDueDate": null,
    "Note": getprefacturaValue(10, 2), //cambia los valores para llamar la nota de la factura
    "ExternalGR": false
    //"AdditionalProperty": AdditionalProperty
  }


  return InvoiceGeneralInformation;
}
function getPaymentSummary(startingRowTaxation) {
  let posTotalFactura=startingRowTaxation+7
  let posMontoNeto=startingRowTaxation+12
  var total_factura = prefactura_sheet.getRange("A"+String(posTotalFactura)).getValue();// por ahora esto no lo utilizamos ya que no hay descuentos
  var monto_neto = prefactura_sheet.getRange("B"+String(posMontoNeto)).getValue();
  //var numeros = new Intl.NumberFormat('es-CO', {maximumFractionDigits:0, style: 'currency', currency: 'COP' }).format(monto);
  var numeros_total = new Intl.NumberFormat().format(total_factura);
  var numeros_neto = new Intl.NumberFormat().format(monto_neto);
  //Browser.msgBox(`${numeros_total}  ${numeros_neto}`);

  Logger.log("total_factura"+total_factura)
  Logger.log("monto_neto"+monto_neto)
  var PaymentTypeTxt = prefactura_sheet.getRange("J2").getValue();
  var PaymentMeansTxt = prefactura_sheet.getRange("J3").getValue();
  var PaymentSummary = {
    "PaymentType": PaymentTypeTxt,
    "PaymentMeans": getPaymentMeans(PaymentMeansTxt),
    "PaymentNote": `Total Factura: $${numeros_total} \r Neto a Pagar  $${numeros_neto}: ${int2word(monto_neto)}`
  }
  return PaymentSummary;
}

function guardarYGenerarInvoice(){

  //obtener el total de prodcutos
  let posicionTotalProductos = prefactura_sheet.getRange("A16").getValue(); // para verificar donde esta el TOTAL
  if (posicionTotalProductos==="Total productos"){
    var cantidadProductos=prefactura_sheet.getRange("B16").getValue();// cantidad total de productos 
  }else{
    let startingRowTax=getTaxSectionStartRow(prefactura_sheet)
    let posicionTotalProductos=startingRowTax-2
    var cantidadProductos=prefactura_sheet.getRange("B"+String(posicionTotalProductos)).getValue();// cantidad total de productos
  }

  let llavesParaLinea=prefactura_sheet.getRange("A14:L14");//llamo los headers 
  llavesParaLinea = slugifyF(llavesParaLinea.getValues()).replace(/\s/g, ''); // Todo en una sola linea
  const llavesFinales =llavesParaLinea.split(",");
  /* Creo que esto se puede cambiar a una manera mas simple, ya que los headers de la fila H7 hatsa N7 nunca van a cambiar */

  let invoiceTaxTotal=[];
  var productoInformation = [];

  Logger.log("cantidadProductos"+cantidadProductos)

  let i = 15 // es 15 debido a que aqui empieza los productos elegidos por el cliente
  do{
    let filaActual = "A" + String(i) + ":I" + String(i);
    let rangoProductoActual=prefactura_sheet.getRange(filaActual);
    let productoFilaActual= String(rangoProductoActual.getValues());
    productoFilaActual=productoFilaActual.split(",");// cojo el producto de la linea actual y se le hace split a toda la info
    Logger.log(productoFilaActual)
    let LineaFactura={};

    for (let j=0;j<11;j++){// original dice que son 11=COL_TOTALES_PREFACTURA deberian ser 10 creo
      LineaFactura[llavesFinales[j]]=productoFilaActual[j]
    }
    Logger.log("LineaFactura "+LineaFactura)

    let Name = LineaFactura['producto'];
    let ItemCode = new Number(LineaFactura['referencia']);
    let MeasureUnitCode = "Sin unidad"
    let Quantity = LineaFactura['cantidad'];
    let Price = LineaFactura['precio unitario'];
    let Amount = parseFloat(LineaFactura['subtotal']);//importe
    let Impuestos = LineaFactura['impuetos']
    let LineChargeTotal = parseFloat(LineaFactura['totaldelinea']);
    let Iva = LineChargeTotal-Amount;
    let descuento=LineaFactura["descuento"];
    let retencion=LineaFactura["retencion"];
    Logger.log("descuento "+descuento)
    Logger.log("retencion "+retencion)

    
    if (descuento==""){
      Logger.log("hay un producto con descuento vacio")
      descuento=0
    }
    if(retencion==""){
      retencion=0
    }
    if(reCargoEqui==""){
      reCargoEqui=0
    }


    //IVA
    let ItemTaxesInformation = [];//taxes del producto en si
    let percent = convertToPercentage(LineaFactura["iva"]); //aqui deberia de calcular el porcentaje pero como todavia no tengo IVA solo por ahora no
    Logger.log("percent "+percent)
    let ivaTaxInformation = {
      Id: "01",//Id
      TaxEvidenceIndicator: false,
      TaxableAmount: Amount,
      TaxAmount: Impuestos,
      Percent: percent,
      BaseUnitMeasure: "",
      PerUnitAmount: "",
      Descuento:descuento,
      Retencion:retencion,
      RecgEquivalencia:reCargoEqui
    };

    ItemTaxesInformation.push(ivaTaxInformation);
    invoiceTaxTotal.push(ivaTaxInformation);

    let LineExtensionAmount = Amount;
    let LineTotalTaxes = Impuestos;

    let productoI = {//aqui organizamos todos los parametros necesarios para 
      ItemReference: ItemCode,
      Name: Name,
      Quatity: new Number(Quantity),
      Price: new Number(Price),
      LineAllowanceTotal: 0.0,
      LineChargeTotal: 0.0,// que pasa aca ?
      LineTotalTaxes: LineTotalTaxes,
      LineTotal: LineChargeTotal,
      LineExtensionAmount: LineExtensionAmount,
      MeasureUnitCode: MeasureUnitCode,
      FreeOFChargeIndicator: false,
      AdditionalReference: [],
      AdditionalProperty: [],
      TaxesInformation: ItemTaxesInformation,
      AllowanceCharge: []
    };
    productoInformation.push(productoI);//agregamos el producto actual a la lista total 
    i++;
  }while(i<(15+cantidadProductos));

  //estos es dinamico, verificar donde va el total cargo y descuento
  const posicionOriginalTotalFactura = prefactura_sheet.getRange("A31").getValue(); // para verificar donde esta el TOTAL
  let rangeFacturaTotal=""
  let rangeTotales=""
  let rangeImpuestosValor=""
  let cargoTotal=0
  let descuentoTotal=0
  let cargoFactura=0
  let descuentoFactura=0

  let startingRowTaxation=getTaxSectionStartRow(prefactura_sheet)
  if (posicionOriginalTotalFactura==="Total factura"){
    rangeImpuestosValor=prefactura_sheet.getRange(22,1,1,3);
    rangeTotales=prefactura_sheet.getRange(29,1,1,4);//va a cambiar
    rangeFacturaTotal=prefactura_sheet.getRange("K25")
    cargoFactura=prefactura_sheet.getRange("D17").getValue()
    descuentoFactura=prefactura_sheet.getRange("D18").getValue()
    
  }else{
    let rowImpuestosValor=startingRowTaxation+7//va a cambiar
    let rowTotales=startingRowTaxation+10
    let rowTotalFactura=startingRowTaxation+12
    let rowCargoFactura=startingRowTaxation-2
    let rowDescuentoFactura=startingRowTaxation-1
    rangeImpuestosValor=prefactura_sheet.getRange(rowImpuestosValor,1,1,3); //va a cambiar
    rangeTotales=prefactura_sheet.getRange(rowTotales,1,1,4);
    rangeFacturaTotal=prefactura_sheet.getRange(rowTotalFactura,2);//(maxRows-1) porque no necesito el total
    cargoFactura=prefactura_sheet.getRange("B"+String(rowCargoFactura)).getValue()
    descuentoFactura=prefactura_sheet.getRange("B"+String(rowDescuentoFactura)).getValue()
  }

  if(cargoFactura==""){
    cargoFactura=0
  }

  if(descuentoFactura==""){
    descuentoFactura=0
  }
  
  let totalesValores=String(rangeTotales.getValues())
  totalesValores=totalesValores.split(",")
  Logger.log("totalesValores"+totalesValores)
  cargoTotal=totalesValores[2]
  descuentoTotal=totalesValores[3]
  Logger.log("cargoTotal "+cargoTotal)
  Logger.log("descuentoTotal "+descuentoTotal)
  Logger.log("cargoFactura "+cargoFactura)
  Logger.log("descuentoFactura "+descuentoFactura)
  // aqui cambia con respecto al original, aqui deberia de cambiar el segundo parametro creo, seria con respecto a un j el cual seria la cantidad de ivas que hay
  let facturaTotalesBaseImponilbe=String(rangeImpuestosValor.getValues());
  facturaTotalesBaseImponilbe=facturaTotalesBaseImponilbe.split(",");
  Logger.log("facturaTotales "+facturaTotalesBaseImponilbe)
  let TotalFactura=rangeFacturaTotal.getValue()

  /*Aqui cambia por completo, por ahora solo voy a dejar los parametros en numeros x 
  ,  solo coinciden el base imponible he IVA */
  let pfSubTotal = parseFloat(facturaTotalesBaseImponilbe[0]);//base imponible
  let pfIVA = parseFloat(facturaTotalesBaseImponilbe[2]);//IVA
  let pfImpoconsumo = 0;
  let pfTotal = parseFloat(facturaTotalesBaseImponilbe[0]+facturaTotalesBaseImponilbe[2]);
  let pfRefuente = 0;
  let pfReteICA = 0;
  let pfReteIVA = 0;
  let pfTRetenciones = 0; 
  let pfAnticipo = descuentoTotal;
  let pfTPagar = 0;

  //Aqui seguiria el texto, pero en el de carlos nunca lo llama 
  let facturaTotales=String(rangeImpuestosValor.getValues());
  let invoice_total = {
    "LineExtensionAmount": pfSubTotal,
    "TaxExclusiveAmount": pfSubTotal,
    "TaxInclusiveAmount": pfTotal,
    "AllowanceTotalAmount": 0,
    "GeneralChargeTotalAmount": cargoFactura,
    "ChargeTotalAmount": cargoTotal,
    "GeneralPrePaidAmount": descuentoFactura,
    "PrePaidAmount": pfAnticipo,
    "PayableAmount": TotalFactura ,// antes era (pfTotal - pfAnticipo) 
    "totalRet":totalesValores[0],
    "totalCargoEqui":totalesValores[1]
  }


  let cliente = prefactura_sheet.getRange("B2").getValue();
  let InvoiceGeneralInformation = getInvoiceGeneralInformation();
  let CustomerInformation = getCustomerInformation(cliente);// tal ves que por ahora no llame al cliente
  
  let sheetDatosEmisor=spreadsheet.getSheetByName('Datos de emisor');
  let userId = String(sheetDatosEmisor.getRange("B11").getValue());
  let companyId = String(sheetDatosEmisor.getRange("B3").getValue());
  let PaymentSummary=getPaymentSummary(startingRowTaxation)

  let fechParaNuevoInvoice=ConvertirFecha("vacio")
  let fechaVencdioParaNuevoInvoice=ConvertirFecha("pago")

  let PercentSurchargeEquivalence;
  let PercentageRetention;

  if(totalesValores[0]==="" || totalesValores[0]===0||totalesValores[0]===null){
//futuro para calcuclar bien estos valores
  }else{

  }

  
  calcularPorcentaje()
  let nuevoInvoiceResumido=JSON.stringify({
    "file": "base64",
    "Document": {
      "fileName": "nombre documento",
      "userId": userId,
      "companyId": companyId,
      "invoice": {
        "invoiceType": false,
        "contactName": String(cliente),
        "numeroIdentificacion": String(CustomerInformation["Identification"]),
        "invoiceDate": String(fechParaNuevoInvoice),
        "numberInvoice": String(InvoiceGeneralInformation["InvoiceNumber"]),
        "taxableAmount": String(parseFloat(facturaTotalesBaseImponilbe[0])),
        "Percent": "0",
        "taxAmount": String(parseFloat(facturaTotalesBaseImponilbe[2])),
        "surchargeAmount": "el valor no se debe de reportar",
        "surchargeValue": "el valor no se debe de reportar",
        "PercentSurchargeEquivalence": "0",
        "PercentageRetention": "0",
        "IRPFValue": "el valor no se debe de reportar",
        "invoiceTotal": String(TotalFactura),
        "payDate":String(fechaVencdioParaNuevoInvoice),
        "PaymentType": String(PaymentSummary["PaymentType"]),
        "Observations": String(InvoiceGeneralInformation["note"])
      }
    }
  }
  );
  Logger.log(invoice_total)
  let invoice = JSON.stringify({
    CustomerInformation: CustomerInformation,
    InvoiceGeneralInformation: InvoiceGeneralInformation,
    Delivery: getDelivery(),
    AdditionalDocuments: getAdditionalDocuments(),
    AdditionalProperty: getAdditionalProperty(),
    PaymentSummary: PaymentSummary, //por ahora esto leugo se cambia la funcion getPaymentSummary para que cumpla los parametros
    ItemInformation: productoInformation,
    //Invoice_Note: invoice_note,
    InvoiceTaxTotal: invoiceTaxTotal,
    InvoiceAllowanceCharge: [],
    InvoiceTotal: invoice_total
  });
  Logger.log(invoice)
  Logger.log(nuevoInvoiceResumido)

  let nameString = prefactura_sheet.getRange("B2").getValue();
  let numeroFactura = JSON.stringify(InvoiceGeneralInformation.InvoiceNumber);
  let fecha =ObtenerFecha();
  let codigoCliente=prefactura_sheet.getRange("B3").getValue();
  listadoestado_sheet.appendRow(["vacio", "vacio","vacio" , fecha,"vacio" ,numeroFactura ,nameString ,codigoCliente,"vacio" ,"vacio" ,"representacion" ,"Vacio", String(invoice),String(nuevoInvoiceResumido)]);
  
  SpreadsheetApp.getUi().alert("Factura generada y guardada satisfactoriamente, aguarde unos segundos");
  
}

function calcularPorcentaje(valor, total) {
  return (valor / total) * 100;
}

function showCustomDialog() {
  var html = HtmlService.createHtmlOutputFromFile('postFactura')
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Elige una opcion');
}


function ConvertirFecha(opcion) {
  
  // Llama a la función ObtenerFecha para obtener la fecha formateada
  let fechaFormateada = ObtenerFecha(opcion);
  
  // Divide la fecha en día, mes y año
  let [dia, mes, año] = fechaFormateada.split("/");

  // Reorganiza la fecha en formato YYYY-MM-DD
  let fechaConvertida = `${año}-${mes}-${dia}`;

  return fechaConvertida;
}

function SumarDiasAFecha(dias) {
  // Obtiene la fecha en formato yyyy-MM-dd
  let fechaConvertida = ConvertirFecha();
  
  // Descompone la fecha en año, mes y día
  let [año, mes, dia] = fechaConvertida.split("-").map(Number);

  // Crea un objeto Date con los valores de año, mes y día
  let fecha = new Date(año, mes - 1, dia); // mes - 1 porque los meses en Date son indexados desde 0

  // Suma el número de días a la fecha
  fecha.setDate(fecha.getDate() + dias);

  // Formatea la nueva fecha en formato yyyy-MM-dd
  let nuevoAño = fecha.getFullYear();
  let nuevoMes = ("0" + (fecha.getMonth() + 1)).slice(-2); // Asegura dos dígitos para el mes
  let nuevoDia = ("0" + fecha.getDate()).slice(-2); // Asegura dos dígitos para el día

  let nuevaFecha = `${nuevoAño}-${nuevoMes}-${nuevoDia}`;

  return nuevaFecha;
}





//--------------------------------------------------------------------------------------------//
function obtenerDatosFactura(factura){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ListadoEstado');
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
  var invoiceColIndex = 5; // Columna F (indexada desde 0)
  var jsonColIndex = 12; // Columna M (indexada desde 0)
  Logger.log("factura "+factura)
  Logger.log("data length "+data.length)
  Logger.log(typeof(factura))
  //Logger.log("data +"+data)
  for (var i = 1; i < data.length; i++) { // Comienza en 1 para saltar la fila de encabezado
    Logger.log(data[i][invoiceColIndex])
    Logger.log(typeof(data[i][invoiceColIndex]))
    if (data[i][invoiceColIndex] == factura) {
      var jsonData = data[i][jsonColIndex];
      Logger.log("jsondata "+jsonData)
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
              Logger.log('IVA: ' + key + '%');
              
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

function testWriteNumeroIdentificacionToPlantilla() {
  var invoiceNumber = '192'; // Reemplaza con el número de factura deseado
  Logger.log(obtenerDatosFactura(invoiceNumber));
}

function resetPlantilla() {
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copia de Plantilla');

  // Borrar información de productos
  var colProductos = "A";
  var lineaProductos = 19;
  limpiarTablas(colProductos, lineaProductos);

  var colBases = "E";
  var lineaBases = 30;
  limpiarTablas(colBases, lineaBases);
  
  // Borrar información del cliente
  targetSheet.getRange('B12').clearContent();
  targetSheet.getRange('B13').clearContent();
  targetSheet.getRange('B14').clearContent();
  targetSheet.getRange('B15').clearContent();
  targetSheet.getRange('B16').clearContent();
  targetSheet.getRange('K12').clearContent();
  targetSheet.getRange('K13').clearContent();
  targetSheet.getRange('K14').clearContent();
  targetSheet.getRange('J15').clearContent();
  
  // Borrar valor a pagar, nota de pago y observaciones
  targetSheet.getRange('B41').clearContent();
  targetSheet.getRange('A45').clearContent();
  targetSheet.getRange('A50').clearContent();
  
  // Borrar total de items, descuentos y cargos
  targetSheet.getRange('B21').clearContent();
  targetSheet.getRange('A24').clearContent();
  targetSheet.getRange('D24').clearContent();
  
  // Borrar totales
  targetSheet.getRange('A32').clearContent();
  targetSheet.getRange('F32').clearContent();
  targetSheet.getRange('I32').clearContent();
  targetSheet.getRange('A36').clearContent();
  targetSheet.getRange('D36').clearContent();
  targetSheet.getRange('G36').clearContent();
  targetSheet.getRange('K36').clearContent();
  
}

function limpiarTablas(columna, linea){
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copia de Plantilla');
  var primeraFila = targetSheet.getRange(linea+":"+linea);
  primeraFila.clearContent();
  linea++;
  while (!targetSheet.getRange(columna+linea).isBlank()) {
    targetSheet.deleteRow(linea);
  }
}

function sacarColumnaFila(celda){
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Celdas Plantilla');
  var celdaDestino = hoja.getRange(celda).getValue();
  var match = celdaDestino.match(/([A-Z]+)(\d+)/);
  if (match) {
    var columna = match[1];  // 'B'
    var fila = parseInt(match[2], 10);  // 21
    
    return [columna, fila];
  } else {
    Logger.log('No se pudo dividir la referencia de celda.');
  }
}

function pruebaSacar(){
  var lista = sacarColumnaFila("E18")
  Logger.log(lista)
}

function subirFactura(nombre, pdfBlob) {
  var folder = DriveApp.getFolderById(folderId);
  var file = folder.createFile(pdfBlob.setName(`Factura ${nombre}.pdf`));
  var id = file.getId();
  return id;
}

function crearCarpeta() {
  var folder = DriveApp.createFolder("MisFacturas");
  var id = folder.getId();
  hojaDatosEmisor.getRange("B14").setValue(id);
}