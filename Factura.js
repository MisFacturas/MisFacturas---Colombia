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
var folderId = hojaDatosEmisor.getRange("B14").getValue();
var paisesCodigos = {
  "Afganistán": "AF",
  "Albania": "AL",
  "Alemania": "DE",
  "Andorra": "AD",
  "Angola": "AO",
  "Antigua y Barbuda": "AG",
  "Arabia Saudita": "SA",
  "Argelia": "DZ",
  "Argentina": "AR",
  "Armenia": "AM",
  "Australia": "AU",
  "Austria": "AT",
  "Azerbaiyán": "AZ",
  "Bahamas": "BS",
  "Bangladés": "BD",
  "Barbados": "BB",
  "Baréin": "BH",
  "Bélgica": "BE",
  "Belice": "BZ",
  "Benín": "BJ",
  "Bielorrusia": "BY",
  "Birmania": "MM",
  "Bolivia": "BO",
  "Bosnia y Herzegovina": "BA",
  "Botsuana": "BW",
  "Brasil": "BR",
  "Brunéi": "BN",
  "Bulgaria": "BG",
  "Burkina Faso": "BF",
  "Burundi": "BI",
  "Bután": "BT",
  "Cabo Verde": "CV",
  "Camboya": "KH",
  "Camerún": "CM",
  "Canadá": "CA",
  "Catar": "QA",
  "Chad": "TD",
  "Chile": "CL",
  "China": "CN",
  "Chipre": "CY",
  "Ciudad del Vaticano": "VA",
  "Colombia": "CO",
  "Comoras": "KM",
  "Corea del Norte": "KP",
  "Corea del Sur": "KR",
  "Costa de Marfil": "CI",
  "Costa Rica": "CR",
  "Croacia": "HR",
  "Cuba": "CU",
  "Dinamarca": "DK",
  "Dominica": "DM",
  "Ecuador": "EC",
  "Egipto": "EG",
  "El Salvador": "SV",
  "Emiratos Árabes Unidos": "AE",
  "Eritrea": "ER",
  "Eslovaquia": "SK",
  "Eslovenia": "SI",
  "España": "ES",
  "Estados Unidos": "US",
  "Estonia": "EE",
  "Etiopía": "ET",
  "Filipinas": "PH",
  "Finlandia": "FI",
  "Fiyi": "FJ",
  "Francia": "FR",
  "Gabón": "GA",
  "Gambia": "GM",
  "Georgia": "GE",
  "Ghana": "GH",
  "Granada": "GD",
  "Grecia": "GR",
  "Guatemala": "GT",
  "Guyana": "GY",
  "Guinea": "GN",
  "Guinea ecuatorial": "GQ",
  "Guinea-Bisáu": "GW",
  "Haití": "HT",
  "Honduras": "HN",
  "Hungría": "HU",
  "India": "IN",
  "Indonesia": "ID",
  "Irak": "IQ",
  "Irán": "IR",
  "Irlanda": "IE",
  "Islandia": "IS",
  "Islas Marshall": "MH",
  "Islas Salomón": "SB",
  "Israel": "IL",
  "Italia": "IT",
  "Jamaica": "JM",
  "Japón": "JP",
  "Jordania": "JO",
  "Kazajistán": "KZ",
  "Kenia": "KE",
  "Kirguistán": "KG",
  "Kiribati": "KI",
  "Kosovo": "XK",
  "Kuwait": "KW",
  "Laos": "LA",
  "Lesoto": "LS",
  "Letonia": "LV",
  "Líbano": "LB",
  "Liberia": "LR",
  "Libia": "LY",
  "Liechtenstein": "LI",
  "Lituania": "LT",
  "Luxemburgo": "LU",
  "Macedonia del Norte": "MK",
  "Madagascar": "MG",
  "Malasia": "MY",
  "Malaui": "MW",
  "Maldivas": "MV",
  "Malí": "ML",
  "Malta": "MT",
  "Marruecos": "MA",
  "Mauricio": "MU",
  "Mauritania": "MR",
  "México": "MX",
  "Micronesia": "FM",
  "Moldavia": "MD",
  "Mónaco": "MC",
  "Mongolia": "MN",
  "Montenegro": "ME",
  "Mozambique": "MZ",
  "Namibia": "NA",
  "Nauru": "NR",
  "Nepal": "NP",
  "Nicaragua": "NI",
  "Níger": "NE",
  "Nigeria": "NG",
  "Noruega": "NO",
  "Nueva Zelanda": "NZ",
  "Omán": "OM",
  "Países Bajos": "NL",
  "Pakistán": "PK",
  "Palaos": "PW",
  "Panamá": "PA",
  "Papúa Nueva Guinea": "PG",
  "Paraguay": "PY",
  "Perú": "PE",
  "Polonia": "PL",
  "Portugal": "PT",
  "Reino Unido": "GB",
  "República Centroafricana": "CF",
  "República Checa": "CZ",
  "República del Congo": "CG",
  "República Democrática del Congo": "CD",
  "República Dominicana": "DO",
  "Ruanda": "RW",
  "Rumania": "RO",
  "Rusia": "RU",
  "Samoa": "WS",
  "San Cristóbal y Nieves": "KN",
  "San Marino": "SM",
  "San Vicente y las Granadinas": "VC",
  "Santa Lucía": "LC",
  "Santo Tomé y Príncipe": "ST",
  "Senegal": "SN",
  "Serbia": "RS",
  "Seychelles": "SC",
  "Sierra Leona": "SL",
  "Singapur": "SG",
  "Siria": "SY",
  "Somalia": "SO",
  "Sri Lanka": "LK",
  "Suazilandia": "SZ",
  "Sudáfrica": "ZA",
  "Sudán": "SD",
  "Sudán del Sur": "SS",
  "Suecia": "SE",
  "Suiza": "CH",
  "Surinam": "SR",
  "Tailandia": "TH",
  "Tanzania": "TZ",
  "Tayikistán": "TJ",
  "Timor Oriental": "TL",
  "Togo": "TG",
  "Tonga": "TO",
  "Trinidad y Tobago": "TT",
  "Túnez": "TN",
  "Turkmenistán": "TM",
  "Turquía": "TR",
  "Tuvalu": "TV",
  "Ucrania": "UA",
  "Uganda": "UG",
  "Uruguay": "UY",
  "Uzbekistán": "UZ",
  "Vanuatu": "VU",
  "Venezuela": "VE",
  "Vietnam": "VN",
  "Yemen": "YE",
  "Yibuti": "DJ",
  "Zambia": "ZM",
  "Zimbabue": "ZW"
};

var diccionarioCaluclarIva={
  "0.21": "21,00",
  "0.1": "1,00",
  "0.05": "5,00",
  "0.04": "4,00",
  "0": 0
}

function verificarEstadoValidoFactura() {
  // en esta funcion se debe de verificar si el numero de factura ya fue utiliazado en alguna otra factura
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  
  // función que verifica si una factura cumple con los requisitos mínimos para guardar
  let estaValido = true;

  let clienteActual = hojaFactura.getRange("B2").getValue();
  let informacionFactura = hojaFactura.getRange(2, 7, 5, 1).getValues();

  // Crear una lista combinada
  let listaCombinada = [clienteActual];  // Añadir clienteActual al array
  for (let i = 0; i < informacionFactura.length; i++) {
    listaCombinada.push(informacionFactura[i][0]);  // Añadir cada valor de informacionFactura
  }

  // Recorrer 
  for (let i = 0; i < listaCombinada.length; i++) {
    Logger.log(listaCombinada[i])
    if(listaCombinada[i]===""){
      estaValido=false
    }
  }

  let totalProductos=hojaFactura.getRange("A23").getValue();

  if (totalProductos==="Total productos"){
    // no hay necesidad de encontrar TOTAL PRODUCTOS si no esta, porque eso implica que si anadio asi sea 1 prodcuto
    let valorTotalProductos=hojaFactura.getRange("B23").getValue();
    if(valorTotalProductos===0){
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
  }else{
    Logger.log("Factura no valida")
  }
  

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
    Logger.log("entra a dictInformacionProducto")
    dictInformacionProducto = obtenerInformacionProducto(producto);
  }

  Logger.log("Pasa verificacion de producto")
  Logger.log("Number(taxSectionStartRow-1) "+Number(taxSectionStartRow-1))

  if(Number(taxSectionStartRow)===24){
    let totalProductos=hojaFactura.getRange("B23").getValue();
    if(totalProductos===5){
      hojaFactura.insertRowAfter(20)
      taxSectionStartRow += 1
    }
    Logger.log("lastProductRow dentro de coso ===24" +lastProductRow)
    for(let i =productStartRow;i<21;i++){
      let valorProducto= hojaFactura.getRange("A"+String(i)).getValue();
      if(valorProducto===""){
        //em estado deafult, sea agrega dentro de las primeras 6 lienas
        hojaFactura.getRange("A"+String(i)).setValue(producto);//producto
        hojaFactura.getRange("B" + String(i)).setValue(dictInformacionProducto["codigo Producto"]);//referencia
        hojaFactura.getRange("D" + String(i)).setValue(dictInformacionProducto["valor Unitario"]);//valor unitario sin iva
        hojaFactura.getRange("E" + String(i)).setValue(dictInformacionProducto["precio Con Iva"]);//precio con IVA
        hojaFactura.getRange("C"+String(i)).setValue(cantidad);//cantidad
        break
      }
    }
  }else{
    Logger.log("lastProductRow dentro de coso neuvo" +lastProductRow)
    let rowParaAgregar=Number(lastProductRow-2)
    Logger.log("rowParaAgregar"+rowParaAgregar)
    hojaFactura.getRange("A"+String(rowParaAgregar)).setValue(producto);//producto
    hojaFactura.getRange("B" + String(rowParaAgregar)).setValue(dictInformacionProducto["codigo Producto"]);//referencia
    hojaFactura.getRange("D" + String(rowParaAgregar)).setValue(dictInformacionProducto["valor Unitario"]);//valor unitario sin iva
    hojaFactura.getRange("E" + String(rowParaAgregar)).setValue(dictInformacionProducto["precio Con Iva"]);//precio con IVA
    hojaFactura.getRange("C"+String(rowParaAgregar)).setValue(cantidad);//cantidad
    //agg fila
    //tal vez aca aumntar el tax csoso para el bug
    hojaFactura.insertRowAfter(rowParaAgregar);
    hojaFactura.getRange("F"+String(rowParaAgregar)).setValue("=C"+String(rowParaAgregar)+"*D"+String(rowParaAgregar))
    hojaFactura.getRange("G"+String(rowParaAgregar)).setValue("=C"+String(rowParaAgregar)+"*E"+String(rowParaAgregar))
    taxSectionStartRow += 1
  }

  updateTotalProductCounter(hojaFactura, productStartRow, taxSectionStartRow);
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
  var fechaEmision = hojaFactura.getRange("G4").getValue();
  var estado = "Creada";
  var informacionCliente = getCustomerInformation(cliente);
  var nif = informacionCliente.Identification;

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

  var celdaNIF = hojaListado.getRange("C" + newRow);
  celdaNIF.setValue(nif);
  celdaNIF.setHorizontalAlignment('center');
  celdaNIF.setBorder(true, true, true, true, null, null, null, null);

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
  guardarIdArchivo(idArchivo, numeroFactura);

  var html = HtmlService.createHtmlOutputFromFile('postFactura')
    .setTitle('Menú');
  SpreadsheetApp.getUi()
    .showSidebar(html);
  

}

function guardarIdArchivo(idArchivo, numeroFactura) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Facturas ID');
  var lastRow = hoja.getLastRow();
  var newRow = lastRow + 1;
  hoja.getRange("A" + newRow).setValue(numeroFactura).setBorder(true, true, true, true, null, null, null, null);
  hoja.getRange("B" + newRow).setValue(idArchivo).setBorder(true, true, true, true, null, null, null, null);

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
  //Prueba
  //var lista = DriveApp.getFilesByName("Factura.pdf");
  //var nuevoId = lista.next().getId();

  var pdf = DriveApp.getFileById(idAsociado);
  var link = pdf.getDownloadUrl();
  //Logger.log(nuevoId);
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
  var sheet = ss.getSheetByName('Plantilla');

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

  hojaFactura.getRange("G3").setValue("")//hora
  hojaFactura.getRange("G4").setValue("")//fecha
  hojaFactura.getRange("G5").setValue("")//forma pago
  hojaFactura.getRange("G6").setValue(0)//dias vencimiento

  hojaFactura.getRange("B10").setValue("")//Osbervaciones
  hojaFactura.getRange("B11").setValue("")//IBAN
  hojaFactura.getRange("D11").setValue("")//Nota de pago 


  //productos
  const productStartRow = 15;
  let taxSectionStartRow = getTaxSectionStartRow(hojaFactura);
  const lastProductRow = getLastProductRow(hojaFactura, productStartRow, taxSectionStartRow);
  for(let i=productStartRow+1;i<=lastProductRow;i++){
    Logger
    hojaFactura.deleteRow(i);
  }

  hojaFactura.getRange("B15").setValue("")//producto
  hojaFactura.getRange("C15").setValue("")//cantidad
  hojaFactura.getRange("A15").setValue("")//referncia
  hojaFactura.getRange("G15").setValue("")//IVA%
  hojaFactura.getRange("H15").setValue("")//descuento
  hojaFactura.getRange("I15").setValue("")//retencion
  hojaFactura.getRange("J15").setValue("")//recargo
  hojaFactura.getRange("B16").setValue("")//tptal producto
  hojaFactura.getRange("B17").setValue("")//carrgos
  hojaFactura.getRange("B18").setValue("")//descuentos
}


function inicarFacturaNueva(){
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  let hojaInfoUsuario= spreadsheet.getSheetByName('Datos de emisor');
  let IABN=hojaInfoUsuario.getRange("B9").getValue()
  limpiarHojaFactura();

  hojaFactura.getRange("B11").setValue(IABN)
  generarNumeroFactura(); 
  obtenerFechaYHoraActual();
}

function limpiarYEliminarFila(numeroFila,hoja,hojaTax){
  //funcion para el boton que se va a agregar al final del producto
  if (numeroFila>20 && numeroFila<hojaTax){
    hoja.deleteRow(numeroFila)
  }else{
    hoja.getRange("A"+String(numeroFila)).setValue("");//producto
    hoja.getRange("B"+String(numeroFila)).setValue("");//ref
    hoja.getRange("C"+String(numeroFila)).setValue("");//cantidad
    hoja.getRange("D"+String(numeroFila)).setValue(0);//CON IVa
    hoja.getRange("E"+String(numeroFila)).setValue(0);//sin iva
    //sheet.getRange("C"+String(posicionTaxInfo)).setValue(valorEnPorcentaje);
  }
}

function verificarYCopiarContacto(e) {
  let hojaFacturas = e.source.getSheetByName('Factura');
  let hojaContactos = e.source.getSheetByName('Clientes');
  let celdaEditada = e.range;



  let nombreContacto = celdaEditada.getValue();
  let ultimaColumnaPermitida = 20; // Columna del estado en la hoja de contactos
  let datosARetornar = ["B", "O","M","L","N","Q"]; // Columnas que quiero de la hoja de contactos


  if (nombreContacto==="Cliente"){
    Logger.log("Estado default")
  }else{
    let listaConInformacion = obtenerInformacionCliente(nombreContacto);
    if (listaConInformacion["Estado"]==="No Valido"){
      SpreadsheetApp.getUi().alert("Error: El contacto seleccionado no es válido.");
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

  let fecha = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");
  let hora= Utilities.formatDate(new Date(), "GMT+1", "HH:mm:ss");

  sheet.getRange("G4").setValue(fecha)
  sheet.getRange("G3").setValue(hora)

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
        sheet.getRange("D14").setValue(data[i][2]);  // Valor unitario
        sheet.getRange("E14").setValue(data[i][4]);  // Otros datos,  segun sea necesario
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
    "Currency": "EUR",
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
function getPaymentSummary(posicionTotalFactura) {
  var total_factura = prefactura_sheet.getRange("A"+String(posicionTotalFactura)).getValue();// por ahora esto no lo utilizamos ya que no hay descuentos
  var monto_neto = prefactura_sheet.getRange("B"+String(posicionTotalFactura+1)).getValue();
  //var numeros = new Intl.NumberFormat('es-CO', {maximumFractionDigits:0, style: 'currency', currency: 'COP' }).format(monto);
  var numeros_total = new Intl.NumberFormat().format(total_factura);
  var numeros_neto = new Intl.NumberFormat().format(monto_neto);
  //Browser.msgBox(`${numeros_total}  ${numeros_neto}`);

  Logger.log("total_factura"+total_factura)
  Logger.log("monto_neto"+monto_neto)
  var PaymentTypeTxt = prefactura_sheet.getRange("G5").getValue();
  var PaymentMeansTxt = prefactura_sheet.getRange("E4").getValue();
  var PaymentSummary = {
    "PaymentType": PaymentTypeTxt,
    "PaymentMeans": "PaymentMeansTxt: No hay medio de pago",//a qui habia getPaymentMeans(PaymentMeansTxt)
    "PaymentNote": `Total Factura: $${numeros_total} \r Neto a Pagar  $${numeros_neto}: ${int2word(monto_neto)}`
  }
  return PaymentSummary;
}

function guardarYGenerarInvoice(){

  //obtener el total de prodcutos
  const posicionTotalProductos = prefactura_sheet.getRange("A23").getValue(); // para verificar donde esta el TOTAL
  if (posicionTotalProductos==="Total productos"){
    Logger.log("entra al primer if de json")
    var cantidadProductos=prefactura_sheet.getRange("B23").getValue();// cantidad total de productos 
  }else{
    const maxRows = prefactura_sheet.getLastRow();
    for(let i = 24;i<maxRows;i++){// 24 - porque 23 es el estado en donde deberia de estar el total prodcutos 
      let informacionCelda=prefactura_sheet.getRange("A"+String(i)).getValue();
      Logger.log("i"+i)
      Logger.log("informacionCelda"+informacionCelda)
      if(informacionCelda==="Total productos"){
        var cantidadProductos=prefactura_sheet.getRange("B"+String(i)).getValue();// cantidad total de productos 
        
      }
    }

  }

  let llavesParaLinea=prefactura_sheet.getRange("A14:G14");//llamo los headers 
  llavesParaLinea = slugifyF(llavesParaLinea.getValues()).replace(/\s/g, ''); // Todo en una sola linea
  const llavesFinales =llavesParaLinea.split(",");
  /* Creo que esto se puede cambiar a una manera mas simple, ya que los headers de la fila H7 hatsa N7 nunca van a cambiar */

  let invoiceTaxTotal=[];
  var productoInformation = [];
  Logger.log("cantidadProductos"+cantidadProductos)
  let saltarEspaciosEnBlanco=false
  let buffer=0
  if (cantidadProductos<5){
    Logger.log("entra a saltarEspaciosEnBlanco")
    saltarEspaciosEnBlanco=true
    buffer=5
  }
  let i = 15 // es 15 debido a que aqui empieza los productos elegidos por el cliente
  do{
    let filaActual = "A" + String(i) + ":G" + String(i);
    let rangoProductoActual=prefactura_sheet.getRange(filaActual);
    let productoFilaActual= String(rangoProductoActual.getValues());
    productoFilaActual=productoFilaActual.split(",");// cojo el producto de la linea actual y se le hace split a toda la info
    Logger.log(productoFilaActual)
    let LineaFactura={};

    for (let j=0;j<7;j++){// original dice que son 11=COL_TOTALES_PREFACTURA deberian ser 10 creo, en el nuevo son 7 tal vez 8
      LineaFactura[llavesFinales[j]]=productoFilaActual[j]
    }
    Logger.log(LineaFactura)

    let Name = LineaFactura['producto'];
    let ItemCode = new Number(LineaFactura['referencia']);
    let MeasureUnitCode = "Sin unidad"
    let Quantity = LineaFactura['cantidad'];
    let Price = LineaFactura['siniva'];
    let Amount = parseFloat(LineaFactura['importe']);//importe
    let ImpoConsumo = 1// no es un parametro para empresas espanolas
    let LineChargeTotal = parseFloat(LineaFactura['totaldelinea']);
    let Iva = LineChargeTotal-Amount;
    Logger.log("Name +")
    Logger.log("saltarEspaciosEnBlanco" +saltarEspaciosEnBlanco)
    Logger.log(i)
    if(Name==="" && saltarEspaciosEnBlanco){
      Logger.log("entra dentro del continue")
      i++
      continue

    }

    //IVA
    let ItemTaxesInformation = [];//taxes del producto en si
    let percent = parseFloat(((Iva / Amount) * 100).toFixed(1)); //aqui deberia de calcular el porcentaje pero como todavia no tengo IVA solo por ahora no
    let ivaTaxInformation = {
      Id: "01",//Id
      TaxEvidenceIndicator: false,
      TaxableAmount: Amount,
      TaxAmount: Iva,
      Percent: percent,
      BaseUnitMeasure: "",
      PerUnitAmount: ""
    };

    ItemTaxesInformation.push(ivaTaxInformation);
    invoiceTaxTotal.push(ivaTaxInformation);

    let LineExtensionAmount = Amount;
    let LineTotalTaxes = Iva + ImpoConsumo;

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
  }while(i<(15+cantidadProductos+buffer));//esos 5 son de buffer por si hay espacio entre los prodcutos

  /* Aqui empieza el proceso de coger el precio total de la facutra OJO en nuestro caso se agrupan por % de iva, entonces cambia
  algo mucho */
  

  //pasos para poder procesar todos los valores totales de la facutra agrupados por iva
  // let k=13;
  // do{

  //   let rangeLineaFacturaTotal=prefactura_sheet.getRange("A"+String(k)+":D"+String(k));
  //   let lineaFacturaTotal=String(rangeLineaFacturaTotal.getValues());
  //   lineaFacturaTotal=lineaFacturaTotal.split(",")
  //   //comaprador para que cuando encuentre un vacio se salga porque significa que ya acabo de leer
  //   let baseImponible=lineaFacturaTotal[0];
  //   let porcentajeIVA=lineaFacturaTotal[1];
  //   let IVA=lineaFacturaTotal[2];
  //   let total=lineaFacturaTotal[3];

  //   let invoice_total_2 = {
  //     "baseImponible": baseImponible,
  //     "porcentajeIVA": pfSubporcentajeIVATotal,
  //     "IVA": IVA,
  //     "total": total,
  //   }
  //   Logger.log(invoice_total_2)

  //   k++
  // }while(k<20);


  //estos es dinamico, verificar donde va el total cargo y descuento
  const posicionOriginalTotalFactura = prefactura_sheet.getRange("A32").getValue(); // para verificar donde esta el TOTAL
  let rangeFacturaTotal=""
  let cargo=0
  let descuento=0
  let posicionTotalFactura=31
  if (posicionOriginalTotalFactura==="Total factura"){
    rangeFacturaTotal=prefactura_sheet.getRange(31,1,1,4);
    cargo = prefactura_sheet.getRange("G24").getValue();
    descuento=prefactura_sheet.getRange("G25").getValue();
  }else{
    const maxRows = prefactura_sheet.getLastRow();//creo que maxrow siempre va a hacer la maxima, por ende es donde esta el total
    rangeFacturaTotal=prefactura_sheet.getRange((maxRows-1),1,1,4);//(maxRows-1) porque no necesito el total
    cargo = prefactura_sheet.getRange("G"+String(maxRows-8)).getValue();//(maxRows-8)  y -7 porque es donde deberia estar descuento y cargos
    descuento=prefactura_sheet.getRange("G"+String(maxRows-7)).getValue();
    posicionTotalFactura=maxRows-1
  }


  // aqui cambia con respecto al original, aqui deberia de cambiar el segundo parametro creo, seria con respecto a un j el cual seria la cantidad de ivas que hay
  let facturaTotal=String(rangeFacturaTotal.getValues());
  facturaTotal=facturaTotal.split(",");
  Logger.log(facturaTotal)



  /*Aqui cambia por completo, por ahora solo voy a dejar los parametros en numeros x 
  ,  solo coinciden el base imponible he IVA */
  let pfSubTotal = parseFloat(facturaTotal[0]);//base imponible
  let pfIVA = parseFloat(facturaTotal[2]);//IVA
  let pfImpoconsumo = 0;
  let pfTotal = parseFloat(facturaTotal[3]);
  let pfRefuente = 0;
  let pfReteICA = 0;
  let pfReteIVA = 0;
  let pfTRetenciones = 0; 
  let pfAnticipo = descuento;
  let pfTPagar = 0;

  // if (pfRefuente > 0) {
  //   let Percent = parseFloat((pfRefuente / pfSubTotal * 100).toFixed(2));
  //   let retefuente_taxinformation = {
  //     Id: "06",//Id,
  //     TaxEvidenceIndicator: true,
  //     TaxableAmount: pfSubTotal,
  //     TaxAmount: pfRefuente,
  //     Percent: Percent,
  //     BaseUnitMeasure: "",
  //     PerUnitAmount: ""
  //   };
  //   invoiceTaxTotal.push(retefuente_taxinformation);
  // };

  // if (pfReteICA > 0) {
  //   let Factor = datos_sheet.getRange("B8").getValue();
  //   let PercentReteICA = (Factor * 100).toFixed(3);
  //   let invoice_ReteICA = {
  //     Id: "07",//Id,
  //     TaxEvidenceIndicator: true,
  //     TaxableAmount: pfSubTotal,
  //     TaxAmount: pfReteICA,
  //     Percent: parseFloat(PercentReteICA),
  //     BaseUnitMeasure: "",
  //     PerUnitAmount: ""
  //   };
  //   invoiceTaxTotal.push(invoice_ReteICA);
  // }

  // if (pfReteIVA > 0) {
  //   let FactorReteIva = pfReteIVA / pfSubTotal;
  //   let PercentReteIVA = (FactorReteIva * 100).toFixed(2);
  //   let invoice_reteIVA = {
  //     Id: "05",
  //     TaxEvidenceIndicator: true,
  //     TaxableAmount: pfSubTotal,
  //     TaxAmount: pfReteIVA,
  //     Percent: parseFloat(PercentReteIVA),
  //     BaseUnitMeasure: "",
  //     PerUnitAmount: ""
  //   };
  //   invoiceTaxTotal.push(invoice_reteIVA);
  // }

  //Aqui seguiria el texto, pero en el de carlos nunca lo llama 

  let invoice_total = {
    "LineExtensionAmount": pfSubTotal,
    "TaxExclusiveAmount": pfSubTotal,
    "TaxInclusiveAmount": pfTotal,
    "AllowanceTotalAmount": 0,
    "ChargeTotalAmount": cargo,
    "PrePaidAmount": pfAnticipo,
    "PayableAmount": (pfTotal-pfAnticipo+cargo) // antes era (pfTotal - pfAnticipo) 
  }


  let cliente = prefactura_sheet.getRange("B2").getValue();
  let InvoiceGeneralInformation = getInvoiceGeneralInformation();
  let CustomerInformation = getCustomerInformation(cliente);// tal ves que por ahora no llame al cliente

  let invoice = JSON.stringify({
    CustomerInformation: CustomerInformation,
    InvoiceGeneralInformation: InvoiceGeneralInformation,
    Delivery: getDelivery(),
    AdditionalDocuments: getAdditionalDocuments(),
    AdditionalProperty: getAdditionalProperty(),
    PaymentSummary: getPaymentSummary(posicionTotalFactura), //por ahora esto leugo se cambia la funcion getPaymentSummary para que cumpla los parametros
    ItemInformation: productoInformation,
    //Invoice_Note: invoice_note,
    InvoiceTaxTotal: invoiceTaxTotal,
    InvoiceAllowanceCharge: [],
    InvoiceTotal: invoice_total
  });
  Logger.log(invoice)

  let nameString = prefactura_sheet.getRange("B2").getValue();
  let numeroFactura = JSON.stringify(InvoiceGeneralInformation.InvoiceNumber);
  let fecha = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");
  let codigoCliente=prefactura_sheet.getRange("B3").getValue();
  listadoestado_sheet.appendRow(["vacio", "vacio","vacio" , fecha,"vacio" ,numeroFactura ,nameString ,codigoCliente,"vacio" ,"vacio" ,"representacion" ,"Vacio", String(invoice)]);
  
  
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
          var nif = invoiceData.CustomerInformation.Identification;
          var codigo = invoiceData.CustomerInformation.AdditionalAccountID;
          var direccion = invoiceData.CustomerInformation.AddressLine;
          var telefono = invoiceData.CustomerInformation.Telephone;
          var poblacion = invoiceData.CustomerInformation.CityName;
          var provincia = invoiceData.CustomerInformation.SubdivisionName;
          var pais = invoiceData.CustomerInformation.CountryName;
          var fechaEmision = invoiceData.CustomerInformation.DV;
          var formaPago = invoiceData.PaymentSummary.PaymentType;
          var listaProductos = invoiceData.ItemInformation;
          var numeroProductos = 0;
          var totalFacturaJSON = parseFloat(invoiceData.InvoiceTotal.PayableAmount);
          var valorPagar = int2word(totalFacturaJSON) //arreglar
          var notaPago = invoiceData.PaymentSummary.PaymentNote;
          var observaciones = invoiceData.InvoiceGeneralInformation.Note;

          var filasInsertadas = 0;
          var filasInsertadasPorProductos = 0;
          var grupoIva = {};

          var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Plantilla'); // Hoja donde quieres insertar el NIF
          if (!targetSheet) {
            targetSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Plantilla');
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
            celdaPrecioUnitario.setNumberFormat('€#,##0.00')
            
            var celdaIva = targetSheet.getRange('H'+numeroCelda);
            celdaIva.setBorder(true,true,true,true,null,null,null,null);
            celdaIva.setValue((listaProductos[j].TaxesInformation[0].Percent)/100);
            celdaIva.setNumberFormat('0.0%');
            celdaIva.setHorizontalAlignment('center');

            
            var celdaImporte = targetSheet.getRange('I'+numeroCelda);
            celdaImporte.setBorder(true,true,true,true,null,null,null,null);
            celdaImporte.setValue(listaProductos[j].LineExtensionAmount);
            celdaImporte.setHorizontalAlignment('normal');
            celdaImporte.setNumberFormat('€#,##0.00')

            var producto = listaProductos[j]
            //crea un diccionario que la llave sea el % de iva y el valor sea el total de la linea
            
            if (grupoIva.hasOwnProperty(producto.TaxesInformation[0].Percent)) {
              grupoIva[producto.TaxesInformation[0].Percent] += producto.TaxesInformation[0].TaxableAmount;
            } else {
              grupoIva[producto.TaxesInformation[0].Percent] = producto.TaxesInformation[0].TaxableAmount;
            }
          }
          var contador = 0;
          var auxiliarFilasInsertadas = filasInsertadas;
          for (var key in grupoIva) {
            if (grupoIva.hasOwnProperty(key)) {
              var numeroCelda = 27 + auxiliarFilasInsertadas;
              if (contador > 0) {
                targetSheet.insertRowAfter(numeroCelda);
                targetSheet.getRange('A'+(numeroCelda+1)+':D'+(numeroCelda+1)).merge();
                targetSheet.getRange('F'+(numeroCelda+1)+':G'+(numeroCelda+1)).merge();
                targetSheet.getRange('H'+(numeroCelda+1)+':I'+(numeroCelda+1)).merge();
                filasInsertadas += 1;
                auxiliarFilasInsertadas += 1;
              } else {
                auxiliarFilasInsertadas += 1;
              }
              var celdaBaseImponible = targetSheet.getRange('A'+numeroCelda);
              celdaBaseImponible.setBorder(true,true,true,true,null,null,null,null);
              celdaBaseImponible.setValue(grupoIva[key]);
              celdaBaseImponible.setNumberFormat('€#,##0.00');
              celdaBaseImponible.setHorizontalAlignment('normal');
              
              var celdaPorcentajeIva = targetSheet.getRange('E'+numeroCelda);
              celdaPorcentajeIva.setBorder(true,true,true,true,null,null,null,null);
              celdaPorcentajeIva.setValue(key/100);
              celdaPorcentajeIva.setNumberFormat('0.0%');
              celdaPorcentajeIva.setHorizontalAlignment('center');
              
              var celdaIVA = targetSheet.getRange('F'+numeroCelda);
              celdaIVA.setBorder(true,true,true,true,null,null,null,null);
              celdaIVA.setFormula('=A'+numeroCelda+'*E'+numeroCelda);
              celdaIVA.setNumberFormat('€#,##0.00');
              celdaIVA.setHorizontalAlignment('normal');
              
              var celdaTotal = targetSheet.getRange('H'+numeroCelda);
              celdaTotal.setBorder(true,true,true,true,null,null,null,null);
              celdaTotal.setFormula('=A'+numeroCelda+'+F'+numeroCelda);
              celdaTotal.setNumberFormat('€#,##0.00');
              celdaTotal.setHorizontalAlignment('normal');

              contador += 1;
              Logger.log('IVA: ' + key + '%');
            }
          }

          //Extaccion celdas de datos cliente
          var clienteCeldaHoja = hojaCeldas.getRange('E3').getValue();
          var nifCeldaHoja = hojaCeldas.getRange('E4').getValue();
          var codigoCeldaHoja = hojaCeldas.getRange('E8').getValue();
          var direccionCeldaHoja = hojaCeldas.getRange('E5').getValue();
          var telefonoCeldaHoja = hojaCeldas.getRange('E7').getValue();
          var poblacionCeldaHoja = hojaCeldas.getRange('E6').getValue();
          var fechaEmisionCeldaHoja = hojaCeldas.getRange('E9').getValue();
          var formaPagoCeldaHoja = hojaCeldas.getRange('E10').getValue();

          //factura
          var celdaNumFactura = targetSheet.getRange('A9');
          //Datos Cliente
          var clienteCell = targetSheet.getRange(clienteCeldaHoja);
          var nifCell = targetSheet.getRange(nifCeldaHoja);
          var codigoCell = targetSheet.getRange(codigoCeldaHoja);
          var direccionCell = targetSheet.getRange(direccionCeldaHoja);
          var telefonoCell = targetSheet.getRange(telefonoCeldaHoja);
          var poblacionCell = targetSheet.getRange(poblacionCeldaHoja);
          var fechaEmisionCell = targetSheet.getRange(fechaEmisionCeldaHoja);
          var formaPagoCell = targetSheet.getRange(formaPagoCeldaHoja);
          var valorPagarCell = targetSheet.getRange('B'+(34+filasInsertadas));
          var notaPagoCell = targetSheet.getRange('A'+(38+filasInsertadas));
          var observacionesCell = targetSheet.getRange('A'+(43+filasInsertadas));
          var totalItemsCell = targetSheet.getRange('B'+(21+filasInsertadasPorProductos));
          var descuentosCell = targetSheet.getRange('A'+(32+filasInsertadas));
          var cargosCell = targetSheet.getRange('C'+(32+filasInsertadas));
          var sumaBaseImponible = targetSheet.getRange('B'+(29+filasInsertadas));
          var sumaImpIva = targetSheet.getRange('F'+(29+filasInsertadas));
          var sumaTotal = targetSheet.getRange('H'+(29+filasInsertadas));

          celdaNumFactura.setValue("FACTURA DE VENTA NO. "+facturaNumero);
          clienteCell.setValue(cliente);
          nifCell.setValue(nif);
          codigoCell.setValue(codigo);
          direccionCell.setValue(direccion);
          telefonoCell.setValue(telefono);
          // Ajustar la forma en que se ve el pais - IMPORTANTE
          if (poblacion == "" || provincia == "" || pais == "") {
            var columnaPoblacion = poblacionCell.getColumn();
            var filaPoblacion = poblacionCell.getRow();
            targetSheet.getRange(filaPoblacion, columnaPoblacion-1).clearContent();
          } else {
            poblacionCell.setValue(poblacion+', '+provincia+', '+pais);
          }
          fechaEmisionCell.setValue(fechaEmision);
          formaPagoCell.setValue(formaPago);
          valorPagarCell.setValue(valorPagar);
          notaPagoCell.setValue(notaPago);
          observacionesCell.setValue(observaciones);
          totalItemsCell.setValue(numeroProductos);
          descuentosCell.setValue(0);
          cargosCell.setValue(0);
          sumaBaseImponible.setFormula('=SUM(A'+(27+numeroProductos-1)+':A'+(28+filasInsertadas-1)+')');
          sumaImpIva.setFormula('=SUM(F'+(27+numeroProductos-1)+':F'+(28+filasInsertadas-1)+')');
          sumaTotal.setFormula('=SUM(H'+(27+numeroProductos-1)+':H'+(28+filasInsertadas-1)+')');
          
          
          var itemCellPrueba = targetSheet.getRange('A19')
          var hojaEnBlanco = clienteCell.isBlank() || formaPagoCell.isBlank() || itemCellPrueba.isBlank() || celdaBaseImponible.isBlank();
          while (hojaEnBlanco) {
            sleep(1000);
            hojaEnBlanco = clienteCell.isBlank() || formaPagoCell.isBlank() || itemCellPrueba.isBlank() || celdaBaseImponible.isBlank();
          }

          if (!hojaEnBlanco){
            var pdfFactura = generatePdfFromPlantilla();
            var id = subirFactura(facturaNumero, pdfFactura);
            resetPlantilla();
            return id;
          }
          

        } catch (e) {
          Logger.log('Error parsing JSON for row ' + (i + 1) + ': ' + e.message);
        }
      }
    }
  }
  Logger.log('Invoice number ' + factura + ' not found.');
}

function testWriteNIFToPlantilla() {
  var invoiceNumber = '192'; // Reemplaza con el número de factura deseado
  Logger.log(obtenerDatosFactura(invoiceNumber));
}

function resetPlantilla() {
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Plantilla');

  // Borrar información de productos
  var colProductos = "A";
  var lineaProductos = 19;
  limpiarTablas(colProductos, lineaProductos);

  var colBases = "E";
  var lineaBases = 27;
  limpiarTablas(colBases, lineaBases);
  
  // Borrar información del cliente
  targetSheet.getRange('B12').clearContent();
  targetSheet.getRange('B13').clearContent();
  targetSheet.getRange('B14').clearContent();
  targetSheet.getRange('B15').clearContent();
  targetSheet.getRange('B16').clearContent();
  targetSheet.getRange('H14').clearContent();
  targetSheet.getRange('H15').clearContent();
  targetSheet.getRange('H12').clearContent();
  targetSheet.getRange('H13').clearContent();
  
  // Borrar valor a pagar, nota de pago y observaciones
  targetSheet.getRange('B34').clearContent();
  targetSheet.getRange('A38').clearContent();
  targetSheet.getRange('A43').clearContent();
  
  // Borrar total de items, descuentos y cargos
  targetSheet.getRange('B21').clearContent();
  targetSheet.getRange('A32').clearContent();
  targetSheet.getRange('C32').clearContent();
  
  
}

function limpiarTablas(columna, linea){
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Plantilla');
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
  var folder = DriveApp.createFolder("FacturasApp");
  var id = folder.getId();
  hojaDatosEmisor.getRange("B14").setValue(id);
}