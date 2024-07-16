var spreadsheet = SpreadsheetApp.getActive();
var prefactura_sheet = spreadsheet.getSheetByName('Factura');
var unidades_sheet = spreadsheet.getSheetByName('Unidades');
var datos_sheet = spreadsheet.getSheetByName('Datos2');


function onOpen() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  // https://developers.google.com/apps-script/guides/menus

  ui.createMenu('Sidebar')
    .addItem('Sidebar', 'showSidebar')
    .addToUi()
    .addItem('Producto', 'showPreProductos')
    .addToUi();


  return;
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('main')
      .setTitle('Menú prueba');
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

function showMenuFactura() {
  var html = HtmlService.createHtmlOutputFromFile('menuFactura')
    .setTitle('Menú Factura');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function openClientesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Clientes");
  SpreadsheetApp.setActiveSheet(sheet);
}

function openProductosSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Productos");
  SpreadsheetApp.setActiveSheet(sheet);
}

function showEnviarEmail() {
  var html = HtmlService.createHtmlOutputFromFile('enviarEmail')
    .setTitle('Enviar Email');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}


function processForm(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Productos");
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    
    const codigoReferencia = data.codigoReferencia;
    const nombre = data.nombre;
    const valorUnitario = parseFloat(data.valorUnitario);
    const iva = parseFloat(data.iva) / 100;
    const precioConIva = valorUnitario * (1 + iva);
    const impuestos = valorUnitario * iva;
    
    sheet.getRange(newRow, 1).setValue(codigoReferencia);
    sheet.getRange(newRow, 2).setValue(nombre);
    sheet.getRange(newRow, 3).setValue(valorUnitario);
    // Establece el IVA y formatea la celda como porcentaje
    const ivaCell = sheet.getRange(newRow, 4);
    ivaCell.setValue(iva); // Establece el valor del IVA como decimal
    ivaCell.setNumberFormat('0.00%'); // Formatea la celda como porcentaje con dos decimales
    sheet.getRange(newRow, 5).setValue(precioConIva); // Guarda el precio con IVA
    sheet.getRange(newRow, 6).setValue(impuestos); // Guarda el valor de los impuestos
    
    

    return "Datos guardados correctamente";
  } catch (error) {
    return "Error al guardar los datos: " + error.message;
  }
}

function generatePdfFromFactura() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Factura');
  
  if (!sheet) {
    throw new Error('La hoja Factura no existe.');
  }
  
  // Crear una nueva hoja de cálculo temporal
  var tempSpreadsheet = SpreadsheetApp.create('TempSpreadsheet');
  var tempSheet = tempSpreadsheet.getActiveSheet();
  
  // Copiar la hoja Factura a la hoja temporal
  sheet.copyTo(tempSpreadsheet);
  var newSheet = tempSpreadsheet.getSheets()[1];  // La hoja copiada es la segunda hoja
  tempSpreadsheet.deleteSheet(tempSheet);  // Borrar la hoja inicial que se crea con el nuevo archivo
  newSheet.setName('Factura');  // Renombrar la hoja copiada
  
  // Generar el PDF
  var pdf = DriveApp.getFileById(tempSpreadsheet.getId()).getAs('application/pdf').setName('Factura.pdf');
  
  // Borrar la hoja de cálculo temporal
  DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
  
  return pdf;
}

function getPdfUrl() {
  var pdfBlob = generatePdfFromFactura();
  var base64Data = Utilities.base64Encode(pdfBlob.getBytes());
  var contentType = pdfBlob.getContentType();
  var name = pdfBlob.getName();
  return `data:${contentType};base64,${base64Data}`;
}

function sendPdfByEmail(email) {
  var pdfFile = generatePdfFromFactura();
  var subject = 'Factura';
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


function onEdit(e){
  let hojaActual = e.source.getActiveSheet();
  verificarTipoDeDatos(e);

  if (hojaActual.getName()==="Factura"){

    let celdaEditada = e.range;
    let columnaContactos = 2; // Ajusta según sea necesario
    let rowContactos= 1;

    if (celdaEditada.getColumn() === columnaContactos || celdaEditada.getRow() === rowContactos) {
      //celda de elegir contacto
      Logger.log("No se editó un contacto válido");
      verificarYCopiarContacto(e);

    }else{


    }



    verificarYCopiarContacto(e);

  }else if(hojaActual.getName()==="Clientes"){
    verificarDatosObligatorios(e);

  }
}

function slugifyF (str) {
  var map = {
      '-' : ' ',
      '-' : '_',
      'a' : 'á|à|ã|â|À|Á|Ã|Â',
      'e' : 'é|è|ê|É|È|Ê',
      'i' : 'í|ì|î|Í|Ì|Î',
      'o' : 'ó|ò|ô|õ|Ó|Ò|Ô|Õ',
      'u' : 'ú|ù|û|ü|Ú|Ù|Û|Ü',
      'c' : 'ç|Ç',
      'n' : 'ñ|Ñ'
  };
  
  str = String(str)
  str = str.toLowerCase();
  
  for (var pattern in map) {
      str = str.replace(new RegExp(map[pattern], 'g'), pattern);
  };

  return str;
};
function getAdditionalDocuments() {
  //Browser.msgBox('getAddtionionalDocuments');
  var AdditionalDocuments = {
    "OrderReference": "",
    "DespatchDocumentReference": "",
    "ReceiptDocumentReference": "",
    "AdditionalDocument": []
  }
  return AdditionalDocuments;
}

var centenas = ['', 'Ciento ', 'Doscientos ', 'Trescientos ', 'Cuatrocientos ', 'Quinientos ', 'Seiscientos ', 
'Setecientos ', 'Ochocientos ', 'Novecientos ']

var decenas1 = ['Diez ', 'Once ', 'Doce ', 'Trece ', 'Catorce ', 'Quince ', 'Dieciseis ', 'Diecisiete ', 
'Dieciocho ', 'Diecinueve ']

var decenas2 = ['', 'Diez', 'Veinte ', 'Treinta ', 'Cuarenta ', 'Cincuenta ', 'Sesenta ', 'Setenta', 'Ochenta ', 'Noventa ']
var unidades = ['', 'Un ', 'Dos ', 'Tres ', 'Cuatro ', 'Cinco ', 'Seis ', 'Siete ', 'Ocho ', 'Nueve ']

function getPaymentMeans(PaymentMeansTxt) {
  switch (PaymentMeansTxt) {
    case 'Instrumento no definido':
      var PaymentMeans = 1;
      break;
    case 'Crédito ACH':
      var PaymentMeans = 2;
      break;
    case 'Débito ACH':
      var PaymentMeans = 3;
      break;
    case 'Reversión débito de demanda ACH':
      var PaymentMeans = 4;
      break;
    case 'Reversión crédito de demanda ACH':
      var PaymentMeans = 5;
      break;
    case 'Crédito de demanda ACH':
      var PaymentMeans = 6;
      break;
    case 'Débito de demanda ACH':
      var PaymentMeans = 7;
      break;
    case 'Mantener':
      var PaymentMeans = 8;
      break;
    case 'Clearing Nacional o Regional':
      var PaymentMeans = 9;
    case 'Efectivo':
      var PaymentMeans = 10;
      break;
    case 'Reversión Crédito Ahorro':
      var PaymentMeans = 11;
      break;
    case 'Reversión Débito Ahorro':
      var PaymentMeans = 12;
      break;
    case 'Crédito Ahorro':
      var PaymentMeans = 13;
      break;
    case 'Débito Ahorro':
      var PaymentMeans = 14;
      break;
    case 'Bookentry Crédito':
      var PaymentMeans = 15;
      break;
    case 'Bookentry Débito':
      var PaymentMeans = 16;
      break;
    case 'Concentración de la demanda en efectivo/Desembolso Crédito (CCD)':
      var PaymentMeans = 17;
      break;
    case 'Concentración de la demanda en efectivo/Desembolso (CCD) débito':
      var PaymentMeans = 18;
      break;
    case 'Crédito Pago negocio corporativo (CTP)':
      var PaymentMeans = 19;
      break;
    case 'Cheque':
      var PaymentMeans = 20;
      break;
    case 'Proyecto bancario':
      var PaymentMeans = 21;
      break;
    case 'Proyecto bancario certificado':
      var PaymentMeans = 22;
      break;
    case 'Cheque bancario':
      var PaymentMeans = 23;
      break;
    case 'Nota cambiaria esperando aceptación':
      var PaymentMeans = 24;
      break;
    case 'Cheque certificado':
      var PaymentMeans = 25;
      break;
    case 'Cheque Local':
      var PaymentMeans = 26;
      break;
    case 'Débito Pago Neogcio Corporativo (CTP)':
      var PaymentMeans = 27;
      break;
    case 'Crédito Negocio Intercambio Corporativo (CTX)':
      var PaymentMeans = 28;
      break;
    case 'Débito Negocio Intercambio Corporativo (CTX)':
      var PaymentMeans = 29;
      break;
    case 'Transferecia Crédito':
      var PaymentMeans = 30;
      break;
    case 'Transferencia Débito':
      var PaymentMeans = 31;
      break;
    case 'Concentración Efectivo/Desembolso Crédito plus (CCD+)':
      var PaymentMeans = 32;
      break;
    case 'Concentración Efectivo/Desembolso Débito plus (CCD+)':
      var PaymentMeans = 33;
      break;
    case 'Pago y depósito pre acordado (PPD)':
      var PaymentMeans = 34;
      break;
    case 'Concentración efectivo ahorros/Desembolso Crédito (CCD)':
      var PaymentMeans = 35;
      break;
    case 'Concentración efectivo ahorros / Desembolso Drédito (CCD)':
      var PaymentMeans = 36;
      break;
    case 'Pago Negocio Corporativo Ahorros Crédito (CTP)':
      var PaymentMeans = 37;
      break;
    case 'Pago Neogcio Corporativo Ahorros Débito (CTP)':
      var PaymentMeans = 38;
    default:
      Logger.log("Error: PaymentMeans");
      var PaymentMeans=100
  }
  return PaymentMeans;

}



function getPaymentType(PaymentTypeTxt) { 
  switch (PaymentTypeTxt) {
    case 'Contado':
      var PaymentType = 1;
      break;
    case 'Credito':
      var PaymentType = 2;
      break;
    default:
      Logger.log('Error: getPaymentType');
      Browser.msgBox("Oops! PaymentType");
  }
  return PaymentType;
}

function unos(n) {
    if (n == 0) {
        return '';
    }
    else {
        return unidades[n];
    }
}

function dieces(n) {
    var decena = Math.floor (n / 10);
    var unidad = n % 10;
    //Logger.log(`Decena: ${decena}`);
    switch (true) {
        case ((n % 10) == 0)://decena exacta
            return (decenas2[n / 10])
        case ((11 <= n) && (n <= 19))://
            return (decenas1[(n % 10)])
        case (Math.floor(n / 10) == 2)://22 a 29
            return `Veinti${unos(unidad).toLowerCase()}`
        case (0 <= n && n < 10)://unidad
            return (unos(n % 10))
        default://31 a 99
            var letras = `${decenas2[decena]} y ${unos(unidad)}`;
            return (letras)
    }
}

function cienes(n) {
    if (n == 100) {
        return 'Cien '
    }
    if (!n / 100) {
        return dieces(n)
    }
    else {
        return (centenas[Math.floor(n / 100)] + dieces(n % 100))
    }
}

function int2word(n) {
    megas = Math.floor(n / 1000 / 1000);
    kilos = Math.floor((n - megas * 1000000) / 1000);
    ones = n - megas * 1000000 - kilos * 1000;

    letras = ''
    if (megas >= 1) {
        if (megas == 1) {
            letras = letras + 'Un Millon '
        }
        else {
            letras = letras + cienes(megas) + ' Millones '
        }
    }
    if (kilos >= 1) {
        if (kilos == 1) {
            letras = letras + 'Mil'
        }
        else {
            letras = letras + cienes(kilos) + 'Mil '
        }
    }

    if (ones >= 1) {
        if (ones == 1) {
            letras = letras + 'Un '
        }
        else {
            letras = letras + cienes(ones)
        }
    }

    return letras;
}
function getAdditionalProperty() {
  //Browser.msgBox('getAddtionionalProperty');
  var AdditionalProperty = [];
  return AdditionalProperty;
}

function getCustomerInformation(customer) {
  /*esta funcion debe de cambiar para obtener son los datos directamente de la hoja cliente */


  var cell = datos_sheet.getRange("B50");
  //Browser.msgBox(customer);
  cell.setValue(customer);
  //Browser.msgBox("getCustomerInformation()");
  var range = datos_sheet.getRange("D50");
  var Customer = range.getValue();

  var range = datos_sheet.getRange("E50");
  var CustomerCode = range.getValue();

  range = datos_sheet.getRange("C51");
  var IdentificationType = range.getValue();

  range = datos_sheet.getRange("B52");
  var Identification = range.getValue();

  range = datos_sheet.getRange("B53");
  var DV = range.getValue();

  range = datos_sheet.getRange("B60");
  var Address = range.getValue().split('|');
  var AddressLine = Address[0];
  var AddressPostalZone = Address[1];

  range = datos_sheet.getRange("C60");
  var CityID = range.getValue();

  range = datos_sheet.getRange("B62");
  var Telephone = range.getValue();

  switch (datos_sheet.getRange("C1").getValue()) {
    case "Pruebas":
      var range = datos_sheet.getRange("E1");
      break;
    case "Produccion":
      var range = datos_sheet.getRange("B63");
      break;
    default:
      Logger.log("Oops!...Error Ambiente")
      return;
  }
  var Email = range.getValue();
  //Browser.msgBox(Email);


  range = datos_sheet.getRange("B64");
  var WebSiteURI = range.getValue();


  if (IdentificationType == "#NUM!") {
    Browser.msgBox("ERROR: Seleccione Tipo de Identificacion en Clientes")
    return;
  }

  var CustomerInformation = {
    "IdentificationType": IdentificationType,
    "Identification": Identification,//.toString(),
    "DV": DV,
    "RegistrationName": Customer,
    "CountryCode": "CO",
    "CountryName": "Colombia",
    "SubdivisionCode": datos_sheet.getRange("D60").getValue(),// 11, //Codigo de Municipio
    "SubdivisionName": datos_sheet.getRange("G60").getValue(),//"Bogotá", //Nombre de Departamente
    "CityCode": datos_sheet.getRange("E60").getValue(),//11001,
    "CityName": datos_sheet.getRange("F60").getValue(),//"Bogotá, D.C.",
    "AddressLine": String(AddressLine),
    "PostalZone": String(AddressPostalZone),
    "Email": Email,
    "CustomerCode": CustomerCode,
    "Telephone": Telephone,
    "WebSiteURI": WebSiteURI,
    "AdditionalAccountID": String(datos_sheet.getRange("C55").getValue()),//"1",//1, //1: Juridica, 2: Natural
    "TaxLevelCodeListName": String(datos_sheet.getRange("C54").getValue()),//"48" Impuesto sobre las ventas IVA 49 – No responsable de impuesto sobre las ventas IVA
    "TaxSchemeCode": String(datos_sheet.getRange("D54").getValue()),
    "TaxSchemeName": String(datos_sheet.getRange("E54").getValue()),
    "FiscalResponsabilities": String(datos_sheet.getRange("F54").getValue()).replace(/[|]/g, ';'),

    "PartecipationPercent": 100,
    "AdditionalCustomer": []


  }
  return CustomerInformation;
}

function getdatosValueA1(range) {
  var range = datos_sheet.getRange(range);
  return range.getValue();
}

function getDelivery() {
  var row = getdatosValueA1("C50");
  //Browser.msgBox(row);

  var Delivery = {
    "AddressLine": "",//getdatosValueA1("B61"),
    "CountryCode": "",//"CO",
    "CountryName": "",//"Colombia",
    "SubdivisionCode": "",//getdatosValueA1("D61"),//Departamento Codigo
    "SubdivisionName": "",//getdatosValueA1("G61"),///Departamento Nombre
    "CityCode": "",//getdatosValueA1("E61"),//Codigo Municipio
    "CityName": "",//getdatosValueA1("F61"),//Nombre Municip
    "ContactPerson": "",
    "DeliveryDate": "",
    "DeliveryCompany": ""
  };
  return Delivery;

}

function getMeasureUnitCode(measureName){
  var range = unidades_sheet.getRange("E1");
  
  var formula = `=DGET($A$1:$B$1104,A$1,{"Descripcion";"=${measureName}"})`;
  range.setValue(formula);

  return range.getValue();
}


function verificarTipoDeDatos(e){
    /*Funcion que verificar que celda o grupo de celdas editada
  y verifica su valor para saber si es valido 
  Input: e objeto que actua como una instancia del sheet editado 
  Output: no tiene output pero regresa un mensaje en caso de que sea erroneo el tipo de dato*/

  let sheet = e.range.getSheet();

  if(sheet.getName()==="Clientes"){//aca filtro de hoja, por cada hoja verifica cosas distintas
    let numIdentificacion = sheet.getRange("D2:D1000");
    let contacto = sheet.getRange("A2:A1000");
    let codigoContacto = sheet.getRange("B2:B1000");
    let primerNombre = sheet.getRange("H2:H1000");
    let segundoNombre = sheet.getRange("I2:I1000");
    let primeraApellido = sheet.getRange("J2:J1000");
    let segundoApellido = sheet.getRange("K2:K1000");
    let pais = sheet.getRange("l2:l1000");
    let codigoPostal =sheet.getRange("N2:N1000");
    let telefono =sheet.getRange("O2:O1000");
    let sitioWeb=codigoPostal =sheet.getRange("P2:P1000");
    let email =sheet.getRange("Q2:Q1000");
    let editedCell = e.range;

    esCeldaEnRango(numIdentificacion,editedCell,undefined,e);
    esCeldaEnRango(contacto,editedCell,"string",e);
    esCeldaEnRango(codigoContacto,editedCell,undefined,e);
    esCeldaEnRango(primerNombre,editedCell,"string",e);
    esCeldaEnRango(segundoNombre,editedCell,"string",e);
    esCeldaEnRango(primeraApellido,editedCell,"string",e);
    esCeldaEnRango(segundoApellido,editedCell,"string",e);
    esCeldaEnRango(pais,editedCell,"string",e)
    esCeldaEnRango(codigoPostal,editedCell,undefined,e);
    esCeldaEnRango(telefono,editedCell,undefined,e);
    esCeldaEnRango(sitioWeb,editedCell,"string",e)
    esCeldaEnRango(email,editedCell,"string",e)
  }
}

function esCeldaEnRango(range,editedCell,tipoDato='number',e){
      if (editedCell.getRow() >= range.getRow() && 
        editedCell.getRow() <= range.getLastRow() &&
        editedCell.getColumn() >= range.getColumn() &&
        editedCell.getColumn() <= range.getLastColumn()){
          let value = e.value;
          if(typeof value ==="undefined"){// no funciona value===null || value ==="null" || value ===''
            Logger.log("Ingreso algo vacio")
          }else{
            let newValue = convertirANumero(value);
            if (typeof newValue!==tipoDato){
              SpreadsheetApp.getUi().alert("Error: Solo se permite "+tipoDato+" en este rango");
              e.range.setValue("");
            }else{
              Logger.log("Ingreso el tipo de valor correcto")
            }
          }
        }
}

function convertirANumero(value){

  let number=Number(value);
  if(!isNaN(number)){
    return number;
  }else{
    return value;
  }

}

function getsheetValueA1(sheet, column, row) {
  var cell = column + row;
  var range = sheet.getRange(cell);
  return range.getValue();
}


function getsheetValue(sheet, column, row) {
  var range = sheet.getRange(column, row);
  return range.getValue();
}

function updatesheetValueA1(sheet, column, row, value) {
  var cell = column + row;
  var range = sheet.getRange(cell);
  range.setValue(value);
  return
}

