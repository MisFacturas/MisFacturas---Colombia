var spreadsheet = SpreadsheetApp.getActive();
//let unidades_sheet = spreadsheet.getSheetByName('Unidades');
//let datos_sheet = spreadsheet.getSheetByName('Datos2');

// directorio alejandro C:\\Users\\catan\\OneDrive\\Documents\\Work\\Appsheets\\MisFacturasApp
// directorio sebastian C:\\Users\\elfue\\Documents\\MisFacturasApp


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

function openHistorialSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Historial Facturas");
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

function showClientes() {
  var html = HtmlService.createHtmlOutputFromFile('menuCliente')
    .setTitle('Menu cliente');
  SpreadsheetApp.getUi()
    .showSidebar(html);
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

function generatePdfFromPlantilla() {
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
    '&top_margin=0.00' +  // Margen superior
    '&bottom_margin=0.00' +  // Margen inferior
    '&left_margin=0.00' +  // Margen izquierdo
    '&right_margin=0.00' +  // Margen derecho
    '&horizontal_alignment=CENTER' +  // Alineación horizontal
    '&vertical_alignment=MIDDLE';  // Alineación vertical

  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });

  var pdfBlob = response.getBlob().setName('Plantilla.pdf');
  return pdfBlob;
}

function getPdfUrl() {
  var pdfBlob = generatePdfFromPlantilla();
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
    let rowEditada=celdaEditada.getRow();
    let colEditada=celdaEditada.getColumn();

    let columnaContactos = 3; // Ajusta según sea necesario
    let rowContactos= 2;
    

    const productStartRow = 15; // prodcutos empeiza aca
    const productEndColumn = 8; //   procutos terminan en column H
    let taxSectionStartRow = getTaxSectionStartRow(hojaActual); // Assuming products end at column H
    //Logger.log("taxSectionStartRow "+taxSectionStartRow)

    if (colEditada === columnaContactos && rowEditada === rowContactos) {
      //celda de elegir contacto
      Logger.log("No se editó un contacto válido");
      verificarYCopiarContacto(e);
      obtenerFechaYHoraActual(hojaActual)
      generarNumeroFactura(hojaActual)

    }
    else if (rowEditada >= productStartRow && colEditada == 2 && rowEditada<taxSectionStartRow) {//asegurar que si sea dentro del espacio permititdo(donde empieza el taxinfo)
      const lastProductRow = getLastProductRow(hojaActual, productStartRow, taxSectionStartRow);
      Logger.log("lastProductRow "+lastProductRow)
      const nextRow = lastProductRow + 1;
      //Logger.log("entra al primer else if")
      //Logger.log("next row "+nextRow)
      Logger.log("taxSectionStartRow "+taxSectionStartRow) 

      //PRoceso para ingresar la info del producto
      let valorCelda=celdaEditada.getValue();
      let dictInformacionProducto= obtenerInformacionProducto(valorCelda);
      

      // Insertar una nueva row 
      let diferencia =Math.abs(taxSectionStartRow-lastProductRow)
      if(diferencia<=2 && rowEditada ===lastProductRow){
        hojaActual.getRange("C"+String(rowEditada)).setValue(dictInformacionProducto["codigo Producto"]);//referencia
        hojaActual.getRange("E"+String(rowEditada)).setValue(dictInformacionProducto["valor Unitario"]);//valor unitario sin iva
        hojaActual.getRange("F"+String(rowEditada)).setValue(dictInformacionProducto["precio Con Iva"]);//precio con IVA
        Logger.log("Entra a la comparacion de taxSectionStartRow-lastProductRow")
        Logger.log("Differencia"+diferencia)
        hojaActual.insertRowAfter(lastProductRow);//tal vez aca aumntar el tax csoso para el bug
        taxSectionStartRow+=1
        calcularImporteYTotal(hojaActual, rowEditada);
      }else if (lastProductRow < taxSectionStartRow) {//erores ? deberia de ser la ultima valida 
        // insertar cosas del producto en la hoja
        
        hojaActual.getRange("C"+String(rowEditada)).setValue(dictInformacionProducto["codigo Producto"]);//referencia
        hojaActual.getRange("E"+String(rowEditada)).setValue(dictInformacionProducto["valor Unitario"]);//valor unitario sin iva
        hojaActual.getRange("F"+String(rowEditada)).setValue(dictInformacionProducto["precio Con Iva"]);//precio con IVA
        //calcular importe y total de linea apenas se ingrese el valor de cantidad

        
        Logger.log("Entra al segundo if dnetro del else if ")
        Logger.log("")
        calcularImporteYTotal(hojaActual, rowEditada);
      }
    }else if (rowEditada >= productStartRow && colEditada == 4 && rowEditada<taxSectionStartRow){// edita celda cantidad
      //calcular Importe y Total de linea
      calcularImporteYTotal(hojaActual, rowEditada);

    }
    //calcularTaxInformation(celdaEditada,productStartRow,taxSectionStartRow);
    updateTotalProductCounter(hojaActual, productStartRow, taxSectionStartRow,celdaEditada);//tengo que revisar esto 

  }else if(hojaActual.getName()==="Clientes"){
    verificarDatosObligatorios(e);

  }
}

function calcularImporteYTotal(hojaActual, rowEditada) {
  Logger.log("rowEditada"+rowEditada)
  let producto = hojaActual.getRange("B" + String(rowEditada)).getValue(); // Obtiene el producto en la línea seleccionada
  let dictInformacionProducto = obtenerInformacionProducto(producto);
  let cantidadProducto = hojaActual.getRange("D" + String(rowEditada)).getValue(); // Asume que la cantidad está en la columna D
  Logger.log("producto"+producto)
  Logger.log("cantidadProducto"+cantidadProducto)
  let importe = cantidadProducto * dictInformacionProducto["valor Unitario"];
  let totalDeLinea = cantidadProducto * dictInformacionProducto["precio Con Iva"];

  hojaActual.getRange("G" + String(rowEditada)).setValue(importe);
  hojaActual.getRange("H" + String(rowEditada)).setValue(totalDeLinea);
}

function getLastProductRow(sheet, productStartRow, taxSectionStartRow) {
  //tal vez se lo tira si no agrega en orden
  let lastProductRow = productStartRow;

  for (let row = productStartRow; row < taxSectionStartRow; row++) {
    if (sheet.getRange(row, 2).getValue() !== '') { 
      lastProductRow = row;
    }
  }
  //aqui arrelgar error que se agrega una nueva linea cuando hay espacio arriba
  return lastProductRow;
}

function getTaxSectionStartRow(sheet) {
  //obtiene la row donde esta la seccion de taxinformation
  const maxRows = sheet.getMaxRows();
  //Logger.log("get tac section star row")
  //Logger.log(maxRows)

  // obtenemos la row donde esta la "base imponible, llegando asi al principio "
  for (let row = 22; row <= maxRows; row++) { // 22 porque su row predetermmiado es ese
    if (sheet.getRange(row, 2).getValue() === 'Base imponible') {

      Logger.log("dentro de getTax row "+row)
      return row;
    }
  }

  // esto puede causar errores
  return maxRows + 1;
}

function updateTotalProductCounter(sheet, productStartRow, taxSectionStartRow,celdaEditada) {
  let totalProducts = 0;
  let rowEdited= celdaEditada.getRow();
  
  //toca revisar creo que cuando hay un producto con un espacio en el medio no teien encuenta y se sale 
  limpiarDict();
  // calcualr cuando no hay cantidad
  for (let row = productStartRow; row < taxSectionStartRow; row++) {
    let prodcutoActual=sheet.getRange(row, 2).getValue()
    if(prodcutoActual===""){
        Logger.log("PRODUCTO VACIO")
    }else{
      totalProducts++;
      let dictInformacionProducto= obtenerInformacionProducto(prodcutoActual);
      //Logger.log("dictInformacionProducto"+dictInformacionProducto)
      let porcientoIVA=dictInformacionProducto["porciento Iva"];

      Logger.log("porcientoIVA "+porcientoIVA)
      if (porcientoIVA in diccionarioCaluclarIva){
        Logger.log("entra a coger el importe")
        let importeActual=sheet.getRange("G"+String(row)).getValue();
        Logger.log("importeActual "+importeActual)
        Logger.log("Row"+ row)
        diccionarioCaluclarIva[porcientoIVA]+=importeActual;
      }
    }

    // if (sheet.getRange(row, 2).getValue() !== '') { // Assuming product names are in column B
    //   totalProducts++;
    // }
  }

  Logger.log("Obtener llaves del dict")
  let llavesDiccionarioProducto=Object.keys(diccionarioCaluclarIva);
  let posicionTaxInfo=taxSectionStartRow+1;//tal vez +1 > row?
  Logger.log("posicionTaxInfo "+posicionTaxInfo)
  Logger.log("llavesDiccionarioProducto"+llavesDiccionarioProducto)
  for(let k=0;k<llavesDiccionarioProducto.length;k++){
    let llaveActual =llavesDiccionarioProducto[k];
    let valorllave=diccionarioCaluclarIva[llaveActual];
    Logger.log("llaveActual "+llaveActual)
    Logger.log("valorllave"+ valorllave)
    if(valorllave===0){
      //revisar que ya se halla borrado de la lista de total taxes, ya que esto implica que no hay ningun prodcuto con este %de IVA
      let RangeIVAActivos=sheet.getRange(posicionTaxInfo,3,6)// 3 porque es donde esta el IVA
      let IVAsActivos=RangeIVAActivos.getValues();
      Logger.log("IVAsActivos sin String " +IVAsActivos)
      let IVAsActivos2=String(RangeIVAActivos.getValues());
      Logger.log("IVAsActivos CON String " +IVAsActivos2)


      continue
    }else{
      sheet.getRange("B"+String(posicionTaxInfo)).setValue(valorllave);
      let valorEnPorcentaje=(llaveActual * 100) + '%';
      sheet.getRange("C"+String(posicionTaxInfo)).setValue(valorEnPorcentaje);
      sheet.getRange("C"+String(posicionTaxInfo)).setNumberFormat("0.00%");
      Logger.log("SetnumberFormat?")
      posicionTaxInfo++;
    }
  }



  // Set the total products count in cell B27
  sheet.getRange('H13').setValue(totalProducts);
}

function limpiarDict(){
  Logger.log("Limpiar el dict")
  diccionarioCaluclarIva={
    "0.21": 0,
    "0.1": 0,
    "0.05": 0,
    "0.04": 0,
    "0": 0
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
  var decena = Math.floor(n / 10);
  var unidad = n % 10;
  switch (true) {
      case ((n % 10) == 0):
          return (decenas2[n / 10]);
      case ((11 <= n) && (n <= 19)):
          return (decenas1[(n % 10)]);
      case (Math.floor(n / 10) == 2):
          return `Veinti${unos(unidad).toLowerCase()}`;
      case (0 <= n && n < 10):
          return (unos(n % 10));
      default:
          var letras = `${decenas2[decena]} y ${unos(unidad)}`;
          return (letras);
  }
}

function cienes(n) {
  if (n == 100) {
      return 'Cien ';
  }
  if (n < 100) {
      return dieces(n);
  }
  else {
      return (centenas[Math.floor(n / 100)] + dieces(n % 100));
  }
}

function int2word(n) {
  var euros = Math.floor(n);
  var centimos = Math.round((n - euros) * 100);

  var megas = Math.floor(euros / 1000 / 1000);
  var kilos = Math.floor((euros - megas * 1000000) / 1000);
  var ones = euros - megas * 1000000 - kilos * 1000;

  var letras = '';
  if (megas >= 1) {
      if (megas == 1) {
          letras = letras + 'Un Millón ';
      } else {
          letras = letras + cienes(megas) + ' Millones ';
      }
  }
  if (kilos >= 1) {
      if (kilos == 1) {
          letras = letras + 'Mil ';
      } else {
          letras = letras + cienes(kilos) + 'Mil ';
      }
  }

  if (ones >= 1) {
      if (ones == 1) {
          letras = letras + 'Un ';
      } else {
          letras = letras + cienes(ones);
      }
  }

  if (centimos > 0) {
      letras = letras  +'Euros'+ `Con ${cienes(centimos)}Céntimos`;
  }

  return letras;
}

function getAdditionalProperty() {
  var AdditionalProperty = [];
  return AdditionalProperty;
}

function getdatosValueA1(range) {
  var range = datos_sheet.getRange(range);
  return range.getValue();
}

function getDelivery() {
  var row = getdatosValueA1("C50");

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
    let provincia=sheet.getRange("M2:M1000");
    let poblacion=sheet.getRange("N2:N1000");
    let direccion =sheet.getRange("O2:O1000");
    let codigoPostal =sheet.getRange("P2:P1000");
    let telefono =sheet.getRange("Q2:Q1000");
    let sitioWeb =sheet.getRange("R2:R1000");
    let email =sheet.getRange("S2:S1000");
    let editedCell = e.range;

    esCeldaEnRango(numIdentificacion,editedCell,undefined,e);
    esCeldaEnRango(contacto,editedCell,"string",e);
    esCeldaEnRango(codigoContacto,editedCell,undefined,e);
    esCeldaEnRango(primerNombre,editedCell,"string",e);
    esCeldaEnRango(segundoNombre,editedCell,"string",e);
    esCeldaEnRango(primeraApellido,editedCell,"string",e);
    esCeldaEnRango(segundoApellido,editedCell,"string",e);
    esCeldaEnRango(pais,editedCell,"string",e)
    esCeldaEnRango(provincia,editedCell,"string",e)
    esCeldaEnRango(poblacion,editedCell,"string",e)
    esCeldaEnRango(direccion,editedCell,"string",e)
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

