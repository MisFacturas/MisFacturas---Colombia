PREFACTURA_ROW = 3;
PREFACTURA_COLUMN = 2;
COL_TOTALES_PREFACTURA = 11;// K
FILA_INICIAL_PREFACTURA = 8;
COLUMNA_FINAL = 50;
ADDITIONAL_ROWS = 3 + 3; //(Personalizacion)
var spreadsheet = SpreadsheetApp.getActive();
var prefactura_sheet = spreadsheet.getSheetByName('Factura');
var unidades_sheet = spreadsheet.getSheetByName('Unidades');

function verificarYCopiarContacto(e) {
  let hojaFacturas = e.source.getSheetByName('Factura');
  let hojaContactos = e.source.getSheetByName('Clientes');
  let celdaEditada = e.range;

  // aqui se define las celdas en la hoja de facturas donde van a poner los datos ojo que importa el orden
  let celdasDestino = ["C1", "D1","E1","B2","C2"];

  let nombreContacto = celdaEditada.getValue();
  let ultimaColumnaPermitida = 18; // Columna del estado en la hoja de contactos
  let datosARetornar = ["C", "D", "O", "M", "L"]; // Columnas que quiero de la hoja de contactos

  // Busca el contacto en la hoja de contactos
  let rangoContactos = hojaContactos.getRange(2, 1, hojaContactos.getLastRow() - 1, hojaContactos.getLastColumn());
  let valoresContactos = rangoContactos.getValues();

  for (let i = 0; i < valoresContactos.length; i++) {
    if (valoresContactos[i][0] === nombreContacto) {
      let estadoContacto = valoresContactos[i][ultimaColumnaPermitida - 1];
      Logger.log(estadoContacto);
      if (estadoContacto === "Valido") {
        // Copia los datos de las columnas deseadas en las celdas especificas
        for (let j = 0; j < datosARetornar.length; j++) {
          let columnaIndex = hojaContactos.getRange(datosARetornar[j] + (i + 2)).getColumn();
          let valor = hojaContactos.getRange(i + 2, columnaIndex).getValue();
          hojaFacturas.getRange(celdasDestino[j]).setValue(valor);
        }
      } else {
        // no es valido
        SpreadsheetApp.getUi().alert("Error: El contacto seleccionado no es válido.");
      }
      return;
    }
  }

  // Si no se encuentra el contacto
  SpreadsheetApp.getUi().alert("Error: El contacto seleccionado no se encontró.");
}


function generarNumeroFactura(sheet){
  let max=1000000;
  let min=1;
  let numero= Math.floor(Math.random() * (max - min + 1)) + min;
  sheet.getRange(1,5).setValue(numero);
}

function obtenerFechaYHoraActual(sheet){ 
  let fecha = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");
  let hora= Utilities.formatDate(new Date(), "GMT+1", "HH:mm:ss");

  sheet.getRange(2,5).setValue(fecha)
  sheet.getRange(3,5).setValue(hora)
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
  var range = datos_sheet.getRange("B7");
  var InvoiceAuthorizationNumber = range.getValue();
  //
  range = prefactura_sheet.getRange("F5");//dias de vencimiento
  var DaysOff = range.getValue();

  var invoice_number = getprefacturaValue(1, 6);//cambiamos los valores para llamar el numero de factura
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
    "Note": getprefacturaValue(8, 6),//cambia los valroes parak llamar la nota de la factura 
    "ExternalGR": false
    //"AdditionalProperty": AdditionalProperty
  }


  return InvoiceGeneralInformation;
}
function getPaymentSummary(num_items, pfAnticipo) {
  var total_factura = prefactura_sheet.getRange(FILA_INICIAL_PREFACTURA + 10 + num_items, COL_TOTALES_PREFACTURA).getValue();
  var monto_neto = prefactura_sheet.getRange(FILA_INICIAL_PREFACTURA + 16 + num_items, COL_TOTALES_PREFACTURA).getValue();
  //var numeros = new Intl.NumberFormat('es-CO', {maximumFractionDigits:0, style: 'currency', currency: 'COP' }).format(monto);
  var numeros_total = new Intl.NumberFormat().format(total_factura);
  var numeros_neto = new Intl.NumberFormat().format(monto_neto);
  //Browser.msgBox(`${numeros_total}  ${numeros_neto}`);
  if (pfAnticipo > 0)
    var PaymentNote = `Total Factura: $${numeros_total} \rSaldo Factura  $${numeros_neto}: ${int2word(monto_neto)}Pesos M/L`;
  else
    var PaymentNote = `Total Factura: $${numeros_total} \r Neto a Pagar  $${numeros_neto}: ${int2word(monto_neto)}Pesos M/L`;
  ;

  var PaymentTypeTxt = prefactura_sheet.getRange("C5").getValue();
  var PaymentMeansTxt = prefactura_sheet.getRange("B5").getValue();
  var PaymentSummary = {
    "PaymentType": getPaymentType(PaymentTypeTxt),
    "PaymentMeans": getPaymentMeans(PaymentMeansTxt),
    "PaymentNote": `Total Factura: $${numeros_total} \r Neto a Pagar  $${numeros_neto}: ${int2word(monto_neto)}Pesos M/L`
  }
  return PaymentSummary;
}
function sendInvoice() {


  // 	https://misfacturas.cenet.ws/integrationAPI_2/api/InsertInvoice?SchemaID=31&IDNumber=800176901&TemplateID=73


  var item_information = [];
  range = prefactura_sheet.getRange("C3");
  var num_items = range.getValue();


  var invoice_tax_total = [];
  //Iva Impocoinsumo InvoiceTaxTotal Retefuente ReteICA ReteIVA
  var invoice_taxes = ['Iva', 'Impocoinsumo', 'Retefuente', 'ReteICA', 'ReteIVA'];

  //Creacion de keys: arreglo para el diccionario de LineaFactura
  var keys_range_str = "A" + String(FILA_INICIAL_PREFACTURA - 1) + ":K" + String(FILA_INICIAL_PREFACTURA - 1);
  Logger.log(keys_range_str)
  var range_keys = prefactura_sheet.getRange(keys_range_str);
  var key_list = slugifyF(range_keys.getValues()).replace(/\s/g, '');
  Logger.log(key_list);
  var keys = key_list.split(',');
  //Logger.log(keys);//[producto, codigoitem, unidad, cantidad, preciounitario, %descuento, subtotallinea, iva, impoconsumo, retefuente, totallinea]

  //var InvoiceTaxTotal = [];  
  function addTaxToInvoice(itt, tax) {//invoice total tax
    Browser.msgBox(itt[0].Id);
    var indice = itt.indexOf(t => (t.Id == tax.Id) && (t.Percent == tax.Percent));
    Browser.msgBox(indice);
    if (indice == -1)
      InvoiceTaxTotal.push(tax);
    else
      InvoiceTaxTotal[indice].TaxableAmount += tax.TaxableAmount
  };

  var i = FILA_INICIAL_PREFACTURA;
  do {
    var fila = "A" + String(i) + ":K" + String(i);
    range = prefactura_sheet.getRange(fila);
    var list_item = String(range.getValues());
    //Browser.msgBox(list_item);
    // ABC Only: 
    //Producto	Código Item	Unidad	Cantidad	SIN IVA	CON IVA	Sub Total Linea	IVA	ImpoConsumo	ReteFuente	Total Linea
    //Producto	Código Item	Unidad	Cantidad	Precio Unitario	% Descuento	Sub Total Linea	IVA	ImpoConsumo	ReteFuente	Total Linea
    //[Producto No 2 (Con IVA), 2, Unidad, 1, 1000, 0, 1000, 190, 0, 0, 1190]

    arr = list_item.split(',');
    var LineaFactura = {};
    for (var j = 0; j < COL_TOTALES_PREFACTURA; j++) {
      LineaFactura[keys[j]] = arr[j];
    }

    var Name = LineaFactura['producto'];
    var ItemCode = new Number(LineaFactura['codigoitem']);
    var MeasureUnitCode = getMeasureUnitCode(LineaFactura['unidad']);
    //var FreeOFChargeIndicator = false;//True si el ítem es un regalo que no genera contraprestación y por ende no es una venta. False si no es un regalo.
    var Quantity = LineaFactura['cantidad'];
    var Price = LineaFactura['siniva'];
    //var coniva = LineaFactura['coniva'];
    var Amount = parseFloat(LineaFactura['subtotallinea']);
    var Iva = parseFloat(LineaFactura['iva']);
    var ImpoConsumo = parseFloat(LineaFactura['impoconsumo']);
    //var ReteFuente = parseFloat(LineaFactura['retefuente']);
    var LineChargeTotal = parseFloat(LineaFactura['totallinea']);




    ItemTaxesInformation = [];
    //IVA
    var Percent = parseFloat(((Iva / Amount) * 100).toFixed(1));
    //Browser.msgBox(Percent);
    var iva_taxinformation = {
      Id: "01",//Id,
      TaxEvidenceIndicator: false,
      TaxableAmount: Amount,
      TaxAmount: Iva,
      Percent: Percent,
      BaseUnitMeasure: "",
      PerUnitAmount: ""
    };
    ItemTaxesInformation.push(iva_taxinformation);
    invoice_tax_total.push(iva_taxinformation);

    var LineExtensionAmount = Amount;
    var LineTotalTaxes = Iva + ImpoConsumo;//+ ReteFuente; //?ReteFuente?
    //var LineTotal = LineChargeTotal;//new Number(Quantity * Price + LineTotalTaxes);

    var item_i = {
      ItemReference: ItemCode,
      Name: Name,
      Quatity: new Number(Quantity),
      Price: new Number(Price),
      LineAllowanceTotal: 0.0,
      LineChargeTotal: 0.0,
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
    item_information.push(item_i);
    //Browser.msgBox(JSON.stringify (item_information));
    i++;
  } while (i < (FILA_INICIAL_PREFACTURA + num_items));

  //prefactura_sheet.getRange(FILA_INICIAL_PREFACTURA + 10 + num_items, COL_TOTALES_PREFACTURA - 1 ).getValue();

  range = prefactura_sheet.getRange(FILA_INICIAL_PREFACTURA + num_items + 1 + ADDITIONAL_ROWS, COL_TOTALES_PREFACTURA, 10, 1);
  var list_pfTotales = String(range.getValues());
  var array_pfTotales = list_pfTotales.split(',');

  var pfSubTotal = parseFloat(array_pfTotales[0]);
  var pfIVA = parseFloat(array_pfTotales[1]);
  var pfImpoconsumo = parseFloat(array_pfTotales[2]);
  var pfTotal = parseFloat(array_pfTotales[3]);
  var pfRefuente = parseFloat(array_pfTotales[4]);
  var pfReteICA = parseFloat(array_pfTotales[5]);
  var pfReteIVA = parseFloat(array_pfTotales[6]);
  var pfTRetenciones = parseFloat(array_pfTotales[7]);
  var pfAnticipo = parseFloat(array_pfTotales[8]);
  var pfTPagar = parseFloat(array_pfTotales[9]);
  /*
  if (pfIVA > 0){// Improvement: IndexOf?
    //Browser.msgBox('IVA hay')
    //invoice_tax_total.push(invoice_iva);
  };
    
  if (pfImpoconsumo > 0)
    invoice_tax_total.push(invoice_impoconsumo);
 
    */

  if (pfRefuente > 0) {
    var Percent = parseFloat((pfRefuente / pfSubTotal * 100).toFixed(2));
    //Browser.msgBox(`Hay ReteFte%: ${Percent}`);
    var retefuente_taxinformation = {
      Id: "06",//Id,
      TaxEvidenceIndicator: true,
      TaxableAmount: pfSubTotal,
      TaxAmount: pfRefuente,
      Percent: Percent,
      BaseUnitMeasure: "",
      PerUnitAmount: ""
    };
    invoice_tax_total.push(retefuente_taxinformation);
  };

  if (pfReteICA > 0) {
    var Factor = datos_sheet.getRange("B8").getValue();
    var PercentReteICA = (Factor * 100).toFixed(3);
    //Browser.msgBox('ReteICA ' + PercentReteICA);
    var invoice_ReteICA = {
      Id: "07",//Id,
      TaxEvidenceIndicator: true,
      TaxableAmount: pfSubTotal,
      TaxAmount: pfReteICA,
      Percent: parseFloat(PercentReteICA),
      BaseUnitMeasure: "",
      PerUnitAmount: ""
    };
    invoice_tax_total.push(invoice_ReteICA);
  }

  if (pfReteIVA > 0) {
    var FactorReteIva = pfReteIVA / pfSubTotal;
    var PercentReteIVA = (FactorReteIva * 100).toFixed(2);
    //Browser.msgBox(`Hay ReteIVA sobre ${pfSubTotal} de ${pfReteIVA} es decir ${PercentReteIVA}% `);
    var invoice_reteIVA = {
      Id: "05",
      TaxEvidenceIndicator: true,
      TaxableAmount: pfSubTotal,
      TaxAmount: pfReteIVA,
      Percent: parseFloat(PercentReteIVA),
      BaseUnitMeasure: "",
      PerUnitAmount: ""
    };
    invoice_tax_total.push(invoice_reteIVA);
  }

  texto = getprefacturaValue(2, 4);
  var invoice_note = {
    "Note": texto
  };

  var invoice_total = {
    "LineExtensionAmount": pfSubTotal,
    "TaxExclusiveAmount": pfSubTotal,
    "TaxInclusiveAmount": pfTotal,
    "AllowanceTotalAmount": 0,
    "ChargeTotalAmount": 0,
    "PrePaidAmount": pfAnticipo,
    "PayableAmount": (pfTotal - pfAnticipo)
  }

  var customer = prefactura_sheet.getRange("B1").getValue();
  var InvoiceGeneralInformation = getInvoiceGeneralInformation();
  var CustomerInformation = getCustomerInformation(customer);

  var invoice = JSON.stringify({
    CustomerInformation: CustomerInformation,
    InvoiceGeneralInformation: InvoiceGeneralInformation,
    Delivery: getDelivery(),
    AdditionalDocuments: getAdditionalDocuments(),
    AdditionalProperty: getAdditionalProperty(),
    PaymentSummary: getPaymentSummary(num_items),
    ItemInformation: item_information,
    //Invoice_Note: invoice_note,
    InvoiceTaxTotal: invoice_tax_total,
    InvoiceAllowanceCharge: [],
    InvoiceTotal: invoice_total
  });


  range = datos_sheet.getRange("B11");
  var token = range.getValue().slice(1, -1);
  //Logger.log(token);
  var authorization = 'misfacturas ' + token;
  Logger.log(invoice);//Browser.msgBox(invoice);

  var headers = {
    'Content-Type': 'application/json',
    //Access-Control-Allow-Origin': HOST,
    'Authorization': authorization
  };

  var options = {
    'muteHttpExceptions': true,//default:false
    'method': 'post',
    'headers': headers,
    'payload': invoice
  };
  range = datos_sheet.getRange("C1");
  var ambiente = range.getValue();
  var numeroFactura = JSON.stringify(InvoiceGeneralInformation.InvoiceNumber);
  var stringToShowHeader = `Emisión Factura ${numeroFactura} en ${ambiente}`;
  //var nameString = JSON.stringify(CustomerInformation.RegistrationName).slice(1, -1);
  var nameString = prefactura_sheet.getRange("B1").getValue();
  var emailString = JSON.stringify(CustomerInformation.Email).slice(1, -1);
  var stringToShow2 = `Cliente: ${nameString}\nEmail:${emailString}`;
  var responseDialog = Browser.msgBox(`${stringToShowHeader}`, stringToShow2, Browser.Buttons.OK_CANCEL);
  /*
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert(
    stringToShowHeader,
    stringToShow2,
    ui.ButtonSet.YES_NO);
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    ui.alert('Confirmation received.');
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Permission denied.');
  };
  */
  if (responseDialog == "ok") {
    Logger.log(`User clicked OK: Cliente: Invoice ${numeroFactura} to be issued`);
  } else {
    Logger.log('The user clicked "Cancel" or the dialog\'s close button.');
    return;
  }

  var response = UrlFetchApp.fetch(url, options);
  //Browser.msgBox(response.getResponseCode());
  //Browser.msgBox(response.getContent());


  var id_documento = response.getContentText().substring(15, 51);

  if (id_documento == '')
    if (response.getResponseCode() == 200) {
      Browser.msgBox("Factura No Emitida: Renovar Token");
      return;
    }


  switch (response.getResponseCode()) {
    case 200:
      Browser.msgBox("200: Documento Enviado: " + id_documento);
      break;       
    case 400:
      Browser.msgBox("400: " + response.getContentText());
      Logger.log(response.getContentText());
      return;
      break;
    default:
      Browser.msgBox("Error " + String(response.getResponseCode()));
      return;
      break;
  }

  //actualizacion Numero PreInvoice Invoice
  updateprefacturaValue(PREFACTURA_ROW, PREFACTURA_COLUMN, getprefacturaValue(PREFACTURA_ROW, PREFACTURA_COLUMN) + 1);




  //var factura = SpreadsheetApp.openById(response);

  Logger.log('documento: ' + id_documento);
  //listadoestado_sheet.appendRow([id_documento]);
  //var factura_sheet= SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet();
  //factura_sheet.setName(id_documento);

  // https://developers.google.com/apps-script/reference/spreadsheet/sheet#getDataRange()
  /*var numRows = prefactura_sheet.getLastRow();
  var numColumns = prefactura_sheet.getLastColumn();
  Browser.msgBox(`${numRows} ${numColumns}`);*/

  //Toma de valores personalizados
  var i = 20; // Indice en la hoja de personalizacion
  var celda_ultimo_item = FILA_INICIAL_PREFACTURA + num_items - 1;
  var j = celda_ultimo_item + 2 + ADDITIONAL_ROWS; // Al nivel de SubTotal: Inicio Campos Personalizados
  var lastRow_personalizacion = personalizacion_sheet.getLastRow();
  var count_elementos_personalizacion = lastRow_personalizacion - i + 1;
  if (count_elementos_personalizacion > 0) {
    range = prefactura_sheet.getRange(j, 1, count_elementos_personalizacion, 2);
    var values = range.getValues();
    var campos_personales = "";
    for (var row in values) {
      for (var col in values[row]) {
        //Browser.msgBox(values[row]);
        campos_personales += values[row][col];
        campos_personales += ',';
      }
    }
    //Browser.msgBox(values[0]);

  }

  listadoestado_sheet.appendRow([id_documento, campos_personales.slice(0, -1), , , , , , , , , , datos_sheet.getRange("C1").getValue(), String(invoice)]);


  return;
}

function guardarYGenerarInvoice(){
  const cantidadProductos = prefactura_sheet.getRange("I4").getValue(); // cantidad total de productos 
  let llavesParaLinea=prefactura_sheet.getRange("H7:N7");//llamo los headers 
  llavesParaLinea = slugifyF(llavesParaLinea.getValues()).replace(/\s/g, ''); // Todo en una sola linea
  const llavesFinales =llavesParaLinea.split(",");
  /* Creo que esto se puede cambiar a una manera mas simple, ya que los headers de la fila H7 hatsa N7 nunca van a cambiar */

  let invoiceTaxTotal=[]
  var productoInformation = [];

  let i = 8 // es 8 debido a que aqui empieza los productos elegidos por el cliente
  do{
    let filaActual = "H" + String(i) + ":N" + String(i);
    Logger.log(filaActual)
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
    let Iva = 1 //aqui toca calcular o traer el iva desde producto, por ahora en 1
    let ImpoConsumo = 0// no esta ni en el original ni aca
    let LineChargeTotal = parseFloat(LineaFactura['totaldelinea']);


    //IVA
    let ItemTaxesInformation = [];
    let Percent = 44; //aqui deberia de calcular el porcentaje pero como todavia no tengo IVA solo por ahora no
    let ivaTaxInformation = {
      Id: "01",//Id
      TaxEvidenceIndicator: false,
      TaxableAmount: Amount,
      TaxAmount: Iva,
      Percent: Percent,
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
      LineChargeTotal: 0.0,
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
  }while(i<(8+cantidadProductos));

  /* Aqui empieza el proceso de coger el precio total de la facutra OJO en nuestro caso se agrupan por % de iva, entonces cambia
  algo mucho */
  

  let rangeFacturaTotal=prefactura_sheet.getRange(13,1,1,4);// aqui cambia con respecto al original
  let facturaTotal=String(rangeFacturaTotal.getValues());
  facturaTotal=facturaTotal.split(",");


  /*Aqui cambia por completo, por ahora solo voy a dejar los parametros en numeros x 
  ,  solo coinciden el base imponible he IVA */
  let pfSubTotal = parseFloat(facturaTotal[0]);//base imponible
  let pfIVA = parseFloat(facturaTotal[2]);//IVA
  let pfImpoconsumo = 22;
  let pfTotal = 22;
  let pfRefuente = 0;
  let pfReteICA = 0;
  let pfReteIVA = 44;
  let pfTRetenciones = 33; 
  let pfAnticipo = 55;
  let pfTPagar = 66;

  if (pfRefuente > 0) {
    let Percent = parseFloat((pfRefuente / pfSubTotal * 100).toFixed(2));
    let retefuente_taxinformation = {
      Id: "06",//Id,
      TaxEvidenceIndicator: true,
      TaxableAmount: pfSubTotal,
      TaxAmount: pfRefuente,
      Percent: Percent,
      BaseUnitMeasure: "",
      PerUnitAmount: ""
    };
    invoiceTaxTotal.push(retefuente_taxinformation);
  };

  if (pfReteICA > 0) {
    let Factor = datos_sheet.getRange("B8").getValue();
    let PercentReteICA = (Factor * 100).toFixed(3);
    let invoice_ReteICA = {
      Id: "07",//Id,
      TaxEvidenceIndicator: true,
      TaxableAmount: pfSubTotal,
      TaxAmount: pfReteICA,
      Percent: parseFloat(PercentReteICA),
      BaseUnitMeasure: "",
      PerUnitAmount: ""
    };
    invoiceTaxTotal.push(invoice_ReteICA);
  }

  if (pfReteIVA > 0) {
    let FactorReteIva = pfReteIVA / pfSubTotal;
    let PercentReteIVA = (FactorReteIva * 100).toFixed(2);
    let invoice_reteIVA = {
      Id: "05",
      TaxEvidenceIndicator: true,
      TaxableAmount: pfSubTotal,
      TaxAmount: pfReteIVA,
      Percent: parseFloat(PercentReteIVA),
      BaseUnitMeasure: "",
      PerUnitAmount: ""
    };
    invoiceTaxTotal.push(invoice_reteIVA);
  }

  //Aqui seguiria el texto, pero en el de carlos nunca lo llama 

  let invoice_total = {
    "LineExtensionAmount": pfSubTotal,
    "TaxExclusiveAmount": pfSubTotal,
    "TaxInclusiveAmount": pfTotal,
    "AllowanceTotalAmount": 0,
    "ChargeTotalAmount": 0,
    "PrePaidAmount": pfAnticipo,
    "PayableAmount": (pfTotal - pfAnticipo)
  }


  let cliente = prefactura_sheet.getRange("B1").getValue();
  let InvoiceGeneralInformation = getInvoiceGeneralInformation();
  let CustomerInformation = getCustomerInformation(cliente);// tal ves que por ahora no llame al cliente

  let invoice = JSON.stringify({
    CustomerInformation: CustomerInformation,
    InvoiceGeneralInformation: InvoiceGeneralInformation,
    Delivery: getDelivery(),
    AdditionalDocuments: getAdditionalDocuments(),
    AdditionalProperty: getAdditionalProperty(),
    PaymentSummary: 44444, //por ahora esto leugo se cambia la funcion getPaymentSummary para que cumpla los parametros
    ItemInformation: productoInformation,
    //Invoice_Note: invoice_note,
    InvoiceTaxTotal: invoiceTaxTotal,
    InvoiceAllowanceCharge: [],
    InvoiceTotal: invoice_total
  });
  //merge con main
  Logger.log(invoice)

  
  
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
          
          var cliente = invoiceData.CustomerInformation.RegistrationName;
          var nif = invoiceData.CustomerInformation.Identification;
          var codigo = invoiceData.CustomerInformation.AdditionalAccountID;
          var direccion = invoiceData.CustomerInformation.AddressLine;
          var telefono = invoiceData.CustomerInformation.Telephone;
          var poblacion = invoiceData.CustomerInformation.CityName;
          var provincia = invoiceData.CustomerInformation.SubdivisionName;
          var pais = invoiceData.CustomerInformation.CountryName;
          var fechaEmision = invoiceData.Delivery.DeliveryDate;
          //var formaPago = invoiceData.


          
          var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Plantilla'); // Hoja donde quieres insertar el NIF
          if (!targetSheet) {
            targetSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Plantilla');
          }
          
          var clienteCell = targetSheet.getRange('C12');
          var nifCell = targetSheet.getRange('C13');
          var codigoCell = targetSheet.getRange('C14');
          var direccionCell = targetSheet.getRange('C15');

          clienteCell.setValue(cliente);
          nifCell.setValue(nif);
          codigoCell.setValue(codigo);
          direccionCell.setValue(direccion);
          
          Logger.log(`NIF written for invoice ${factura} at row ${i + 1}`);
          return;
        } catch (e) {
          Logger.log('Error parsing JSON for row ' + (i + 1) + ': ' + e.message);
        }
      }
    }
  }
  Logger.log('Invoice number ' + factura + ' not found.');
}

function testWriteNIFToPlantilla() {
  var invoiceNumber = 'FE946'; // Reemplaza con el número de factura deseado
  writeNIFToPlantilla(invoiceNumber);
}