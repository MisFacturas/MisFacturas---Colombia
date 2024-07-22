PREFACTURA_ROW = 3;
PREFACTURA_COLUMN = 2;
COL_TOTALES_PREFACTURA = 11;// K
FILA_INICIAL_PREFACTURA = 8;
COLUMNA_FINAL = 50;
ADDITIONAL_ROWS = 3 + 3; //(Personalizacion)
var spreadsheet = SpreadsheetApp.getActive();
var prefactura_sheet = spreadsheet.getSheetByName('Factura2');
var unidades_sheet = spreadsheet.getSheetByName('Unidades');
var listadoestado_sheet = spreadsheet.getSheetByName('ListadoEstado');

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
      hojaFacturas.getRange("C3").setValue(listaConInformacion["Código cliente"]);
    }
  }




  // // Busca el contacto en la hoja de contactos
  // let rangoContactos = hojaContactos.getRange(2, 1, hojaContactos.getLastRow() - 1, hojaContactos.getLastColumn());
  // let valoresContactos = rangoContactos.getValues();

  // for (let i = 0; i < valoresContactos.length; i++) {
  //   if (valoresContactos[i][0] === nombreContacto) {
  //     let estadoContacto = valoresContactos[i][ultimaColumnaPermitida - 1];
  //     Logger.log(estadoContacto);
  //     if (estadoContacto === "Valido") {
  //       // Copia los datos de las columnas deseadas de manera vertical
  //       for (let j = 0; j < datosARetornar.length; j++) {
  //         let columna = hojaContactos.getRange(datosARetornar[j] + (i + 2)).getValue();
  //         Logger.log(datosARetornar[j] + (i + 2))
  //         hojaFacturas.getRange("B2").offset(j, 0).setValue(columna); //  aquí se puede ajustar la celda de inicio y el desplazamiento vertical
  //       }
  //     } else {
  //       // Muestra un mensaje si el contacto no es válido
  //       SpreadsheetApp.getUi().alert("Error: El contacto seleccionado no es válido.");
  //     }
  //     return;
  //   }
  // }
  
  // Si no se encuentra el contacto
  //SpreadsheetApp.getUi().alert("Error: El contacto seleccionado no se encontró.");
}


function generarNumeroFactura(sheet){
  let max=1000000;
  let min=1;
  let numero= Math.floor(Math.random() * (max - min + 1)) + min;
  sheet.getRange("G2").setValue(numero);
}

function obtenerFechaYHoraActual(sheet){ 
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
  var range = datos_sheet.getRange("B7");//Resolución Autorización
  var InvoiceAuthorizationNumber = range.getValue();
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
    "Note": getprefacturaValue(8, 3),//cambia los valroes parak llamar la nota de la factura 
    "ExternalGR": false
    //"AdditionalProperty": AdditionalProperty
  }


  return InvoiceGeneralInformation;
}
function getPaymentSummary(num_items, pfAnticipo) {
  var total_factura = prefactura_sheet.getRange(FILA_INICIAL_PREFACTURA + 10 + num_items, COL_TOTALES_PREFACTURA).getValue();// por ahora esto no lo utilizamos ya que no hay descuentos
  var monto_neto = prefactura_sheet.getRange("B23").getValue();
  //var numeros = new Intl.NumberFormat('es-CO', {maximumFractionDigits:0, style: 'currency', currency: 'COP' }).format(monto);
  var numeros_total = new Intl.NumberFormat().format(total_factura);
  var numeros_neto = new Intl.NumberFormat().format(monto_neto);
  //Browser.msgBox(`${numeros_total}  ${numeros_neto}`);
  if (pfAnticipo > 0)
    var PaymentNote = `Total Factura: $${numeros_total} \rSaldo Factura  $${numeros_neto}: ${int2word(monto_neto)}Pesos M/L`;
  else
    var PaymentNote = `Total Factura: $${numeros_total} \r Neto a Pagar  $${numeros_neto}: ${int2word(monto_neto)}Pesos M/L`;
  ;

  var PaymentTypeTxt = prefactura_sheet.getRange("F4").getValue();
  var PaymentMeansTxt = prefactura_sheet.getRange("E4").getValue();
  var PaymentSummary = {
    "PaymentType": "getPaymentType: No hay tipo de pago",
    "PaymentMeans": PaymentMeansTxt,//a qui habia getPaymentMeans(PaymentMeansTxt)
    "PaymentNote": `Total Factura: $${numeros_neto} \r Neto a Pagar  $${numeros_neto}: ${int2word(monto_neto)}`
  }
  return PaymentSummary;
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
  }while(i<(8+cantidadProductos));

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



  let rangeFacturaTotal=prefactura_sheet.getRange(20,1,1,4);// aqui cambia con respecto al original, aqui deberia de cambiar el segundo parametro creo, seria con respecto a un j el cual seria la cantidad de ivas que hay
  let facturaTotal=String(rangeFacturaTotal.getValues());
  facturaTotal=facturaTotal.split(",");
  Logger.log(facturaTotal)



  /*Aqui cambia por completo, por ahora solo voy a dejar los parametros en numeros x 
  ,  solo coinciden el base imponible he IVA */
  let pfSubTotal = parseFloat(facturaTotal[0]);//base imponible
  let pfIVA = parseFloat(facturaTotal[2]);//IVA
  let pfImpoconsumo = 22;
  let pfTotal = parseFloat(facturaTotal[3]);
  let pfRefuente = 0;
  let pfReteICA = 0;
  let pfReteIVA = 44;
  let pfTRetenciones = 33; 
  let pfAnticipo = 55;
  let pfTPagar = 66;

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
    "ChargeTotalAmount": 0,
    "PrePaidAmount": pfAnticipo,
    "PayableAmount": pfTotal // antes era (pfTotal - pfAnticipo) 
  }


  let cliente = prefactura_sheet.getRange("C2").getValue();
  let InvoiceGeneralInformation = getInvoiceGeneralInformation();
  let CustomerInformation = getCustomerInformation(cliente);// tal ves que por ahora no llame al cliente

  let invoice = JSON.stringify({
    CustomerInformation: CustomerInformation,
    InvoiceGeneralInformation: InvoiceGeneralInformation,
    Delivery: getDelivery(),
    AdditionalDocuments: getAdditionalDocuments(),
    AdditionalProperty: getAdditionalProperty(),
    PaymentSummary: getPaymentSummary(0,0), //por ahora esto leugo se cambia la funcion getPaymentSummary para que cumpla los parametros
    ItemInformation: productoInformation,
    //Invoice_Note: invoice_note,
    InvoiceTaxTotal: invoiceTaxTotal,
    InvoiceAllowanceCharge: [],
    InvoiceTotal: invoice_total
  });
  Logger.log(invoice)

  let nameString = prefactura_sheet.getRange("C2").getValue();
  let numeroFactura = JSON.stringify(InvoiceGeneralInformation.InvoiceNumber);
  let fecha = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");
  listadoestado_sheet.appendRow(["vacio", "vacio","vacio" , fecha,"vacio" ,numeroFactura ,nameString , "falta","vacio" ,"vacio" ,"representacion" ,"Vacio", String(invoice)]);
  
  
}

function guardarInvoice(invoice){

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
          var formaPago = invoiceData.PaymentSummary.PaymentMeans;
          var listaProductos = invoiceData.ItemInformation;
          var numeroProductos = 0;
          var valorPagar = invoiceData.PaymentSummary.PaymentNote;
          var notaPago = invoiceData.PaymentSummary.PaymentNote;
          var observaciones = invoiceData.InvoiceGeneralInformation.Note;

          var filasInsertadas = 0;
          var grupoIva = {};

          var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Plantilla'); // Hoja donde quieres insertar el NIF
          if (!targetSheet) {
            targetSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Plantilla');
          }
          
          for (var j = 0; j < listaProductos.length; j++) {
            numeroProductos += 1;
            var numeroCelda = 22 + j;
            if (numeroProductos > 1) {
              targetSheet.insertRowAfter(numeroCelda);
              filasInsertadas += 1;
            }
            var celdaItem = targetSheet.getRange('C'+numeroCelda);
            celdaItem.setBorder(true,true,true,true,null,null,null,null);
            celdaItem.setValue(numeroProductos);

            var celdaReferencia = targetSheet.getRange('D'+numeroCelda);
            celdaReferencia.setBorder(true,true,true,true,null,null,null,null);
            celdaReferencia.setValue(listaProductos[j].ItemReference);

            var celdaDespricion = targetSheet.getRange('E'+numeroCelda);
            celdaDespricion.setBorder(true,true,true,true,null,null,null,null);
            celdaDespricion.setValue(listaProductos[j].Name);
            
            var celdaCantidad = targetSheet.getRange('F'+numeroCelda);
            celdaCantidad.setBorder(true,true,true,true,null,null,null,null);
            celdaCantidad.setValue(listaProductos[j].Quatity);
            
            var celdaPrecioUnitario = targetSheet.getRange('G'+numeroCelda);
            celdaPrecioUnitario.setBorder(true,true,true,true,null,null,null,null);
            celdaPrecioUnitario.setValue(listaProductos[j].Price);
            
            var celdaIva = targetSheet.getRange('H'+numeroCelda);
            celdaIva.setBorder(true,true,true,true,null,null,null,null);
            celdaIva.setValue((listaProductos[j].TaxesInformation[0].Percent)/100);
            celdaIVA.setNumberFormat('0.0%');

            
            var celdaImporte = targetSheet.getRange('I'+numeroCelda);
            celdaImporte.setBorder(true,true,true,true,null,null,null,null);
            celdaImporte.setValue(listaProductos[j].LineExtensionAmount);

            var producto = listaProductos[j]
            //crea un diccionario que la llave sea el % de iva y el valor sea el total de la linea
            
            if (grupoIva.hasOwnProperty(producto.TaxesInformation[0].Percent)) {
              grupoIva[producto.TaxesInformation[0].Percent] += producto.TaxesInformation[0].TaxableAmount;
            } else {
              grupoIva[producto.TaxesInformation[0].Percent] = producto.TaxesInformation[0].TaxableAmount;
            }
          }
          var contador = 0;
          for (var key in grupoIva) {
            if (grupoIva.hasOwnProperty(key)) {
              var numeroCelda = 30 + filasInsertadas;
              if (contador > 0) {
                targetSheet.insertRowAfter(numeroCelda);
                filasInsertadas += 1;
              }
              var celdaBaseImponible = targetSheet.getRange('C'+numeroCelda);
              celdaBaseImponible.setBorder(true,true,true,true,null,null,null,null);
              celdaBaseImponible.setValue(grupoIva[key]);
              
              var celdaPorcentajeIva = targetSheet.getRange('E'+numeroCelda);
              celdaPorcentajeIva.setBorder(true,true,true,true,null,null,null,null);
              celdaPorcentajeIva.setValue(key/100);
              celdaPorcentajeIva.setNumberFormat('0.0%');
              
              var celdaIVA = targetSheet.getRange('G'+numeroCelda);
              celdaIVA.setBorder(true,true,true,true,null,null,null,null);
              celdaIVA.setFormula('=C'+numeroCelda+'*E'+numeroCelda);
              
              var celdaTotal = targetSheet.getRange('I'+numeroCelda);
              celdaTotal.setBorder(true,true,true,true,null,null,null,null);
              celdaTotal.setFormula('=C'+numeroCelda+'+G'+numeroCelda);

              contador += 1;
            }
          }


          var clienteCell = targetSheet.getRange('C12');
          var nifCell = targetSheet.getRange('C13');
          var codigoCell = targetSheet.getRange('C14');
          var direccionCell = targetSheet.getRange('C15');
          var telefonoCell = targetSheet.getRange('C16');
          var poblacionCell = targetSheet.getRange('C17');
          var provinciaCell = targetSheet.getRange('C18');
          var paisCell = targetSheet.getRange('C19');
          var fechaEmisionCell = targetSheet.getRange('H12');
          var formaPagoCell = targetSheet.getRange('H13');
          var valorPagarCell = targetSheet.getRange('C'+(36+filasInsertadas));
          var notaPagoCell = targetSheet.getRange('B'+(41+filasInsertadas));
          var observacionesCell = targetSheet.getRange('B'+(47+filasInsertadas));
          var totalItemsCell = targetSheet.getRange('C'+(24+filasInsertadas));
          var descuentosCell = targetSheet.getRange('C'+(34+filasInsertadas));
          var cargosCell = targetSheet.getRange('E'+(34+filasInsertadas));


          clienteCell.setValue(cliente);
          nifCell.setValue(nif);
          codigoCell.setValue(codigo);
          direccionCell.setValue(direccion);
          telefonoCell.setValue(telefono);
          poblacionCell.setValue(poblacion);
          provinciaCell.setValue(provincia);
          paisCell.setValue(pais);
          fechaEmisionCell.setValue(fechaEmision);
          formaPagoCell.setValue(formaPago);
          valorPagarCell.setValue(valorPagar);
          notaPagoCell.setValue(notaPago);
          observacionesCell.setValue(observaciones);
          totalItemsCell.setValue(numeroProductos);
          descuentosCell.setValue(0);
          cargosCell.setValue(0);
          
          Logger.log(grupoIva);
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
  var invoiceNumber = 'FE947'; // Reemplaza con el número de factura deseado
  obtenerDatosFactura(invoiceNumber);
}

function resetPlantilla() {
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Plantilla');
  
  // Borrar información del cliente
  targetSheet.getRange('C12').clearContent();
  targetSheet.getRange('C13').clearContent();
  targetSheet.getRange('C14').clearContent();
  targetSheet.getRange('C15').clearContent();
  targetSheet.getRange('C16').clearContent();
  targetSheet.getRange('C17').clearContent();
  targetSheet.getRange('C18').clearContent();
  targetSheet.getRange('C19').clearContent();
  targetSheet.getRange('H12').clearContent();
  targetSheet.getRange('H13').clearContent();
  
  // Borrar valor a pagar, nota de pago y observaciones
  targetSheet.getRange('C36').clearContent();
  targetSheet.getRange('B41').clearContent();
  targetSheet.getRange('B47').clearContent();
  
  // Borrar total de items, descuentos y cargos
  targetSheet.getRange('C24').clearContent();
  targetSheet.getRange('C34').clearContent();
  targetSheet.getRange('E34').clearContent();
  
  // Borrar productos y reestablecer filas insertadas
  for (var i = 22; i < targetSheet.getLastRow(); i++) {
    var rowRange = targetSheet.getRange(i, 3, 1, 7); // Columnas C a I
    rowRange.clearContent();
    rowRange.setBorder(false, false, false, false, null, null, null, null);
  }
  
  // Borrar bases imponibles, porcentajes de IVA, IVA y totales
  for (var j = 30; j < targetSheet.getLastRow(); j++) {
    var ivaRowRange = targetSheet.getRange(j, 3, 1, 7); // Columnas C a I
    ivaRowRange.clearContent();
    ivaRowRange.setBorder(false, false, false, false, null, null, null, null);
  }
  
  // Eliminar filas adicionales que se hayan insertado
  var originalLastRow = 50; // Ajusta este valor según el número de filas original en la hoja Plantilla
  var currentLastRow = targetSheet.getLastRow();
  if (currentLastRow > originalLastRow) {
    targetSheet.deleteRows(originalLastRow + 1, currentLastRow - originalLastRow);
  }
}
