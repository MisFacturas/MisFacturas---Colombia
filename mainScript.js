var spreadsheet = SpreadsheetApp.getActive();

function onOpen() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();


  ui.createAddonMenu()
    .addItem('Inicio', 'showSidebar')
    .addToUi();

  console.log("setActiveSheet Inicio");

  var sheet = ss.getSheetByName("Inicio");
  SpreadsheetApp.setActiveSheet(sheet);

  console.log("onOpenReturning");
  return;

}

function pruebaLogo() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Datos de emisor");
  var celdaLogo = hoja.getRange("B12").getValue();
  hoja.getRange("B20").setValue(celdaLogo);
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('main')
    .setTitle('Menú');
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

function openDatosEmisorSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Datos de emisor");
  SpreadsheetApp.setActiveSheet(sheet);
}
function showConfiguracion() {
  var html = HtmlService.createHtmlOutputFromFile('menuConfiguracion')
    .setTitle('Datos emisor');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function showDesvincular() {
  var html = HtmlService.createHtmlOutputFromFile('menuDesvincular')
    .setTitle('Desvincular cuenta');
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

function showNuevaFactura() {
  var html = HtmlService.createHtmlOutputFromFile('nuevaFactura').setTitle("Nueva factura")
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function showAgregarProdcuto() {
  var html = HtmlService.createHtmlOutputFromFile('menuAgregarProducto').setTitle("Agregar Producto")
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

function showVincularCuenta() {
  var html = HtmlService.createHtmlOutputFromFile('menuVincular')
    .setTitle('Vincular cuenta');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function showEliminarInfo() {
  var html = HtmlService.createHtmlOutputFromFile('menuEliminarInfo')
    .setTitle('Eliminar informacion');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}


function eliminarTotalidadInformacion() {
  let hojaDatos = spreadsheet.getSheetByName('Datos');
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let hojaProductos = spreadsheet.getSheetByName('Productos');
  let hojaClientes = spreadsheet.getSheetByName("Clientes");
  let hojaListadoEstado = spreadsheet.getSheetByName('ListadoEstado');
  let ClientesInvalidos = spreadsheet.getSheetByName('ClientesInvalidos');


  borrarInfoHoja(hojaProductos)
  borrarInfoHoja(hojaClientes)
  borrarInfoHoja(hojaListadoEstado)
  borrarInfoHoja(ClientesInvalidos)
  borrarInfoHoja(hojaDatosEmisor)
  hojaDatos.getRange("Q11").setValue(0)
  hojaDatos.getRange("F47").setValue("")
  hojaDatosEmisor.getRange("B13").setBackground('#FFC7C7')
  hojaDatosEmisor.getRange("B13").setValue("Desvinculado")
  SpreadsheetApp.getUi().alert('Informacion eliminada correctamente');


}
function borrarInfoHoja(hoja) {
  let lastrow = Number(hoja.getLastRow())
  let nombreHoja = hoja.getSheetName()
  Logger.log("borrarInfoHoja")
  Logger.log("nombreHoja " + nombreHoja)
  Logger.log("lastrow " + lastrow)
  if (nombreHoja === "Datos de emisor") {
    Logger.log("Hoja es datos emisor ")
    hoja.getRange(1, 2, 12).setValue("")
    for (let i = 18; i < 50; i++) {
      hoja.getRange(i, 1, 1, 6).setValue("")
    }
  } else if (nombreHoja === "Clientes") {
    Logger.log("Hoja es clientes")
    hoja.deleteRows(3, lastrow)
    let maxRows = hoja.getMaxRows()
    Logger.log("maxRows " + maxRows)
    let dif = 1000 - maxRows
    Logger.log("dif " + dif)
    hoja.insertRows(maxRows, dif)
  }
  else {
    Logger.log("else")
    hoja.deleteRows(2, lastrow)
    let maxRows = hoja.getMaxRows()
    Logger.log("maxRows " + maxRows)
    let dif = 1000 - maxRows
    Logger.log("dif " + dif)
    hoja.insertRows(maxRows, dif)
  }
}
function mensajeBorrarInfoError() {
  Logger.log("Error borrar info")
  SpreadsheetApp.getUi().alert('Si deseas eliminar toda la informacion de misfacturas asegurate de escribir ELIMINAR en el campo');
}
function DesvincularMisfacturas() {
  Logger.log("Desvincular")
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let hojaDatos = spreadsheet.getSheetByName('Datos');
  hojaDatosEmisor.getRange("B13").setBackground('#FFC7C7')
  hojaDatosEmisor.getRange("B13").setValue("Desvinculado")
  hojaDatos.getRange("F47").setValue("")
  SpreadsheetApp.getUi().alert('Haz desvinculado exitosamente misfacturas');
  for (let i = 18; i < 50; i++) {
    hojaDatosEmisor.getRange(i, 1, 1, 6).setValue("")
  }
}
function mensajeErrorDesvincularMisfacturas() {
  Logger.log("Error Desvincular")
  SpreadsheetApp.getUi().alert('Si deseas desvincular misfacturas asegurate de escribir DESVINCULAR en el campo');

}

function openProductosSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Productos");
  SpreadsheetApp.setActiveSheet(sheet);
}


function processForm(data) {
  let existe = verificarIdentificacionUnica(data.codigoReferencia, "Productos")
  if (existe) {
    SpreadsheetApp.getUi().alert("El codigo de referencia ya existe, por favor poner un codigo de referencia unico");
    throw new Error('por favor poner un Numero de Identificacion unico');
  }
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Productos");
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    //Crea las variables para guardar los datos del producto
    const tipo = data.tipoProducto;
    const codigoReferencia = data.codigoReferencia;
    const nombre = data.nombre;
    const precioUnitario = parseFloat(data.precioUnitario);
    const unidadDeMedida = data.unidadDeMedida;
    const unidadMedidaCheckbox = data.unidadMedidaCheckbox;
    const referenciaAdicional = data.referenciaAdicional;
    const numeroReferenciaAdicional = referenciaAdicionalCodigos[referenciaAdicional];
    let iva = "";
    let tarifaIva = String(data.tarifaIva) + "%";
    if (data.tarifaIva !== "0") {
      iva = "IVA";
    }
    let inc = "";
    let tarifaInc = String(data.tarifaInc) + "%";
    if (data.tarifaInc !== "0") {
      inc = "INC";
    }
    const retencionConcepto = data.tarifaReteRenta;
    Logger.log("retencionConcepto " + retencionConcepto)
    let tarifaRetencion = String(validarTarifaRetencion(retencionConcepto) + "%");

    //Asigna los valores a los campos en el sheet
    //Tipo
    sheet.getRange(newRow, 1).setValue(tipo);
    sheet.getRange(newRow, 1).setHorizontalAlignment('center');
    //Codigo Referencia
    sheet.getRange(newRow, 2).setValue(codigoReferencia);
    sheet.getRange(newRow, 2).setHorizontalAlignment('center');
    //Nombre
    sheet.getRange(newRow, 3).setValue(nombre);
    sheet.getRange(newRow, 3).setHorizontalAlignment('center');
    //Referencia Adicional
    sheet.getRange(newRow, 4).setValue(referenciaAdicional);
    //Codigo Referencia Adicional
    sheet.getRange(newRow, 5).setValue(numeroReferenciaAdicional);

    //Precio Unitario
    sheet.getRange(newRow, 6).setValue(precioUnitario);
    sheet.getRange(newRow, 6).setHorizontalAlignment('normal');
    sheet.getRange(newRow, 6).setNumberFormat('$#,##0');
    //Unidad de Medida
    Logger.log("unidadMedidaCheckbox " + unidadMedidaCheckbox)
    Logger.log("unidadDeMedida " + unidadDeMedida)
    if (unidadMedidaCheckbox) {
      sheet.getRange(newRow, 7).setValue("Unidad");
    } else {
    sheet.getRange(newRow, 7).setValue(unidadDeMedida);
    }
    //Columna IVA
    sheet.getRange(newRow, 8).setValue(iva);
    //Tarifa IVA (formatea la celda como porcentaje)
    const tarifaIVA = sheet.getRange(newRow, 9);
    tarifaIVA.setHorizontalAlignment('center');
    tarifaIVA.setValue(tarifaIva); // Establece el valor del IVA como decimal
    //Columna INC
    sheet.getRange(newRow, 10).setValue(inc);
    //Tarifa INC (formatea la celda como porcentaje)  
    const tarifaINC = sheet.getRange(newRow, 11);
    tarifaINC.setHorizontalAlignment('center');
    tarifaINC.setValue(tarifaInc); // Establece el valor del IVA como decimal
    //Precio impuesto
    precioImpuesto = precioUnitario * (parseFloat(data.tarifaIva) / 100) + precioUnitario * (parseFloat(data.tarifaInc) / 100);
    sheet.getRange(newRow, 12).setValue(precioImpuesto);
    //Retencion concepto
    sheet.getRange(newRow, 13).setValue(retencionConcepto);
    //Tarifa Retencion (formatea la celda como porcentaje)
    const tarifaRetencionCell = sheet.getRange(newRow, 14);
    tarifaRetencionCell.setValue(tarifaRetencion); // Establece el valor del IVA como decimal
    tarifaRetencionCell.setNumberFormat('0%'); // Formatea la celda como porcentaje con dos decimales
    //Valor Retencion
    valorRetencion = precioUnitario * (parseFloat(tarifaRetencion) / 100);
    sheet.getRange(newRow, 15).setValue(valorRetencion);

    let referenciaUnica = nombre + "-" + codigoReferencia
    sheet.getRange(newRow, 16).setValue(referenciaUnica);
    sheet.getRange(newRow, 16).setHorizontalAlignment('normal');


    return SpreadsheetApp.getUi().alert("Nuevo producto creado satisfactoriamente");
  } catch (error) {
    return SpreadsheetApp.getUi().alert("Error al guardar los datos: " + error.message);
  }
}


function convertToPercentage(value) {
  return (value * 100).toFixed(2);
}

function onEdit(e) {
  let hojaActual = e.source.getActiveSheet();
  let nombreHoja = hojaActual.getName();

  if (nombreHoja === "Factura") {
    let celdaEditada = e.range;
    let rowEditada = celdaEditada.getRow();
    let colEditada = celdaEditada.getColumn();
    let columnaClientes = 2; // Ajusta según sea necesario
    let rowClientes = 2;

    const productStartRow = 15; // prodcutos empeiza aca
    let cargosDescuentosStartRow = getcargosDescuentosStartRow(hojaActual); // Assuming products end at column L
    let posRowTotalProductos = cargosDescuentosStartRow - 2//poscion (row) de Total productos

    if (colEditada === columnaClientes && rowEditada === rowClientes) {
      verificarYCopiarCliente(e);
      ponerFechaYHoraActual();
    }
    if (rowEditada == 3 && colEditada == 8) {
      let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
      let numeroAutorizacion = celdaEditada.getValue();
      let consecutivoFactura = 0;
      for (i = 18; i <= 20; i++) {
        if (hojaDatosEmisor.getRange(i, 1).getValue() == numeroAutorizacion) {
          consecutivoFactura = hojaDatosEmisor.getRange(i, 5).getValue();
          break;
        }
      }
      hojaActual.getRange("H2").setValue(consecutivoFactura);
    }

    else if (rowEditada >= productStartRow && (colEditada == 2 || colEditada == 3) && rowEditada < posRowTotalProductos) {
      let i = rowEditada;
      let productoFilaI = factura_sheet.getRange("B" + String(i)).getValue()
      let dictInformacionProducto = obtenerInformacionProducto(productoFilaI);
      let cantidadProducto = factura_sheet.getRange("C" + String(i)).getValue()

      if (cantidadProducto === "") {
        cantidadProducto = 0
        factura_sheet.getRange("A" + String(i)).setValue(dictInformacionProducto["codigo Producto"])
        factura_sheet.getRange("D" + String(i)).setValue(dictInformacionProducto["precio Unitario"])//precio unitario
        factura_sheet.getRange("G" + String(i)).setValue(dictInformacionProducto["tarifa IVA"])//%IVA
        factura_sheet.getRange("H" + String(i)).setValue(dictInformacionProducto["tarifa INC"])//%INC

      } else {
        factura_sheet.getRange("A" + String(i)).setValue(dictInformacionProducto["codigo Producto"])
        factura_sheet.getRange("D" + String(i)).setValue(dictInformacionProducto["precio Unitario"])//precio unitario
        factura_sheet.getRange("E" + String(i)).setValue("=D" + String(i) + "*C" + String(i) + "-(D" + String(i) + "*C" + String(i) + ")*" + "I" + String(i) + "+J" + String(i))//Subtotal teniendo en cuenta descuentos y cargos
        factura_sheet.getRange("F" + String(i)).setValue("=E" + String(i) + "*" + dictInformacionProducto["tarifa IVA"] + "+E" + String(i) + "*" + String(dictInformacionProducto["tarifa INC"]))//Impuestos
        factura_sheet.getRange("G" + String(i)).setValue(dictInformacionProducto["tarifa IVA"])//%IVA
        factura_sheet.getRange("H" + String(i)).setValue(dictInformacionProducto["tarifa INC"])//%INC
        cargos = Number(factura_sheet.getRange("J" + String(i)).getValue())//Cargos
        factura_sheet.getRange("K" + String(i)).setValue(dictInformacionProducto["valor Retencion"] * factura_sheet.getRange("E" + String(i)).getValue())//Retencion
        factura_sheet.getRange("L" + String(i)).setValue("=E" + String(i) + "+F" + String(i))
      }

      let lastRowProducto = cargosDescuentosStartRow - 3;
      if (lastRowProducto === productStartRow) {
        // //ESTADO DEAFULT no se hace nada
        let rowParaTotales = getTotalesLinea(hojaActual);
        hojaActual.getRange("K" + String(rowParaTotales)).setValue("=L15")
        hojaActual.getRange("L" + String(rowParaTotales)).setValue("=K" + String(rowParaTotales) + "-J" + String(rowParaTotales))
        let productoFilaI = factura_sheet.getRange("B15").getValue()
        let dictInformacionProducto = obtenerInformacionProducto(productoFilaI);
        let tarifaINC = dictInformacionProducto["tarifa INC"]
        let tarifaIVA = dictInformacionProducto["tarifa IVA"]
        if (tarifaINC !== 0 || tarifaIVA !== 0) {
          hojaActual.getRange("B" + String(rowParaTotales)).setValue("=E" + String(lastRowProducto))
          let impuestosSeccionStartRow = getLastCargoDescuentoRow(hojaActual) + 4
          if (tarifaINC !== 0) {
            hojaActual.getRange("A" + String(impuestosSeccionStartRow)).setValue("INC")//tipo impuesto
            hojaActual.getRange("B" + String(impuestosSeccionStartRow)).setValue("=H" + String(lastRowProducto))//tarifa
            hojaActual.getRange("C" + String(impuestosSeccionStartRow)).setValue("=E" + String(lastRowProducto))//base grabable
            hojaActual.getRange("E" + String(impuestosSeccionStartRow)).setValue("=C" + String(impuestosSeccionStartRow) + "*B" + String(impuestosSeccionStartRow))//total impuesto
          } else {

          }
          if (tarifaIVA !== 0) {
            if (tarifaINC !== 0) {
              hojaActual.insertRowAfter(impuestosSeccionStartRow)
              impuestosSeccionStartRow += 1
            }
            hojaActual.getRange("A" + String(impuestosSeccionStartRow)).setValue("IVA")//tipo impuesto
            hojaActual.getRange("B" + String(impuestosSeccionStartRow)).setValue("=G" + String(lastRowProducto))//tarifa
            hojaActual.getRange("C" + String(impuestosSeccionStartRow)).setValue("=E" + String(lastRowProducto))//base grabable
            hojaActual.getRange("E" + String(impuestosSeccionStartRow)).setValue("=C" + String(impuestosSeccionStartRow) + "*G" + String(lastRowProducto))//total impuesto
          }
        }
        calcularDescuentosCargosYTotales(lastRowProducto, cargosDescuentosStartRow, hojaActual)

      } else {

        calcularImpuestos(hojaActual, lastRowProducto, cargosDescuentosStartRow)
        calcularDescuentosCargosYTotales(lastRowProducto, cargosDescuentosStartRow, hojaActual)
      }

      updateTotalProductCounter(lastRowProducto, productStartRow, hojaActual, cargosDescuentosStartRow)

    } else if ((colEditada == 9 || colEditada == 10) && rowEditada >= productStartRow && rowEditada < posRowTotalProductos) {
      //verificar descuentos
      let cargosDescuentosStartRow = getcargosDescuentosStartRow(hojaActual);
      let lastRowProducto = cargosDescuentosStartRow - 3;
      calcularDescuentosCargosYTotales(lastRowProducto, cargosDescuentosStartRow, hojaActual)

    } else if ((colEditada == 2 || colEditada == 3 || colEditada == 4) && rowEditada > posRowTotalProductos) {
      let cargosDescuentosStartRow = getcargosDescuentosStartRow(hojaActual);
      let lastRowProducto = cargosDescuentosStartRow - 3;
      calcularDescuentosCargosYTotales(lastRowProducto, cargosDescuentosStartRow, hojaActual)
    } else if (colEditada == 7 && rowEditada == 6) {
      // Entra a verificar días de vencimiento
      let valorDiasVencimiento = celdaEditada.getValue();

      // Verifica si es un entero positivo
      if (!Number.isInteger(valorDiasVencimiento) || valorDiasVencimiento <= 0) {
        // Muestra una alerta
        SpreadsheetApp.getUi().alert('El valor de días de vencimiento debe ser un entero positivo.');

        // Restablece el valor a 0
        celdaEditada.setValue(0);
      }
    } else if (colEditada == 10 && rowEditada == 4) {
      //Verifica la moneda
      let moneda = celdaEditada.getValue();
      if (moneda != "Pesos Colombianos") {
        hojaActual.getRange(rowEditada + 1, colEditada).setBackground('#FFC7C7');
        ponerFechaTasaDeCambio();
      } else {
        hojaActual.getRange(rowEditada + 1, colEditada).setBackground('#FFFFFF');
        hojaActual.getRange(rowEditada + 1, colEditada).setValue("");
        hojaActual.getRange(rowEditada + 2, colEditada).setValue("");
      }
    }



  } else if (nombreHoja === "Clientes") {

    let celdaEditada = e.range;
    let hojaCliente = e.source.getActiveSheet();

    let rowEditada = celdaEditada.getRow();
    let colEditada = celdaEditada.getColumn();
    let colTipoDePersona = 2
    let tipoPersona = obtenerTipoDePersona(e);

    if (colEditada == 10 && rowEditada > 1) {
      Logger.log("entro a ver si el edit es en numero")
      let numeroIdentificacion = hojaCliente.getRange(rowEditada, colEditada).getValue()
      Logger.log("num i" + numeroIdentificacion)
      let existe = verificarIdentificacionUnica(numeroIdentificacion, "Clientes", true, rowEditada)
      if (existe) {
        SpreadsheetApp.getUi().alert("El numero de identificacion ya existe, por favor elegir otro numero unico");
        celdaEditada.setValue("");
        verificarDatosObligatorios(e, tipoPersona)
        throw new Error('por favor poner un Numero de Identificacion unico');
      }
    }
    var hojaDatos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos');
    hojaDatos.getRange("L101").setValue("=Clientes!N" + rowEditada);
    // Agregar regla de validación de datos
    let rangoValidacion = hojaCliente.getRange("O" + rowEditada);
    let regla = SpreadsheetApp.newDataValidation()
      .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRange("Datos!$M$101:$M$367"))
      .setAllowInvalid(false)
      .build();
    rangoValidacion.setDataValidation(regla);

    verificarDatosObligatorios(e, tipoPersona)
    agregarCodigoIdentificador(e, tipoPersona)

  } else if (nombreHoja === "Productos") {

    let celdaEditada = e.range;
    let hojaProductos = e.source.getActiveSheet();

    let rowEditada = celdaEditada.getRow();
    let colEditada = celdaEditada.getColumn();
    let codigoReferencia = hojaProductos.getRange(rowEditada, colEditada).getValue()
    let existe = verificarIdentificacionUnica(codigoReferencia, "Productos", true, rowEditada)
    if (existe) {
      SpreadsheetApp.getUi().alert("El codigo de referencia ya existe, por favor elegir otro numero unico");
      hojaProductos.getRange(rowEditada, 2).setValue("");
      throw new Error('por favor poner un codigo de referencia unico');
    }
    let tipoPersona = '';
    verificarDatosObligatoriosProductos(e);
    agregarCodigoIdentificador(e, tipoPersona);
  }
}

function calcularImpuestos(hojaActual, lastRowProducto, cargosDescuentosStartRow) {

}


function calcularDescuentosCargosYTotales(lastRowProducto, cargosDescuentosStartRow, hojaActual) {


  let rowParaTotales = getTotalesLinea(hojaActual);

  //subtotal 
  hojaActual.getRange("A" + String(rowParaTotales)).setValue("=SUM(E15:E" + String(lastRowProducto) + ")")

  //base grabable
  const rangoDatos = hojaActual.getRange("E15:F" + String(lastRowProducto)).getValues();
  const condicion = 0;
  let suma = 0;
  rangoDatos.forEach(fila => {
    if (fila[1] !== condicion) {
      suma += fila[0];
    }
  });
  hojaActual.getRange("B" + String(rowParaTotales)).setValue(suma)

  //Seccion Cargos y Descuentos
  let lastCargoDescuentoRow = getLastCargoDescuentoRow(hojaActual);
  let rowSeccionCargosYDescuentos = cargosDescuentosStartRow;
  let totalDescuentosSeccionCargosyDescuentos = calcularCargYDescu(hojaActual, rowSeccionCargosYDescuentos, lastCargoDescuentoRow);

  //Seccion Impuestos
  let firstRowImpuestos = lastCargoDescuentoRow + 4;
  let lastRowImpuestos = getTotalesLinea(hojaActual) + 3;


  //impuestos
  hojaActual.getRange("C" + String(rowParaTotales)).setValue("=SUM(F15:F" + String(lastRowProducto) + ")")
  //subtotal mas impuestos
  hojaActual.getRange("D" + String(rowParaTotales)).setValue("=A" + String(rowParaTotales) + "+C" + String(rowParaTotales))

  //retenciones
  hojaActual.getRange("E" + String(rowParaTotales)).setValue("=SUM(K15:K" + String(lastRowProducto) + ")")

  //descuentos
  //let descuentosPorProductos = calcularDescuentos(hojaActual, lastRowProducto)
  //Logger.log("descuentosPorProductos " + totalDescuentosSeccionCargosyDescuentos.descuentos)
  //hojaActual.getRange("F" + String(rowParaTotales)).setValue(descuentosPorProductos + Number(totalDescuentosSeccionCargosyDescuentos.descuentos))
  hojaActual.getRange("F" + String(rowParaTotales)).setValue(Number(totalDescuentosSeccionCargosyDescuentos.descuentos))

  //cargos
  //hojaActual.getRange("H" + String(rowParaTotales)).setValue("=SUM(J15:J" + String(lastRowProducto) + ")+" + totalDescuentosSeccionCargosyDescuentos.cargos)
  hojaActual.getRange("H" + String(rowParaTotales)).setValue(totalDescuentosSeccionCargosyDescuentos.cargos)


  //total
  hojaActual.getRange("K" + String(rowParaTotales)).setValue("=D" + String(rowParaTotales) + "-F" + String(rowParaTotales) + "+H" + String(rowParaTotales))
  //neto a pagar
  hojaActual.getRange("L" + String(rowParaTotales)).setValue("=K" + String(rowParaTotales) + "-J" + String(rowParaTotales))
}

function calcularDescuentos(hojaActual, lastRowProducto) {
  let rangoDatos = hojaActual.getRange("E15:K" + String(lastRowProducto)).getValues();
  let resultado = 0;

  rangoDatos.forEach(fila => {
    //suma = subtotal + impuestos + cargos + retencion
    let suma = Number(fila[0]) + Number(fila[1]) + Number(fila[5]) + Number(fila[6]);
    //producto = suma * descuento
    let producto = suma * Number(fila[4]);
    resultado += producto;

  });

  Logger.log("resultado " + resultado)
  return resultado;

}

function getLastProductRow(sheet, productStartRow, cargosDescuentosStartRow) {

  //retorna el numero de fila exacta donde esta el ulitmo producto agregado
  // si no encuntra producto agg si solo tiene un producto retorna el mismo productStartRow 
  let lastProductRow = productStartRow;

  for (let row = productStartRow; row < cargosDescuentosStartRow; row++) {

    let valorCeldaActual = sheet.getRange(row, 1).getValue()
    if (valorCeldaActual === "Total filas") {
      return lastProductRow
    } else {
      lastProductRow = row;
    }
  }
  return lastProductRow;
}
function getLastCargoDescuentoRow(sheet) {
  //obtiene la row donde esta el final de la seccion de cargos y descuentos
  const lastRow = sheet.getLastRow();
  let row = 21

  for (row; row < lastRow; row++) {
    if (sheet.getRange(row, 1).getValue() === 'Tipo Impuesto') {
      return row - 3;
    }
  }
}
function getTotalesLinea(sheet) {
  //obtiene la row donde esta la linea de totales
  const lastRow = sheet.getLastRow();
  let row = 25

  for (row; row < lastRow; row++) {
    if (sheet.getRange(row, 1).getValue() === 'Subtotal') {
      return row + 1;
    }
  }
}

function getcargosDescuentosStartRow(sheet) {
  const lastRow = sheet.getLastRow();
  let row = 14

  for (row; row < lastRow; row++) {
    if (sheet.getRange(row, 1).getValue() === 'Cargos y/o Descuentos') {
      return row;
    }
  }
}

function calcularCargYDescu(hojaActual, rowSeccionCargosYDescuentos, lastCargoDescuentoRow) {
  let totalDescuentosSeccionCargosyDescuentos = { cargos: 0, descuentos: 0 };
  let totalesRow = getTotalesLinea(hojaActual);
  for (let i = rowSeccionCargosYDescuentos + 2; i < lastCargoDescuentoRow + 1; i++) {
    let celdaValorPorcentaje = hojaActual.getRange("C" + String(i)).getValue()
    if (hojaActual.getRange("A" + String(i)).getValue() === "Cargo") {
      if (String(celdaValorPorcentaje).includes("%")) {
        porcentaje = celdaValorPorcentaje.replace("%", "") / 100
        let subtotal = hojaActual.getRange("A" + String(totalesRow)).getValue();
        let base = hojaActual.getRange("D" + String(i)).setValue(subtotal);
        hojaActual.getRange("E" + String(i)).setValue("=" + subtotal + "*" + porcentaje)

      } else {
        hojaActual.getRange("D" + String(i)).setValue("N/A")
        hojaActual.getRange("E" + String(i)).setValue(celdaValorPorcentaje)
      }
      totalDescuentosSeccionCargosyDescuentos.cargos += hojaActual.getRange("E" + String(i)).getValue();
    }
    else {
      if (String(celdaValorPorcentaje).includes("%")) {
        porcentaje = celdaValorPorcentaje.replace("%", "") / 100
        let subtotal = hojaActual.getRange("A" + String(totalesRow)).getValue();
        let base = hojaActual.getRange("D" + String(i)).setValue(subtotal);
        hojaActual.getRange("E" + String(i)).setValue("=" + subtotal + "*" + porcentaje)

      } else {
        hojaActual.getRange("D" + String(i)).setValue("N/A")
        hojaActual.getRange("E" + String(i)).setValue(celdaValorPorcentaje)
      }
      totalDescuentosSeccionCargosyDescuentos.descuentos += hojaActual.getRange("E" + String(i)).getValue();
    }

  }
  return totalDescuentosSeccionCargosyDescuentos;
}

function updateTotalProductCounter(lastRowProducto, productStartRow, hojaActual, cargosDescuentosStartRow) {
  let totalProducts = 0;

  for (let i = productStartRow; i <= lastRowProducto; i++) {
    if (hojaActual.getRange("B" + String(i)).getValue() != "") {
      totalProducts++
    }
  }

  let rowTotalProductos = cargosDescuentosStartRow - 2
  hojaActual.getRange("B" + String(rowTotalProductos)).setValue(totalProducts)

}

function verificarIdentificacionUnica(codigo, nombreHoja, inHoja, numRow) {
  // Clientes o Productos
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  if (codigo == "") {
    return false
  } else if (nombreHoja === "Clientes") {
    try {
      let columnaNumIdentificacionC = 10;
      let lastActiveRow = sheet.getLastRow();
      let rangeNumeroIdentificaciones;
      if (inHoja) {
        rangeNumeroIdentificaciones = sheet.getRange(2, columnaNumIdentificacionC, lastActiveRow);
      } else {
        rangeNumeroIdentificaciones = sheet.getRange(2, columnaNumIdentificacionC, lastActiveRow);
      }
      let NumerosIdentificacion = String(rangeNumeroIdentificaciones.getValues());
      NumerosIdentificacion = NumerosIdentificacion.split(",")

      // Verificar si el código ya existe
      if (NumerosIdentificacion.includes(String(codigo))) {
        var posicionArreglo = NumerosIdentificacion.indexOf(String(codigo)) + 2;
        if (posicionArreglo !== numRow){
          Logger.log(`posicion en el arreglo ${NumerosIdentificacion.indexOf(String(codigo))} y en numero de fila ${numRow}`)
          Logger.log("El Num identificion ya existe.");
          return true
        }
      } else {
        Logger.log("El Num identificion no existe.");
        return false
      }
    } catch (error) {
      Logger.log("Error al verificar el Num identificion: " + error.message);
    }
  } else if (nombreHoja === "Productos") {
    try {
      let columnaNumIdentificacionP = 2;
      let lastActiveRow = sheet.getLastRow();
      let codigoNumero = Number(codigo);
      let rangeCodigoReferencia = sheet.getRange(2, columnaNumIdentificacionP, lastActiveRow);
      let datos = rangeCodigoReferencia.getValues().flat().map(Number);
      for (let i = 0; i < datos.length; i++) {
        if (datos[i] === codigoNumero && i + 2 !== numRow) {
          Logger.log(`El código "${codigoNumero}" ya existe en la hoja "${nombreHoja}".`);
          return true;
        }
      }
      return false;
    } catch (error) {
      Logger.log("Error al verificar el codigo: " + error.message);
      return false;
    }
  }
}

function agregarCodigoIdentificador(e, tipoPersona) {
  let hoja = e.source.getActiveSheet();
  let range = e.range;
  let rowEditada = range.getRow();


  if (hoja.getName() == "Clientes") {
    let estadoActual = hoja.getRange(rowEditada, 1).getValue()
    if (estadoActual == "Valido") {
      if (tipoPersona === "Natural") {
        let nombre = hoja.getRange(rowEditada, 5).getValue()
        let apellido = hoja.getRange(rowEditada, 7).getValue()
        let numeroIdentificacion = hoja.getRange(rowEditada, 10).getValue()
        let identificadorUnico = nombre + " " + apellido + "-" + numeroIdentificacion
        hoja.getRange(rowEditada, 23).setValue(identificadorUnico)
        let rangoValidacion = hoja.getRange("O" + rowEditada);
        rangoValidacion.clearDataValidations();

      } else if (tipoPersona === "Juridica") {
        let nombre = hoja.getRange(rowEditada, 4).getValue()
        let numeroIdentificacion = hoja.getRange(rowEditada, 10).getValue()
        let identificadorUnico = nombre + "-" + numeroIdentificacion
        hoja.getRange(rowEditada, 23).setValue(identificadorUnico)
        let rangoValidacion = hoja.getRange("O" + rowEditada);
        rangoValidacion.clearDataValidations();
      }
    }
  } else if (hoja.getName() == "Productos") {
    let nombre = hoja.getRange(rowEditada, 3).getValue()
    let numeroIdentificacion = hoja.getRange(rowEditada, 2).getValue()
    let identificadorUnico = nombre + "-" + numeroIdentificacion
    hoja.getRange(rowEditada, 16).setValue(identificadorUnico)
  }
}

function limpiarDict() {
  Logger.log("Limpiar el dict")
  diccionarioCaluclarIva = {
    "0.21": 0,
    "0.1": 0,
    "0.05": 0,
    "0.04": 0,
    "0": 0
  }
}

function slugifyF(str) {
  var map = {
    '-': ' ',
    '-': '_',
    'a': 'á|à|ã|â|À|Á|Ã|Â',
    'e': 'é|è|ê|É|È|Ê',
    'i': 'í|ì|î|Í|Ì|Î',
    'o': 'ó|ò|ô|õ|Ó|Ò|Ô|Õ',
    'u': 'ú|ù|û|ü|Ú|Ù|Û|Ü',
    'c': 'ç|Ç',
    'n': 'ñ|Ñ'
  };

  str = String(str)
  str = str.toLowerCase();

  for (var pattern in map) {
    str = str.replace(new RegExp(map[pattern], 'g'), pattern);
  };

  return str;
};
function getAdditionalDocuments() {
  var AdditionalDocuments = {
    "OrderReference": "",
    "DespatchDocumentReference": "",
    "ReceiptDocumentReference": "",
    "AdditionalDocument": []
  }
  return AdditionalDocuments;
}


function getAdditionalProperty() {
  var AdditionalProperty = [];
  return AdditionalProperty;
}

function getDelivery() {
  var Delivery = {
    "AddressLine": "",
    "CountryCode": "",
    "CountryName": "",
    "SubdivisionCode": "",
    "SubdivisionName": "",
    "CityCode": "",
    "CityName": "",
    "ContactPerson": "",
    "DeliveryDate": "",
    "DeliveryCompany": ""
  };
  return Delivery;

}
