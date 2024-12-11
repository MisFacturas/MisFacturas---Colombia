var spreadsheet = SpreadsheetApp.getActive();

function buscarUnidadesDeMedidaProducto(terminoBusqueda) {
  let unidadesDeMedida = datos_sheet.getRange(35, 3, 365, 1).getValues();
  var resultados = [];
  if (terminoBusqueda === "") {
    return resultados
  }
  // Recorre los valores obtenidos
  for (var i = 0; i < unidadesDeMedida.length; i++) {
    var valor = unidadesDeMedida[i][0]; // Accede al primer (y único) valor de cada fila

    // Comprueba si el valor coincide con el término de búsqueda
    if (valor.toLowerCase().indexOf(terminoBusqueda.toLowerCase()) !== -1) {
      resultados.push(valor); // Añade el valor a la lista de resultados si coincide
    }
  }

  // Devuelve los resultados
  return resultados;
}

function verificarDatosObligatoriosProductos(e) {
  let sheet = e.source.getActiveSheet();
  let range = e.range;
  let rowEditada = range.getRow();
  let colEditada = range.getColumn();
  let ultimaColumnaPermitida = 9;
  let columnasObligatorias = [1, 2, 3, 4, 6, 7, 13];
  let estadosDefault = [""];
  let todasLasColumnas = [1, 2, 3, 4, 5, 6, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15];

  if (rowEditada > 1 && colEditada <= ultimaColumnaPermitida) {
    let estaCompleto = true;
    let estaVacioOPredeterminado = true;

    // Borrar el color de fondo de todas las celdas obligatorias antes de la verificación
    for (let i = 0; i < todasLasColumnas.length; i++) {
      sheet.getRange(rowEditada, todasLasColumnas[i]).setBackground(null);
    }

    // Verificar celdas obligatorias
    for (let i = 0; i < columnasObligatorias.length; i++) {
      let valorDeCelda = sheet.getRange(rowEditada, columnasObligatorias[i]).getValue();
      if (estadosDefault.includes(valorDeCelda)) {
        estaCompleto = false;
        sheet.getRange(rowEditada, columnasObligatorias[i]).setBackground('#FFC7C7'); // Resaltar en rojo claro
      } else {
        estaVacioOPredeterminado = false;
      }
    }
    tarifaIVA = sheet.getRange("I" + String(rowEditada)).getValue();
    tarifaINC = sheet.getRange("K" + String(rowEditada)).getValue();
    Logger.log(tarifaIVA);
    Logger.log(tarifaINC);
    if (tarifaINC !== "" || tarifaIVA !== "") {
      let precioImpuesto = sheet.getRange("F" + String(rowEditada)).getValue()*tarifaIVA + sheet.getRange("F" + String(rowEditada)).getValue()*tarifaINC;
      sheet.getRange("L"  + String(rowEditada)).setValue(precioImpuesto);
    }
    let referenciaAdicional = sheet.getRange("D" + String(rowEditada)).getValue()
    let codigoRefAdicional = referenciaAdicionalCodigos[referenciaAdicional]
    sheet.getRange(rowEditada, 5).setValue(codigoRefAdicional);
    let tipoRetencion = sheet.getRange("M" + String(rowEditada)).getValue();
    if (tipoRetencion !== "" || tipoRetencion !== "No Aplica") {
      let tarifaRetencion = reteRentaValores[tipoRetencion];
      sheet.getRange(rowEditada, 14).setValue(tarifaRetencion+"%");
      let valorRetencion = sheet.getRange("F" + String(rowEditada)).getValue()*(tarifaRetencion/100);
      sheet.getRange(rowEditada, 15).setValue(valorRetencion);
    } else {
      sheet.getRange(rowEditada, 14).setValue("");
      sheet.getRange(rowEditada, 15).setValue("");
      sheet.getRange(rowEditada, 13).setBackground(null);
    }
  }
}

// a cambiar cuando se pregunte y agg los otros porcinetos
function obtenerInformacionProducto(producto) {
  let celdaProducto = datos_sheet.getRange("I11");
  celdaProducto.setValue(producto);

  let codigoProducto = datos_sheet.getRange("H11").getValue();
  let precioUnitario = datos_sheet.getRange("J11").getValue();
  let tarifaIVA = datos_sheet.getRange("K11").getValue();
  let tarifaINC = datos_sheet.getRange("L11").getValue();
  let precioImpuesto = datos_sheet.getRange("L11").getValue();
  let tarifaRetencion = datos_sheet.getRange("M11").getValue();
  let valorRetencion = datos_sheet.getRange("N11").getValue();


  let informacionProducto = {
    "codigo Producto": codigoProducto,
    "precio Unitario": precioUnitario,
    "tarifa IVA": tarifaIVA,
    "tarifa INC": tarifaINC,
    "precio Impuesto": precioImpuesto,
    "tarifa Retencion": tarifaRetencion,
    "valor Retencion": valorRetencion
  };
  if (informacionProducto["tarifa IVA"] == "") {
    informacionProducto["tarifa IVA"] = 0;
  }
  if (informacionProducto["tarifa INC"] == "") {
    informacionProducto["tarifa INC"] = 0;
  }




  return informacionProducto;
}

function buscarProductos(terminoBusqueda) {
  var spreadsheet = SpreadsheetApp.getActive();
  var hojaProductos = spreadsheet.getSheetByName('Productos');
  var ultimaFila = hojaProductos.getLastRow();
  var valores = hojaProductos.getRange(2, 16, ultimaFila - 1, 1).getValues();

  // Filtrar los productos que coincidan con el término de búsqueda
  var productosFiltrados = valores
    .map(function (row) { return row[0]; })
    .filter(function (producto) {
      // Verificar que 'producto' es una cadena antes de llamar a 'toLowerCase'
      return typeof producto === 'string' && producto.toLowerCase().includes(terminoBusqueda.toLowerCase());
    });

  return productosFiltrados;
}

function validarImpuestos(impuestos, tarifaIva, tarifaInc) {
  let tarifaImpuestos = 0;
  if (impuestos === "IVA") {
    tarifaImpuestos = tarifaIva;
  } else {
    tarifaImpuestos = tarifaInc;
  }
  return tarifaImpuestos;
}

function validarTipoRetencion(tarifaReteIva, tarifaReteRenta) {
  let tipoRetencion = "";
  if (tarifaReteIva !== "") {
    tipoRetencion = "Retencion sobre el IVA";
  } else if (tarifaReteRenta !== "") {
    tipoRetencion = tarifaReteRenta;
  }
  return tipoRetencion;
}

function validarTarifaRetencion(tarifaReteIva, tarifaReteRenta) {
  let tarifaRetencion = 0;
  if (tarifaReteIva !== "") {
    tarifaRetencion = tarifaReteIva;
  } else if (tarifaReteRenta !== "") {
    tarifaRetencion = reteRentaValores[tarifaReteRenta];
  }
  return tarifaRetencion;
}

function buscarUnidadesDeMedida(terminoBusqueda) {
  var spreadsheet = SpreadsheetApp.getActive();
  var hojaDatos = spreadsheet.getSheetByName('Datos');
  var valores = hojaDatos.getRange(35, 3, 399, 1).getValues();

  // Filtrar los productos que coincidan con el término de búsqueda
  var productosFiltrados = valores
    .map(function(row) { return row[0]; })
    .filter(function(producto) {
      // Verificar que 'producto' es una cadena antes de llamar a 'toLowerCase'
      return typeof producto === 'string' && producto.toLowerCase().includes(terminoBusqueda.toLowerCase());
    });

  return productosFiltrados;
}

 