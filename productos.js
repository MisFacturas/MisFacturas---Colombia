
function buscarUnidadesDeMedidaProducto(terminoBusqueda) {
  let spreadsheet = SpreadsheetApp.getActive();
  let datos_sheet = spreadsheet.getSheetByName('Datos');
  let unidadesDeMedida = datos_sheet.getRange(35, 2, 365, 1).getValues();
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
  let ultimaColumnaPermitida = 15;
  let columnasObligatorias = [1, 2, 3, 4, 6, 7];
  let todasLasColumnas = [1, 2, 3, 4, 5, 6, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15];
  let valido = false;
  var contadorFaltantes = 6;



  if (rowEditada >= 2 && colEditada <= ultimaColumnaPermitida) {
    let estaCompleto = true;
    let estaVacioOPredeterminado = true;

    // Borrar el color de fondo de todas las celdas obligatorias antes de la verificación
    for (let i = 0; i < todasLasColumnas.length; i++) {
      sheet.getRange(rowEditada, todasLasColumnas[i]).setBackground(null);
    }

    // Verificar celdas obligatorias
    for (let i = 0; i < columnasObligatorias.length; i++) {

      let valorDeCelda = sheet.getRange(rowEditada, columnasObligatorias[i]).getValue();
      if (valorDeCelda === "") {
        estaCompleto = false;
        sheet.getRange(rowEditada, columnasObligatorias[i]).setBackground('#FFC7C7'); // Resaltar en rojo claro
      } else {
        contadorFaltantes--;
        estaVacioOPredeterminado = false;
      }
    }
    tarifaIVA = sheet.getRange("I" + String(rowEditada)).getValue();
    tarifaINC = sheet.getRange("K" + String(rowEditada)).getValue();
    Logger.log(tarifaIVA);
    Logger.log(tarifaINC);
    if (tarifaINC !== "" || tarifaIVA !== "") {
      let precioImpuesto = sheet.getRange("F" + String(rowEditada)).getValue() * tarifaIVA + sheet.getRange("F" + String(rowEditada)).getValue() * tarifaINC;
      sheet.getRange("L" + String(rowEditada)).setValue(precioImpuesto);
      sheet.getRange("L" + String(rowEditada)).setBackground('#d9d9d9');

    }
    let referenciaAdicional = sheet.getRange("D" + String(rowEditada)).getValue()
    let codigoRefAdicional = referenciaAdicionalCodigos[referenciaAdicional]
    sheet.getRange(rowEditada, 5).setValue(codigoRefAdicional);
    let tipoRetencion = sheet.getRange("M" + String(rowEditada)).getValue();
    if (tipoRetencion !== "" || tipoRetencion !== "No Aplica") {
      let tarifaRetencion = reteRentaValores[tipoRetencion];
      sheet.getRange(rowEditada, 14).setValue(tarifaRetencion + "%");
      let valorRetencion = sheet.getRange("F" + String(rowEditada)).getValue() * (tarifaRetencion / 100);
      sheet.getRange(rowEditada, 15).setValue(valorRetencion);
      sheet.getRange(rowEditada, 14).setBackground('#d9d9d9');
    } else {
      sheet.getRange(rowEditada, 14).setValue("");
      sheet.getRange(rowEditada, 15).setValue("");
      sheet.getRange(rowEditada, 13).setBackground(null);
    }
  }
  Logger.log("contador faltantes :" + contadorFaltantes);
  if (contadorFaltantes === 0) {
    valido = true;
  }
  return valido;
}

function obtenerInformacionProducto(producto) {
  let spreadsheet = SpreadsheetApp.getActive();
  let datos_sheet = spreadsheet.getSheetByName('Datos');
  let celdaProducto = datos_sheet.getRange("I11");
  celdaProducto.setValue(producto);

  let codigoProducto = datos_sheet.getRange("H11").getValue();
  let precioUnitario = datos_sheet.getRange("J11").getValue();
  let tarifaIVA = datos_sheet.getRange("K11").getValue();
  let tarifaINC = datos_sheet.getRange("L11").getValue();
  let precioImpuesto = datos_sheet.getRange("M11").getValue();
  let tarifaRetencion = datos_sheet.getRange("N11").getValue();
  let valorRetencion = datos_sheet.getRange("O11").getValue();


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

function validarTarifaRetencion(tarifaReteRenta) {
  let tarifaRetencion = reteRentaValores[tarifaReteRenta];
  return tarifaRetencion;
}

function buscarUnidadesDeMedida(terminoBusqueda) {
  var spreadsheet = SpreadsheetApp.getActive();
  var hojaDatos = spreadsheet.getSheetByName('Datos');
  var valores = hojaDatos.getRange(35, 3, 399, 1).getValues();

  // Filtrar los productos que coincidan con el término de búsqueda
  var productosFiltrados = valores
    .map(function (row) { return row[0]; })
    .filter(function (producto) {
      // Verificar que 'producto' es una cadena antes de llamar a 'toLowerCase'
      return typeof producto === 'string' && producto.toLowerCase().includes(terminoBusqueda.toLowerCase());
    });

  return productosFiltrados;
}

function buscarRetencion(idProducto) {
  let respuesta = [];
  var spreadsheet = SpreadsheetApp.getActive();
  var hojaDatos = spreadsheet.getSheetByName('Datos');
  hojaDatos.getRange("H14").setValue(idProducto);
  var nombreRetencion = hojaDatos.getRange("I14").getValue();
  respuesta.push(nombreRetencion);
  var porcentajeRetencion = hojaDatos.getRange("J14").getValue();
  respuesta.push(porcentajeRetencion);
  return respuesta;
}

