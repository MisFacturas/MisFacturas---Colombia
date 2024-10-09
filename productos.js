var spreadsheet = SpreadsheetApp.getActive();

function buscarUnidadesDeMedida(terminoBusqueda){
  let unidadesDeMedida=datos_sheet.getRange(35,3,365,1).getValues();
  var resultados = [];
  if(terminoBusqueda===""){
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

function saveProductData(formData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Productos');
  if (!sheet) {
    throw new Error('La hoja "Productos" no existe.');
  }

  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(2, 1, lastRow, 12).getValues(); // Obtener desde la columna B hasta la S (19 columnas)

  let emptyRow = 0;
  for (let i = 0; i < dataRange.length; i++) {
    const row = dataRange[i];
    const isEmpty = row.every(cell => cell === ''); // Verificar si todas las celdas de la fila están vacías

    if (isEmpty) {
      emptyRow = i + 2; // i + 2 porque dataRange empieza en la fila 2
      break;
    }
  }

  if (emptyRow === 0) {
    emptyRow = lastRow + 1; // Si no se encontró ninguna fila vacía, usar la siguiente fila después de la última
  }

  const values = [
    formData.codigoReferencia,
    formData.nombre,
    formData.referenciaAdicional,
    formData.precioUnitario,
    formData.unidadDeMedida,
    formData.impuestos,
    formData.tarifaImpuestos,
    formData.retencion,
    formData.tarifaRetencion,
  ];

  sheet.getRange(emptyRow, 1, 1, values.length).setValues([values]);
  SpreadsheetApp.getUi().alert("Nuevo producto generado satisfactoriamente");
}


// a cambiar cuando se pregunte y agg los otros porcinetos
function obtenerInformacionProducto(producto) {
    let celdaProducto = datos_sheet.getRange("I11");
    Logger.log("producto dentro de obtener "+producto)
    celdaProducto.setValue(producto);
  
  
  
    let codigoProducto = datos_sheet.getRange("H11").getValue();
    let precioUnitario = datos_sheet.getRange("J11").getValue();
    let tarifaImpuesto = datos_sheet.getRange("K11").getValue();
    let precioImpuesto = datos_sheet.getRange("L11").getValue();
    let tarifaRetencion = datos_sheet.getRange("M11").getValue();
    let valorRetencion=datos_sheet.getRange("N11").getValue();



    let informacionProducto = {
      "codigo Producto": codigoProducto,
      "precio Unitario": precioUnitario,
      "tarifa Impuesto": tarifaImpuesto,
      "precio Impuesto": precioImpuesto,
      "tarifa Retencion": tarifaRetencion,
      "valor Retencion": valorRetencion
    };
  
    return informacionProducto;
  }

  function buscarProductos(terminoBusqueda) {
    var spreadsheet = SpreadsheetApp.getActive();
    var hojaProductos = spreadsheet.getSheetByName('Productos');
    var ultimaFila = hojaProductos.getLastRow();
    var valores = hojaProductos.getRange(1, 2, ultimaFila - 1, 1).getValues();
  
    // Filtrar los productos que coincidan con el término de búsqueda
    var productosFiltrados = valores
      .map(function(row) { return row[0]; })
      .filter(function(producto) {
        return producto.toLowerCase().includes(terminoBusqueda.toLowerCase());
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

  function validarTipoRetencion(retencion, tarifaReteRenta) {
    let tipoRetencion = "";
    if(retencion === "ReteIva"){
      tipoRetencion = "Retencion sobre el IVA";
    }else {
      tipoRetencion = tarifaReteRenta;
    }
    return tipoRetencion;
  }

  function validarTarifaRetencion(retencion, tarifaReteIva, tarifaReteRenta) {
    let tarifaRetencion = 0;
    if (retencion === "ReteIva") {
      tarifaRetencion = tarifaReteIva;
    } else {
      tarifaRetencion = reteRentaValores[tarifaReteRenta];
    }
    return tarifaRetencion;
  }

  function validarReferenciaAdicional(referenciaAdicional) {
    let numeroReferenciaAdicional = 0;
    if (referenciaAdicional === "UNSPSC") {
      numeroReferenciaAdicional = 1;
    } else if (referenciaAdicional === "GTIN") {
      numeroReferenciaAdicional = 10;
    } else if (referenciaAdicional === "Partida Arancelarias") {
      numeroReferenciaAdicional = 20;
    } else if (referenciaAdicional === "Estándar de adopción del contribuyente") {
      numeroReferenciaAdicional = 999;
    } else if (referenciaAdicional === "No Aplica") {
      numeroReferenciaAdicional = 0;
    }
    return numeroReferenciaAdicional;
  }
  