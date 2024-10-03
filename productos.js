var spreadsheet = SpreadsheetApp.getActive();

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
      "impuestos": tarifaImpuesto,
      "precio impuesto": precioImpuesto,
      "retencion": impuestos,
      "descuentos": descunetos,
      "tarifa Retencion": tarifaRetencion,

    };
  
    return informacionProducto;
  }

  function buscarProductos(terminoBusqueda) {
    var spreadsheet = SpreadsheetApp.getActive();
    var hojaProductos = spreadsheet.getSheetByName('Productos');
    var ultimaFila = hojaProductos.getLastRow();
    var valores = hojaProductos.getRange(2, 2, ultimaFila - 1, 1).getValues();
  
    // Filtrar los productos que coincidan con el término de búsqueda
    var productosFiltrados = valores
      .map(function(row) { return row[0]; })
      .filter(function(producto) {
        return producto.toLowerCase().includes(terminoBusqueda.toLowerCase());
      });
  
    return productosFiltrados;
  }
   

  