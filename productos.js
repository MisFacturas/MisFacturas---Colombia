var spreadsheet = SpreadsheetApp.getActive();

// a cambiar cuando se pregunte y agg los otros porcinetos
function obtenerInformacionProducto(producto) {
    let celdaProducto = datos_sheet.getRange("I11");
    Logger.log("producto dentro de obtener "+producto)
    celdaProducto.setValue(producto);
  
  
  
    let codigoProducto = datos_sheet.getRange("H11").getValue();
    let valorUnitario = datos_sheet.getRange("J11").getValue();
    let porcientoIva = String(datos_sheet.getRange("K11").getValue());
    let precioConIva = datos_sheet.getRange("L11").getValue();
    let impuestos = datos_sheet.getRange("M11").getValue();
    let descunetos=datos_sheet.getRange("N11").getValue();
    let retencion=datos_sheet.getRange("O11").getValue();
    let RecgEquivalencia=datos_sheet.getRange("P11").getValue();
    // Logger.log("Dentro de funcion dict porcientoIva "+ porcientoIva)
    // Logger.log("Dentro de funcion dict porcientoIva sin string"+ datos_sheet.getRange("K11").getValue())
    

    let informacionProducto = {
      "codigo Producto": codigoProducto,
      "valor Unitario": valorUnitario,
      "IVA": porcientoIva,
      "precio Con Iva": precioConIva,
      "impuestos": impuestos,
      "descuentos": descunetos,
      "retencion":retencion,
      "Recargo de equivalencia":RecgEquivalencia


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
   

  