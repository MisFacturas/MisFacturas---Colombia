function obtenerInformacionProducto(producto) {
    let celdaProducto = datos_sheet.getRange("I11");
    celdaProducto.setValue(producto);
  
  
  
    let codigoProducto = datos_sheet.getRange("H11").getValue();
    let valorUnitario = datos_sheet.getRange("J11").getValue();
    let porcientoIva = String(datos_sheet.getRange("K11").getValue());
    let precioConIva = datos_sheet.getRange("L11").getValue();
    let impuestos = datos_sheet.getRange("M11").getValue();

    
    let porcentajeNumerico = parseFloat(porcientoIva.replace('%', ''));
    Logger.log("porcentajeNumerico"+porcentajeNumerico)
    let valorConIva = (valorUnitario * porcentajeNumerico) / 100;

    let informacionProducto = {
      "codigo Producto": codigoProducto,
      "valor Unitario": valorUnitario,
      "porciento Iva": porcientoIva,
      "precio Con Iva": precioConIva,
      "impuestos": impuestos,
      "precio con IVA":valorConIva
    };
  
    return informacionProducto;
  }
  