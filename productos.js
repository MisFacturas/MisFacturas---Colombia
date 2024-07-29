var spreadsheet = SpreadsheetApp.getActive();
function obtenerInformacionProducto(producto) {
    let celdaProducto = datos_sheet.getRange("I11");
    Logger.log("producto dentro de obtener "+producto)
    celdaProducto.setValue(producto);
  
  
  
    let codigoProducto = datos_sheet.getRange("H11").getValue();
    let valorUnitario = datos_sheet.getRange("J11").getValue();
    let porcientoIva = String(datos_sheet.getRange("K11").getValue());
    let precioConIva = datos_sheet.getRange("L11").getValue();
    let impuestos = datos_sheet.getRange("M11").getValue();
    // Logger.log("Dentro de funcion dict porcientoIva "+ porcientoIva)
    // Logger.log("Dentro de funcion dict porcientoIva sin string"+ datos_sheet.getRange("K11").getValue())
    

    let informacionProducto = {
      "codigo Producto": codigoProducto,
      "valor Unitario": valorUnitario,
      "porciento Iva": porcientoIva,
      "precio Con Iva": precioConIva,
      "impuestos": impuestos
    };
  
    return informacionProducto;
  }

function buscarProductos(){
  var spreadsheet = SpreadsheetApp.getActive();
  var hojaProductos = spreadsheet.getSheetByName('Productos');
  let ultimaRowActiva=hojaProductos.getMaxRows();
  let valorListaProductos=hojaProductos.getRange(2,2,ultimaRowActiva,1).getValues()
  var listaProductos = valores.map(function(row) {
    return row[0];
  }).join(',');

  const resultBox= document.querySelector(".result-box");
  const inputBox=document.getElementById("search-box");

  inputBox.onkeyup = function(){
    let result=[];
    let input=inputBox.value;
    if(input.length){
      result = listaProductos.filter((keyword)=>{
        
      });
    }
  }

}
  

  