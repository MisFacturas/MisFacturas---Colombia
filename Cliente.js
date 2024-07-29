var spreadsheet = SpreadsheetApp.getActive();
var datos_sheet = spreadsheet.getSheetByName('Datos');
var factura_sheet= spreadsheet.getSheetByName("Factura2")


function obtenerTipoDePersona(e){
  let sheet = e.source.getActiveSheet();
  let range = e.range;
  let rowEditada = range.getRow();
  let colEditada = 4;

  let tipoPersona =sheet.getRange(rowEditada,colEditada).getValue()
  return tipoPersona
}

function saveClientData(formData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clientes');
  if (!sheet) {
    throw new Error('La hoja "Clientes" no existe.');
  }

  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(2, 1, lastRow, 1).getValues(); // Obtener la columna A desde la fila 2 hasta la última

  let emptyRow = 0;
  for (let i = 0; i < dataRange.length; i++) {
    if (dataRange[i][0] === '') { // Si la celda está vacía, esta es la fila vacía
      emptyRow = i + 2; // i + 2 porque dataRange empieza en la fila 2
      break;
    }
  }

  if (emptyRow === 0) {
    emptyRow = lastRow + 1; // Si no se encontró ninguna fila vacía, usar la siguiente fila después de la última
  }

  const values = [
    formData.tipoContacto,
    formData.tipoPersona,
    formData.tipoDocumento,
    formData.numeroIdentificacion,
    formData.codigoContacto,
    formData.regimen,
    formData.nombreComercial,
    formData.primerNombre,
    formData.segundoNombre,
    formData.primerApellido,
    formData.segundoApellido,
    formData.pais,
    formData.provincia,
    formData.poblacion,
    formData.direccion,
    formData.codigoPostal,
    formData.telefono,
    formData.sitioWeb,
    formData.email
  ];

  sheet.getRange(emptyRow, 1, 1, values.length).setValues([values]);
}


function verificarDatosObligatorios(e, tipoPersona) {
  let sheet = e.source.getActiveSheet();
  let range = e.range;
  let rowEditada = range.getRow();
  let colEditada = range.getColumn();
  let ultimaColumnaPermitida = 21; // Actualizado para reflejar el número de columnas
  let columnasObligatorias = [];
  let todasLasColumnas = [ 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20,21];

  if (tipoPersona === "") {
    Logger.log("Vacio hizo edicion no en tipoPersona, cogemos el viejo");
    tipoPersona = sheet.getRange("D" + String(rowEditada)).getValue(); // Columna 4 para Tipo Persona
  }

  if (tipoPersona === "Autonomo") {
    columnasObligatorias = [2,3, 4, 5, 6, 8, 10, 12, 14, 17, 18,19, 21]; // Incluyendo "Nombre cliente" (columna 2)
  } else if (tipoPersona === "Empresa") {
    columnasObligatorias = [2,3, 4, 5, 6, 8,9, 14, 17, 18,19, 21]; // Incluyendo "Nombre cliente" (columna 2)
  } else {
    Logger.log("Vacio tipo de persona");
  }
  
  let estadosDefault = ["", "Tipo Documento", "Regimen", "Tipo de persona"]; // Aquí otros estados predeterminados si es necesario

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

    // Actualizar el estado en la primera columna
    if (estaVacioOPredeterminado) {
      sheet.getRange(rowEditada, 1).clearContent(); // Limpiar contenido de "Estado"
    } else {
      let status = estaCompleto ? "Valido" : "No Valido";
      sheet.getRange(rowEditada, 1).setValue(status); // Establecer valor en "Estado"
    }
  }
}


function crearContacto(){
  Logger.log("imprima algo")
  showClientes()

}

function getCustomerInformation(customer) {
  /*esta funcion debe de cambiar para obtener son los datos directamente de la hoja cliente */
  // ojo de donde esta cogiendo el datosheet ?

  let celdaCliente = datos_sheet.getRange("H2");
  celdaCliente.setValue(customer);


  // var range = datos_sheet.getRange("D50");
  // var Customer = range.getValue();

  var range = datos_sheet.getRange("I2");
  var CustomerCode = range.getValue();

  //range = datos_sheet.getRange("C51");// aqui agarra es el numero mas no el tipo en si
  //var IdentificationType = range.getValue();
  let IdentificationType=datos_sheet.getRange("J2").getValue();

  range = datos_sheet.getRange("K2");
  var Identification = range.getValue();//numero de identificacion

  
  var DV = 0;//no existe en espana, predeterminado 0

  range = datos_sheet.getRange("T2");
  var Address = range.getValue();// aqui lo dividia entre 2 por el psotalcode
  
  

  range = datos_sheet.getRange("S2");//cambie en vez de ciudad pais, porque en espana no hay parametro ciudad
  var CityID = range.getValue();

  range = datos_sheet.getRange("V2");
  var Telephone = range.getValue();

  // switch (datos_sheet.getRange("C1").getValue()) {
  //   case "Pruebas":
  //     var range = datos_sheet.getRange("E1");
  //     break;
  //   case "Produccion":
  //     var range = datos_sheet.getRange("B63");
  //     break;
  //   default:
  //     Logger.log("Oops!...Error Ambiente")
  //     return;
  // }
  var range = datos_sheet.getRange("X2");
  var Email = range.getValue();
  //Browser.msgBox(Email);


  range = datos_sheet.getRange("W2");
  var WebSiteURI = range.getValue();


  if (IdentificationType == "#NUM!") {
    Browser.msgBox("ERROR: Seleccione Tipo de Identificacion en Clientes")
    return;
  }

  var CustomerInformation = {
    "IdentificationType": IdentificationType,
    "Identification": Identification,//.toString(),
    "DV": DV,
    "RegistrationName": customer,
    "CountryCode": "ES",
    "CountryName": "España",
    "SubdivisionCode": "En España no se como funcionan codigo  de provinica",// 11, //Codigo de Municipio
    "SubdivisionName": datos_sheet.getRange("AA2").getValue(),// provicnica
    "CityCode": "Hay dos codigos postales, este solo existe para colombia",
    "CityName": datos_sheet.getRange("Z2").getValue(),//polbacion
    "AddressLine": String(Address),
    "PostalZone": datos_sheet.getRange("U2").getValue(),//Confundido con el codigo postal hay 2, de recepcion y de 
    "Email": Email,
    "CustomerCode": CustomerCode,
    "Telephone": Telephone,
    "WebSiteURI": WebSiteURI,
    "AdditionalAccountID": "Numero que representa el tipo de persona, en España no se sabe si se utiliza o no",//"1",//1, //1: Juridica, 2: Natural
    "TaxLevelCodeListName": "numero que representa unos impuestos, no se si en España exista",//"48" Impuesto sobre las ventas IVA 49 – No responsable de impuesto sobre las ventas IVA
    "TaxSchemeCode": "Numero que representa algo, no se si en España exista ",
    "TaxSchemeName": "Numero que representa algo, no se si en España exista ",
    "FiscalResponsabilities": "Responsabiliades fiscales, no se si en España exista",

    "PartecipationPercent": 100,
    "AdditionalCustomer": []


  }
  return CustomerInformation;
}

function obtenerInformacionCliente(cliente) {
  let celdaCliente = datos_sheet.getRange("H2");
  celdaCliente.setValue(cliente);



  let codigoContacto = datos_sheet.getRange("I2").getValue();
  let direccion = datos_sheet.getRange("T2").getValue();
  let pais = datos_sheet.getRange("S2").getValue();
  let provincia = datos_sheet.getRange("AA2").getValue();
  let poblacion = datos_sheet.getRange("Z2").getValue();
  let telefono = datos_sheet.getRange("V2").getValue();
  let estado = datos_sheet.getRange("Y2").getValue();

  let ubicacion = poblacion + ", " + provincia + ", " + pais;

  let informacionCliente = {
    "Código cliente": codigoContacto,
    "Dirección": direccion,
    "Ubicación": ubicacion,
    "Teléfono": telefono,
    "Estado": estado
  };

  return informacionCliente;
}

