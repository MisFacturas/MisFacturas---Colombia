var spreadsheet = SpreadsheetApp.getActive();
let datos_sheet = spreadsheet.getSheetByName('Datos');
let factura_sheet= spreadsheet.getSheetByName("Factura2")


function verificarDatosObligatorios(e) {

  //falta verificar datos en facturas cuando genero factura
  let sheet = e.source.getActiveSheet();
  let range = e.range;
  let rowEditada = range.getRow();
  let colEditada = range.getColumn();
  let ultimaColumnaPermitida = 18;
  let columnasObligatorias = [1, 2, 3,4,5,6];
  let estadosDefault = ["", "Tipo Documento","Regimen","Tipo de persona"]; // aqui otros estados predeterminados si es necesario


  if (rowEditada > 1 && colEditada <= ultimaColumnaPermitida) {
    let estaCompleto = true;
    let estaVacioOPredeterminado = true;

    for (let i = 0; i < columnasObligatorias.length; i++) {
      let valorDeCelda = sheet.getRange(rowEditada, columnasObligatorias[i]).getValue();
      if (estadosDefault.includes(valorDeCelda)) {
        estaCompleto = false;
      } else {
        estaVacioOPredeterminado = false;
      }
    }

    if (estaVacioOPredeterminado) {
      sheet.getRange(rowEditada, ultimaColumnaPermitida).clearContent();
    } else {
      let status = estaCompleto ? "Valido" : "No Valido";
      sheet.getRange(rowEditada, ultimaColumnaPermitida).setValue(status);
    }
  }
}

function getCustomerInformation(customer) {
  /*esta funcion debe de cambiar para obtener son los datos directamente de la hoja cliente */
  // ojo de donde esta cogiendo el datosheet ?

  let celdaCliente = datos_sheet.getRange("H2");
  celdaCliente.setValue(customer);


  // var range = datos_sheet.getRange("D50");
  // var Customer = range.getValue();

  var range = datos_sheet.getRange("E50");
  var CustomerCode = range.getValue();//llamar otra vez al cliente ?

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
    "SubdivisionCode": "En España no existen municipios",// 11, //Codigo de Municipio
    "SubdivisionName": datos_sheet.getRange("B7").getValue(),//"Bogotá", //Nombre de Departamente //solo aparece en la facutra privincia
    "CityCode": "Hay dos codigos postales, este solo existe para colombia",
    "CityName": factura_sheet.getRange("B8").getValue(),//"Bogotá, D.C.",
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