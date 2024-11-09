var spreadsheet = SpreadsheetApp.getActive();
var datos_sheet = spreadsheet.getSheetByName('Datos');
var factura_sheet= spreadsheet.getSheetByName("Factura")

function showNuevaCliente() {
  var html = HtmlService.createHtmlOutputFromFile('menuAgregarCliente').setTitle("Nuevo Cliente")
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function showInactivarCliente() {
  var html = HtmlService.createHtmlOutputFromFile('menuInactivarCliente').setTitle("Inactivar Cliente")
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function showActivarCliente() {
  var html = HtmlService.createHtmlOutputFromFile('menuActivarCliente').setTitle("Activar Cliente")
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function inactivarCliente(cliente){
  let hojaClientesInactivos=spreadsheet.getSheetByName('ClientesInvalidos');
  let hojaClietnes=spreadsheet.getSheetByName("Clientes")
  Logger.log(cliente)
  datos_sheet.getRange("H2").setValue(cliente)
  

  let rowDelCliente=datos_sheet.getRange("G2").getValue();
  let rowMaximaClientesInactivos=hojaClientesInactivos.getLastRow()+1;
  let rowMaximaClientes=hojaClietnes.getLastRow()+1;

  let codigoCliente=datos_sheet.getRange("I2").getValue();
  let tipoDoc=datos_sheet.getRange("J2").getValue();
  let numIdentificacion=datos_sheet.getRange("K2").getValue();
  let tipoPersona=datos_sheet.getRange("L2").getValue();
  let regimen=datos_sheet.getRange("M2").getValue();
  let nombreComercial=datos_sheet.getRange("N2").getValue();
  let primerNombre=datos_sheet.getRange("O2").getValue();
  let segundoNombre=datos_sheet.getRange("P2").getValue();
  let primerApellido=datos_sheet.getRange("Q2").getValue();
  let segundoApellido=datos_sheet.getRange("R2").getValue();
  let pais=datos_sheet.getRange("S2").getValue();
  let direccion=datos_sheet.getRange("T2").getValue();
  let codigoPostal=datos_sheet.getRange("U2").getValue();
  let telefono=datos_sheet.getRange("V2").getValue();
  let sitioWeb=datos_sheet.getRange("W2").getValue();
  let email=datos_sheet.getRange("X2").getValue();
  let estado=datos_sheet.getRange("Y2").getValue();
  let municipio=datos_sheet.getRange("AZ2").getValue();
  let departamento=datos_sheet.getRange("AA2").getValue();
  let detallesTributarios=datos_sheet.getRange("AB2").getValue();
  let responsabilidadFiscal=datos_sheet.getRange("AC2").getValue();
  let tipoTercero=datos_sheet.getRange("AD2").getValue();



  // Proceso para agregar a la hoja de clientes inactivos
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 1).setValue(estado);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 2).setValue(cliente);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 3).setValue(tipoTercero);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 4).setValue(tipoPersona);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 5).setValue(tipoDoc);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 6).setValue(numIdentificacion);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 7).setValue(codigoCliente);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 8).setValue(regimen);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 9).setValue(nombreComercial);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 10).setValue(primerNombre);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 11).setValue(segundoNombre);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 12).setValue(primerApellido);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 13).setValue(segundoApellido);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 14).setValue(pais);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 15).setValue(departamento);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 16).setValue(municipio);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 17).setValue(direccion);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 18).setValue(codigoPostal);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 19).setValue(telefono);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 20).setValue(sitioWeb);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 21).setValue(email);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 22).setValue(detallesTributarios);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 23).setValue(responsabilidadFiscal);


  //eliminar cliente de la hoja clientes

  hojaClietnes.deleteRow(rowDelCliente)
  hojaClietnes.insertRowAfter(rowMaximaClientes)
}

function activarCliente(cliente){
  let hojaClientesInactivos=spreadsheet.getSheetByName('ClientesInvalidos');
  let hojaClietnes=spreadsheet.getSheetByName("Clientes")
  Logger.log(cliente)
  datos_sheet.getRange("I6").setValue(cliente)

  let rowDelCliente=datos_sheet.getRange("G6").getValue();
  let rowMaximaClientesInactivos=hojaClientesInactivos.getLastRow()+1;
  let rowMaximaClientes=hojaClietnes.getLastRow()+1;

  let estado = datos_sheet.getRange('H6').getValue();
  let tipoTercero = datos_sheet.getRange('J6').getValue();
  let tipoPersona = datos_sheet.getRange('K6').getValue();
  let tipoDoc = datos_sheet.getRange('L6').getValue();
  let numIdentificacion = datos_sheet.getRange('M6').getValue();
  let codigoCliente = datos_sheet.getRange('N6').getValue();
  let regimen = datos_sheet.getRange('O6').getValue();
  let nombreComercial = datos_sheet.getRange('P6').getValue();
  let primerNombre = datos_sheet.getRange('Q6').getValue();
  let segundoNombre = datos_sheet.getRange('R6').getValue();
  let primerApellido = datos_sheet.getRange('S6').getValue();
  let segundoApellido = datos_sheet.getRange('T6').getValue();
  let pais = datos_sheet.getRange('U6').getValue();
  let departamento = datos_sheet.getRange('V6').getValue();
  let municipio = datos_sheet.getRange('W6').getValue();
  let direccion = datos_sheet.getRange('X6').getValue();
  let codigoPostal = datos_sheet.getRange('Y6').getValue();
  let telefono = datos_sheet.getRange('Z6').getValue();
  let sitioWeb = datos_sheet.getRange('AA6').getValue();
  let email = datos_sheet.getRange('AB6').getValue();
  let detallesTributarios = datos_sheet.getRange('AC6').getValue();
  let responsabilidadFiscal = datos_sheet.getRange('AD6').getValue();
  
  hojaClietnes.getRange(rowMaximaClientes, 1).setValue(estado);
  hojaClietnes.getRange(rowMaximaClientes, 2).setValue(cliente);
  hojaClietnes.getRange(rowMaximaClientes, 3).setValue(tipoTercero);
  hojaClietnes.getRange(rowMaximaClientes, 4).setValue(tipoPersona);
  hojaClietnes.getRange(rowMaximaClientes, 5).setValue(tipoDoc);
  hojaClietnes.getRange(rowMaximaClientes, 6).setValue(numIdentificacion);
  hojaClietnes.getRange(rowMaximaClientes, 7).setValue(codigoCliente);
  hojaClietnes.getRange(rowMaximaClientes, 8).setValue(regimen);
  hojaClietnes.getRange(rowMaximaClientes, 9).setValue(nombreComercial);
  hojaClietnes.getRange(rowMaximaClientes, 10).setValue(primerNombre);
  hojaClietnes.getRange(rowMaximaClientes, 11).setValue(segundoNombre);
  hojaClietnes.getRange(rowMaximaClientes, 12).setValue(primerApellido);
  hojaClietnes.getRange(rowMaximaClientes, 13).setValue(segundoApellido);
  hojaClietnes.getRange(rowMaximaClientes, 14).setValue(pais);
  hojaClietnes.getRange(rowMaximaClientes, 15).setValue(departamento);
  hojaClietnes.getRange(rowMaximaClientes, 16).setValue(municipio);
  hojaClietnes.getRange(rowMaximaClientes, 17).setValue(direccion);
  hojaClietnes.getRange(rowMaximaClientes, 18).setValue(codigoPostal);
  hojaClietnes.getRange(rowMaximaClientes, 19).setValue(telefono);
  hojaClietnes.getRange(rowMaximaClientes, 20).setValue(sitioWeb);
  hojaClietnes.getRange(rowMaximaClientes, 21).setValue(email);
  hojaClietnes.getRange(rowMaximaClientes, 22).setValue(detallesTributarios);
  hojaClietnes.getRange(rowMaximaClientes, 23).setValue(responsabilidadFiscal);

  hojaClientesInactivos.deleteRow(rowDelCliente)
  hojaClientesInactivos.insertRowAfter(rowMaximaClientesInactivos)
}
function buscarClientes(terminoBusqueda,hojaA) {
  var resultados = [];

  if(hojaA==="Inactivar"){
    var sheet = spreadsheet.getSheetByName('Clientes');
  }else{

    var sheet = spreadsheet.getSheetByName('ClientesInvalidos');
    var ultimaFila = sheet.getLastRow(); 
    var valores = sheet.getRange(2, 2, ultimaFila - 1, 1).getValues();

    for (var i = 0; i < valores.length; i++) {
      var valor = valores[i][0]; // Accede al primer (y único) valor de cada fila
      resultados.push(valor);}
      
    return resultados
}
  
  var ultimaFila = sheet.getLastRow(); 
  var valores = sheet.getRange(2, 2, ultimaFila - 1, 1).getValues(); // `ultimaFila - 1` porque empieza en la fila 2


  if(terminoBusqueda===""){
    return resultados
  }
  // Recorre los valores obtenidos
  for (var i = 0; i < valores.length; i++) {
    var valor = valores[i][0]; // Accede al primer (y único) valor de cada fila
    
    // Comprueba si el valor coincide con el término de búsqueda
    if (valor.toLowerCase().indexOf(terminoBusqueda.toLowerCase()) !== -1) {
      resultados.push(valor); // Añade el valor a la lista de resultados si coincide
    }
  }
  
  // Devuelve los resultados
  return resultados;
}
function buscarPaises(terminoBusqueda){
  let paises=datos_sheet.getRange(25,1,170,1).getValues();
  var resultados = [];
  if(terminoBusqueda===""){
    return resultados
  }
  // Recorre los valores obtenidos
  for (var i = 0; i < paises.length; i++) {
    var valor = paises[i][0]; // Accede al primer (y único) valor de cada fila
    
    // Comprueba si el valor coincide con el término de búsqueda
    if (valor.toLowerCase().indexOf(terminoBusqueda.toLowerCase()) !== -1) {
      resultados.push(valor); // Añade el valor a la lista de resultados si coincide
    }
  }
  
  // Devuelve los resultados
  return resultados;
}


function agregarPaises(){
  const paises = [
    "Afganistán",
    "Albania",
    "Alemania",
    "Andorra",
    "Angola",
    "Antigua y Barbuda",
    "Arabia Saudita",
    "Argelia",
    "Argentina",
    "Armenia",
    "Australia",
    "Austria",
    "Azerbaiyán",
    "Bahamas",
    "Bangladés",
    "Barbados",
    "Baréin",
    "Bélgica",
    "Belice",
    "Benín",
    "Bielorrusia",
    "Birmania",
    "Bolivia",
    "Bosnia y Herzegovina",
    "Botsuana",
    "Brasil",
    "Brunéi",
    "Bulgaria",
    "Burkina Faso",
    "Burundi",
    "Bután",
    "Cabo Verde",
    "Camboya",
    "Camerún",
    "Canadá",
    "Catar",
    "Chad",
    "Chile",
    "China",
    "Chipre",
    "Ciudad del Vaticano",
    "Colombia",
    "Comoras",
    "Corea del Norte",
    "Corea del Sur",
    "Costa de Marfil",
    "Costa Rica",
    "Croacia",
    "Cuba",
    "Dinamarca",
    "Dominica",
    "Ecuador",
    "Egipto",
    "El Salvador",
    "Emiratos Árabes Unidos",
    "Eritrea",
    "Eslovaquia",
    "Eslovenia",
    "España",
    "Estados Unidos",
    "Estonia",
    "Etiopía",
    "Filipinas",
    "Finlandia",
    "Fiyi",
    "Francia",
    "Gabón",
    "Gambia",
    "Georgia",
    "Ghana",
    "Granada",
    "Grecia",
    "Guatemala",
    "Guyana",
    "Guinea",
    "Guinea ecuatorial",
    "Guinea-Bisáu",
    "Haití",
    "Honduras",
    "Hungría",
    "India",
    "Indonesia",
    "Irak",
    "Irán",
    "Irlanda",
    "Islandia",
    "Islas Marshall",
    "Islas Salomón",
    "Israel",
    "Italia",
    "Jamaica",
    "Japón",
    "Jordania",
    "Kazajistán",
    "Kenia",
    "Kirguistán",
    "Kiribati",
    "Kosovo",
    "Kuwait",
    "Laos",
    "Lesoto",
    "Letonia",
    "Líbano",
    "Liberia",
    "Libia",
    "Liechtenstein",
    "Lituania",
    "Luxemburgo",
    "Macedonia del Norte",
    "Madagascar",
    "Malasia",
    "Malaui",
    "Maldivas",
    "Malí",
    "Malta",
    "Marruecos",
    "Mauricio",
    "Mauritania",
    "México",
    "Micronesia",
    "Moldavia",
    "Mónaco",
    "Mongolia",
    "Montenegro",
    "Mozambique",
    "Namibia",
    "Nauru",
    "Nepal",
    "Nicaragua",
    "Níger",
    "Nigeria",
    "Noruega",
    "Nueva Zelanda",
    "Omán",
    "Países Bajos",
    "Pakistán",
    "Palaos",
    "Panamá",
    "Papúa Nueva Guinea",
    "Paraguay",
    "Perú",
    "Polonia",
    "Portugal",
    "Reino Unido",
    "República Centroafricana",
    "República Checa",
    "República del Congo",
    "República Democrática del Congo",
    "República Dominicana",
    "Ruanda",
    "Rumania",
    "Rusia",
    "Samoa",
    "San Cristóbal y Nieves",
    "San Marino",
    "San Vicente y las Granadinas",
    "Santa Lucía",
    "Santo Tomé y Príncipe",
    "Senegal",
    "Serbia",
    "Seychelles",
    "Sierra Leona",
    "Singapur",
    "Siria",
    "Somalia",
    "Sri Lanka",
    "Suazilandia",
    "Sudáfrica",
    "Sudán",
    "Sudán del Sur",
    "Suecia",
    "Suiza",
    "Surinam",
    "Tailandia",
    "Tanzania",
    "Tayikistán",
    "Timor Oriental",
    "Togo",
    "Tonga",
    "Trinidad y Tobago",
    "Túnez",
    "Turkmenistán",
    "Turquía",
    "Tuvalu",
    "Ucrania",
    "Uganda",
    "Uruguay",
    "Uzbekistán",
    "Vanuatu",
    "Venezuela",
    "Vietnam",
    "Yemen",
    "Yibuti",
    "Zambia",
    "Zimbabue"
  ];
  let Paragg=0
  for(let i=25;i<paises.length;i++){
    datos_sheet.getRange("A"+String(i)).setValue(paises[Paragg])
    Paragg++
  }
 }

function obtenerTipoDePersona(e){
  let sheet = e.source.getActiveSheet();
  let range = e.range;
  let rowEditada = range.getRow();
  let colEditada = 4;

  let tipoPersona =sheet.getRange(rowEditada,colEditada).getValue()
  Logger.log(tipoPersona)
  return tipoPersona
  
}

function saveClientData(formData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clientes');
  if (!sheet) {
    throw new Error('La hoja "Clientes" no existe.');
  }

  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(2, 2, lastRow, 19).getValues(); // Obtener desde la columna B hasta la S (19 columnas)

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
    formData.nombreCliente,
    formData.tipoTercero,
    formData.tipoPersona,
    formData.tipoDocumento,
    formData.numeroIdentificacion,
    formData.codigoCliente,
    formData.regimen,
    formData.nombreComercial,
    formData.primerNombre,
    formData.segundoNombre,
    formData.primerApellido,
    formData.segundoApellido,
    formData.pais,
    formData.departamento,
    formData.municipio,
    formData.direccion,
    formData.codigoPostal,
    formData.telefono,
    formData.sitioWeb,
    formData.email,
    formData.detallesTributarios,
    formData.responsabilidadFiscal
  ];

  sheet.getRange(emptyRow, 2, 1, values.length).setValues([values]);

  sheet.getRange(emptyRow, 1,).setValue("Valido");

  SpreadsheetApp.getUi().alert("Nuevo cliente generado satisfactoriamente");
}



function verificarDatosObligatorios(e, tipoPersona) {
  let sheet = e.source.getActiveSheet();
  let range = e.range;
  let rowEditada = range.getRow();
  let colEditada = range.getColumn();
  let ultimaColumnaPermitida = 21; // Actualizado para reflejar el número de columnas
  let columnasObligatorias = [];
  let todasLasColumnas = [ 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23];

  if (tipoPersona === "") {
    Logger.log("Vacio hizo edicion no en tipoPersona, cogemos el viejo");
    tipoPersona = sheet.getRange("D" + String(rowEditada)).getValue(); // Columna 4 para Tipo Persona
  }

  if (tipoPersona === "Natural") {
    columnasObligatorias = [2, 3, 4, 5, 6,7, 8, 10, 12, 14, 17, 18,19, 21]; // Incluyendo "Nombre cliente" (columna 2)
  } else if (tipoPersona === "Juridica") {
    columnasObligatorias = [2, 3, 4, 5, 6, 7, 8, 9, 14, 17, 18,19, 21, 22, 23]; // Incluyendo "Nombre cliente" (columna 2)
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


function crearCliente(){
  showNuevaCliente()

}

function getCustomerInformation(cliente) {
  let celdaCliente = datos_sheet.getRange("H2");
  celdaCliente.setValue(cliente);

  //Codigo de cliente
  var codigoCliente = datos_sheet.getRange("I2").getValue();

  //Tipo de identificacion 
  let tipoIdentificacion = datos_sheet.getRange("J2").getValue();

  //Numero de identificacion
  var numeroIdentificacion = datos_sheet.getRange("K2").getValue();//numero de identificacion
  
  //Digito de verificacion
  var DV = 0;

  //Direccion
  var direccion = datos_sheet.getRange("T2").getValue();
  
  //Telefono
  var telefono = datos_sheet.getRange("V2").getValue();

  //Email
  var email = datos_sheet.getRange("X2").getValue();

  //Sitio web
  var webSiteURI = datos_sheet.getRange("W2").getValue();

  //Pais
  var paisCliente = datos_sheet.getRange("S2").getValue();

  //Departamento
  var departamentoCliente= datos_sheet.getRange("AA2").getValue();

  //Municipio
  var municipioCliente= (datos_sheet.getRange("Z2").getValue()).toUpperCase();

  //Codigo tipo persona
  var tipoPersona = datos_sheet.getRange("L2").getValue();

  //Codigo regimen
  var regimen = datos_sheet.getRange("M2").getValue();

  //Detalles tributarios
  var detallesTributarios = datos_sheet.getRange("AB2").getValue();

  //Responsabilidad fiscal
  var responsabilidadFiscal = datos_sheet.getRange("AC2").getValue();


  if (tipoIdentificacion == "#NUM!") {
    Browser.msgBox("ERROR: Seleccione Tipo de Identificacion en Clientes")
    return;
  }


  var CustomerInformation = {
    "IdentificationType": Number(tiposDocumento[tipoIdentificacion]),
    "Identification": Number(numeroIdentificacion),
    "DV": DV,
    "RegistrationName": cliente,
    "CountryCode": paisesCodigos[paisCliente],
    "CountryName": paisCliente,
    "SubdivisionCode": String(departamentosCodigos[departamentoCliente]),
    "SubdivisionName": departamentoCliente,
    "CityCode": String(municipiosCodigos[municipioCliente]),
    "CityName": municipioCliente,
    "AddressLine": String(direccion),
    "Telephone": String(telefono),
    "Email": email,
    "CustomerCode": String(codigoCliente),
    "AdditionalAccountID": Number(tiposPersona[tipoPersona]),
    "TaxLevelCodeListName": codigosRegimenes[regimen],
    "PostalZone": String(datos_sheet.getRange("U2").getValue()),
    "TaxSchemeCode": detallesTributariosLib[detallesTributarios],
    "TaxSchemeName": detallesTributarios,
    "FiscalResponsabilities": responsabilidadFiscalLib[responsabilidadFiscal],
    "PartecipationPercent": 100,
    "AdditionalCustomer": []
  }
  return CustomerInformation;
}

function obtenerInformacionCliente(cliente) {
  let celdaCliente = datos_sheet.getRange("H2");
  celdaCliente.setValue(cliente);

  let codigoCliente = datos_sheet.getRange("I2").getValue();
  let direccion = datos_sheet.getRange("T2").getValue();
  let pais = datos_sheet.getRange("S2").getValue();
  let departamento = datos_sheet.getRange("AA2").getValue();
  let municipio = datos_sheet.getRange("Z2").getValue();
  let telefono = datos_sheet.getRange("V2").getValue();
  let estado = datos_sheet.getRange("Y2").getValue();

  let ubicacion = municipio + ", " + departamento + ", " + pais;

  let informacionCliente = {
    "Código cliente": codigoCliente,
    "Dirección": direccion,
    "Ubicación": ubicacion,
    "Teléfono": telefono,
    "Estado": estado
  };

  return informacionCliente;
}

