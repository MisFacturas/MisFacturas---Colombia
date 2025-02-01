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

//Esta funcion se llama desde el sheets
function crearCliente() {
  showNuevaClienteV2()
}

function showActivarCliente() {
  var html = HtmlService.createHtmlOutputFromFile('menuActivarCliente').setTitle("Activar Cliente")
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function inactivarCliente(cliente) {
  let spreadsheet = SpreadsheetApp.getActive();
  let datos_sheet = spreadsheet.getSheetByName('Datos');
  let hojaClientesInactivos = spreadsheet.getSheetByName('ClientesInvalidos');
  let hojaClientes = spreadsheet.getSheetByName("Clientes")
  Logger.log(cliente)
  datos_sheet.getRange("H2").setValue(cliente)


  let rowDelCliente = datos_sheet.getRange("G2").getValue();
  let rowMaximaClientesInactivos = hojaClientesInactivos.getLastRow() + 1;
  let rowMaximaClientes = hojaClientes.getLastRow() + 1;

  let codigoCliente = datos_sheet.getRange("I2").getValue();
  let tipoDoc = datos_sheet.getRange("J2").getValue();
  let numIdentificacion = datos_sheet.getRange("K2").getValue();
  let tipoPersona = datos_sheet.getRange("L2").getValue();
  let regimen = datos_sheet.getRange("M2").getValue();
  let nombreComercial = datos_sheet.getRange("N2").getValue();
  let primerNombre = datos_sheet.getRange("O2").getValue();
  let segundoNombre = datos_sheet.getRange("P2").getValue();
  let primerApellido = datos_sheet.getRange("Q2").getValue();
  let segundoApellido = datos_sheet.getRange("R2").getValue();
  let pais = datos_sheet.getRange("S2").getValue();
  let direccion = datos_sheet.getRange("T2").getValue();
  let codigoPostal = datos_sheet.getRange("U2").getValue();
  let telefono = datos_sheet.getRange("V2").getValue();
  let sitioWeb = datos_sheet.getRange("W2").getValue();
  let email = datos_sheet.getRange("X2").getValue();
  let estado = datos_sheet.getRange("Y2").getValue();
  let municipio = datos_sheet.getRange("AZ2").getValue();
  let departamento = datos_sheet.getRange("AA2").getValue();
  let detallesTributarios = datos_sheet.getRange("AB2").getValue();
  let responsabilidadFiscal = datos_sheet.getRange("AC2").getValue();
  let tipoTercero = datos_sheet.getRange("AD2").getValue();
  let identificadorUnico = datos_sheet.getRange("AE2").getValue();



  // Proceso para agregar a la hoja de clientes inactivos
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 1).setValue(estado);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 2).setValue(tipoTercero);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 3).setValue(tipoPersona);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 4).setValue(nombreComercial);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 5).setValue(primerNombre);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 6).setValue(segundoNombre);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 7).setValue(primerApellido);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 8).setValue(segundoApellido);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 9).setValue(tipoDoc);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 10).setValue(numIdentificacion);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 11).setValue(codigoCliente);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 12).setValue(regimen);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 13).setValue(pais);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 14).setValue(departamento);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 15).setValue(municipio);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 16).setValue(direccion);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 17).setValue(codigoPostal);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 18).setValue(telefono);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 19).setValue(sitioWeb);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 20).setValue(email);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 21).setValue(detallesTributarios);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 22).setValue(responsabilidadFiscal);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 23).setValue(identificadorUnico);


  //eliminar cliente de la hoja clientes

  hojaClientes.deleteRow(rowDelCliente)
  hojaClientes.insertRowAfter(rowMaximaClientes)
  SpreadsheetApp.getUi().alert("El cliente se ha inactivado satisfactoriamente");
}

function activarCliente(cliente) {
  let spreadsheet = SpreadsheetApp.getActive();
  let datos_sheet = spreadsheet.getSheetByName('Datos');
  let hojaClientesInactivos = spreadsheet.getSheetByName('ClientesInvalidos');
  let hojaClientes = spreadsheet.getSheetByName("Clientes")
  Logger.log(cliente)
  datos_sheet.getRange("H6").setValue(cliente)

  let rowDelCliente = datos_sheet.getRange("G6").getValue();
  let rowMaximaClientesInactivos = hojaClientesInactivos.getLastRow() + 1;
  let rowMaximaClientes = hojaClientes.getLastRow() + 1;

  let estado = datos_sheet.getRange('I6').getValue();
  let tipoTercero = datos_sheet.getRange('J6').getValue();
  let tipoPersona = datos_sheet.getRange('K6').getValue();
  let nombreComercial = datos_sheet.getRange('L6').getValue();
  let primerNombre = datos_sheet.getRange('M6').getValue();
  let segundoNombre = datos_sheet.getRange('N6').getValue();
  let primerApellido = datos_sheet.getRange('O6').getValue();
  let segundoApellido = datos_sheet.getRange('P6').getValue();
  let tipoDoc = datos_sheet.getRange('Q6').getValue();
  let numIdentificacion = datos_sheet.getRange('R6').getValue();
  let codigoCliente = datos_sheet.getRange('S6').getValue();
  let regimen = datos_sheet.getRange('T6').getValue();
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
  let identificadorUnico = datos_sheet.getRange('AE6').getValue();

  hojaClientes.getRange(rowMaximaClientes, 1).setValue(estado);
  hojaClientes.getRange(rowMaximaClientes, 2).setValue(tipoTercero);
  hojaClientes.getRange(rowMaximaClientes, 3).setValue(tipoPersona);
  hojaClientes.getRange(rowMaximaClientes, 4).setValue(nombreComercial);
  hojaClientes.getRange(rowMaximaClientes, 5).setValue(primerNombre);
  hojaClientes.getRange(rowMaximaClientes, 6).setValue(segundoNombre);
  hojaClientes.getRange(rowMaximaClientes, 7).setValue(primerApellido);
  hojaClientes.getRange(rowMaximaClientes, 8).setValue(segundoApellido);
  hojaClientes.getRange(rowMaximaClientes, 9).setValue(tipoDoc);
  hojaClientes.getRange(rowMaximaClientes, 10).setValue(numIdentificacion);
  hojaClientes.getRange(rowMaximaClientes, 11).setValue(codigoCliente);
  hojaClientes.getRange(rowMaximaClientes, 12).setValue(regimen);
  hojaClientes.getRange(rowMaximaClientes, 13).setValue(pais);
  hojaClientes.getRange(rowMaximaClientes, 14).setValue(departamento);
  hojaClientes.getRange(rowMaximaClientes, 15).setValue(municipio);
  hojaClientes.getRange(rowMaximaClientes, 16).setValue(direccion);
  hojaClientes.getRange(rowMaximaClientes, 17).setValue(codigoPostal);
  hojaClientes.getRange(rowMaximaClientes, 18).setValue(telefono);
  hojaClientes.getRange(rowMaximaClientes, 19).setValue(sitioWeb);
  hojaClientes.getRange(rowMaximaClientes, 20).setValue(email);
  hojaClientes.getRange(rowMaximaClientes, 21).setValue(detallesTributarios);
  hojaClientes.getRange(rowMaximaClientes, 22).setValue(responsabilidadFiscal);
  hojaClientes.getRange(rowMaximaClientes, 23).setValue(identificadorUnico);

  hojaClientesInactivos.deleteRow(rowDelCliente)
  hojaClientesInactivos.insertRowAfter(rowMaximaClientesInactivos)
}

function buscarClientes(terminoBusqueda, hojaA) {
  let spreadsheet = SpreadsheetApp.getActive();
  var resultados = [];

  if (hojaA === "Inactivar") {
    var sheet = spreadsheet.getSheetByName('Clientes');
    var ultimaFila = sheet.getLastRow();
    var valores = sheet.getRange(2, 1, ultimaFila - 1, 23).getValues(); // Obtener todas las columnas desde la fila 2

    if (terminoBusqueda === "") {
      return resultados;
    }

    // Recorre los valores obtenidos
    for (var i = 0; i < valores.length; i++) {
      var codigoCliente = valores[i][10]; // Columna K
      var nombreComercial = valores[i][3]; // Columna D
      var primerNombre = valores[i][4]; // Columna E
      var tipoDocumento = String(valores[i][8]); // Columna I
      var numeroIdentificacion = String(valores[i][9]); // Columna J
      var identificadorUnico = valores[i][22]; // Columna W

      // Comprueba si el valor coincide con el término de búsqueda
      if (
        (codigoCliente && String(codigoCliente).toLowerCase().indexOf(terminoBusqueda.toLowerCase()) !== -1) ||
        (nombreComercial && nombreComercial.toLowerCase().indexOf(terminoBusqueda.toLowerCase()) !== -1) ||
        (primerNombre && primerNombre.toLowerCase().indexOf(terminoBusqueda.toLowerCase()) !== -1) ||
        (tipoDocumento && tipoDocumento.toLowerCase().indexOf(terminoBusqueda.toLowerCase()) !== -1) ||
        (numeroIdentificacion && numeroIdentificacion.toLowerCase().indexOf(terminoBusqueda.toLowerCase()) !== -1)
      ) {
        resultados.push(identificadorUnico); // Añade el valor de la columna W a la lista de resultados si coincide
      }
    }
  } else {
    var sheet = spreadsheet.getSheetByName('ClientesInvalidos');
    var ultimaFila = sheet.getLastRow();
    var valores = sheet.getRange(2, 1, ultimaFila - 1, 23).getValues(); // Obtener todas las columnas desde la fila 2

    if (terminoBusqueda === "") {
      return resultados;
    }

    // Recorre los valores obtenidos
    for (var i = 0; i < valores.length; i++) {
      var codigoCliente = valores[i][10]; // Columna K
      var nombreComercial = valores[i][3]; // Columna D
      var primerNombre = valores[i][4]; // Columna E
      var tipoDocumento = String(valores[i][9]); // Columna J
      var numeroIdentificacion = String(valores[i][10]); // Columna K
      var identificadorUnico = valores[i][22]; // Columna W

      // Comprueba si el valor coincide con el término de búsqueda
      if (
        (codigoCliente && String(codigoCliente).toLowerCase().indexOf(terminoBusqueda.toLowerCase()) !== -1) ||
        (nombreComercial && nombreComercial.toLowerCase().indexOf(terminoBusqueda.toLowerCase()) !== -1) ||
        (primerNombre && primerNombre.toLowerCase().indexOf(terminoBusqueda.toLowerCase()) !== -1) ||
        (tipoDocumento && tipoDocumento.toLowerCase().indexOf(terminoBusqueda.toLowerCase()) !== -1) ||
        (numeroIdentificacion && numeroIdentificacion.toLowerCase().indexOf(terminoBusqueda.toLowerCase()) !== -1)
      ) {
        resultados.push(identificadorUnico); // Añade el valor de la columna W a la lista de resultados si coincide
      }
    }
  }

  Logger.log(resultados);
  // Devuelve los resultados
  return resultados;
}

function buscarPaises(terminoBusqueda) {
  let spreadsheet = SpreadsheetApp.getActive();
  let datos_sheet = spreadsheet.getSheetByName('Datos');
  let paises = datos_sheet.getRange(25, 1, 195, 1).getValues();
  var resultados = [];
  if (terminoBusqueda === "") {
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

function obtenerTipoDePersona(e) {
  let sheet = e.source.getActiveSheet();
  let range = e.range;
  let rowEditada = range.getRow();
  let colEditada = 3;

  let tipoPersona = sheet.getRange(rowEditada, colEditada).getValue()
  Logger.log(tipoPersona)
  return tipoPersona

}

function saveClientData(formData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clientes');
  if (!sheet) {
    throw new Error('La hoja "Clientes" no existe.');
  }
  let codigosIdetificadores = formData.numeroIdentificacion + "-" + formData.codigoCliente;

  let existe = verificarIdentificacionUnica(codigosIdetificadores, "Clientes", false)
  if (existe === 1) {
    SpreadsheetApp.getUi().alert("El Numero de Identificacion del cliente ya existe, por favor poner un Numero de Identificacion unico");
    throw new Error('por favor poner un Numero de Identificacion unico');
  } else if (existe === 2) {
    SpreadsheetApp.getUi().alert("El Codigo del cliente ya existe, por favor poner un Codigo de cliente unico");
    throw new Error('por favor poner un Codigo de cliente unico');
  }
  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(2, 2, lastRow, 22).getValues(); // Obtener desde la columna B hasta la S (19 columnas)

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
    formData.tipoTercero,
    formData.tipoPersona,
    formData.nombreComercial,
    formData.primerNombre,
    formData.segundoNombre,
    formData.primerApellido,
    formData.segundoApellido,
    formData.tipoDocumento,
    formData.numeroIdentificacion,
    formData.codigoCliente,
    formData.regimen,
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
  let identificadorUnico = "";
  Logger.log(formData.tipoPersona)
  if (formData.tipoPersona === "Natural") {
    identificadorUnico = formData.primerNombre + " " + formData.primerApellido + "-" + formData.numeroIdentificacion;
  } else {
    identificadorUnico = formData.nombreComercial + "-" + formData.numeroIdentificacion;
  }
  sheet.getRange(emptyRow, 23).setValue(identificadorUnico);

  SpreadsheetApp.getUi().alert("Nuevo cliente generado satisfactoriamente");
}

function verificarDatosObligatorios(e, tipoPersona) {
  let sheet = e.source.getActiveSheet();
  let range = e.range;
  let rowEditada = range.getRow();
  let colEditada = range.getColumn();
  let ultimaColumnaPermitida = 22; // Actualizado para reflejar el número de columnas
  let columnasObligatorias = [];
  let todasLasColumnas = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22];

  if (tipoPersona === "") {
    Logger.log("Vacio hizo edicion no en tipoPersona, cogemos el viejo");
    tipoPersona = sheet.getRange("C" + String(rowEditada)).getValue(); // Columna 4 para Tipo Persona
  }

  if (tipoPersona === "Natural") {
    columnasObligatorias = [5, 7, 9, 10, 11, 12, 13, 16, 17, 18, 20, 21, 22];
  } else if (tipoPersona === "Juridica") {
    columnasObligatorias = [4, 9, 10, 11, 12, 13, 16, 17, 18, 20, 21, 22];
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

    // Verificar si el país es Colombia y validar departamento y municipio
    let pais = sheet.getRange(rowEditada, 13).getValue(); // Columna M es la 13
    let departamento = sheet.getRange(rowEditada, 14).getValue(); // Columna N es la 14
    let municipio = sheet.getRange(rowEditada, 15).getValue(); // Columna O es la 15

    if (pais === "Colombia") {
      if (!departamento || !municipio) {
        estaCompleto = false;
        if (!departamento) {
          sheet.getRange(rowEditada, 14).setBackground('#FFC7C7'); // Resaltar en rojo claro
        }
        if (!municipio) {
          sheet.getRange(rowEditada, 15).setBackground('#FFC7C7'); // Resaltar en rojo claro
        }
      }
    }

    // Verificar si el correo electrónico en la columna T es válido
    let email = sheet.getRange(rowEditada, 20).getValue(); // Columna T es la 20
    let emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

    if (!emailRegex.test(email)) {
      estaCompleto = false;
      sheet.getRange(rowEditada, 20).setBackground('#FFC7C7'); // Resaltar en rojo claro
      SpreadsheetApp.getUi().alert('El correo electrónico ingresado no es válido.');
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

function showNuevaClienteV2() {
  var html = HtmlService.createHtmlOutputFromFile('menuAgregarClienteFactura').setTitle("Nuevo Cliente")
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function getCityCode(departamentoCliente, ciudadCliente) {
  switch (departamentoCliente) {
    case "Amazonas":
      return municipiosAmazonas[ciudadCliente];
    case "Antioquia":
      return municipiosAntioquia[ciudadCliente];
    case "Arauca":
      return municipiosArauca[ciudadCliente];
    case "Atlantico":
      return municipiosAtlantico[ciudadCliente];
    case "Bogota":
      return municipiosBogota[ciudadCliente];
    case "Bolivar":
      return municipiosBolivar[ciudadCliente];
    case "Boyaca":
      return municipiosBoyaca[ciudadCliente];
    case "Caldas":
      return municipiosCaldas[ciudadCliente];
    case "Caqueta":
      return municipiosCaqueta[ciudadCliente];
    case "Casanare":
      return municipiosCasanare[ciudadCliente];
    case "Cauca":
      return municipiosCauca[ciudadCliente];
    case "Cesar":
      return municipiosCesar[ciudadCliente];
    case "Choco":
      return municipiosChoco[ciudadCliente];
    case "Cordoba":
      return municipiosCordoba[ciudadCliente];
    case "Cundinamarca":
      return municipiosCundinamarca[ciudadCliente];
    case "Guainia":
      return municipiosGuainia[ciudadCliente];
    case "Guaviare":
      return municipiosGuaviare[ciudadCliente];
    case "Huila":
      return municipiosHuila[ciudadCliente];
    case "La Guajira":
      return municipiosLaGuajira[ciudadCliente];
    case "Magdalena":
      return municipiosMagdalena[ciudadCliente];
    case "Meta":
      return municipiosMeta[ciudadCliente];
    case "Narino":
      return municipiosNarino[ciudadCliente];
    case "Norte de Santander":
      return municipiosNorteDeSantander[ciudadCliente];
    case "Putumayo":
      return municipiosPutumayo[ciudadCliente];
    case "Quindio":
      return municipiosQuindio[ciudadCliente];
    case "Risaralda":
      return municipiosRisaralda[ciudadCliente];
    case "San Andres y Providencia":
      return municipiosSanAndresYProvidencia[ciudadCliente];
    case "Santander":
      return municipiosSantander[ciudadCliente];
    case "Sucre":
      return municipiosSucre[ciudadCliente];
    case "Tolima":
      return municipiosTolima[ciudadCliente];
    case "Valle del Cauca":
      return municipiosValleDelCauca[ciudadCliente];
    case "Vaupes":
      return municipiosVaupes[ciudadCliente];
    case "Vichada":
      return municipiosVichada[ciudadCliente];
    default:
      return null;
  }
}

function getCustomerInformation(cliente) {
  let spreadsheet = SpreadsheetApp.getActive();
  let datos_sheet = spreadsheet.getSheetByName('Datos');
  let celdaCliente = datos_sheet.getRange("H2");
  celdaCliente.setValue(cliente);

  //Codigo de cliente
  var codigoCliente = datos_sheet.getRange("I2").getValue();

  //Tipo de identificacion 
  let tipoIdentificacion = datos_sheet.getRange("J2").getValue();

  //Numero de identificacion
  var numeroIdentificacion = datos_sheet.getRange("K2").getValue();//numero de identificacion

  //Digito de verificacion
  function calcularDV(tipoIdentificacion) {
    var DV = 0;
    if (tipoIdentificacion === "NIT") {
      DV = Number(datos_sheet.getRange("H50").getValue());
    }
    return DV;
  }


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
  var departamentoCliente = datos_sheet.getRange("AA2").getValue();

  //Municipio
  var municipioCliente = (datos_sheet.getRange("Z2").getValue()).toUpperCase();

  // Get city code using the new function
  var cityCode = getCityCode(departamentoCliente, municipioCliente);

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
    "DV": calcularDV(tipoIdentificacion),
    "RegistrationName": cliente,
    "CountryCode": paisesCodigos[paisCliente],
    "CountryName": paisCliente,
    "SubdivisionCode": String(departamentosCodigos[departamentoCliente]),
    "SubdivisionName": departamentoCliente,
    "CityCode": String(cityCode),
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
  let spreadsheet = SpreadsheetApp.getActive();
  let datos_sheet = spreadsheet.getSheetByName('Datos');
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

