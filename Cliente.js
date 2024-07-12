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