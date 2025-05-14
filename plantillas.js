function plantillaVincularMF(inHoja) {
    return `
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,100..900;1,100..900&display=swap');
        body {
          font-family: 'Roboto', sans-serif;
          font-size: 16px;
          margin: 0;
          padding: 0;
        }
        .container {
          padding: 10px;
        }
        .button-container {
          display: flex;
          justify-content: flex-end;
          margin-top: 20px;
        }
        .red-text {
          color: rgb(231, 112, 14);
          font-size: 16px;
          font-family: 'Roboto', sans-serif;
          font-weight: 600;
        }
        button {
          padding: 3px 12px;
          font-family: 'Roboto', sans-serif;
          background-color: rgba(255, 255, 255, 0);
          border: none;
          border-radius: 30px;
          cursor: pointer;
        }
        button:hover {
          background-color:rgba(255, 218, 187, 0.32);
        }
      </style>
      <div class="container">
        <p>Por favor <b>vincule su cuenta</b> para poder generar las facturas.</p>
        <div class="button-container">
          <button onclick="google.script.run.abrirMenuVinculacion(${inHoja}); google.script.host.close()"><p class="red-text">Aceptar<p></button>
        </div>
      </div>
    `;
}

function plantillaResumenFactura(nombreCliente, numeroFactura, impuestos, invoiceTotal) {
  return `
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,100..900;1,100..900&display=swap');
      body {
        font-family: 'Roboto', sans-serif;
        font-size: 16px;
        margin: 0;
        padding: 0;
      }
      .container {
        padding: 10px;
        display: flex;
        flex-direction: column;
        height: auto; /* Cambiar de 100vh a auto */
      }
      .title {
        font-size: 18px;
        font-weight: bold;
        margin-bottom: 10px;
      }
      .info {
        margin-bottom: 10px;
      }
      .info span {
        font-weight: bold;
      }
      .columns {
        display: flex;
        justify-content: space-between;
        gap: 20px; /* Añadir espacio entre las columnas */
      }
      .column {
        flex: 1;
        padding: 10px;
        margin-right: 10px;
        box-sizing: border-box; /* Asegurar que el padding no afecte el ancho total */
      }
      .column:last-child {
        margin-right: 0;
      }
      .button-container {
        display: flex;
        justify-content: space-between;
        margin-top: 20px;
      }
      .red-text {
        color: rgb(231, 112, 14);
        font-size: 16px;
        font-family: 'Roboto', sans-serif;
        font-weight: 600;
      }
      .grey-text {
        color: grey;
        font-size: 16px;
        font-family: 'Roboto', sans-serif;
        font-weight: 600;
      }
      button {
        padding: 3px 12px;
        font-family: 'Roboto', sans-serif;
        background-color: rgba(255, 255, 255, 0);
        border: none;
        border-radius: 30px;
        cursor: pointer;
      }
      button:hover {
        background-color:rgba(255, 218, 187, 0.32);
      }
      ul {
        padding: 0;
        list-style-type: none;
      }
      li {
        display: flex;
        justify-content: space-between;
        padding: 5px 0;
      }
    </style>
    <div class="container">
      <div class="columns">
        <div class="column">
          <div class="info"><span>Nombre del Cliente:</span> ${nombreCliente}</div>
          <div class="info"><span>Número de la Factura:</span> ${numeroFactura}</div>
          <div class="info"><span>Impuestos:</span></div>
          <ul>
            ${impuestos.map(function (impuesto) {
              return `<li><span>${impuesto.tipo} (${impuesto.percent}%):</span> <span>${formatearPesos(impuesto.amount)}</span></li>`;
            }).join('')}
          </ul>
        </div>
        <div class="column">
          <div class="info"><span>Totales de la Factura:</span></div>
          <ul>
            <li><span>Subtotal:</span> <span>${formatearPesos(invoiceTotal.lineExtensionamount)}</span></li>
            <li><span>Impuestos Excluidos:</span> <span>${formatearPesos(invoiceTotal.TaxExclusiveAmount)}</span></li>
            <li><span>Impuestos Incluidos:</span> <span>${formatearPesos(invoiceTotal.TaxInclusiveAmount)}</span></li>
            <li><span>Descuentos:</span> <span>${formatearPesos(invoiceTotal.AllowanceTotalAmount)}</span></li>
            <li><span>Cargos:</span> <span>${formatearPesos(invoiceTotal.ChargeTotalAmount)}</span></li>
            <li><span>Pagos Anticipados:</span> <span>${formatearPesos(invoiceTotal.PrePaidAmount)}</span></li>
            <li><span>Total a Pagar:</span> <span>${formatearPesos(invoiceTotal.PayableAmount)}</span></li>
          </ul>
        </div>
      </div>
      <div class="button-container">
        <button onclick="google.script.run.modificarFactura(); google.script.host.close()"><p class="grey-text">Editar</p></button>
        <button onclick="google.script.run.enviarFacturaHtml(); google.script.host.close()"><p class="red-text">Enviar</p></button>     
      </div>
    </div>
  `;
}

function formatearPesos(valor) {
    return `$${valor.toLocaleString('es-CO')}`;
}

function plantillaCambiarAmbiente() {
  return `
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,100..900;1,100..900&display=swap');
      body {
        font-family: 'Roboto', sans-serif;
        font-size: 16px;
        margin: 0;
        padding: 0;
      }
      .container {
        padding: 20px;
      }
      .title {
        font-size: 18px;
        font-weight: bold;
        margin-bottom: 20px;
      }
      .options {
        display: flex;
        flex-direction: column;
        gap: 15px;
        margin-bottom: 20px;
      }
      .option {
        display: flex;
        align-items: center;
        cursor: pointer;
        padding: 10px;
        border-radius: 5px;
        transition: background-color 0.2s;
      }
      .option:hover {
        background-color: rgba(255, 218, 187, 0.32);
      }
      .option input[type="radio"] {
        margin-right: 10px;
      }
      .button-container {
        display: flex;
        justify-content: flex-end;
        margin-top: 30px;
      }
      .red-text {
        color: rgb(231, 112, 14);
        font-size: 16px;
        font-family: 'Roboto', sans-serif;
        font-weight: 600;
      }
      .grey-text {
        color: grey;
        font-size: 16px;
        font-family: 'Roboto', sans-serif;
        font-weight: 600;
      }
      button {
        padding: 3px 12px;
        font-family: 'Roboto', sans-serif;
        background-color: rgba(255, 255, 255, 0);
        border: none;
        border-radius: 30px;
        cursor: pointer;
        margin-left: 10px;
      }
      button:hover {
        background-color: rgba(255, 218, 187, 0.32);
      }
    </style>
    <div class="container">
      <div class="title">Selecciona el nuevo ambiente:</div>
      <div class="options">
        <label class="option">
          <input type="radio" name="ambiente" value="QA" checked> QA
        </label>
        <label class="option">
          <input type="radio" name="ambiente" value="Preproducción"> Preproducción
        </label>
        <label class="option">
          <input type="radio" name="ambiente" value="Producción"> Producción
        </label>
      </div>
      <div class="button-container">
        <button onclick="google.script.host.close()"><p class="grey-text">Cancelar</p></button>
        <button onclick="confirmarCambio(); google.script.host.close()"><p class="red-text">Confirmar</p></button>
      </div>
    </div>
    <script>
      function confirmarCambio() {
        const selectedOption = document.querySelector('input[name="ambiente"]:checked').value;
        google.script.run.withSuccessHandler(function() {
          google.script.host.close();
        }).aplicarCambioAmbiente(selectedOption);
      }
    </script>
  `;
}