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
          padding: 15px;
        }
        .alert-info {
          background-color: #d1ecf1;
          border: 1px solid #bee5eb;
          border-radius: 5px;
          padding: 12px;
          margin-bottom: 15px;
        }
        .alert-heading {
          color: #0c5460;
          font-size: 16px;
          font-weight: bold;
          margin-bottom: 8px;
        }
        .alert-text {
          color: #0c5460;
          font-size: 14px;
          margin-bottom: 8px;
        }
        .requirement-text {
          color: #721c24;
          font-weight: bold;
          font-size: 14px;
        }
        .reminder-text {
          color: #856404;
          background-color: #fff3cd;
          border: 1px solid #ffeaa7;
          border-radius: 5px;
          padding: 10px;
          margin-bottom: 15px;
          font-size: 14px;
          font-weight: 600;
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
        <div class="reminder-text">
          ⚠️ Recuerda tener activo el plan Google Sheets™ en tu cuenta de misfacturas.
        </div>
        <p>Por favor <b>vincula tu cuenta</b> para poder generar tus facturas.</p>
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
      .btn-orange {
        background-color: #ff6a00;
        border: 2px solid #ff6a00;
        color: #fff;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 140px;
        padding: 10px 0;
        font-size: 1rem;
        border-radius: 8px;
        text-decoration: none;
        margin: 0 10px;
        font-weight: 600;
        transition: background 0.2s, color 0.2s, border 0.2s;
      }
      .btn-orange:hover {
        background-color: #cb4a22;
        border-color: #cb4a22;
        color: #fff;
      }
      .btn-grey-outline {
        background: #fff;
        border: 2px solid #888;
        color: #888;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 140px;
        padding: 10px 0;
        font-size: 1rem;
        border-radius: 8px;
        text-decoration: none;
        margin: 0 10px;
        font-weight: 600;
        transition: background 0.2s, color 0.2s, border 0.2s;
      }
      .btn-grey-outline:hover {
        background: #888;
        color: #fff;
        border-color: #888;
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
        <button class="btn-grey-outline" onclick="google.script.run.modificarFactura(); google.script.host.close()">Editar</button>
        <button class="btn-orange" onclick="google.script.run.enviarFacturaHtml(); google.script.host.close()">Enviar</button>
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

function plantillaEstadoEnRevision() {
  return `
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,100..900;1,100..900&display=swap');
      .estado-enrevision-box {
        background: white;
        border-radius: 10px;
        padding: 30px;
        max-width: 500px;
        width: 100%;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
        font-family: 'Roboto', sans-serif;
        margin: 0 auto;
      }
      .popup-header {
        text-align: center;
        margin-bottom: 20px;
      }
      .popup-icon {
        font-size: 3rem;
        color: #ffc107;
        margin-bottom: 15px;
      }
      .popup-title {
        color: #333;
        font-size: 1.5rem;
        font-weight: bold;
        margin-bottom: 10px;
      }
      .popup-subtitle {
        color: #666;
        font-size: 1.1rem;
        margin-bottom: 20px;
      }
      .popup-body {
        margin-bottom: 25px;
      }
      .popup-text {
        color: #555;
        line-height: 1.6;
        margin-bottom: 15px;
      }
      .popup-highlight {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 5px;
        padding: 15px;
        margin: 15px 0;
      }
      .popup-highlight strong {
        color: #856404;
      }
      .popup-actions {
        text-align: center;
      }
      .btn-orange {
        background-color: #ff6a00;
        border: 2px solid #ff6a00;
        color: #fff;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 180px;
        padding: 12px 0;
        font-size: 1rem;
        border-radius: 8px;
        text-decoration: none;
        margin: 0 10px;
        font-weight: 600;
        transition: background 0.2s, color 0.2s, border 0.2s;
      }
      .btn-orange:hover {
        background-color: #cb4a22;
        border-color: #cb4a22;
        color: #fff;
      }
      .btn-orange-outline {
        background: #fff;
        border: 2px solid #cb4a22;
        color: #cb4a22;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 180px;
        padding: 12px 0;
        font-size: 1rem;
        border-radius: 8px;
        text-decoration: none;
        margin: 0 10px;
        font-weight: 600;
        transition: background 0.2s, color 0.2s, border 0.2s;
      }
      .btn-orange-outline:hover {
        background: #cb4a22;
        color: #fff;
        border-color: #cb4a22;
      }
    </style>
    <div class="estado-enrevision-box">
      <div class="popup-header">
        <div class="popup-icon">
          <i class="icon-24-outlined-action-main-info"></i>
        </div>
      </div>
      <div class="popup-body">
        <p class="popup-text">
          Si tus facturas presentan el estado <strong>"En revisión"</strong>, esto indica que quedaron como <strong>Prefacturas y falta firmarlas</strong>.
        </p>
        <div class="popup-highlight">
          <strong>¿Qué debes hacer?</strong><br>
          Tienes que ir a la web de misfacturas y firmar tus facturas manualmente para completar el proceso de facturación.
        </div>
        <p class="popup-text">
          Una vez que las firmes, el estado cambiará a <strong>"Enviada"</strong> y podrás descargar tus facturas desde el historial.
        </p>
      </div>
      <div class="popup-actions">
        <a href="#" class="btn-orange-outline" onclick="window.open('https://www.misfacturas.com.co', '_blank')">
          Ir a MisFacturas
        </a>
        <a href="#" class="btn-orange" onclick="google.script.host.close()">
          Entendido
        </a>
      </div>
    </div>
  `;
}

function plantillaAvisoDescargaFactura() {
  return `
    <style>
      .modal-msg-container {padding: 30px; max-width: 400px; font-family: Roboto, sans-serif; text-align: center;}
      .btn-orange {background-color: #ff6a00; border: none; color: #fff; display: inline-flex; align-items: center; justify-content: center; width: 140px; padding: 10px 0; font-size: 1rem; border-radius: 8px; text-decoration: none; margin: 20px auto 0 auto; font-weight: 600; transition: background 0.2s, color 0.2s, border 0.2s;}
      .btn-orange:hover {background-color: #cb4a22; border-color: #cb4a22; color: #fff;}
    </style>
    <div class="modal-msg-container">
      <p>Para descargar una factura, ingresa el prefijo y número de la factura en el menú del costado derecho.</p>
      <button class="btn-orange" onclick="google.script.host.close()">Entendido</button>
    </div>
    <script>document.querySelector('.btn-orange').focus();</script>
  `;
}