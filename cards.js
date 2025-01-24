/**
 * Callback for rendering the homepage card.
 * @return {CardService.Card} The card to show to the user.
 */
function onHomepage(e) {
  console.log(e);
  
  console.log("onHomepage Enters");
  var message;

  message = 'Para empezar, haz clic en el menu de la parte superior llamado <b>misfacturas</b> y selecciona la opción <b>Inicio</b>.  '
  message += '<br>Luego podrás navegar en el menú que se desplegará en al parte derecha de tu pantalla';
  return createCard(message, true);
}

/**
 * Creates a card with an image of FacturasApp, and the text.
 * @param {String} text The text to overlay on the image.
 * @param {Boolean} isHomepage True if the card created here is a homepage;
 *      false otherwise. Defaults to false.
 * @return {CardService.Card} The assembled card.
 */
function createCard(text, isHomepage) {
  // Explicitly set the value of isHomepage as false if null or undefined.
  if (!isHomepage) {
    isHomepage = false;
  }
  console.log("createCard");
  
  var imageUrl = 'https://misfacturas.cenet.ws/Publico/google-sheets-resources/images/logoAssistant.png' 
  var image = CardService.newImage()
      .setImageUrl(imageUrl)
      .setAltText('misfacturas');

var multilineDecoratedText = CardService.newDecoratedText()
    .setText(text)
    .setWrapText(true);

  // Create a footer to be shown at the bottom.
  var footer = CardService.newFixedFooter()
      .setPrimaryButton(CardService.newTextButton()
          .setText('Powered by misfacturas.com')
          .setOpenLink(CardService.newOpenLink()
              .setUrl('https://www.misfacturas.com.co/')));

  // Assemble the widgets and return the card.
  var section = CardService.newCardSection()
      .addWidget(image)
      .addWidget(multilineDecoratedText); 

      
  var card = CardService.newCardBuilder()
      .addSection(section)
      .setFixedFooter(footer);

  if (!isHomepage) {
    // Create the header shown when the card is minimized,
    // but only when this card is a contextual card. Peek headers
    // are never used by non-contexual cards like homepages.
    var peekHeader = CardService.newCardHeader()
      .setTitle('Contextual Card')
      .setImageUrl('https://www.gstatic.com/images/icons/material/system/1x/pets_black_48dp.png')
      .setSubtitle(text);
    card.setPeekCardHeader(peekHeader)
  }

  return card.build();
}



  