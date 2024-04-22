function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Articoli Subito')
    .addItem('Leggi Articoli in vendita', 'showPrompt')
    .addToUi();
}

function showPrompt() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'Leggi Articoli in vendita',
    'Inserisci il link del venditore (Esempio: https://www.subito.it/utente/12345678):',
    ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    var idVenditore = text.match(/https:\/\/www\.subito\.it\/utente\/(\d+)/);
    if (idVenditore != null) {
      idVenditore = idVenditore[1];
      getArticoliInVendita(idVenditore);
    }
    else {
      ui.alert('Il link inserito non è valido.');
      return;
    }
  } else if (button == ui.Button.CANCEL) {
    return;
  } else if (button == ui.Button.CLOSE) {
    return;
  }
}

function getArticoliInVendita(idVenditore) {

  // Nell'url è possibile cambiare il parametro "lim=1000" con un altro numero, in base a quanti articoli si vogliono estrarre 
  const response = JSON.parse(UrlFetchApp.fetch("https://www.subito.it/hades/v1/search/items?uid=" + idVenditore + "&lim=1000").getContentText());
  var annunci = response.ads;

  var articoli = [];
  var titolo;
  var prezzo;
  var urlArticolo;
  var dettagli;

  for (let i = 0; i < annunci.length; i++) {
    prezzo = 0;
    titolo = annunci[i].subject;

    // La posizione del prezzo nell'array "features" cambia ad ogni chiamata
    dettagli = annunci[i].features;

    for (let x = 0; x < dettagli.length; x++) {
      if (dettagli[x].label == "Prezzo") {
        prezzo = dettagli[x].values[0].key;
      }
    }
    urlArticolo = annunci[i].urls.default;
    articoli.push([titolo, prezzo, urlArticolo]);
  }

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Foglio1").getRange(2, 1, articoli.length, 3).setValues(articoli);
}
