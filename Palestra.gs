// ====================================
// Versione del 4 Nov 2023
// ====================================

// colonne del foglio di lavoro
const COLUMN_NOME = "NOME"
const CLOUMN_COGNOME = "COGNOME"
const COLUMN_EMAIL = "EMAIL"
const COLUMN_FINEABBONAMENTO = "SCADENZA ABBONAMENTO"
const COLUMN_CERTIFICATOMEDICO = "SCADENZA CERTIFICATO MEDICO"
const COLUMN_MAILABBONAMENTO = "DATA INVIO MAIL ABBONAMENTO"
const COLUMN_MAILCERTIFICATO = "DATA INVIO MAIL CERTIFICATO"

const documentProperties = PropertiesService.getDocumentProperties()

/**
 * The event handler triggered when installing the add-on.
 * @param {Event} e The onInstall event.
 * @see https://developers.google.com/apps-script/guides/triggers#oninstalle
 */
function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Palestra')

      .addItem('Invia Mail Abbonamenti In Scadenza', 'fInviaMailAbbonamentiScaduti')
      .addItem('Invia Mail Certificati Mancanti / In Scadenza', 'fInviaMailCertificati')
      .addSubMenu(ui.createMenu('Impostazioni')
          .addItem('firma sulla mail', 'fFirmaSullaMail')
          .addItem('giorni invio mail', 'fGiorniInvioMail'))
      .addSeparator()
      .addItem('Aggiungi Nuovo Corso', 'fNuovoCorso')
      .addToUi();
}

function fFirmaSullaMail() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Firma Sulla Mail', 'Valore attuale: ' + getMailSignature(), ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    documentProperties.setProperty('MAIL_SIGNATURE', response.getResponseText())
  }
}

function fGiorniInvioMail() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Giorni Invio Mail', 'Valore attuale: ' + getExpiryDays(), ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    documentProperties.setProperty('EXPIRY_DAYS', response.getResponseText())
  }
}

function fNuovoCorso() {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var yourNewSheet = activeSpreadsheet.insertSheet();
    yourNewSheet.setName("Nuovo Corso");
    // intestazioni
    var intestazioni =  [COLUMN_NOME, CLOUMN_COGNOME, COLUMN_EMAIL, COLUMN_FINEABBONAMENTO, COLUMN_CERTIFICATOMEDICO, COLUMN_MAILABBONAMENTO, COLUMN_MAILCERTIFICATO];
    yourNewSheet.appendRow(intestazioni);

    var headersRange = yourNewSheet.getRange(yourNewSheet.getLastRow(), 1, 1, yourNewSheet.getLastColumn())
    headersRange.setFontWeight('bold')
    headersRange.setFontColor('white')
    headersRange.setBackground('#52489C')
    headersRange.setHorizontalAlignment('center')
    yourNewSheet.setFrozenRows(yourNewSheet.getLastRow())
    yourNewSheet.autoResizeRows(1, 1)

    var esempio =  ["mario", "rossi", "mario.rossi@mail.com", new Date(), new Date(), "", ""];
    yourNewSheet.appendRow(esempio);
    yourNewSheet.getRange(2, 4, 1000).setNumberFormat('dd MMM yyyy')
    yourNewSheet.getRange(2, 5, 1000).setNumberFormat('dd MMM yyyy')
    yourNewSheet.getRange(2, 6, 1000).setNumberFormat('dd MMM yyyy')
    yourNewSheet.getRange(2, 7, 1000).setNumberFormat('dd MMM yyyy')

    var ui = SpreadsheetApp.getUi()
    ui.alert('E\' stato creato un nuovo foglio di lavoro col nome "Nuovo Corso": rinominarlo come serve, es. "Corso Boxe delle 19:00"' , ui.ButtonSet.OK)
}

function fInviaMailAbbonamentiScaduti() {
  nonEvidenziareNullaAbbonamenti()  
  evidenzaAbbonamentiInScadenza()

  var ui = SpreadsheetApp.getUi()
  var response = ui.alert('Abbonamenti In Scadenza', 'Inviare la mail per le righe evidenziate?', ui.ButtonSet.YES_NO)
  if (response == ui.Button.YES) {
    mailAbbonamentiInScadenza()
  }
}

function fInviaMailCertificati() {
  nonEvidenziareNullaCertificati()
  evidenzaCertificatiInScadenza()

  var ui = SpreadsheetApp.getUi()
  var response = ui.alert('Certificati Mancanti / In Scadenza', 'Inviare la mail per le righe evidenziate?', ui.ButtonSet.YES_NO)
  if (response == ui.Button.YES) {
    mailCertifcatiMancanti()
    mailCertificatiInScadenza()
  }
}

function fInviaMailAbbonamentiScadutiTutti() {
  var yourNewSheet = SpreadsheetApp.getActiveSpreadsheet();
}

function fInviaMailCertificatiTutti() {
  var yourNewSheet = SpreadsheetApp.getActiveSpreadsheet();
  
}

