// ====================================
// Versione del 4 Nov 2023
// ====================================

const TODAY = new Date(new Date().setHours(0, 0, 0, 0))
// quanti giorni prima della scadenza si considera un abbonamento/certificato "in scadenza"
const DEBUG = false

const MAIL_SIGNATURE_DEAFULT = 'Fit N\' Fight'
const EXPIRY_DAYS_DEAFULT = 7

var COLUMN_NOME_CLEAN = null
var CLOUMN_COGNOME_CLEAN = null
var COLUMN_FINEABBONAMENTO_CLEAN = null
var COLUMN_CERTIFICATOMEDICO_CLEAN = null
var COLUMN_MAILABBONAMENTO_CLEAN = null
var COLUMN_MAILCERTIFICATO_CLEAN = null
var COLUMN_EMAIL_CLEAN = null

function getMailSignature() {
  return documentProperties.getProperty('MAIL_SIGNATURE') ? documentProperties.getProperty('MAIL_SIGNATURE') : MAIL_SIGNATURE_DEAFULT
}

function getExpiryDays() {
  return documentProperties.getProperty('EXPIRY_DAYS') ? documentProperties.getProperty('EXPIRY_DAYS') : EXPIRY_DAYS_DEAFULT
}

function isValidDate(d) {
  if (Object.prototype.toString.call(d) !== "[object Date]") {
    return false
  }
  return !isNaN(d.getTime())
}

function pulisciValore(valoreGrezzo) {
  if (valoreGrezzo == null) return null
  var valorePulito = valoreGrezzo.toUpperCase()
            .replace(/^\\s+/, '')
            .replace(/\\s+$/, '')
            .replace(/\\s+/, ' ')
            .replace(/[^A-Z0-9]+/, '')
  return valorePulito
}

function pulisciData(dataGrezza) {
  if (dataGrezza == undefined) {
    return undefined
  }
  if (isValidDate(dataGrezza)) {
    return dataGrezza
  }

  var rx = /(\d+)\/(\d+)\/(\d+)/
  var datepart = dataGrezza.match(rx)
  if (datepart) {
    return new Date(datepart[3], datepart[2] - 1, datepart[1])
  }
  return undefined
}

function processaFoglio(tipo, sheet, inviaMail, evidenzia, pulisci) {

  if (!sheet) {
    SpreadsheetApp.getUi().alert("Foglio Di Lavoro non inizializzato!")
    return
  }

  Logger.log("=============================================================================================================")
  Logger.log("[INFO] inizio elaborazione foglio '" + sheet.getSheetName() + "' ...")
  Logger.log("=============================================================================================================")

  // TEST START
  if (DEBUG) {
    var response = SpreadsheetApp.getUi().alert('Are you sure you want to process sheet ' + sheet.getSheetName() + '?', SpreadsheetApp.getUi().ButtonSet.YES_NO)
    if (response == SpreadsheetApp.getUi().Button.YES) {
      Logger.log('The user clicked "Yes."');
    } else {
      return
    }
  }
  // TEST END

  var dataRange = sheet.getDataRange().getValues()
  var TROVATO_INTESTAZIONI = false
  var IDX_NOME = undefined
  var IDX_COGNOME = undefined
  var IDX_EMAIL = undefined
  var IDX_ABBONAMENTO = undefined
  var IDX_CERTIFICATO = undefined
  var IDX_DATA_MAIL_ABBONAMENTO = undefined
  var IDX_DATA_MAIL_CERTIFICATO = undefined

  dataRange.forEach(function (row, rowIndex) {
    // controllo scadenze
    if (TROVATO_INTESTAZIONI) {
      Logger.log("[INFO] Processo riga " + row)
      try {
        var NOME = IDX_NOME != undefined ? row[IDX_NOME] : undefined
        var COGNOME = IDX_COGNOME != undefined ? row[IDX_COGNOME] : undefined
        var EMAIL = IDX_EMAIL != undefined ? row[IDX_EMAIL] : undefined
        var DATA_SCADENZA_ABBONAMENTO = IDX_ABBONAMENTO != undefined ? pulisciData(row[IDX_ABBONAMENTO]) : undefined
        var DATA_SCADENZA_CERTIFICATO = IDX_CERTIFICATO != undefined ? pulisciData(row[IDX_CERTIFICATO]) : undefined
        var DATA_MAIL_ABBONAMENTO = IDX_DATA_MAIL_ABBONAMENTO != undefined ? pulisciData(row[IDX_DATA_MAIL_ABBONAMENTO]) : undefined
        var DATA_MAIL_CERTIFICATO = IDX_DATA_MAIL_CERTIFICATO != undefined ? pulisciData(row[IDX_DATA_MAIL_CERTIFICATO]) : undefined
        /*Logger.log("[INFO]" +
          " NOME: " + NOME +
          " EMAIL: " + EMAIL +
          " ABBONAMENTO: " + DATA_SCADENZA_ABBONAMENTO +
          " CERTIFICATO: " + DATA_SCADENZA_CERTIFICATO +
          " DATA_MAIL_ABBONAMENTO: " + DATA_MAIL_ABBONAMENTO +
          " DATA_MAIL_CERTIFICATO: " + DATA_MAIL_CERTIFICATO)*/

        if (tipo == "ABBONAMENTI" && IDX_ABBONAMENTO != undefined && DATA_SCADENZA_ABBONAMENTO != undefined) {

          // =====================================================================                      
          // =====================================================================
          // Evidenzia a manda mail per abbonamenti scaduti
          // =====================================================================
          // =====================================================================

          // nonEvidenziareNulla
          if (pulisci) {
            sheet.getRange(rowIndex + 1, IDX_ABBONAMENTO + 1).setBackground(null)
          }

          var dataScadenzaAbbonamento = new Date(DATA_SCADENZA_ABBONAMENTO)
          var expiryDate = new Date(dataScadenzaAbbonamento.getTime() - getExpiryDays() * (24 * 3600 * 1000))
          var dataMail = (DATA_MAIL_ABBONAMENTO == undefined) ? undefined : new Date(DATA_MAIL_ABBONAMENTO)

          if (dataScadenzaAbbonamento != undefined) {
            if (evidenzia) {
              if (TODAY < expiryDate) {

                // ============================
                // PAGATO OK
                // ============================            
                sheet.getRange(rowIndex + 1, IDX_ABBONAMENTO + 1).setBackground(null)

              } else if ((TODAY >= expiryDate) && (dataMail == undefined || dataMail < expiryDate) && (EMAIL != undefined && EMAIL != "")) {

                // ============================
                // SCADUTO E MAIL 
                // ============================
                sheet.getRange(rowIndex + 1, IDX_ABBONAMENTO + 1).setBackground('orange')
                /*Logger.log("[WARN] EMAIL!!! " +
                  " NOME: " + NOME +
                  " EMAIL: \"" + EMAIL + "\"" +
                  " ABBONAMENTO: " + DATA_SCADENZA_ABBONAMENTO +
                  " expiryDate: " + expiryDate +
                  " DATA_MAIL_ABBONAMENTO: " + DATA_MAIL_ABBONAMENTO)*/

              } else if (TODAY >= expiryDate) {

                // ============================
                // SCADUTO E NO MAIL
                // ============================
                sheet.getRange(rowIndex + 1, IDX_ABBONAMENTO + 1).setBackground('yellow')

              }
            } else if (inviaMail && (TODAY >= expiryDate) && (dataMail == undefined || dataMail < expiryDate)) {

              // ============================
              // MAIL ABBONAMENTO
              // ============================
              if (EMAIL != undefined) {
                if (NOME != undefined && NOME != "") {
                  //var cttLogoBlob = DriveApp.getFileById("1_Le0ofGDsm1iNhGUxTthUSfW0T3XtSoN").getBlob()

                  /*MailApp.sendEmail({
                    name: 'Combat Traning Team A.S.D.',
                    replyTo: 'combattrainingteam@outlook.it',
                    to: EMAIL,
                    subject: "Scandenza Abbonamento",
                    htmlBody: "<p style='font-family:Calibri,Arial,Helvetica,sans-serif,serif,EmojiFont;font-size:14pt'>" +
                    "Ciao " + NOME + " " + COGNOME + ", <br>" + 
                      "Ti informiamo che il tuo <b>Abbonamento</b> sta per Scadere!<br>" +
                      "Ti aspettiamo in Palestra per Rinnovarlo! </p><br>" +
                      "<p style='font-family:Calibri,Arial,Helvetica,sans-serif,serif,EmojiFont;font-size:12pt'>Combat Training Team<br>" +
                      "<img src='cid:cttLogo' width='150' height='150'></p>",
                    inlineImages: { cttLogo: cttLogoBlob },
                  });*/

                  MailApp.sendEmail({
                    name: 'Combat Traning Team A.S.D.',
                    replyTo: 'combattrainingteam@outlook.it',
                    to: EMAIL,
                    subject: "Scandenza Abbonamento",
                    htmlBody: "<p style='font-family:Calibri,Arial,Helvetica,sans-serif,serif,EmojiFont;font-size:14pt'>" +
                    "Ciao " + NOME + " " + COGNOME + ", <br>" + 
                      "Ti informiamo che il tuo <b>Abbonamento</b> sta per Scadere!<br>" +
                      "Ti aspettiamo in Palestra per Rinnovarlo! </p><br>" +
                      "<p style='font-family:Calibri,Arial,Helvetica,sans-serif,serif,EmojiFont;font-size:14pt'>" + getMailSignature() + "<br></p>",
                  });

                  sheet.getRange(rowIndex + 1, IDX_DATA_MAIL_ABBONAMENTO + 1).setValue(TODAY)
                  sheet.getRange(rowIndex + 1, IDX_ABBONAMENTO + 1).setBackground('yellow')
                } else {
                  Logger.log('No NAME for ' + EMAIL)
                }
              } else {
                Logger.log('No MAIL for row ' + (rowIndex + 1))
              }
            }
          }

        } else if (tipo == "CERTIFICATI" && IDX_CERTIFICATO != undefined && DATA_SCADENZA_CERTIFICATO != undefined) {

          // =====================================================================                      
          // =====================================================================
          // Evidenzia a manda mail per certificati scaduti
          // =====================================================================
          // =====================================================================

          // nonEvidenziareNulla
          if (pulisci) {
            sheet.getRange(rowIndex + 1, IDX_CERTIFICATO + 1).setBackground(null)
          }

          var dataScadenzaCertificato = new Date(DATA_SCADENZA_CERTIFICATO)
          var expiryDate = new Date(dataScadenzaCertificato.getTime() - getExpiryDays() * (24 * 3600 * 1000))
          var dataMail = (DATA_MAIL_CERTIFICATO == undefined) ? undefined : new Date(DATA_MAIL_CERTIFICATO)

          if (dataScadenzaCertificato != undefined) {
            if (evidenzia) {
              if (TODAY < expiryDate) {

                // ============================
                // PAGATO OK
                // ============================            
                sheet.getRange(rowIndex + 1, IDX_CERTIFICATO + 1).setBackground(null)

              } else if ((TODAY >= expiryDate) && (dataMail == undefined || dataMail < expiryDate) && (EMAIL != undefined && EMAIL != "")) {

                // ============================
                // SCADUTO E MANDEREBBE MAIL 
                // ============================
                sheet.getRange(rowIndex + 1, IDX_CERTIFICATO + 1).setBackground('orange')

              } else if (TODAY >= expiryDate) {

                // ============================
                // SCADUTO E NO MAIL
                // ============================
                sheet.getRange(rowIndex + 1, IDX_CERTIFICATO + 1).setBackground('yellow')

              }
            } else if (inviaMail && (TODAY >= expiryDate) && (dataMail == undefined || dataMail < expiryDate)) {

              // ============================
              // MAIL CERTIFICATO SCADUTO
              // ============================
              if (EMAIL != undefined) {
                  if (NOME != undefined && NOME != "") {
                  //var cttLogoBlob = DriveApp.getFileById("1_Le0ofGDsm1iNhGUxTthUSfW0T3XtSoN").getBlob()
                  MailApp.sendEmail({
                    name: 'Combat Traning Team A.S.D.',
                    replyTo: 'combattrainingteam@outlook.it',
                    to: EMAIL,
                    subject: "Scandenza Certificato Medico",
                    htmlBody: "<p style='font-family:Calibri,Arial,Helvetica,sans-serif,serif,EmojiFont;font-size:14pt'>" +
                    "Ciao " + NOME + " " + COGNOME + ", <br>" +  
                      "Ti informiamo che il tuo <b>Certificato Medico</b> sta per Scadere!</p><br>" +
                      "<p style='font-family:Calibri,Arial,Helvetica,sans-serif,serif,EmojiFont;font-size:14pt'>" + getMailSignature() + "<br></p>",
                  });

                  sheet.getRange(rowIndex + 1, IDX_DATA_MAIL_CERTIFICATO + 1).setValue(TODAY)
                  sheet.getRange(rowIndex + 1, IDX_CERTIFICATO + 1).setBackground('yellow')
                } else {
                  Logger.log('No NAME for ' + EMAIL)
                }
              } else {
                Logger.log('No EMAIL for row ' + (rowIndex + 1))
              }
            }
          }

        } else if (tipo == "CERTIFICATI_MANCANTI" 
                    && IDX_CERTIFICATO != undefined 
                    && DATA_SCADENZA_CERTIFICATO == undefined 
                    && (
                          (DATA_SCADENZA_ABBONAMENTO != undefined) 
                          || (sheet.getSheetName().toUpperCase().indexOf("OFFERTA SOCIALE") >= 0) 
                          || (sheet.getSheetName().toUpperCase().indexOf("TESSERAMENTI PRIVATE") >= 0)
                       )) {

          // =====================================================================                      
          // =====================================================================
          // Evidenzia a manda mail per certificati mancanti
          // =====================================================================
          // =====================================================================
          
          // nonEvidenziareNulla
          if (pulisci) {
            sheet.getRange(rowIndex + 1, IDX_CERTIFICATO + 1).setBackground(null)
          }
          
          var dataMail = (DATA_MAIL_CERTIFICATO == undefined) ? undefined : new Date(DATA_MAIL_CERTIFICATO)

          if (evidenzia) {

            if (dataMail == undefined && EMAIL != undefined && EMAIL != "") {
              // ============================
              // MANCANTE E MANDEREBBE MAIL 
              // ============================
              sheet.getRange(rowIndex + 1, IDX_CERTIFICATO + 1).setBackground('orange')
            } else {
              // ============================
              // MANCANTE E NO MAIL
              // ============================
              sheet.getRange(rowIndex + 1, IDX_CERTIFICATO + 1).setBackground('yellow')
            }

          } else if (inviaMail && dataMail == undefined) {

            // ============================
            // MAIL CERTIFICATO MANCANTE
            // ============================
            if (EMAIL != undefined && EMAIL != "") {
              if (NOME != null && NOME != "") {
                //var cttLogoBlob = DriveApp.getFileById("1_Le0ofGDsm1iNhGUxTthUSfW0T3XtSoN").getBlob()
                MailApp.sendEmail({
                  name: 'Combat Traning Team A.S.D.',
                  replyTo: 'combattrainingteam@outlook.it',
                  to: EMAIL,
                  subject: "Certificato Medico Mancante",
                  htmlBody: "<p style='font-family:Calibri,Arial,Helvetica,sans-serif,serif,EmojiFont;font-size:14pt'>" +
                    "Ciao " + NOME + " " + COGNOME + ", <br>" + 
                    "Ti informiamo che non hai ancora portato il <b>Certificato Medico</b>!</p><br>" +
                    "<p style='font-family:Calibri,Arial,Helvetica,sans-serif,serif,EmojiFont;font-size:14pt'>" + getMailSignature() + "<br></p>",
                });

                sheet.getRange(rowIndex + 1, IDX_DATA_MAIL_CERTIFICATO + 1).setValue(TODAY)
                sheet.getRange(rowIndex + 1, IDX_CERTIFICATO + 1).setBackground('yellow')
              } else {
                Logger.log("No NAME for " + EMAIL)
              }
            } else {
              Logger.log("No EMAIL for " + (rowIndex+1))
            }
          }

        }
      } catch (e) {
        Logger.log("[WARN] Errore processando riga " + row + ": " + e)
      }
    }

    // =====================================================================                      
    // =====================================================================
    // Recupero indici colonne che contengono dati utili
    // =====================================================================
    // =====================================================================
    if (!TROVATO_INTESTAZIONI && rowIndex < 15) {

      COLUMN_NOME_CLEAN = pulisciValore(COLUMN_NOME)
      CLOUMN_COGNOME_CLEAN = pulisciValore(CLOUMN_COGNOME)
      COLUMN_FINEABBONAMENTO_CLEAN = pulisciValore(COLUMN_FINEABBONAMENTO)
      COLUMN_CERTIFICATOMEDICO_CLEAN = pulisciValore(COLUMN_CERTIFICATOMEDICO)
      COLUMN_MAILABBONAMENTO_CLEAN = pulisciValore(COLUMN_MAILABBONAMENTO)
      COLUMN_MAILCERTIFICATO_CLEAN = pulisciValore(COLUMN_MAILCERTIFICATO)
      COLUMN_EMAIL_CLEAN = pulisciValore(COLUMN_EMAIL)

      //Logger.log("[INFO] Cerco intestazioni in riga " + row)
      for (idx = 0; idx < row.length; idx++) {
        try {
          var valoreGrezzo = row[idx].toString()
          var valorePulito = pulisciValore(valoreGrezzo)

          if (valorePulito) {
            if (valorePulito == COLUMN_NOME_CLEAN) {
              Logger.log("CELL[" + idx + "]: " + row[idx] + " --> NOME")
              IDX_NOME = idx
              TROVATO_INTESTAZIONI = true
            } else if (valorePulito == CLOUMN_COGNOME_CLEAN) {
              Logger.log("CELL[" + idx + "]: " + row[idx] + " --> COGNOME")
              IDX_COGNOME = idx
              TROVATO_INTESTAZIONI = true
            } else if (valorePulito == COLUMN_FINEABBONAMENTO_CLEAN) {
              Logger.log("CELL[" + idx + "]: " + row[idx] + " --> FINE ABBONAMENTO")
              IDX_ABBONAMENTO = idx
              TROVATO_INTESTAZIONI = true
            } else if (valorePulito == COLUMN_CERTIFICATOMEDICO_CLEAN) {
              Logger.log("CELL[" + idx + "]: " + row[idx] + " --> CERTIFICATO MEDICO")
              IDX_CERTIFICATO = idx
              TROVATO_INTESTAZIONI = true
            } else if (valorePulito == COLUMN_MAILABBONAMENTO_CLEAN) {
              Logger.log("CELL[" + idx + "]: " + row[idx] + " --> MAIL ABBONAMENTO")
              IDX_DATA_MAIL_ABBONAMENTO = idx
              TROVATO_INTESTAZIONI = true
            } else if (valorePulito == COLUMN_MAILCERTIFICATO_CLEAN) {
              Logger.log("CELL[" + idx + "]: " + row[idx] + " --> MAIL CERTIFICATO")
              IDX_DATA_MAIL_CERTIFICATO = idx
              TROVATO_INTESTAZIONI = true
            } else if (valorePulito == COLUMN_EMAIL_CLEAN) {
              Logger.log("CELL[" + idx + "]: " + row[idx] + " --> INDIRIZZO MAIL")
              IDX_EMAIL = idx
              TROVATO_INTESTAZIONI = true
            }
          }
        } catch (e) {
          Logger.log("[WARN] Errore ricerca intestazioni su cella " + row[idx] + ": " + e)
        }
      }
      Logger.log("[INFO] Trovate intestazioni: " + TROVATO_INTESTAZIONI)
    }
  });

  Logger.log("[INFO] fine elaborazione foglio '" + sheet.getSheetName() + "'")
}

function mailAbbonamentiInScadenza() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (i = 0; i < sheets.length; i++) {
    processaFoglio("ABBONAMENTI", sheets[i], true, false, false)
  }
}

function mailCertificatiInScadenza() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (i = 0; i < sheets.length; i++) {
    processaFoglio("CERTIFICATI", sheets[i], true, false, false)
  }
}

function mailCertifcatiMancanti() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (i = 0; i < sheets.length; i++) {
    processaFoglio("CERTIFICATI_MANCANTI", sheets[i], true, false, false)
  }
}

function evidenzaAbbonamentiInScadenza() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (i = 0; i < sheets.length; i++) {
    processaFoglio("ABBONAMENTI", sheets[i], false, true, false)
    SpreadsheetApp.flush()
  }
}

function evidenzaCertificatiInScadenza() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (i = 0; i < sheets.length; i++) {
    processaFoglio("CERTIFICATI", sheets[i], false, true, false)
    SpreadsheetApp.flush()
  }
}

function evidenzaCertificatiMancanti() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (i = 0; i < sheets.length; i++) {
    processaFoglio("CERTIFICATI_MANCANTI", sheets[i], false, true, false)
    SpreadsheetApp.flush()
  }
}

function nonEvidenziareNulla() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (i = 0; i < sheets.length; i++) {
    processaFoglio("ABBONAMENTI", sheets[i], false, false, true)
    processaFoglio("CERTIFICATI", sheets[i], false, false, true)
    processaFoglio("CERTIFICATI_MANCANTI", sheets[i], false, false, true)
    SpreadsheetApp.flush()
  }
}

function nonEvidenziareNullaAbbonamenti() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (i = 0; i < sheets.length; i++) {
    processaFoglio("ABBONAMENTI", sheets[i], false, false, true)
    SpreadsheetApp.flush()
  }
}

function nonEvidenziareNullaCertificati() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (i = 0; i < sheets.length; i++) {
    processaFoglio("CERTIFICATI", sheets[i], false, false, true)
    processaFoglio("CERTIFICATI_MANCANTI", sheets[i], false, false, true)
    SpreadsheetApp.flush()
  }
}

