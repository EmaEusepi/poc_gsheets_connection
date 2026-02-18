// ============================================
// Google Apps Script - Batch Cloud Evaluation
// ============================================
//
// INSTALLAZIONE:
// 1. Apri il tuo Google Sheet
// 2. Vai su Estensioni > Apps Script
// 3. Incolla questo codice (in un file separato da quello delle custom function)
// 4. Sostituisci BATCH_API_URL con il tuo endpoint
// 5. Salva e autorizza lo script
// 6. Ricarica il foglio: comparira' il menu "Cloud Calc"
//
// UTILIZZO:
// - Nel foglio "Model", scrivi le formule come TESTO con apostrofo davanti:
//     '=SUM(A1:A10)    → Sheets mostra "=SUM(A1:A10)" senza calcolarla
//     '=A1+B1          → Sheets mostra "=A1+B1" senza calcolarla
//   I valori normali (numeri, testo) si scrivono come sempre.
// - Clicca "Cloud Calc" > "Calcola tutto"
// - I risultati appariranno nel foglio "Results"

// ⚠️ CONFIGURA IL TUO ENDPOINT QUI
var BATCH_API_URL = 'http://18.153.39.218:5000';

/**
 * Aggiunge il menu custom quando si apre il foglio
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cloud Calc')
    .addItem('Calcola tutto', 'evaluateSheet')
    .addToUi();
}

/**
 * Legge il foglio "Model", separa formule (testo che inizia con "=")
 * dai valori, invia tutto al server, e scrive i risultati in "Results".
 *
 * Le formule vanno scritte come testo (con apostrofo: '=SUM(A1:A10))
 * cosi' Sheets non le calcola localmente.
 */
function evaluateSheet() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Leggi il foglio Model
  var model = ss.getSheetByName('Model');
  if (!model) {
    ui.alert('Errore', 'Foglio "Model" non trovato. Crea un foglio chiamato "Model" con le tue formule.', ui.ButtonSet.OK);
    return;
  }

  var dataRange = model.getDataRange();
  if (dataRange.getNumRows() === 0 || dataRange.getNumColumns() === 0) {
    ui.alert('Errore', 'Il foglio "Model" e\' vuoto.', ui.ButtonSet.OK);
    return;
  }

  var startTime = new Date().getTime();

  // Leggiamo solo getValues(): le formule sono testo (scritte con apostrofo)
  var rawValues = dataRange.getValues();

  // 2. Separa formule dai valori
  //    Una cella il cui valore (stringa) inizia con "=" e' una formula-testo.
  var formulas = [];
  var values = [];

  for (var r = 0; r < rawValues.length; r++) {
    var formulaRow = [];
    var valueRow = [];
    for (var c = 0; c < rawValues[r].length; c++) {
      var val = rawValues[r][c];
      if (typeof val === 'string' && val.charAt(0) === '=') {
        // E' una formula scritta come testo
        formulaRow.push(val);
        valueRow.push(null);  // il server calcolera' il valore
      } else {
        // E' un valore diretto
        formulaRow.push('');
        if (val instanceof Date) {
          valueRow.push(val.toISOString());
        } else if (val === '') {
          valueRow.push(null);
        } else {
          valueRow.push(val);
        }
      }
    }
    formulas.push(formulaRow);
    values.push(valueRow);
  }

  // 3. Prepara il payload
  var payload = {
    'formulas': formulas,
    'values': values
  };

  // 4. Invia al server
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    ss.toast('Invio formule al server...', 'Cloud Calc', -1);

    var response = UrlFetchApp.fetch(BATCH_API_URL + '/eval_sheet', options);
    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();
    var contentType = response.getHeaders()['Content-Type'] || '';

    if (contentType.indexOf('application/json') === -1) {
      ui.alert('Errore', 'Il server non ha risposto con JSON (HTTP ' + responseCode + '). Verifica che l\'endpoint sia raggiungibile.', ui.ButtonSet.OK);
      return;
    }

    if (responseCode !== 200) {
      var errorData = JSON.parse(responseBody);
      ui.alert('Errore dal server', errorData.error || 'Errore sconosciuto (HTTP ' + responseCode + ')', ui.ButtonSet.OK);
      return;
    }

    var data = JSON.parse(responseBody);
    var results = data.results;

    if (!results || results.length === 0) {
      ui.alert('Errore', 'Il server ha restituito risultati vuoti.', ui.ButtonSet.OK);
      return;
    }

    // 5. Scrivi risultati nel foglio "Results"
    var resultsSheet = ss.getSheetByName('Results');
    if (!resultsSheet) {
      resultsSheet = ss.insertSheet('Results');
    }
    resultsSheet.clear();

    // Assicurati che la griglia sia rettangolare
    var maxCols = 0;
    for (var r = 0; r < results.length; r++) {
      if (results[r].length > maxCols) maxCols = results[r].length;
    }
    for (var r = 0; r < results.length; r++) {
      while (results[r].length < maxCols) {
        results[r].push('');
      }
    }

    resultsSheet.getRange(1, 1, results.length, maxCols).setValues(results);

    // 6. Report
    var elapsed = new Date().getTime() - startTime;
    var stats = data.stats || {};
    var msg = 'Calcolo completato in ' + (elapsed / 1000).toFixed(1) + 's';
    if (stats.formula_cells !== undefined) {
      msg += '\nFormule calcolate: ' + stats.formula_cells;
    }
    if (stats.total_cells !== undefined) {
      msg += '\nCelle totali: ' + stats.total_cells;
    }
    ss.toast(msg, 'Cloud Calc', 5);

  } catch (error) {
    ui.alert('Errore di connessione', 'Impossibile contattare il server:\n' + error.toString(), ui.ButtonSet.OK);
  }
}
