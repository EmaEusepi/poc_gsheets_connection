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
// WORKFLOW:
// 1. "Scongela formule" (o selezione) -> le formule-testo tornano formule vere
// 2. Modifica, trascina, copia le formule normalmente
// 3. "Congela formule" (o selezione) -> le formule diventano testo, Sheets smette di calcolare
// 4. "Calcola tutto" -> invia le formule congelate al cloud
//    I risultati vanno in un foglio "<nome_foglio>_RES" con lo stesso stile del sorgente

// ⚠️ CONFIGURA IL TUO ENDPOINT QUI
// var BATCH_API_URL = 'http://18.153.39.218:5000';
var BATCH_API_URL = 'http://35.159.123.184:5000';


// ⚠️ PREREQUISITO: Abilita il servizio avanzato "Google Sheets API"
// 1. In Apps Script, vai su Servizi (icona +) nel pannello a sinistra
// 2. Cerca "Google Sheets API" e aggiungilo
// Serve per scrivere formule come testo senza che vengano eseguite.

/**
 * Aggiunge il menu custom quando si apre il foglio
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cloud Calc')
    .addItem('Calcola tutto', 'evaluateSheet')
    .addSeparator()
    .addItem('Congela formule (tutto il foglio)', 'freezeAll')
    .addItem('Scongela formule (tutto il foglio)', 'unfreezeAll')
    .addSeparator()
    .addItem('Congela selezione', 'freezeSelection')
    .addItem('Scongela selezione', 'unfreezeSelection')
    .addToUi();
}

// ============================================
// CONGELA / SCONGELA
// ============================================

/**
 * Congela tutte le formule del foglio attivo:
 * le formule vere (=...) diventano testo ('=...).
 * Sheets smette di calcolarle.
 */
function freezeAll() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  freezeRange_(range);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Formule congelate nel foglio "' + sheet.getName() + '"', 'Cloud Calc', 3
  );
}

/**
 * Scongela tutte le formule del foglio attivo:
 * le stringhe che iniziano con "=" diventano formule vere.
 * ⚠️ Sheets iniziera' a calcolarle!
 */
function unfreezeAll() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  unfreezeRange_(range);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Formule scongelate nel foglio "' + sheet.getName() + '"', 'Cloud Calc', 3
  );
}

/**
 * Congela solo le celle selezionate.
 */
function freezeSelection() {
  var range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('Seleziona un range di celle prima.');
    return;
  }
  freezeRange_(range);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Selezione congelata (' + range.getA1Notation() + ')', 'Cloud Calc', 3
  );
}

/**
 * Scongela solo le celle selezionate.
 */
function unfreezeSelection() {
  var range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('Seleziona un range di celle prima.');
    return;
  }
  unfreezeRange_(range);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Selezione scongelata (' + range.getA1Notation() + ')', 'Cloud Calc', 3
  );
}

/**
 * Congela un range: converte formule vere in stringhe di testo.
 * Legge getFormulas() per trovare le celle con formula,
 * poi scrive il testo della formula come valore (Sheets lo tratta come testo).
 */
function freezeRange_(range) {
  var formulas = range.getFormulas();
  var values = range.getValues();
  var numRows = formulas.length;
  var numCols = formulas[0].length;
  var hasFormulas = false;

  // Costruisci la griglia da scrivere: dove c'e' una formula, metti il testo della formula
  var output = [];
  for (var r = 0; r < numRows; r++) {
    var row = [];
    for (var c = 0; c < numCols; c++) {
      if (formulas[r][c]) {
        row.push(formulas[r][c]);  // es: "=SUM(A1:A10)" come stringa
        hasFormulas = true;
      } else {
        row.push(values[r][c]);
      }
    }
    output.push(row);
  }

  if (!hasFormulas) return false;

  // Usa la Sheets API con valueInputOption RAW per scrivere
  // le formule come testo letterale senza che Sheets le interpreti.
  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var sheetName = range.getSheet().getName();
  var a1 = range.getA1Notation();
  var rangeA1 = "'" + sheetName + "'!" + a1;

  Sheets.Spreadsheets.Values.update(
    { values: output },
    ssId,
    rangeA1,
    { valueInputOption: 'RAW' }
  );

  return true;
}

/**
 * Scongela un range: converte stringhe che iniziano con "=" in formule vere.
 * Usa setFormulas() solo sulle celle che contengono testo-formula.
 */
function unfreezeRange_(range) {
  var values = range.getValues();
  var formulas = range.getFormulas();
  var numRows = values.length;
  var numCols = values[0].length;

  // Dobbiamo scrivere cella per cella perche' setFormula() e setValues()
  // non possono essere mischiati sullo stesso range in un colpo solo.
  // Usiamo un approccio batch: costruiamo una griglia di formule.
  var hasFormulas = false;
  var formulaGrid = [];

  for (var r = 0; r < numRows; r++) {
    var row = [];
    for (var c = 0; c < numCols; c++) {
      var val = values[r][c];
      if (typeof val === 'string' && val.charAt(0) === '=' && !formulas[r][c]) {
        // E' testo che inizia con "=" e NON e' gia' una formula -> scongela
        row.push(val);
        hasFormulas = true;
      } else {
        row.push(null); // non toccare questa cella
      }
    }
    formulaGrid.push(row);
  }

  if (!hasFormulas) return;

  // Scriviamo le formule cella per cella (setFormula accetta una stringa)
  var startRow = range.getRow();
  var startCol = range.getColumn();
  var sheet = range.getSheet();

  for (var r = 0; r < numRows; r++) {
    for (var c = 0; c < numCols; c++) {
      if (formulaGrid[r][c] !== null) {
        sheet.getRange(startRow + r, startCol + c).setFormula(formulaGrid[r][c]);
      }
    }
  }
}

// ============================================
// CALCOLO CLOUD
// ============================================

/**
 * Legge il foglio attivo, separa formule (testo che inizia con "=")
 * dai valori, invia tutto al server, e scrive i risultati in un foglio
 * chiamato "<nome_foglio_attivo>_RES" (copiando anche lo stile).
 */
function evaluateSheet() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Usa il foglio attivo come sorgente
  var sourceSheet = ss.getActiveSheet();
  var sourceName = sourceSheet.getName();
  var resName = sourceName + '_RES';

  // Impedisci di eseguire su un foglio _RES (evita loop)
  if (sourceName.indexOf('_RES') === sourceName.length - 4) {
    ui.alert('Errore', 'Non puoi eseguire "Calcola tutto" su un foglio risultati (_RES). Seleziona il foglio sorgente.', ui.ButtonSet.OK);
    return;
  }

  var dataRange = sourceSheet.getDataRange();
  if (dataRange.getNumRows() === 0 || dataRange.getNumColumns() === 0) {
    ui.alert('Errore', 'Il foglio "' + sourceName + '" e\' vuoto.', ui.ButtonSet.OK);
    return;
  }

  var startTime = new Date().getTime();

  var rawValues = dataRange.getValues();
  var realFormulas = dataRange.getFormulas();

  // 2. Separa formule dai valori
  var formulas = [];
  var values = [];

  for (var r = 0; r < rawValues.length; r++) {
    var formulaRow = [];
    var valueRow = [];
    for (var c = 0; c < rawValues[r].length; c++) {
      var val = rawValues[r][c];
      var realFormula = realFormulas[r][c];

      if (realFormula) {
        formulaRow.push(realFormula);
        valueRow.push(null);
      } else if (typeof val === 'string' && val.charAt(0) === '=') {
        formulaRow.push(val);
        valueRow.push(null);
      } else {
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

    // 5. Crea/ricrea il foglio risultati copiando lo stile dal sorgente
    var resultsSheet = ss.getSheetByName(resName);
    if (resultsSheet) {
      // Elimina il vecchio foglio _RES per ricrearlo con lo stile aggiornato
      ss.deleteSheet(resultsSheet);
    }

    // Copia il foglio sorgente (include formattazione, colori, larghezze, ecc.)
    resultsSheet = sourceSheet.copyTo(ss);
    resultsSheet.setName(resName);

    // 6. Sovrascrivi con i risultati calcolati dal cloud
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

    // Scrivi i valori (sovrascrive formule e testo, mantiene la formattazione)
    resultsSheet.getRange(1, 1, results.length, maxCols).setValues(results);

    // 7. Report
    var elapsed = new Date().getTime() - startTime;
    var stats = data.stats || {};
    var msg = 'Risultati scritti in "' + resName + '" (' + (elapsed / 1000).toFixed(1) + 's)';
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


// ============================================
// CUSTOM FUNCTION BATCH - CLOUD_CALC_BATCH
// ============================================
//
// NAMED FUNCTION (consigliata per semplicita' d'uso):
//   Menu Dati > Named Functions > Aggiungi
//   Nome:       CLOUD
//   Argomenti:  operation, arg1, [arg2], [arg3], [arg4], [arg5], [arg6], [arg7], [arg8]
//   Formula:    =CLOUD_CALC_BATCH(operation, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, ROW(), COLUMN())
//
//   L'utente scrive:  =CLOUD("sum"; A1; A2)
//

// Endpoint batch (porta 5001, server separato)
var BATCH_CALC_URL = 'http://35.159.123.184:5000/batch_calc';

/**
 * Esegue un calcolo cloud con risoluzione batch delle dipendenze.
 *
 * Gli ultimi 2 argomenti DEVONO essere ROW() e COLUMN() (iniettati
 * dalla Named Function oppure scritti a mano).
 *
 * @param {string} operation - Nome dell'operazione (es: "plus", "sum")
 * @param {...any} args - Argomenti per l'operazione, seguiti da ROW() e COLUMN()
 * @return {any} Risultato del calcolo
 * @customfunction
 */
function CLOUD_CALC_BATCH(operation) {
  // -- 1. Separa gli argomenti reali da ROW/COLUMN -----------------------
  var allArgs = [];
  for (var i = 1; i < arguments.length; i++) {
    allArgs.push(arguments[i]);
  }

  // Gli ultimi 2 sono ROW() e COLUMN()
  var colNum = allArgs.pop();
  var rowNum = allArgs.pop();
  var values = allArgs;

  // Ricava il nome della cella chiamante (es: "C3")
  var callingCell = columnToLetter_(colNum) + rowNum;

  // -- 2. Leggi la formula dalla cella per estrarre i riferimenti --------
  var refs = extractRefs_(rowNum, colNum);

  // -- 3. Costruisci gli args come [{ref, value}] -----------------------
  var payload_args = [];
  for (var j = 0; j < values.length; j++) {
    var entry = { value: values[j] === "" ? null : values[j] };
    if (j < refs.length && refs[j] !== '') {
      entry.ref = refs[j];
    }
    payload_args.push(entry);
  }

  // -- 4. IFERROR locale (come nella versione originale) -----------------
  if (typeof operation === 'string' && operation.toLowerCase() === 'iferror') {
    var val = payload_args.length > 0 ? payload_args[0].value : null;
    var fallback = payload_args.length > 1 ? payload_args[1].value : "";
    if (containsBatchSheetError_(val)) {
      return fallback;
    }
    return (val === null || val === undefined) ? "" : val;
  }

  // -- 5. Validazione ----------------------------------------------------
  var validationError = validateBatchArgs_(operation, values);
  if (validationError) {
    return validationError;
  }

  // -- 6. Chiamata al server batch ---------------------------------------
  var payload = {
    cell: callingCell,
    operation: operation,
    args: payload_args
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(BATCH_CALC_URL, options);
    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();
    var contentType = response.getHeaders()['Content-Type'] || '';

    if (contentType.indexOf('application/json') === -1) {
      return '#ERROR: Risposta non JSON (HTTP ' + responseCode + ')';
    }

    if (responseCode === 200) {
      var data = JSON.parse(responseBody);
      var result = data.result;
      if (result === null || result === undefined) return "";
      return result;
    } else {
      var errorData = JSON.parse(responseBody);
      return '#ERROR: ' + errorData.error;
    }
  } catch (error) {
    return '#ERROR: ' + error.toString();
  }
}


// =========================================================================
// Helpers per CLOUD_CALC_BATCH
// =========================================================================

/**
 * Legge la formula della cella (row, col) ed estrae i riferimenti celle
 * passati come argomenti a CLOUD_CALC_BATCH (esclude operation, ROW, COLUMN).
 *
 * Esempio: =CLOUD_CALC_BATCH("sum"; A1; B2; ROW(); COLUMN())
 *   -> ["A1", "B2"]
 */
function extractRefs_(row, col) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var formula = sheet.getRange(row, col).getFormula();
    if (!formula) return [];

    // Estrai il contenuto fra le parentesi piu' esterne della funzione
    var depth = 0;
    var start = -1;
    var end = -1;
    for (var i = 0; i < formula.length; i++) {
      if (formula[i] === '(') {
        if (depth === 0) start = i + 1;
        depth++;
      } else if (formula[i] === ')') {
        depth--;
        if (depth === 0) { end = i; break; }
      }
    }
    if (start === -1 || end === -1) return [];
    var inner = formula.substring(start, end);

    // Splitta rispettando parentesi e stringhe (gestisce sia ; che ,)
    var rawArgs = splitFormulaArgs_(inner);

    // Rimuovi primo arg (operation) e ultimi 2 (ROW(), COLUMN())
    if (rawArgs.length < 3) return [];
    var cellArgs = rawArgs.slice(1, rawArgs.length - 2);

    // Filtra: tieni solo quelli che sembrano riferimenti cella
    var cellRefPattern = /^\$?[A-Z]{1,3}\$?[0-9]+(?::\$?[A-Z]{1,3}\$?[0-9]+)?$/i;
    var refs = [];
    for (var k = 0; k < cellArgs.length; k++) {
      var arg = cellArgs[k].trim();
      if (cellRefPattern.test(arg)) {
        refs.push(arg.replace(/\$/g, '').toUpperCase());
      } else {
        // Valore letterale, nessun ref
        refs.push('');
      }
    }
    return refs;

  } catch (e) {
    // Se non riesce a leggere la formula, ritorna array vuoto
    // (il server calcolera' senza info di dipendenza)
    return [];
  }
}


/**
 * Splitta una stringa di argomenti rispettando parentesi, stringhe e
 * il separatore ; (locale IT) o , (locale EN).
 */
function splitFormulaArgs_(inner) {
  var args = [];
  var depth = 0;
  var inStr = false;
  var current = '';

  for (var i = 0; i < inner.length; i++) {
    var ch = inner[i];
    if (ch === '"') {
      inStr = !inStr;
      current += ch;
    } else if (inStr) {
      current += ch;
    } else if (ch === '(') {
      depth++;
      current += ch;
    } else if (ch === ')') {
      depth--;
      current += ch;
    } else if ((ch === ',' || ch === ';') && depth === 0) {
      args.push(current.trim());
      current = '';
    } else {
      current += ch;
    }
  }
  if (current.trim() !== '') {
    args.push(current.trim());
  }
  return args;
}


/**
 * Converte un numero di colonna in lettera (1 -> A, 27 -> AA, ecc.)
 */
function columnToLetter_(col) {
  var letter = '';
  while (col > 0) {
    var mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}


/**
 * Controlla se un valore e' un errore di Google Sheets.
 */
function containsBatchSheetError_(value) {
  if (typeof value !== 'string') return false;
  var errorPrefixes = ['#REF!', '#VALUE!', '#N/A', '#NULL!', '#NUM!', '#DIV/0!', '#ERROR', '#NAME?'];
  for (var i = 0; i < errorPrefixes.length; i++) {
    if (value.indexOf(errorPrefixes[i]) === 0) return true;
  }
  return false;
}


/**
 * Valida gli argomenti prima della chiamata API.
 */
function validateBatchArgs_(operation, args) {
  if (operation === null || operation === undefined || operation === '') {
    return '#ERROR: Operazione mancante.';
  }
  if (typeof operation !== 'string') {
    return '#ERROR: L\'operazione deve essere una stringa.';
  }
  for (var i = 0; i < args.length; i++) {
    if (containsBatchSheetError_(args[i])) {
      return '#ERROR: Argomento ' + (i + 1) + ' contiene errore Sheets (' + args[i] + ').';
    }
  }
  return null;
}
