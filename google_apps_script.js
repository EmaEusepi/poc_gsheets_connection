// ============================================
// Google Apps Script per Google Sheets
// ============================================
// 
// INSTALLAZIONE:
// 1. Apri il tuo Google Sheet
// 2. Vai su Estensioni > Apps Script
// 3. Incolla questo codice
// 4. Sostituisci API_BASE_URL con il tuo endpoint
// 5. Salva e autorizza lo script

// ⚠️ CONFIGURA IL TUO ENDPOINT QUI
const API_BASE_URL = 'https://tuo-dominio.com/calc';

/**
 * Esegue calcoli su cloud tramite API
 * 
 * @param {string} operation - Nome dell'operazione (es: 'plus', 'if', 'max')
 * @param {...any} args - Argomenti variabili per l'operazione
 * @return {any} Risultato del calcolo
 * @customfunction
 */
function CLOUD_CALC(operation) {
  // Raccogli tutti gli argomenti dopo 'operation'
  var args = [];
  for (var i = 1; i < arguments.length; i++) {
    var arg = arguments[i];
    
    // Gestisci array (range di celle)
    if (Array.isArray(arg)) {
      // Appiattisci array bidimensionali
      arg.forEach(function(row) {
        if (Array.isArray(row)) {
          args = args.concat(row);
        } else {
          args.push(row);
        }
      });
    } else {
      args.push(arg);
    }
  }
  
  // Prepara il payload
  var payload = {
    'operation': operation,
    'args': args
  };
  
  // Opzioni per la richiesta HTTP
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  
  try {
    // Esegui la chiamata API
    var response = UrlFetchApp.fetch(API_BASE_URL, options);
    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();
    
    if (responseCode === 200) {
      var data = JSON.parse(responseBody);
      return data.result;
    } else {
      var errorData = JSON.parse(responseBody);
      return '#ERROR: ' + errorData.error;
    }
  } catch (error) {
    return '#ERROR: ' + error.toString();
  }
}

/**
 * Elenca tutte le operazioni disponibili
 * 
 * @return {string} Lista delle operazioni
 * @customfunction
 */
function CLOUD_CALC_OPERATIONS() {
  var options = {
    'method': 'get',
    'muteHttpExceptions': true
  };
  
  try {
    var response = UrlFetchApp.fetch(API_BASE_URL.replace('/calc', '/operations'), options);
    var data = JSON.parse(response.getContentText());
    return data.operations.join(', ');
  } catch (error) {
    return '#ERROR: ' + error.toString();
  }
}

// ============================================
// ESEMPI DI UTILIZZO NEL FOGLIO:
// ============================================
//
// =CLOUD_CALC("plus", A1, A2)
// =CLOUD_CALC("multiply", B1, B2, B3)
// =CLOUD_CALC("if", A1>10, "Alto", "Basso")
// =CLOUD_CALC("average", C1:C10)
// =CLOUD_CALC("max", D1, D2, D3, D4)
// =CLOUD_CALC("concat", E1, " ", E2)
//
