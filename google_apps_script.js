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
const API_BASE_URL = 'http://18.153.39.218:5000/calc';

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
          row.forEach(function(cell) {
            args.push(cell === "" ? null : cell);
          });
        } else {
          args.push(row === "" ? null : row);
        }
      });
    } else {
      // Celle vuote in Google Sheets arrivano come "" - convertiamo a null
      args.push(arg === "" ? null : arg);
    }
  }

  // IFERROR: gestione locale (non può passare dall'API perché
  // Sheets valuta gli argomenti PRIMA di passarli alla UDF,
  // quindi un errore nell'arg1 arriverebbe come stringa "#DIV/0!" ecc.)
  if (typeof operation === 'string' && operation.toLowerCase() === 'iferror') {
    var value = args.length > 0 ? args[0] : null;
    var fallback = args.length > 1 ? args[1] : "";
    if (containsSheetError_(value)) {
      return fallback;
    }
    return (value === null || value === undefined) ? "" : value;
  }

  // Validazione argomenti prima della chiamata API
  var validationError = validateArgs_(operation, args);
  if (validationError) {
    return validationError;
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
    var contentType = response.getHeaders()['Content-Type'] || '';

    // Verifica che la risposta sia JSON
    if (contentType.indexOf('application/json') === -1) {
      return '#ERROR: Il server non ha risposto con JSON (HTTP ' + responseCode + '). Verifica che l\'endpoint ' + API_BASE_URL + ' sia raggiungibile.';
    }

    if (responseCode === 200) {
      var data = JSON.parse(responseBody);
      var result = data.result;
      if (result === null || result === undefined) {
        return "";
      }
      return result;
    } else {
      var errorData = JSON.parse(responseBody);
      return '#ERROR: ' + errorData.error;
    }
  } catch (error) {
    return '#ERROR: ' + error.toString();
  }
}



/**
 * Appiattisce un range di Google Sheets in un array monodimensionale
 */
function flattenRange_(range) {
  var result = [];
  if (Array.isArray(range)) {
    range.forEach(function(row) {
      if (Array.isArray(row)) {
        row.forEach(function(cell) {
          result.push(cell === "" ? null : cell);
        });
      } else {
        result.push(row === "" ? null : row);
      }
    });
  } else {
    result.push(range === "" ? null : range);
  }
  return result;
}

/**
 * Controlla se un valore e' un errore di Google Sheets.
 * Quando una cella dipendente non e' ancora stata calcolata o e' in errore,
 * Sheets passa stringhe come "#REF!", "#VALUE!", ecc.
 */
function containsSheetError_(value) {
  if (typeof value !== 'string') return false;
  var errorPrefixes = ['#REF!', '#VALUE!', '#N/A', '#NULL!', '#NUM!', '#DIV/0!', '#ERROR', '#NAME?'];
  for (var i = 0; i < errorPrefixes.length; i++) {
    if (value.indexOf(errorPrefixes[i]) === 0) return true;
  }
  return false;
}

/**
 * Valida gli argomenti prima della chiamata API.
 * Restituisce null se tutto OK, altrimenti una stringa di errore.
 */
function validateArgs_(operation, args) {
  // Validazione operazione
  if (operation === null || operation === undefined || operation === '') {
    return '#ERROR: Operazione mancante. Specificare il nome dell\'operazione come primo argomento.';
  }
  if (typeof operation !== 'string') {
    return '#ERROR: L\'operazione deve essere una stringa (es: "plus", "multiply").';
  }

  // Controlla errori di Sheets negli argomenti
  for (var i = 0; i < args.length; i++) {
    if (containsSheetError_(args[i])) {
      return '#ERROR: L\'argomento ' + (i + 1) + ' contiene un errore di Sheets (' + args[i] + '). Verificare le celle di input.';
    }
  }

  return null;
}

/**
 * Controlla se un array (range appiattito) contiene errori di Sheets.
 * Restituisce la prima stringa di errore trovata, o null se tutto OK.
 */
function findSheetErrorInRange_(flatRange) {
  for (var i = 0; i < flatRange.length; i++) {
    if (containsSheetError_(flatRange[i])) {
      return flatRange[i];
    }
  }
  return null;
}

/**
 * Esegue la chiamata API e restituisce il risultato
 */
function callApi_(payload) {
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    var response = UrlFetchApp.fetch(API_BASE_URL, options);
    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();
    var contentType = response.getHeaders()['Content-Type'] || '';

    if (contentType.indexOf('application/json') === -1) {
      return '#ERROR: Il server non ha risposto con JSON (HTTP ' + responseCode + ').';
    }

    if (responseCode === 200) {
      var data = JSON.parse(responseBody);
      var result = data.result;
      if (result === null || result === undefined) {
        return "";
      }
      return result;
    } else {
      var errorData = JSON.parse(responseBody);
      return '#ERROR: ' + errorData.error;
    }
  } catch (error) {
    return '#ERROR: ' + error.toString();
  }
}

/**
 * SUMIFS su cloud: somma condizionale con criteri multipli.
 * Equivalente a SOMMA.PIÙ.SE / SUMIFS di Google Sheets.
 *
 * @param {A1:A100} sum_range - Range dei valori da sommare
 * @param {B1:B100} criteria_range1 - Primo range di criteri
 * @param {string} criteria1 - Primo criterio (es: "metano", ">10", "<>0")
 * @param {C1:C100} criteria_range2 - (Opzionale) Secondo range di criteri
 * @param {string} criteria2 - (Opzionale) Secondo criterio
 * @return {number} Somma dei valori che soddisfano tutti i criteri
 * @customfunction
 */
function CLOUD_SUMIFS(sum_range, criteria_range1, criteria1) {
  // Validazione: sum_range obbligatorio
  if (sum_range === undefined || sum_range === null) {
    return '#ERROR: sum_range mancante. Specificare il range dei valori da sommare.';
  }

  // Validazione: almeno una coppia criteria_range/criteria
  if (arguments.length < 3) {
    return '#ERROR: Servono almeno sum_range, criteria_range e criteria. Forniti solo ' + arguments.length + ' argomenti.';
  }

  // Validazione: argomenti a coppie (criteria_range + criteria)
  if ((arguments.length - 1) % 2 !== 0) {
    return '#ERROR: Gli argomenti dopo sum_range devono essere a coppie (criteria_range, criteria). Numero di argomenti non valido.';
  }

  var flatSumRange = flattenRange_(sum_range);

  // Validazione: errori di Sheets nel sum_range
  var sumRangeError = findSheetErrorInRange_(flatSumRange);
  if (sumRangeError) {
    return '#ERROR: sum_range contiene un errore di Sheets (' + sumRangeError + '). Verificare le celle di input.';
  }

  var criteriaPairs = [];

  // Argomenti a coppie: criteria_range, criteria (a partire dall'indice 1)
  for (var i = 1; i < arguments.length; i += 2) {
    var criteriaRange = flattenRange_(arguments[i]);
    var criteria = arguments[i + 1];

    // Validazione: errori di Sheets nel criteria_range
    var criteriaRangeError = findSheetErrorInRange_(criteriaRange);
    if (criteriaRangeError) {
      return '#ERROR: criteria_range ' + Math.ceil(i / 2) + ' contiene un errore di Sheets (' + criteriaRangeError + '). Verificare le celle di input.';
    }

    // Validazione: criteria_range deve avere la stessa lunghezza di sum_range
    if (criteriaRange.length !== flatSumRange.length) {
      return '#ERROR: criteria_range ' + Math.ceil(i / 2) + ' ha ' + criteriaRange.length + ' elementi, ma sum_range ne ha ' + flatSumRange.length + '. I range devono avere la stessa lunghezza.';
    }

    // Validazione: errori di Sheets nel criterio
    if (containsSheetError_(criteria)) {
      return '#ERROR: Il criterio ' + Math.ceil(i / 2) + ' contiene un errore di Sheets (' + criteria + ').';
    }

    criteriaPairs.push({
      'range': criteriaRange,
      'criteria': criteria === "" ? null : criteria
    });
  }

  var payload = {
    'operation': 'sumifs',
    'sum_range': flatSumRange,
    'criteria_pairs': criteriaPairs
  };

  return callApi_(payload);
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
// ESEMPI DI UTILIZZO NEL FOGLIO (locale IT):
// ============================================
//
// =CLOUD_CALC("plus"; A1; A2)
// =CLOUD_CALC("multiply"; B1; B2; B3)
// =CLOUD_CALC("if"; A1>10; "Alto"; "Basso")
// =CLOUD_CALC("average"; C1:C10)
// =CLOUD_CALC("max"; D1; D2; D3; D4)
// =CLOUD_CALC("concat"; E1; " "; E2)
//
// --- IFERROR ---
// =CLOUD_CALC("iferror"; CLOUD_CALC("divide"; A1; B1); 0)
//
// --- SUMIFS ---
// =CLOUD_SUMIFS(P1:P100; N1:N100; "H Rilevate"; H1:H100; "metano")
//
// Esempio completo (equivale a =SE.ERRORE(SOMMA.PIÙ.SE(...)/SOMMA.PIÙ.SE(...);0)):
// =SE.ERRORE(CLOUD_SUMIFS(P$28:P$4247; $N$28:$N$4247; "H Rilevate"; $H$28:$H$4247; "metano") / CLOUD_SUMIFS(P$28:P$4247; $N$28:$N$4247; "H Rif"; $H$28:$H$4247; "metano"); 0)
//

