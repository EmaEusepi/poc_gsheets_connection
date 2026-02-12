# Cloud Calc API per Google Sheets

API per spostare i calcoli di Google Sheets su cloud, riducendo il carico computazionale sul foglio.

## 🚀 Caratteristiche

- **Argomenti variabili**: Supporta funzioni con qualsiasi numero di argomenti
- **Tipi multipli**: Gestisce numeri, stringhe, booleani
- **Range di celle**: Supporta A1:A10 come argomenti
- **Operazioni condizionali**: `IF`, confronti, operazioni logiche
- **Estendibile**: Facile aggiungere nuove operazioni

## 📋 Operazioni Supportate

### Matematiche
- `plus` - Somma: `=CLOUD_CALC("plus", A1, A2, A3)`
- `minus` - Sottrazione: `=CLOUD_CALC("minus", A1, A2)`
- `multiply` - Moltiplicazione: `=CLOUD_CALC("multiply", A1, A2, A3)`
- `divide` - Divisione: `=CLOUD_CALC("divide", A1, A2)`
- `power` - Potenza: `=CLOUD_CALC("power", 2, 8)`
- `sqrt` - Radice quadrata: `=CLOUD_CALC("sqrt", A1)`
- `abs` - Valore assoluto: `=CLOUD_CALC("abs", A1)`

### Confronto
- `greater` - Maggiore: `=CLOUD_CALC("greater", A1, A2)`
- `less` - Minore: `=CLOUD_CALC("less", A1, A2)`
- `equals` - Uguale: `=CLOUD_CALC("equals", A1, A2)`

### Condizionali
- `if` - Condizionale: `=CLOUD_CALC("if", A1>A2, A1, A2)`

### Aggregate
- `max` - Massimo: `=CLOUD_CALC("max", A1:A10)`
- `min` - Minimo: `=CLOUD_CALC("min", A1:A10)`
- `average` - Media: `=CLOUD_CALC("average", A1:A10)`
- `sum` - Somma: `=CLOUD_CALC("plus", A1:A10)`

### Stringhe
- `concat` - Concatena: `=CLOUD_CALC("concat", A1, " ", B1)`
- `upper` - Maiuscolo: `=CLOUD_CALC("upper", A1)`
- `lower` - Minuscolo: `=CLOUD_CALC("lower", A1)`

## 🔧 Setup Backend

### Locale (Development)

```bash
# Installa dipendenze
pip install -r requirements.txt

# Avvia il server
python cloud_calc_api.py
```

Il server sarà disponibile su `http://localhost:5000`

### Deploy su Cloud

#### Opzione 1: Google Cloud Run (Consigliato)

```bash
# 1. Crea Dockerfile
cat > Dockerfile << 'EOF'
FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY cloud_calc_api.py .
CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 cloud_calc_api:app
EOF

# 2. Deploy
gcloud run deploy cloud-calc-api \
  --source . \
  --platform managed \
  --region europe-west1 \
  --allow-unauthenticated
```

#### Opzione 2: Heroku

```bash
# 1. Crea Procfile
echo "web: gunicorn cloud_calc_api:app" > Procfile

# 2. Deploy
heroku create tuo-nome-app
git init
git add .
git commit -m "Initial commit"
git push heroku main
```

#### Opzione 3: Railway.app

1. Vai su [railway.app](https://railway.app)
2. Collega il tuo repository GitHub
3. Railway rileverà automaticamente Flask e farà il deploy

## 📊 Setup Google Sheets

### Installazione

1. Apri il tuo Google Sheet
2. Vai su **Estensioni** > **Apps Script**
3. Copia il contenuto di `google_apps_script.js`
4. Sostituisci `API_BASE_URL` con il tuo endpoint (es: `https://tua-app.run.app/calc`)
5. Salva il progetto (Ctrl+S)
6. Autorizza lo script quando richiesto

### Utilizzo

```javascript
// Esempi base
=CLOUD_CALC("plus", B1, B2)
=CLOUD_CALC("multiply", A1, A2, A3, A4)

// Condizionali
=CLOUD_CALC("if", B1>B2, B1, B2)
=CLOUD_CALC("if", A1>10, "Alto", "Basso")

// Con range
=CLOUD_CALC("average", A1:A10)
=CLOUD_CALC("max", C1:C20)

// Complessi
=CLOUD_CALC("if", CLOUD_CALC("greater", A1, B1), 
  CLOUD_CALC("plus", A1, 10), 
  CLOUD_CALC("minus", B1, 5))
```

## 🔒 Sicurezza

### Autenticazione API (Opzionale)

Per aggiungere autenticazione via API key:

```python
# Aggiungi all'inizio di cloud_calc_api.py
API_KEY = 'tua-chiave-segreta'

@app.before_request
def check_api_key():
    if request.path != '/health':
        api_key = request.headers.get('X-API-Key')
        if api_key != API_KEY:
            return jsonify({'error': 'Unauthorized'}), 401
```

Nel Google Apps Script:

```javascript
var options = {
  'method': 'post',
  'contentType': 'application/json',
  'headers': {
    'X-API-Key': 'tua-chiave-segreta'
  },
  'payload': JSON.stringify(payload)
};
```

## 🧪 Testing

### Test API locale

```bash
# Test semplice
curl -X POST http://localhost:5000/calc \
  -H "Content-Type: application/json" \
  -d '{"operation": "plus", "args": [5, 3]}'

# Test IF
curl -X POST http://localhost:5000/calc \
  -H "Content-Type: application/json" \
  -d '{"operation": "if", "args": [true, 10, 20]}'

# Lista operazioni
curl http://localhost:5000/operations
```

## ➕ Aggiungere Nuove Operazioni

Modifica il dizionario `OPERATIONS` in `cloud_calc_api.py`:

```python
OPERATIONS = {
    # ... operazioni esistenti ...
    
    # Nuova operazione personalizzata
    'mia_funzione': lambda a, b, c: (a + b) * c,
    
    # Con logica complessa
    'sconto': lambda prezzo, percentuale: prezzo * (1 - percentuale/100),
}
```

Uso in Sheets:
```javascript
=CLOUD_CALC("sconto", A1, 20)  // Applica sconto del 20%
```

## 💡 Vantaggi

1. **Performance**: Calcoli pesanti vengono eseguiti su server potenti
2. **Scalabilità**: Aggiungi server per gestire più richieste
3. **Manutenzione**: Aggiorna la logica senza modificare i fogli
4. **Riuso**: Stessa API per più fogli/utenti
5. **Monitoring**: Traccia l'uso delle funzioni

## 🐛 Troubleshooting

### Errore "Script not authorized"
1. Esegui manualmente una funzione nell'editor Apps Script
2. Autorizza l'accesso quando richiesto

### Errore "Request failed"
- Verifica che l'API sia online: `curl https://tua-app.com/health`
- Controlla che CORS sia abilitato
- Verifica l'URL in `API_BASE_URL`

### Calcoli lenti
- Aumenta i worker Gunicorn: `--workers 4`
- Usa cache per operazioni frequenti
- Considera di implementare batch requests

## 📈 Prossimi Passi

- [ ] Implementare cache Redis per risultati frequenti
- [ ] Aggiungere rate limiting per utente
- [ ] Supportare batch requests (più calcoli in una chiamata)
- [ ] Dashboard per monitoring e analytics
- [ ] Operazioni su date/timestamp
- [ ] Integrazione con database per lookup

## 📄 Licenza

MIT
