"""
Cloud Calc Batch API
====================
Server che accumula le chiamate CLOUD_CALC provenienti da Google Sheets,
ricostruisce il grafo di dipendenze fra celle e le risolve nell'ordine
corretto (topological sort).

Flusso:
1. Ogni cella con CLOUD_CALC invia: cell, operation, args [{ref, value}]
2. Il server accumula le richieste in un batch (finestra di BATCH_WINDOW_S
   secondi senza nuove richieste).
3. Scaduta la finestra, costruisce il grafo di dipendenze, esegue
   topological sort e calcola ogni cella nell'ordine giusto, propagando
   i risultati alle celle dipendenti.
4. Ogni request HTTP riceve la propria risposta.

Avvio:  python cloud_calc_batch_api.py
"""

from __future__ import annotations

from flask import Flask, request, jsonify
from flask_cors import CORS
import threading
import time
import math
import re
import fnmatch

# ---------------------------------------------------------------------------
# Configurazione
# ---------------------------------------------------------------------------
BATCH_WINDOW_S = 2.0   # secondi di silenzio prima di risolvere il batch

app = Flask(__name__)
CORS(app)

# ---------------------------------------------------------------------------
# Operazioni supportate (stesse di cloud_calc_api.py)
# ---------------------------------------------------------------------------
OPERATIONS = {
    'plus':       lambda *a: sum(a),
    'minus':      lambda a, b: a - b,
    'multiply':   lambda *a: math.prod(a),
    'divide':     lambda a, b: a / b if b != 0 else '#DIV/0!',
    'power':      lambda a, b: a ** b,
    'mod':        lambda a, b: a % b,
    'equals':     lambda a, b: a == b,
    'greater':    lambda a, b: a > b,
    'less':       lambda a, b: a < b,
    'greater_equal': lambda a, b: a >= b,
    'less_equal': lambda a, b: a <= b,
    'and':        lambda *a: all(a),
    'or':         lambda *a: any(a),
    'not':        lambda a: not a,
    'if':         lambda cond, t, f: t if cond else f,
    'iferror':    lambda v, fb=0: fb if (isinstance(v, str) and v.startswith('#')) else v,
    'sqrt':       lambda a: math.sqrt(a),
    'abs':        lambda a: abs(a),
    'round':      lambda a, d=0: round(a, int(d)),
    'floor':      lambda a: math.floor(a),
    'ceil':       lambda a: math.ceil(a),
    'max':        lambda *a: max(a),
    'min':        lambda *a: min(a),
    'average':    lambda *a: sum(a) / len(a) if a else 0,
    'count':      lambda *a: len(a),
    'concat':     lambda *a: ''.join(str(x) for x in a),
    'upper':      lambda s: str(s).upper(),
    'lower':      lambda s: str(s).lower(),
    'trim':       lambda s: str(s).strip(),
    'len':        lambda s: len(str(s)),
}

# ---------------------------------------------------------------------------
# Parse helpers
# ---------------------------------------------------------------------------
def parse_value(value):
    if value is None or value == '':
        return None
    if isinstance(value, (int, float, bool)):
        return value
    s = str(value).strip()
    if s.lower() == 'true':
        return True
    if s.lower() == 'false':
        return False
    try:
        return float(s) if '.' in s else int(s)
    except ValueError:
        return s


def normalize_cell(ref):
    """Normalizza un riferimento cella in maiuscolo senza $: '$A$1' -> 'A1'."""
    return ref.replace('$', '').upper().strip()


# ---------------------------------------------------------------------------
# Batch manager
# ---------------------------------------------------------------------------
class BatchManager:
    """Accumula le request in arrivo e le risolve quando la finestra scade."""

    def __init__(self, window_s=BATCH_WINDOW_S):
        self.window_s = window_s
        self._lock = threading.Lock()
        self._batch: dict[str, dict] = {}      # cell -> entry
        self._timer: threading.Timer | None = None

    # -- public API (chiamato dal thread Flask per ogni request) ------------

    def submit(self, cell: str, operation: str, args: list[dict]) -> dict:
        """Aggiunge una cella al batch corrente e blocca fino alla risoluzione.

        Ritorna il dict {'result': ...} oppure {'error': ...}.
        """
        cell = normalize_cell(cell)
        event = threading.Event()
        entry = {
            'cell': cell,
            'operation': operation,
            'args': args,           # [{ref: "A1", value: 10}, ...]
            'event': event,
            'result': None,
        }

        with self._lock:
            self._batch[cell] = entry
            self._reset_timer()

        # Blocca fino a risoluzione (timeout di sicurezza)
        event.wait(timeout=30)

        if entry['result'] is None:
            return {'error': 'Timeout: batch non risolto entro 30s'}
        return entry['result']

    # -- internals ----------------------------------------------------------

    def _reset_timer(self):
        """Resetta (o avvia) il timer della finestra di batch."""
        if self._timer is not None:
            self._timer.cancel()
        self._timer = threading.Timer(self.window_s, self._resolve_batch)
        self._timer.daemon = True
        self._timer.start()

    def _resolve_batch(self):
        """Scaduta la finestra: risolvi tutto il batch."""
        with self._lock:
            batch = self._batch
            self._batch = {}
            self._timer = None

        if not batch:
            return

        cells_in_batch = set(batch.keys())
        print(f"\n[BATCH] Risoluzione di {len(batch)} celle: {sorted(cells_in_batch)}")

        # 1. Costruisci il grafo di dipendenze (solo fra celle nel batch)
        #    deps[cell] = set di celle nel batch da cui dipende
        deps: dict[str, set[str]] = {}
        for cell, entry in batch.items():
            cell_deps = set()
            for arg in entry['args']:
                ref = normalize_cell(arg.get('ref', ''))
                if ref in cells_in_batch and ref != cell:
                    cell_deps.add(ref)
            deps[cell] = cell_deps

        # 2. Topological sort (Kahn's algorithm)
        order = self._topological_sort(deps)
        if order is None:
            # Ciclo nelle dipendenze
            for entry in batch.values():
                entry['result'] = {'error': 'Dipendenza circolare rilevata nel batch'}
                entry['event'].set()
            return

        print(f"[BATCH] Ordine di esecuzione: {order}")

        # 3. Esegui in ordine, propagando i risultati
        computed: dict[str, any] = {}   # cell -> valore calcolato

        for cell in order:
            entry = batch[cell]
            op = entry['operation'].lower()

            # Sostituisci i valori degli argomenti che dipendono da celle
            # gia' calcolate in questo batch
            resolved_args = []
            for arg in entry['args']:
                ref = normalize_cell(arg.get('ref', ''))
                if ref in computed:
                    # Usa il valore appena calcolato
                    resolved_args.append(parse_value(computed[ref]))
                else:
                    resolved_args.append(parse_value(arg.get('value')))

            # Calcola
            try:
                if op not in OPERATIONS:
                    entry['result'] = {'error': f'Operazione sconosciuta: {op}'}
                else:
                    result = OPERATIONS[op](*resolved_args)
                    computed[cell] = result
                    entry['result'] = {'result': result, 'cell': cell}
                    print(f"[BATCH]   {cell} = {result}")
            except Exception as e:
                entry['result'] = {'error': f'Errore calcolo {cell}: {str(e)}'}

            entry['event'].set()

        # Sblocca eventuali celle non nell'ordine (non dovrebbe succedere)
        for entry in batch.values():
            if not entry['event'].is_set():
                entry['result'] = {'error': 'Cella non risolta'}
                entry['event'].set()

    @staticmethod
    def _topological_sort(deps: dict[str, set[str]]) -> list[str] | None:
        """Kahn's algorithm. Ritorna l'ordine o None se c'e' un ciclo."""
        in_degree = {n: 0 for n in deps}
        for node, node_deps in deps.items():
            for d in node_deps:
                if d not in in_degree:
                    in_degree[d] = 0

        # Calcola in-degree (quante celle dipendono da me -> irrilevante,
        # ci serve quante dipendenze ha ogni cella)
        # in_degree[cell] = numero di dipendenze non ancora risolte
        adj: dict[str, list[str]] = {n: [] for n in in_degree}
        for node, node_deps in deps.items():
            in_degree[node] = len(node_deps)
            for d in node_deps:
                adj.setdefault(d, []).append(node)

        queue = [n for n, deg in in_degree.items() if deg == 0]
        order = []
        while queue:
            node = queue.pop(0)
            order.append(node)
            for neighbor in adj.get(node, []):
                in_degree[neighbor] -= 1
                if in_degree[neighbor] == 0:
                    queue.append(neighbor)

        if len(order) != len(in_degree):
            return None  # ciclo
        return order


# Singleton
batch_manager = BatchManager()


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@app.before_request
def _start_timer():
    request._start_time = time.time()


@app.after_request
def _log_request_time(response):
    elapsed = time.time() - getattr(request, '_start_time', time.time())
    print(f"[{request.method} {request.path}] {response.status_code} - {elapsed*1000:.0f}ms")
    return response


@app.route('/batch_calc', methods=['POST'])
def batch_calc():
    """Endpoint per le chiamate batch con dipendenze.

    Payload atteso:
    {
        "cell": "C3",
        "operation": "sum",
        "args": [
            {"ref": "A1", "value": 10},
            {"ref": "A2", "value": 20}
        ]
    }
    """
    try:
        data = request.get_json()
        cell = data.get('cell', '')
        operation = data.get('operation', '')
        args = data.get('args', [])

        if not cell:
            return jsonify({'error': 'cell is required'}), 400
        if not operation:
            return jsonify({'error': 'operation is required'}), 400

        result = batch_manager.submit(cell, operation, args)
        status = 200 if 'result' in result else 500
        return jsonify(result), status

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'healthy', 'mode': 'batch'})


@app.route('/operations', methods=['GET'])
def list_operations():
    return jsonify({'operations': sorted(OPERATIONS.keys())})


if __name__ == '__main__':
    print(f"Batch Cloud Calc API - finestra batch: {BATCH_WINDOW_S}s")
    print("Endpoints:")
    print("  POST /batch_calc  - calcolo con dipendenze (batch)")
    print("  GET  /health")
    print("  GET  /operations")
    app.run(host='0.0.0.0', port=5001, debug=False, threaded=True)
