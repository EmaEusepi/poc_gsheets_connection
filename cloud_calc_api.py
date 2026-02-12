from flask import Flask, request, jsonify
from flask_cors import CORS
import operator
import math

app = Flask(__name__)
CORS(app)  # Necessario per chiamate da Google Sheets

# Dizionario delle operazioni supportate
OPERATIONS = {
    # Operazioni matematiche base
    'plus': lambda *args: sum(args),
    'minus': lambda a, b: a - b,
    'multiply': lambda *args: math.prod(args),
    'divide': lambda a, b: a / b if b != 0 else '#DIV/0!',
    'power': lambda a, b: a ** b,
    'mod': lambda a, b: a % b,
    
    # Operazioni di confronto
    'equals': lambda a, b: a == b,
    'greater': lambda a, b: a > b,
    'less': lambda a, b: a < b,
    'greater_equal': lambda a, b: a >= b,
    'less_equal': lambda a, b: a <= b,
    
    # Operazioni logiche
    'and': lambda *args: all(args),
    'or': lambda *args: any(args),
    'not': lambda a: not a,
    
    # Operazioni condizionali
    'if': lambda condition, true_val, false_val: true_val if condition else false_val,
    
    # Funzioni matematiche
    'sqrt': lambda a: math.sqrt(a),
    'abs': lambda a: abs(a),
    'round': lambda a, decimals=0: round(a, int(decimals)),
    'floor': lambda a: math.floor(a),
    'ceil': lambda a: math.ceil(a),
    
    # Funzioni aggregate
    'max': lambda *args: max(args),
    'min': lambda *args: min(args),
    'average': lambda *args: sum(args) / len(args) if args else 0,
    'count': lambda *args: len(args),
    
    # Funzioni stringa
    'concat': lambda *args: ''.join(str(x) for x in args),
    'upper': lambda s: str(s).upper(),
    'lower': lambda s: str(s).lower(),
    'trim': lambda s: str(s).strip(),
    'len': lambda s: len(str(s)),
}

def parse_value(value):
    """Converte il valore nel tipo appropriato"""
    if value is None or value == '':
        return None
    
    # Prova a convertire in numero
    if isinstance(value, (int, float, bool)):
        return value
    
    value_str = str(value).strip()
    
    # Booleani
    if value_str.lower() == 'true':
        return True
    if value_str.lower() == 'false':
        return False
    
    # Numeri
    try:
        if '.' in value_str:
            return float(value_str)
        return int(value_str)
    except ValueError:
        return value_str

@app.route('/calc', methods=['POST', 'GET'])
def calculate():
    try:
        # Supporta sia GET che POST
        if request.method == 'POST':
            data = request.get_json()
        else:
            data = {
                'operation': request.args.get('operation'),
                'args': request.args.getlist('args')
            }
        
        operation = data.get('operation', '').lower()
        args_raw = data.get('args', [])
        
        # Valida l'operazione
        if not operation:
            return jsonify({'error': 'Operation is required'}), 400
        
        if operation not in OPERATIONS:
            return jsonify({
                'error': f'Unknown operation: {operation}',
                'available': list(OPERATIONS.keys())
            }), 400
        
        # Parse degli argomenti
        args = [parse_value(arg) for arg in args_raw]
        
        # Esegui l'operazione
        result = OPERATIONS[operation](*args)
        
        return jsonify({
            'result': result,
            'operation': operation,
            'args': args
        })
    
    except TypeError as e:
        return jsonify({
            'error': f'Invalid number of arguments: {str(e)}',
            'operation': operation
        }), 400
    except Exception as e:
        return jsonify({
            'error': str(e),
            'operation': operation
        }), 500

@app.route('/operations', methods=['GET'])
def list_operations():
    """Elenca tutte le operazioni disponibili"""
    return jsonify({
        'operations': list(OPERATIONS.keys())
    })

@app.route('/health', methods=['GET'])
def health():
    """Health check endpoint"""
    return jsonify({'status': 'healthy'})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
