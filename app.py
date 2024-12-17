
from flask import Flask, jsonify, request
import pandas as pd

app = Flask(__name__)

# Percorso del file Excel (deve essere aggiornato con il tuo percorso locale)
EXCEL_FILE = "C:\PRIVATO\Lavoro\Aheda - C M base r1.xlsx"
SHEET_NAME = "IA Progetti"

# Funzione per leggere il foglio Excel
def read_excel_data():
    # Legge i dati dal foglio "IA Progetti"
    try:
        data = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, engine='openpyxl')
        data.columns = ["x", "Pos", "Macroattività", "x", "Jr Servizi", "Dev Servizi", "An Servizi", "Sr Servizi", "Expert Servizi",
                        "Tech. 1", "Giorni", "Costo", "Prezzo"]
        data = data.drop(columns=["x"])
        data = data.drop(index=range(0,4))
        json = pd.DataFrame(data)
        json.to_json("output.json", orient="records", indent=4)
        return data.to_dict(orient="records")  # Restituisce i dati come lista di dizionari
    except Exception as e:
        return {"error": str(e)}

#Funzione per cambiare le intestazioni del dizionario
#def formatData(data):


# Endpoint per ottenere i dati del foglio "IA Progetti"
@app.route('/api/get-ia-progetti', methods=['GET'])
def get_ia_progetti():
    data = read_excel_data()
    if "error" in data:
        return jsonify({"status": "error", "message": data["error"]}), 500
    return jsonify({"status": "success", "data": data})

# Endpoint di esempio per filtrare i dati in base a una colonna
@app.route('/api/filter-ia-progetti', methods=['GET'])
def filter_ia_progetti():
    column = request.args.get('column')  # Nome della colonna da filtrare
    value = request.args.get('value')    # Valore da cercare

    if not column or not value:
        return jsonify({"status": "error", "message": "Specifica 'column' e 'value' come parametri GET"}), 400

    data = read_excel_data()
    if "error" in data:
        return jsonify({"status": "error", "message": data["error"]}), 500

    # Filtro dei dati
    filtered_data = [row for row in data if str(row.get(column, "")).lower() == value.lower()]
    return jsonify({"status": "success", "filtered_data": filtered_data})

# Configura l'app Flask per ascoltare su 0.0.0.0 e sulla porta specificata
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))  # Usa la variabile PORT o la porta 5000 come default
    app.run(host='0.0.0.0', port=port)
