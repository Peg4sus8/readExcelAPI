import os
from io import BytesIO
import requests
from flask import Flask, jsonify, request
import pandas as pd

app = Flask(__name__)

# Percorso del file Excel (deve essere aggiornato con il tuo percorso locale)
EXCEL_URL = "https://aizoon365.sharepoint.com/:x:/r/sites/GESTIONEMATRICECOMPETENZE/_layouts/15/Doc.aspx?sourcedoc=%7BB2B301FF-0A60-4F04-AC67-B805AB48E080%7D&file=Aheda%20-%20C%20M%20base%20r1.xlsx&action=default&mobileredirect=true"
SHEET_NAME = "IA Progetti"

# Funzione per leggere il foglio Excel
def read_excel_data():
    try:
        # Scarica il file Excel online
        response = requests.get(EXCEL_URL)
        response.raise_for_status()  # Controlla eventuali errori HTTP

        # Converte il contenuto scaricato in un oggetto BytesIO
        excel_data = BytesIO(response.content)

        # Legge i dati dal foglio Excel specifico
        data = pd.read_excel(excel_data, sheet_name=SHEET_NAME, engine='openpyxl')

        # Rinomina le colonne
        data.columns = ["x", "Pos", "Macroattività", "x", "Jr Servizi", "Dev Servizi", "An Servizi", "Sr Servizi", "Expert Servizi",
                        "Tech. 1", "Giorni", "Costo", "Prezzo"]

        # Rimuove le colonne non necessarie
        data = data.drop(columns=["x"])

        # Rimuove le prime 4 righe
        data = data.drop(index=range(0, 4))

        # Esporta i dati in formato JSON
        data.to_json("output.json", orient="records", indent=4)

        # Restituisce i dati come lista di dizionari
        return data.to_dict(orient="records")
    except Exception as e:
        # In caso di errore, restituisce un dizionario con il messaggio di errore
        return {"error": str(e)}

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
