from flask import Flask, render_template, request, jsonify
from datetime import datetime
import openpyxl
from openpyxl import load_workbook
import os

app = Flask(__name__)

EXCEL_FILE = os.path.join(os.path.dirname(__file__), "stars 3.xlsx")

NOMS = ["Chloé", "Dorian", "Francky", "Jeremy", "Laurence", "Pouncho !"]

@app.route('/')
def index():
    return render_template('index.html', noms=NOMS)

@app.route('/sauvegarder', methods=['POST'])
def sauvegarder():
    data = request.json
    nom = data['nom']
    date_str = data['date']
    selections = data['selections']

    try:
        date = datetime.strptime(date_str, "%Y-%m-%d")
        jour_str = date.strftime("%d/%m/%Y")
    except:
        return jsonify({"success": False, "message": "Date invalide"})

    wb = load_workbook(EXCEL_FILE)
    sheet = wb.active  # On utilise la feuille "Planning"

    # Trouver la ligne du nom + date
    row = None
    for r in range(2, 500):
        if sheet.cell(row=r, column=1).value == jour_str and sheet.cell(row=r, column=2).value == nom:
            row = r
            break

    if not row:
        # Créer une nouvelle ligne si elle n'existe pas
        row = sheet.max_row + 1
        sheet.cell(row=row, column=1, value=jour_str)
        sheet.cell(row=row, column=2, value=nom)

    # Mise à jour des disponibilités
    if selections.get("Matin") == "DISPO":
        sheet.cell(row=row, column=3, value="D")
    if selections.get("Après-midi") == "DISPO":
        sheet.cell(row=row, column=4, value="D")
    if selections.get("Soir") == "DISPO":
        sheet.cell(row=row, column=5, value="D")

    wb.save(EXCEL_FILE)
    return jsonify({"success": True, "message": f"✅ Enregistré le {date.strftime('%d/%m/%Y')}"})

if __name__ == '__main__':
    print("🚀 Application démarrée avec fichier propre")
    app.run(host='0.0.0.0', port=5000, debug=True)
