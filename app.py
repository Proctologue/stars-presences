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
    nom = data.get('nom')
    date_str = data.get('date')
    selections = data.get('selections', {})

    if not nom or not date_str:
        return jsonify({"success": False, "message": "Données manquantes"})

    try:
        date = datetime.strptime(date_str, "%Y-%m-%d")
        sheet_name = {1:"Janvier",2:"Février",3:"Mars",4:"Avril",5:"Mai",6:"Juin",
                      7:"Juillet",8:"Aout",9:"Septembre",10:"Octobre",11:"Novembre",12:"Decembre"}[date.month]
        jour = date.day
    except:
        return jsonify({"success": False, "message": "Date invalide"})

    wb = load_workbook(EXCEL_FILE)
    sheet = wb[sheet_name]

    # Recherche nom
    row_nom = None
    for r in range(1, 100):
        for c in range(1, 6):
            if str(sheet.cell(row=r, column=c).value or "").strip() == nom:
                row_nom = r
                break
        if row_nom: break

    if not row_nom:
        return jsonify({"success": False, "message": f"Nom '{nom}' non trouvé"})

    # Recherche jour
    col_m = None
    for r in range(1, 30):
        for c in range(1, 400):
            val = sheet.cell(row=r, column=c).value
            if val and str(val).strip().isdigit() and int(float(val)) == jour:
                col_m = c
                break
        if col_m: break

    if not col_m:
        return jsonify({"success": False, "message": f"Jour {jour} non trouvé"})

    # Enregistrement
    for slot, choix in selections.items():
        valeur = "D" if choix == "DISPO" else "X"
        if slot == "Matin":
            sheet.cell(row=row_nom, column=col_m, value=valeur)
        elif slot == "Après-midi":
            sheet.cell(row=row_nom, column=col_m+1, value=valeur)
        elif slot == "Soir":
            sheet.cell(row=row_nom, column=col_m+2, value=valeur)

    wb.save(EXCEL_FILE)
    return jsonify({"success": True, "message": f"✅ Enregistré le {jour}/{date.month}"})

if __name__ == '__main__':
    print("🚀 Application minimale démarrée")
    app.run(host='0.0.0.0', port=5000, debug=True)
    
