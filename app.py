from flask import Flask, render_template, request, jsonify
from datetime import datetime
import openpyxl
from openpyxl import load_workbook
import os

app = Flask(__name__)

# Chemin du fichier Excel
EXCEL_FILE = os.path.join(os.path.dirname(__file__), "stars 3.xlsx")

mois_feuilles = {
    1:"Janvier", 2:"Février", 3:"Mars", 4:"Avril", 5:"Mai", 6:"Juin",
    7:"Juillet", 8:"Aout", 9:"Septembre", 10:"Octobre", 11:"Novembre", 12:"Decembre"
}

def get_noms():
    try:
        wb = load_workbook(EXCEL_FILE, data_only=True)
        sheet = wb["Noms des employés"]
        noms = []
        for r in range(1, 30):
            cell = sheet.cell(row=r, column=1).value
            if cell:
                nom = str(cell).strip()
                if nom and nom not in ["Noms des employés", ""]:
                    noms.append(nom)
        return noms
    except:
        return ["Chloé", "Dorian", "Francky", "Jeremy", "Laurence", "Pouncho !"]

@app.route('/')
def index():
    return render_template('index.html', noms=get_noms())

def trouver_colonne_jour(sheet, jour):
    for row in range(1, 25):
        for col in range(1, 400):
            try:
                val = sheet.cell(row=row, column=col).value
                if val is None:
                    continue
                if isinstance(val, (int, float)):
                    if int(val) == jour:
                        return col
                elif str(val).strip().isdigit():
                    if int(str(val).strip()) == jour:
                        return col
            except:
                continue
    return None

@app.route('/sauvegarder', methods=['POST'])
def sauvegarder():
    data = request.json
    nom = data['nom']
    date_str = data['date']
    selections = data['selections']

    try:
        date = datetime.strptime(date_str, "%Y-%m-%d")
        sheet_name = mois_feuilles[date.month]
        jour = date.day
    except:
        return jsonify({"success": False, "message": "Date invalide"})

    wb = load_workbook(EXCEL_FILE)
    sheet = wb[sheet_name]

    # Recherche du nom
    row_nom = None
    nom_clean = str(nom).strip().lower()
    for r in range(1, 100):
        for c in range(1, 6):
            cell = sheet.cell(row=r, column=c).value
            if cell and str(cell).strip().lower() == nom_clean:
                row_nom = r
                break
        if row_nom:
            break

    if not row_nom:
        return jsonify({"success": False, "message": f"Nom '{nom}' non trouvé"})

    col_m = trouver_colonne_jour(sheet, jour)
    if not col_m:
        return jsonify({"success": False, "message": f"Jour {jour} non trouvé dans {sheet_name}"})

    # Mise à jour
    for slot, choix in selections.items():
        valeur = "D" if choix == "DISPO" else "X"
        if slot == "Matin":
            sheet.cell(row=row_nom, column=col_m, value=valeur)
        elif slot == "Après-midi":
            sheet.cell(row=row_nom, column=col_m+1, value=valeur)
        elif slot == "Soir":
            sheet.cell(row=row_nom, column=col_m+2, value=valeur)

    wb.save(EXCEL_FILE)
    return jsonify({"success": True, "message": f"✅ Enregistré pour le {jour}/{date.month}"})

@app.route('/planning')
def get_planning():
    try:
        wb = load_workbook(EXCEL_FILE, data_only=True)
        html = "<h1>📊 Planning Complet 2026</h1>"
        
        for sheet_name in mois_feuilles.values():
            if sheet_name not in wb:
                continue
            sheet = wb[sheet_name]
            html += f"<h2>{sheet_name}</h2>"
            html += "<table border='1' style='border-collapse:collapse; width:100%;'>"
            html += "<tr><th>Nom</th>"
            
            # En-têtes jours
            for c in range(2, 100, 3):
                jour = sheet.cell(row=8, column=c).value
                if jour and str(jour).isdigit():
                    html += f"<th>{jour}</th>"
            html += "</tr>"
            
            # Données des stars
            for r in range(11, 30):
                nom = sheet.cell(row=r, column=1).value
                if not nom:
                    continue
                html += f"<tr><td><b>{nom}</b></td>"
                for c in range(2, 100, 3):
                    m = sheet.cell(row=r, column=c).value or ""
                    am = sheet.cell(row=r, column=c+1).value or ""
                    s = sheet.cell(row=r, column=c+2).value or ""
                    html += f"<td>{m} {am} {s}</td>"
                html += "</tr>"
            html += "</table><br><br>"
        
        return html
    except Exception as e:
        return f"<h2>Erreur : {str(e)}</h2>"

if __name__ == '__main__':
    print("🚀 Application démarrée sur http://127.0.0.1:5000")
    app.run(host='0.0.0.0', port=5000, debug=True)
