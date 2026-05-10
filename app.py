from flask import Flask, render_template, request, jsonify
from datetime import datetime
import openpyxl
from openpyxl import load_workbook
import os

app = Flask(__name__)

# Chemin plus sûr pour le fichier
EXCEL_FILE = os.path.join(os.path.dirname(__file__), "stars 3.xlsx")

mois_feuilles = {1:"Janvier",2:"Février",3:"Mars",4:"Avril",5:"Mai",6:"Juin",
                 7:"Juillet",8:"Aout",9:"Septembre",10:"Octobre",11:"Novembre",12:"Decembre"}

def get_noms():
    # On force les noms en attendant de corriger la lecture
    noms = ["Chloé", "Dorian", "Francky", "Jeremy", "Laurence", "Pouncho !"]
    print(f"✅ Noms forcés : {noms}")
    return noms
@app.route('/')
def index():
    return render_template('index.html', noms=get_noms())

def trouver_colonne_jour(sheet, jour):
    print(f"🔍 Recherche intensive du jour {jour}...")
    
    # On scanne une très large zone
    for row in range(1, 25):
        for col in range(1, 400):
            try:
                cell = sheet.cell(row=row, column=col)
                val = cell.value
                
                if val is None:
                    continue
                    
                # Nettoyage agressif
                if isinstance(val, (int, float)):
                    val_clean = int(val)
                else:
                    val_str = str(val).strip()
                    if val_str.isdigit():
                        val_clean = int(val_str)
                    else:
                        continue
                
                if val_clean == jour:
                    print(f"✅ JOUR {jour} TROUVÉ ! Ligne {row} - Colonne {col}")
                    return col  # On retourne la colonne M (début du trio)
            except:
                continue
    
    print(f"❌ Jour {jour} IMPOSSIBLE à trouver")
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

    # === RECHERCHE ULTRA TOLÉRANTE DU NOM ===
    row_nom = None
    nom_clean = str(nom).strip().lower()
    print(f"🔍 Recherche de '{nom}' (nettoyé: '{nom_clean}') dans {sheet_name}")

    for r in range(1, 100):  # On cherche très large
        cell = sheet.cell(row=r, column=1).value
        if cell is None:
            continue
        cell_clean = str(cell).strip().lower()
        
        if cell_clean == nom_clean or cell_clean.startswith(nom_clean) or nom_clean in cell_clean:
            row_nom = r
            print(f"✅ NOM TROUVÉ à la ligne {r} : {cell}")
            break

    if not row_nom:
        # On essaie aussi les colonnes voisines au cas où
        for r in range(1, 100):
            for c in range(1, 5):
                cell = sheet.cell(row=r, column=c).value
                if cell and str(cell).strip().lower() == nom_clean:
                    row_nom = r
                    print(f"✅ Trouvé dans colonne {c}, ligne {r}")
                    break
            if row_nom:
                break

    if not row_nom:
        print("❌ Nom toujours pas trouvé malgré tout")
        return jsonify({"success": False, "message": f"Nom '{nom}' non trouvé dans le planning du mois"})
    col_m = trouver_colonne_jour(sheet, jour)
    if not col_m:
        return jsonify({"success": False, "message": "Jour non trouvé dans le mois"})

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
        data = {}
        
        for sheet_name in mois_feuilles.values():
            if sheet_name in wb:
                sheet = wb[sheet_name]
                rows = []
                for r in range(10, 30):  # Lignes des stars
                    row = []
                    for c in range(1, 40):
                        val = sheet.cell(row=r, column=c).value
                        row.append(val if val is not None else "")
                    if any(row):  # Si la ligne n'est pas vide
                        rows.append(row)
                data[sheet_name] = rows
        return jsonify({"success": True, "data": data})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})
if __name__ == '__main__':
    print("🚀 Application démarrée sur http://127.0.0.1:5000")
    app.run(host='0.0.0.0', port=5000, debug=True)
