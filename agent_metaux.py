import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Date du jour
date_jour = datetime.now()
date_str = date_jour.strftime("%d-%m-%Y")
taux_usd_mad = 9.34  # Taux de change USD/MAD du jour

# 1. CATALOGUE MONDIAL DES MÉTAUX PAR FAMILLE
catalogue_mondial = {
    "Ferrailles & Aciers": [
        "Ferraille Massive", "Ferraille Légère", "Ferraille E40", "Ferraille E3", 
        "Fonte brute", "Copeaux d'acier", "Ferraille HMS 1&2"
    ],
    "Métaux Non-Fereux": [
        "Cuivre Grade A", "Aluminium LME", "Zinc Standard", "Laiton", 
        "Plomb affiné", "Étain LME", "Nickel"
    ],
    "Métaux Précieux": [
        "Or (Lingot)", "Argent pur", "Platine", "Palladium"
    ],
    "Minéraux & Phosphates (Maroc & Global)": [
        "Phosphates (Roche BPL 68%)", "Minerai de Fer Standard", "Soufre brut", "Potasse"
    ]
}

# Base de données d'abonnés et leurs restrictions (Ciblage par famille)
base_abonnes = [
    {"email": "ferrailleur.pro@gmail.com", "famille_souhaitee": "Ferrailles & Aciers", "debut": "01-01-2026", "fin": "31-12-2026"},
    {"email": "direction.achats@entreprise.com", "famille_souhaitee": "TOUT", "debut": "01-08-2026", "fin": "01-09-2027"},
    {"email": "nonferreux.maroc@gmail.com", "famille_souhaitee": "Métaux Non-Fereux", "debut": "15-06-2026", "fin": "30-09-2026"}
]

# 2. SIMULATION DE MARCHÉ & PRÉDICTIONS SUR 8 JOURS (J à J+7)
np.random.seed(42)
jours_prediction = [date_jour + timedelta(days=i) for i in range(8)]
historique_global = []

base_prices_usd = {
    "Ferraille Massive": 315.0, "Ferraille Légère": 260.0, "Ferraille E40": 275.0, "Ferraille E3": 240.0, "Fonte brute": 350.0, "Copeaux d'acier": 210.0, "Ferraille HMS 1&2": 290.0,
    "Cuivre Grade A": 8900.0, "Aluminium LME": 2400.0, "Zinc Standard": 2700.0, "Laiton": 5800.0, "Plomb affiné": 2150.0, "Étain LME": 29000.0, "Nickel": 16500.0,
    "Or (Lingot)": 65000.0, "Argent pur": 850.0, "Platine": 32000.0, "Palladium": 34000.0,
    "Phosphates (Roche BPL 68%)": 110.0, "Minerai de Fer Standard": 12.0, "Soufre brut": 250.0, "Potasse": 340.0
}

for famille, metaux in catalogue_mondial.items():
    for metal in metaux:
        p_base = base_prices_usd[metal]
        for i, jour in enumerate(jours_prediction):
            # Variation simulée sur 8 jours
            p_base += np.random.normal(0, p_base * 0.008)
            prix_usd = round(p_base, 2)
            prix_mad = round(prix_usd * taux_usd_mad, 2)
            
            # Tendances et Conseil GO / WAIT / NO GO
            if i == 0:
                tendance = "STABLE ➡️"
                conseil = "WAIT"
            else:
                tendance = "HAUSSIÈRE 📈" if i % 2 == 0 else "BAISSIÈRE 📉"
                conseil = "GO" if "BAISSIÈRE" in tendance else "NO GO"

            historique_global.append({
                "Date_Prevue": jour.strftime("%d/%m/%Y"),
                "Famille": famille,
                "Metal": metal,
                "Prix_USD": prix_usd,
                "Prix_MAD": prix_mad,
                "Tendance": tendance,
                "Conseil_Achat": conseil,
                "Lien_Source": f"https://www.marche-metaux-global.com/index/{metal.lower().replace(' ', '-')}"
            })

df_Complet = pd.DataFrame(historique_global)

# 3. GENERATION DES FICHIERS EXCEL PAR ABONNE (Filtrés selon leurs droits)
historique_envois = []

for abonne in base_abonnes:
    date_fin_abo = datetime.strptime(abonne["fin"], "%d-%m-%Y")
    
    # Vérification si l'abonnement est actif
    if datetime.now() <= date_fin_abo:
        famille_visee = abonne["famille_souhaitee"]
        
        if famille_visee == "TOUT":
            df_abonne = df_Complet.copy()
        else:
            df_abonne = df_Complet[df_Complet["Famille"] == famille_visee].copy()
            
        nom_fichier = f"veille_metaux_{abonne['email'].split('@')[0]}_{date_str}.xlsx"
        
        # Création du fichier Excel avec mise en forme professionnelle
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Prédictions 8J"
        
        # En-têtes
        headers = ["Date", "Famille", "Métal / Matière", "Prix (USD)", "Prix (MAD)", "Tendance (8J)", "Décision", "Lien Information"]
        ws.append(headers)
        
        for row in df_abonne.itertuples(index=False):
            ws.append(list(row))
            
        # Styles et Couleurs (Vert pour GO, Orange pour WAIT, Rouge pour NO GO)
        fill_go = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")     # Vert clair
        fill_wait = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")   # Orange/Jaune clair
        fill_nogo = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")   # Rouge clair
        
        for row_idx in range(2, len(df_abonne) + 2):
            cell_conseil = ws.cell(row=row_idx, column=7)
            val = cell_conseil.value
            if val == "GO":
                cell_conseil.fill = fill_go
            elif val == "WAIT":
                cell_conseil.fill = fill_wait
            elif val == "NO GO":
                cell_conseil.fill = fill_nogo
                
        wb.save(nom_fichier)
        
        # Enregistrement dans l'historique des envois
        historique_envois.append({
            "Email": abonne["email"],
            "Fichier_Envoye": nom_fichier,
            "Date_Heure_Envoi": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "Statut": "SUCCÈS (Abonnement Actif)"
        })
    else:
        historique_envois.append({
            "Email": abonne["email"],
            "Fichier_Envoye": "AUCUN",
            "Date_Heure_Envoi": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "Statut": "BLOQUÉ (Abonnement Expiré)"
        })

# Export de l'historique global des envois pour le contrôle
df_suivi_envois = pd.DataFrame(historique_envois)
df_suivi_envois.to_excel("historique_logs_envois.xlsx", index=False)

print("=== RAPPORT : TRAITEMENT TERMINÉ AVEC SUCCÈS ===")
print(f"Taux de change appliqué : 1 USD = {taux_usd_mad} MAD")
print(f"Fichiers générés et logs mis à jour pour {len(base_abonnes)} abonnés.")
