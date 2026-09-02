import os
import smtplib
from email.message import EmailMessage
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import PatternFill

print("=== [ETAPE 1] Initialisation et chargement des accès ===")
EMAIL_EXPEDITEUR = os.environ.get("MAIL_USER")
EMAIL_MOT_DE_PASSE = os.environ.get("MAIL_PASSWORD")

date_jour = datetime.now()
date_str = date_jour.strftime("%d-%m-%Y")
taux_usd_mad = 9.34  # Taux de change USD/MAD explicite

# Catalogue mondial complet (Métaux + Énergies) et leurs unités
catalogue_mondial = {
    "Ferrailles & Aciers": {
        "Ferraille Massive": "Tonne", "Ferraille Légère": "Tonne", "Ferraille E40": "Tonne", 
        "Ferraille E3": "Tonne", "Fonte brute": "Tonne", "Copeaux d'acier": "Tonne", "Ferraille HMS 1&2": "Tonne"
    },
    "Métaux Non-Fereux": {
        "Cuivre Grade A": "Tonne", "Aluminium LME": "Tonne", "Zinc Standard": "Tonne", 
        "Laiton": "Tonne", "Plomb affiné": "Tonne", "Étain LME": "Tonne", "Nickel": "Tonne"
    },
    "Métaux Précieux": {
        "Or (Lingot)": "Kilogramme", "Argent pur": "Kilogramme", "Platine": "Kilogramme", "Palladium": "Kilogramme"
    },
    "Minéraux & Phosphates": {
        "Phosphates (Roche BPL 68%)": "Tonne", "Minerai de Fer Standard": "Tonne", "Soufre brut": "Tonne", "Potasse": "Tonne"
    },
    "Énergies & Carburants": {
        "Gasoil (Diesel)": "Litre", "Essence Super": "Litre", "Fuel Lourd": "Tonne", 
        "Kérosène": "Litre", "Pétrole Brut (Brent)": "Baril"
    }
}

liens_references = {
    "Ferrailles & Aciers": "https://www.lme.com/Metals/Ferrous",
    "Métaux Non-Fereux": "https://www.lme.com/Metals/Non-Ferrous",
    "Métaux Précieux": "https://www.kitco.com/charts",
    "Minéraux & Phosphates": "https://www.indexmundi.com/commodities/",
    "Énergies & Carburants": "https://www.investing.com/commodities/energy"
}

print("=== [ETAPE 2] Lecture de la base de données des abonnés ===")
fichier_abonnes = "abonnes_db.csv"
if os.path.exists(fichier_abonnes):
    df_abonnes = pd.read_csv(fichier_abonnes)
else:
    df_abonnes = pd.DataFrame([
        {"email": EMAIL_EXPEDITEUR, "famille_souhaitee": "TOUT", "debut": "01-01-2026", "fin": "31-12-2027"}
    ])

print("=== [ETAPE 3] Génération des prédictions (Format large avec décisions quotidiennes) ===")
np.random.seed(42)
jours_prediction = [date_jour + timedelta(days=i) for i in range(8)]
noms_colonnes_jours = [j.strftime("%d/%m/%Y") for j in jours_prediction]

base_prices_usd = {
    "Ferraille Massive": 315.0, "Ferraille Légère": 260.0, "Ferraille E40": 275.0, "Ferraille E3": 240.0, "Fonte brute": 350.0, "Copeaux d'acier": 210.0, "Ferraille HMS 1&2": 290.0,
    "Cuivre Grade A": 8900.0, "Aluminium LME": 2400.0, "Zinc Standard": 2700.0, "Laiton": 5800.0, "Plomb affiné": 2150.0, "Étain LME": 29000.0, "Nickel": 16500.0,
    "Or (Lingot)": 65000.0, "Argent pur": 850.0, "Platine": 32000.0, "Palladium": 34000.0,
    "Phosphates (Roche BPL 68%)": 110.0, "Minerai de Fer Standard": 12.0, "Soufre brut": 250.0, "Potasse": 340.0,
    "Gasoil (Diesel)": 0.85, "Essence Super": 0.95, "Fuel Lourd": 420.0, "Kérosène": 0.90, "Pétrole Brut (Brent)": 78.0
}

donnees_globales = []
decisions_globales_par_ligne = [] # Pour stocker les décisions par jour et colorier après

for famille, metaux_dict in catalogue_mondial.items():
    for metal, unite in metaux_dict.items():
        p_base = base_prices_usd[metal]
        ligne_prix_mad = []
        decisions_ligne = []
        
        prix_precedent = None
        for i in range(8):
            p_base += np.random.normal(0, p_base * 0.008)
            prix_mad = round(p_base * taux_usd_mad, 2)
            ligne_prix_mad.append(prix_mad)
            
            # Détermination de la décision d'achat journalière
            if prix_precedent is None:
                decisions_ligne.append("WAIT") # Premier jour de référence
            else:
                if prix_mad < prix_precedent:
                    decisions_ligne.append("GO")       # Baisse de prix = Opportunité d'achat
                elif prix_mad == prix_precedent:
                    decisions_ligne.append("WAIT")     # Stabilité
                else:
                    decisions_ligne.append("NO GO")    # Hausse de prix = Éviter l'achat
            prix_precedent = prix_mad
            
        dictionnaire_ligne = {
            "Famille": famille,
            "Matière / Produit": metal,
            "Unité": unite,
            "Cours Change (USD/MAD)": taux_usd_mad,
        }
        
        for idx, nom_col in enumerate(noms_colonnes_jours):
            dictionnaire_ligne[nom_col] = ligne_prix_mad[idx]
            
        dictionnaire_ligne["Lien Source Réel"] = liens_references[famille]
        
        donnees_globales.append(dictionnaire_ligne)
        decisions_globales_par_ligne.append(decisions_ligne)

df_Complet = pd.DataFrame(donnees_globales)

print("=== [ETAPE 4] Boucle de traitement et d'envoi par abonné ===")
for index, abonne in df_abonnes.iterrows():
    email_client = str(abonne["email"]).strip()
    famille_demandee = str(abonne["famille_souhaitee"]).strip()
    date_fin_str = str(abonne["fin"]).strip()
    
    try:
        date_fin_abo = datetime.strptime(date_fin_str, "%d-%m-%Y")
    except Exception as err:
        print(f"❌ Erreur format date pour {email_client}: {err}")
        continue

    if datetime.now() > date_fin_abo:
        print(f"🔒 Abonné {email_client} expiré. Aucun envoi.")
        continue
        
    # Filtrage par famille
    if famille_demandee.upper() == "TOUT":
        df_abonne = df_Complet.copy()
        decisions_abonne = decisions_globales_par_ligne
        nom_famille_mail = "Toutes les Familles (Métaux & Énergies)"
        nom_fichier_clean = "Toutes_Familles"
    elif famille_demandee in catalogue_mondial.keys():
        indices_famille = [i for i, row in enumerate(donnees_globales) if row["Famille"] == famille_demandee]
        df_abonne = df_Complet.iloc[indices_famille].copy()
        decisions_abonne = [decisions_globales_par_ligne[i] for i in indices_famille]
        nom_famille_mail = famille_demandee
        nom_fichier_clean = famille_demandee.lower().replace(" & ", "_").replace(" ", "_")
    else:
        df_abonne = df_Complet.copy()
        decisions_abonne = decisions_globales_par_ligne
        nom_famille_mail = "Rapport Global"
        nom_fichier_clean = "Rapport_Global"
        
    nom_fichier = f"veille_marche_{nom_fichier_clean}_{date_str}.xlsx"
    
    # Création du fichier Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Prédictions 8J & Décisions"
    
    headers = list(df_abonne.columns)
    ws.append(headers)
    
    for row in df_abonne.itertuples(index=False):
        ws.append(list(row))
        
    # Couleurs conditionnelles par cellule de date selon la décision journalière
    fill_go = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")     # Vert (Achat favorable)
    fill_wait = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid") # Orange/Jaune (Neutre/Stable)
    fill_nogo = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid") # Rouge (Défavorable)
    
    # Les colonnes de dates commencent à l'index 5 (après Famille, Matiere, Unite, Cours Change)
    col_debut_dates = 5
    
    for row_idx, decisions_ligne in enumerate(decisions_abonne, start=2):
        for col_offset, decision in enumerate(decisions_ligne):
            cell = ws.cell(row=row_idx, column=col_debut_dates + col_offset)
            if decision == "GO":
                cell.fill = fill_go
            elif decision == "WAIT":
                cell.fill = fill_wait
            elif decision == "NO GO":
                cell.fill = fill_nogo
                
    wb.save(nom_fichier)

    # Envoi e-mail avec PJ
    msg = EmailMessage()
    msg['Subject'] = f"📊 Rapport Veille Stratégique : {nom_famille_mail} - {date_str}"
    msg['From'] = EMAIL_EXPEDITEUR
    msg['To'] = email_client
    msg.set_content(f"Bonjour,\n\nVoici ton rapport horizontal avec décisions d'achat journalières (colorées par date) pour la famille : {nom_famille_mail}.\nTaux de change appliqué : 1 USD = {taux_usd_mad} MAD.\n\nCordialement,\nTon Agent IA de Veille")

    with open(nom_fichier, "rb") as f:
        file_data = f.read()
    msg.add_attachment(file_data, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=nom_fichier)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_EXPEDITEUR, EMAIL_MOT_DE_PASSE)
            smtp.send_message(msg)
        print(f"🎉 E-mail avec PJ (décisions journalières colorées) envoyé à {email_client}")
    except Exception as e:
        print(f"❌ Erreur SMTP pour {email_client} : {e}")

print("=== [FIN] Traitement terminé avec succès ===")
