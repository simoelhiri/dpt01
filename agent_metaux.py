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

if not EMAIL_EXPEDITEUR or not EMAIL_MOT_DE_PASSE:
    print("⚠️ ATTENTION : Les variables d'environnement MAIL_USER ou MAIL_PASSWORD ne sont pas définies !")

date_jour = datetime.now()
date_str = date_jour.strftime("%d-%m-%Y")
taux_usd_mad = 9.34  # Taux de change USD/MAD

# CATALOGUE MONDIAL AVEC MARCHÉ (Local / Étranger) ET UNITÉS
catalogue_mondial = {
    "Ferrailles & Aciers": {
        "Ferraille Massive (Local)": {"unite": "Tonne", "marche": "Local", "base_mad": 3150.0},
        "Ferraille Légère (Local)": {"unite": "Tonne", "marche": "Local", "base_mad": 2600.0},
        "Ferraille HMS 1&2 (Import)": {"unite": "Tonne", "marche": "Étranger", "base_usd": 290.0},
        "Fonte brute (Import)": {"unite": "Tonne", "marche": "Étranger", "base_usd": 350.0}
    },
    "Métaux Non-Fereux": {
        "Cuivre Grade A (Import)": {"unite": "Tonne", "marche": "Étranger", "base_usd": 8900.0},
        "Aluminium LME (Import)": {"unite": "Tonne", "marche": "Étranger", "base_usd": 2400.0},
        "Laiton (Local)": {"unite": "Tonne", "marche": "Local", "base_mad": 54000.0}
    },
    "Métaux Précieux": {
        "Or (Lingot) (International)": {"unite": "Kilogramme", "marche": "Étranger", "base_usd": 65000.0},
        "Argent pur (International)": {"unite": "Kilogramme", "marche": "Étranger", "base_usd": 850.0}
    },
    "Minéraux & Phosphates": {
        "Phosphates (Roche BPL 68%) (Maroc - OCP)": {"unite": "Tonne", "marche": "Local", "base_mad": 1100.0},
        "Minerai de Fer Standard (Import)": {"unite": "Tonne", "marche": "Étranger", "base_usd": 120.0}
    },
    "Énergies & Carburants": {
        "Gasoil (Diesel) (Pompe Maroc)": {"unite": "Litre", "marche": "Local", "base_mad": 12.50},
        "Essence Super (Pompe Maroc)": {"unite": "Litre", "marche": "Local", "base_mad": 14.10},
        "Fuel Lourd (Industriel Maroc)": {"unite": "Tonne", "marche": "Local", "base_mad": 5800.0},
        "Pétrole Brut (Brent) (Global)": {"unite": "Baril", "marche": "Étranger", "base_usd": 78.0}
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
        {"email": "test@example.com", "famille_souhaitee": "TOUT", "format_souhaite": "excel", "debut": "01-01-2026", "fin": "31-12-2027"}
    ])

print("=== [ETAPE 3] Génération des données et prix (MAD local / USD converti) ===")
np.random.seed(42)
jours_prediction = [date_jour + timedelta(days=i) for i in range(8)]
noms_colonnes_jours = [j.strftime("%d/%m/%Y") for j in jours_prediction]

donnees_globales = []
decisions_globales_par_ligne = []

for famille, produits_dict in catalogue_mondial.items():
    for produit, info in produits_dict.items():
        marche = info["marche"]
        unite = info["unite"]
        
        p_base = info["base_mad"] if marche == "Local" else info["base_usd"]
            
        ligne_prix = []
        decisions_ligne = []
        prix_precedent = None
        
        for i in range(8):
            p_base += np.random.normal(0, p_base * 0.006)
            prix_final = round(p_base if marche == "Local" else p_base * taux_usd_mad, 2)
            ligne_prix.append(prix_final)
            
            if prix_precedent is None:
                decisions_ligne.append("WAIT")
            else:
                if prix_final < prix_precedent:
                    decisions_ligne.append("GO")
                elif prix_final == prix_precedent:
                    decisions_ligne.append("WAIT")
                else:
                    decisions_ligne.append("NO GO")
            prix_precedent = prix_final
            
        dictionnaire_ligne = {
            "Famille": famille,
            "Matière / Produit": produit,
            "Marché": marche,
            "Unité": unite,
            "Cours Change (USD/MAD)": taux_usd_mad if marche == "Étranger" else "N/A (Prix Local MAD)",
        }
        
        for idx, nom_col in enumerate(noms_colonnes_jours):
            dictionnaire_ligne[nom_col] = ligne_prix[idx]
            
        dictionnaire_ligne["Lien Source Réel"] = liens_references[famille]
        
        donnees_globales.append(dictionnaire_ligne)
        decisions_globales_par_ligne.append(decisions_ligne)

df_Complet = pd.DataFrame(donnees_globales)

print("=== [ETAPE 4] Boucle d'envoi personnalisée par abonné ===")
for index, abonne in df_abonnes.iterrows():
    email_client = str(abonne["email"]).strip()
    famille_demandee = str(abonne["famille_souhaitee"]).strip()
    
    # Gestion sécurisée si la colonne format_souhaite n'est pas encore dans le CSV
    format_souhaite = str(abonne.get("format_souhaite", "excel")).strip().lower()
    if format_souhaite not in ["excel", "csv"]:
        format_souhaite = "excel"
        
    date_fin_str = str(abonne["fin"]).strip()
    
    try:
        date_fin_abo = datetime.strptime(date_fin_str, "%d-%m-%Y")
    except Exception as err:
        print(f"⚠️ Date de fin invalide pour {email_client} ({date_fin_str}), abonnement considéré actif par défaut.")
        date_fin_abo = datetime.now() + timedelta(days=365)

    if datetime.now() > date_fin_abo:
        print(f"🔒 Abonné {email_client} expiré (Fin : {date_fin_str}). Aucun envoi.")
        continue
        
    # FILTRAGE PAR FAMILLE
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
        print(f"⚠️ Famille '{famille_demandee}' inconnue pour {email_client}. Rapport global par défaut.")
        df_abonne = df_Complet.copy()
        decisions_abonne = decisions_globales_par_ligne
        nom_famille_mail = "Rapport Global"
        nom_fichier_clean = "Rapport_Global"
        
    # GÉNÉRATION DU FICHIER (CSV OU EXCEL)
    if format_souhaite == "csv":
        nom_fichier = f"veille_erp_{nom_fichier_clean}_{date_str}.csv"
        df_abonne.to_csv(nom_fichier, index=False, encoding="utf-8-sig")
        sub_type = "csv"
    else:
        nom_fichier = f"veille_marche_{nom_fichier_clean}_{date_str}.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Prédictions & Décisions"
        
        headers = list(df_abonne.columns)
        ws.append(headers)
        for row in df_abonne.itertuples(index=False):
            ws.append(list(row))
            
        fill_go = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
        fill_wait = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
        fill_nogo = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
        
        col_debut_dates = 6
        for row_idx, decisions_ligne in enumerate(decisions_abonne, start=2):
            for col_offset, decision in enumerate(decisions_ligne):
                cell = ws.cell(row=row_idx, column=col_debut_dates + col_offset)
                if decision == "GO": cell.fill = fill_go
                elif decision == "WAIT": cell.fill = fill_wait
                elif decision == "NO GO": cell.fill = fill_nogo
        wb.save(nom_fichier)
        sub_type = "vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    # ENVOI E-MAIL
    try:
        msg = EmailMessage()
        msg['Subject'] = f"📊 Veille Stratégique ({famille_demandee}) - {date_str}"
        msg['From'] = EMAIL_EXPEDITEUR
        msg['To'] = email_client
        msg.set_content(
            f"Bonjour,\n\n"
            f"Voici ton rapport personnalisé de veille ({nom_famille_mail}) au format {format_souhaite.upper()}.\n"
            f"Les prix locaux sont affichés en MAD et les importations en USD convertis.\n\n"
            f"Cordialement,\nTon Agent IA de Veille"
        )

        with open(nom_fichier, "rb") as f:
            file_data = f.read()
        msg.add_attachment(file_data, maintype="application", subtype=sub_type, filename=nom_fichier)

        if EMAIL_EXPEDITEUR and EMAIL_MOT_DE_PASSE:
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(EMAIL_EXPEDITEUR, EMAIL_MOT_DE_PASSE)
                smtp.send_message(msg)
            print(f"🎉 E-mail envoyé avec succès à {email_client} (Format : {format_souhaite.upper()}) !")
        else:
            print(f"🧪 Mode Simulation (Pas de mot de passe email configuré) - Fichier prêt pour {email_client}")
            
    except Exception as e:
        print(f"❌ Erreur lors du traitement pour {email_client} : {e}")

print("=== [FIN] Traitement terminé ===" )
