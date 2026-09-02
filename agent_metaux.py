import os
import smtplib
from email.message import EmailMessage
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import PatternFill

print("=== [ETAPE 1] Initialisation ===")
EMAIL_EXPEDITEUR = os.environ.get("MAIL_USER")
EMAIL_MOT_DE_PASSE = os.environ.get("MAIL_PASSWORD")

date_jour = datetime.now()
date_str = date_jour.strftime("%d-%m-%Y")
taux_usd_mad = 9.34  # Taux de change USD/MAD

# CATALOGUE RICHE AVEC DOUBLE DISPONIBILITÉ (LOCAL / ETRANGER POUR LA MEME REFERENCE)
catalogue_mondial = {
    "Ferrailles & Aciers": {
        "Ferraille HMS 1&2": {
            "unite": "Tonne", 
            "base_local_mad": 4050.0, 
            "base_etranger_usd": 403.0,  # 403 * 9.34 = ~3764.02 MAD (brut)
            "source": "LME Ferrous / Platts Scrap Index - https://www.lme.com/Metals/Ferrous"
        },
        "Ferraille Légère": {
            "unite": "Tonne", 
            "base_local_mad": 2600.0, 
            "base_etranger_usd": 275.0,
            "source": "Argus Media Ferrous Scrap - https://www.argusmedia.com"
        },
        "Fonte brute": {
            "unite": "Tonne", 
            "base_local_mad": 3800.0, 
            "base_etranger_usd": 350.0,
            "source": "Fastmarkets Pig Iron - https://www.fastmarkets.com"
        }
    },
    "Métaux Non-Fereux": {
        "Cuivre Grade A": {
            "unite": "Tonne", 
            "base_local_mad": 85000.0, 
            "base_etranger_usd": 8900.0,
            "source": "London Metal Exchange (LME) Cuivre - https://www.lme.com/Metals/Non-Ferrous"
        },
        "Aluminium LME": {
            "unite": "Tonne", 
            "base_local_mad": 24500.0, 
            "base_etranger_usd": 2400.0,
            "source": "LME Aluminium - https://www.lme.com/Metals/Non-Ferrous"
        }
    },
    "Métaux Précieux": {
        "Or (Lingot)": {
            "unite": "Kilogramme", 
            "base_local_mad": 620000.0, 
            "base_etranger_usd": 65000.0,
            "source": "Kitco Gold Index - https://www.kitco.com/charts"
        }
    },
    "Minéraux & Phosphates": {
        "Phosphates (Roche BPL 68%)": {
            "unite": "Tonne", 
            "base_local_mad": 1100.0, 
            "base_etranger_usd": 115.0,
            "source": "OCP / IndexMundi Phosphates - https://www.indexmundi.com/commodities/"
        }
    },
    "Énergies & Carburants": {
        "Gasoil (Diesel)": {
            "unite": "Litre", 
            "base_local_mad": 12.50, 
            "base_etranger_usd": 1.35,
            "source": "Ministère de la Transition Énergétique Maroc / Platts - https://www.investing.com/commodities/energy"
        },
        "Pétrole Brut (Brent)": {
            "unite": "Baril", 
            "base_local_mad": 750.0, 
            "base_etranger_usd": 78.0,
            "source": "Investing.com Brent - https://www.investing.com/commodities/brent-oil"
        }
    }
}

print("=== [ETAPE 2] Lecture de la base de données des abonnés ===")
fichier_abonnes = "abonnes_db.csv"
if os.path.exists(fichier_abonnes):
    df_abonnes = pd.read_csv(fichier_abonnes)
else:
    df_abonnes = pd.DataFrame([
        {"email": EMAIL_EXPEDITEUR, "famille_souhaitee": "TOUT", "format_souhaite": "excel", "debut": "01-01-2026", "fin": "31-12-2027"}
    ])

print("=== [ETAPE 3] Génération des simulations sur 8 jours (Comparatif Local vs Étranger) ===")
np.random.seed(42)
jours_prediction = [date_jour + timedelta(days=i) for i in range(8)]
noms_colonnes_jours = [j.strftime("%d/%m/%Y") for j in jours_prediction]

donnees_excel_globales = []
donnees_csv_globales = []
decisions_globales_excel = []

for famille, produits_dict in catalogue_mondial.items():
    for produit, info in produits_dict.items():
        unite = info["unite"]
        source = info["source"]
        
        b_local = info["base_local_mad"]
        b_etranger = info["base_etranger_usd"]
        
        prix_locaux = []
        prix_etrangers = []
        decisions_ligne = []
        
        dernier_choix = None
        
        for i in range(8):
            b_local += np.random.normal(0, b_local * 0.005)
            b_etranger += np.random.normal(0, b_etranger * 0.005)
            
            p_loc = round(b_local, 2)
            p_etr = round(b_etranger * taux_usd_mad, 2)  # Converti en MAD net
            
            prix_locaux.append(p_loc)
            prix_etrangers.append(p_etr)
            
            # Aide à la décision simple : si Local < Étranger, c'est GO Local, sinon NO GO
            if p_loc <= p_etr:
                decisions_ligne.append("GO LOCAL")
            else:
                decisions_ligne.append("GO ETRANGER")

        # 1. Structure pour Excel (Format Large : une colonne par jour)
        dict_excel = {
            "Famille": famille,
            "Matière / Produit": produit,
            "Unité": unite,
            "Marché Comparé": "LOCAL (MAD)",
        }
        for idx, col_j in enumerate(noms_colonnes_jours):
            dict_excel[col_j] = prix_locaux[idx]
        dict_excel["Source Référence"] = source
        donnees_excel_globales.append(dict_excel)

        dict_excel_etr = {
            "Famille": famille,
            "Matière / Produit": produit,
            "Unité": unite,
            "Marché Comparé": "ETRANGER (Converti MAD)",
        }
        for idx, col_j in enumerate(noms_colonnes_jours):
            dict_excel_etr[col_j] = prix_etrangers[idx]
        dict_excel_etr["Source Référence"] = source
        donnees_excel_globales.append(dict_excel_etr)
        
        decisions_globales_excel.append(decisions_ligne)
        decisions_globales_excel.append(decisions_ligne) # Double pour les 2 lignes

        # 2. Structure pour CSV / ERP (Format Long / Tidy : 1 ligne par jour et par marché)
        for idx, col_j in enumerate(noms_colonnes_jours):
            donnees_csv_globales.append({
                "Famille": famille,
                "Matiere": produit,
                "Unite": unite,
                "Marche": "Local",
                "Date": col_j,
                "Prix_MAD": prix_locaux[idx],
                "Source": source
            })
            donnees_csv_globales.append({
                "Famille": famille,
                "Matiere": produit,
                "Unite": unite,
                "Marche": "Etranger",
                "Date": col_j,
                "Prix_MAD": prix_etrangers[idx],
                "Source": source
            })

df_Complet_Excel = pd.DataFrame(donnees_excel_globales)
df_Complet_Csv = pd.DataFrame(donnees_csv_globales)

print("=== [ETAPE 4] Boucle d'envoi personnalisée par abonné ===")
for index, abonne in df_abonnes.iterrows():
    email_client = str(abonne["email"]).strip()
    famille_demandee = str(abonne["famille_souhaitee"]).strip()
    format_souhaite = str(abonne.get("format_souhaite", "excel")).strip().lower()
    
    if format_souhaite not in ["excel", "csv"]:
        format_souhaite = "excel"
        
    date_fin_str = str(abonne["fin"]).strip()
    try:
        date_fin_abo = datetime.strptime(date_fin_str, "%d-%m-%Y")
    except:
        date_fin_abo = datetime.now() + timedelta(days=365)

    if datetime.now() > date_fin_abo:
        print(f"🔒 Abonné {email_client} expiré. Aucun envoi.")
        continue
        
    # FILTRAGE PAR FAMILLE
    if famille_demandee.upper() == "TOUT":
        df_abonne_excel = df_Complet_Excel.copy()
        df_abonne_csv = df_Complet_Csv.copy()
        nom_fichier_clean = "Toutes_Familles"
    elif famille_demandee in catalogue_mondial.keys():
        df_abonne_excel = df_Complet_Excel[df_Complet_Excel["Famille"] == famille_demandee].copy()
        df_abonne_csv = df_Complet_Csv[df_Complet_Csv["Famille"] == famille_demandee].copy()
        nom_fichier_clean = famille_demandee.lower().replace(" & ", "_").replace(" ", "_")
    else:
        df_abonne_excel = df_Complet_Excel.copy()
        df_abonne_csv = df_Complet_Csv.copy()
        nom_fichier_clean = "Rapport_Global"
        
    # GÉNÉRATION FICHIER SELON FORMAT DEMANDÉ
    if format_souhaite == "csv":
        nom_fichier = f"veille_erp_{nom_fichier_clean}_{date_str}.csv"
        df_abonne_csv.to_csv(nom_fichier, index=False, encoding="utf-8-sig")
        sub_type = "csv"
    else:
        nom_fichier = f"veille_marche_{nom_fichier_clean}_{date_str}.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Comparatif Local vs Étranger"
        
        headers = list(df_abonne_excel.columns)
        ws.append(headers)
        for row in df_abonne_excel.itertuples(index=False):
            ws.append(list(row))
            
        # Coloration légère des cellules de prix
        fill_local = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid") # Vert tendre
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=2):
            marche_cell = ws.cell(row=row_idx, column=4).value
            if marche_cell and "LOCAL" in str(marche_cell):
                for col_c in range(5, ws.max_column):
                    ws.cell(row=row_idx, column=col_c).fill = fill_local
                    
        wb.save(nom_fichier)
        sub_type = "vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    # ENVOI E-MAIL
    try:
        msg = EmailMessage()
        msg['Subject'] = f"📊 Veille Stratégique & Comparatif Local/Étranger ({famille_demandee}) - {date_str}"
        msg['From'] = EMAIL_EXPEDITEUR
        msg['To'] = email_client
        msg.set_content(
            f"Bonjour,\n\n"
            f"Veuillez trouver ci-joint votre rapport de veille stratégique ({famille_demandee}) au format {format_souhaite.upper()}.\n"
            f"Les données intègrent désormais le comparatif côte à côte entre les cours locaux et internationaux (convertis en MAD au taux de {taux_usd_mad}), "
            f"ainsi que les liens de sources officielles.\n\n"
            f"Cordialement,\nVotre Agent IA de Veille Marchés"
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
            print(f"🧪 Simulation - Fichier prêt pour {email_client}")
            
    except Exception as e:
        print(f"❌ Erreur envoi {email_client} : {e}")

print("=== [FIN] Traitement terminé avec succès ===")
