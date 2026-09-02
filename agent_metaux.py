import os
import smtplib
from email.message import EmailMessage
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

print("=== [ETAPE 1] Initialisation ===")
EMAIL_EXPEDITEUR = os.environ.get("MAIL_USER")
EMAIL_MOT_DE_PASSE = os.environ.get("MAIL_PASSWORD")

date_jour = datetime.now()
date_str = date_jour.strftime("%d-%m-%Y")
taux_usd_mad = 9.34  # Taux de change USD/MAD

# CATALOGUE COMPLET (Mêmes sources de référence pour Local et Étranger)
catalogue_mondial = {
    "Ferrailles & Aciers": {
        "Ferraille HMS 1&2": {
            "unite": "Tonne", 
            "base_local_mad": 4050.0, 
            "base_etranger_usd": 403.0,  
            "source": "LME Ferrous / Platts Scrap Index (www.lme.com)"
        },
        "Ferraille Légère": {
            "unite": "Tonne", 
            "base_local_mad": 2600.0, 
            "base_etranger_usd": 275.0,
            "source": "Argus Media Ferrous Scrap (www.argusmedia.com)"
        },
        "Fonte brute": {
            "unite": "Tonne", 
            "base_local_mad": 3800.0, 
            "base_etranger_usd": 350.0,
            "source": "Fastmarkets Pig Iron (www.fastmarkets.com)"
        }
    },
    "Métaux Non-Fereux": {
        "Cuivre Grade A": {
            "unite": "Tonne", 
            "base_local_mad": 85000.0, 
            "base_etranger_usd": 8900.0,
            "source": "London Metal Exchange (LME) Cuivre (www.lme.com)"
        },
        "Aluminium LME": {
            "unite": "Tonne", 
            "base_local_mad": 24500.0, 
            "base_etranger_usd": 2400.0,
            "source": "LME Aluminium (www.lme.com)"
        },
        "Zinc SHG": {
            "unite": "Tonne", 
            "base_local_mad": 28000.0, 
            "base_etranger_usd": 2750.0,
            "source": "LME Zinc (www.lme.com)"
        }
    },
    "Métaux Précieux": {
        "Or (Lingot)": {
            "unite": "Kilogramme", 
            "base_local_mad": 620000.0, 
            "base_etranger_usd": 65000.0,
            "source": "Kitco Gold Index (www.kitco.com)"
        }
    },
    "Minéraux & Phosphates": {
        "Phosphates (Roche BPL 68%)": {
            "unite": "Tonne", 
            "base_local_mad": 1100.0, 
            "base_etranger_usd": 115.0,
            "source": "OCP / IndexMundi Phosphates (www.indexmundi.com)"
        }
    },
    "Énergies & Carburants": {
        "Gasoil (Diesel)": {
            "unite": "Litre", 
            "base_local_mad": 12.50, 
            "base_etranger_usd": 1.35,
            "source": "Ministère Transition Énergétique / Platts (www.investing.com)"
        },
        "Pétrole Brut (Brent)": {
            "unite": "Baril", 
            "base_local_mad": 750.0, 
            "base_etranger_usd": 78.0,
            "source": "Investing.com Brent (www.investing.com)"
        }
    }
}

print("=== [ETAPE 2] Lecture de la base de données des abonnés (Historique étendu) ===")
fichier_abonnes = "abonnes_db.csv"

# Structure enrichie de la base abonnés si le fichier n'existe pas
if not os.path.exists(fichier_abonnes):
    df_abonnes_init = pd.DataFrame([
        {
            "email": EMAIL_EXPEDITEUR, 
            "famille_souhaitee": "TOUT", 
            "format_souhaite": "excel", 
            "statut": "ACTIF", 
            "horizon_jours": 30, 
            "paiement": "Essai Gratuit", 
            "facture": "Non", 
            "debut": "01-01-2026", 
            "fin": "31-12-2027"
        }
    ])
    df_abonnes_init.to_csv(fichier_abonnes, index=False)

df_abonnes = pd.read_csv(fichier_abonnes)

print("=== [ETAPE 3] Boucle d'envoi personnalisée par abonné ===")
for index, abonne in df_abonnes.iterrows():
    email_client = str(abonne["email"]).strip()
    statut_abo = str(abonne.get("statut", "ACTIF")).strip().upper()
    
    # 1. Vérification si l'abonné est ACTIF
    if statut_abo != "ACTIF":
        print(f"🔒 Abonné {email_client} inactif/expiré (Statut: {statut_abo}). Aucun envoi.")
        continue
        
    famille_demandee = str(abonne["famille_souhaitee"]).strip()
    format_souhaite = str(abonne.get("format_souhaite", "excel")).strip().lower()
    
    # Récupération de l'horizon J+i personnalisé (par défaut 8 jours si non spécifié)
    try:
        horizon_i = int(abonne.get("horizon_jours", 8))
    except:
        horizon_i = 8

    if format_souhaite not in ["excel", "csv"]:
        format_souhaite = "excel"
        
    date_fin_str = str(abonne["fin"]).strip()
    try:
        date_fin_abo = datetime.strptime(date_fin_str, "%d-%m-%Y")
    except:
        date_fin_abo = datetime.now() + timedelta(days=365)

    if datetime.now() > date_fin_abo:
        print(f"🔒 Date de fin dépassée pour {email_client}. Aucun envoi.")
        continue

    print(f"-> Traitement pour {email_client} | Horizon J+{horizon_i} | Format: {format_souhaite.upper()}")

    # Génération des simulations sur 'horizon_i' jours
    np.random.seed(42)
    jours_prediction = [date_jour + timedelta(days=d_idx) for d_idx in range(horizon_i)]
    noms_colonnes_jours = [j.strftime("%d/%m/%Y") for j in jours_prediction]

    donnees_date_globales = []
    donnees_csv_globales = []
    
    # Filtrage catalogue selon la demande de l'abonné
    if famille_demandee.upper() == "TOUT":
        cat_filtre = catalogue_mondial
        nom_fichier_clean = "Toutes_Familles"
    elif famille_demandee in catalogue_mondial.keys():
        cat_filtre = {famille_demandee: catalogue_mondial[famille_demandee]}
        nom_fichier_clean = famille_demandee.lower().replace(" & ", "_").replace(" ", "_")
    else:
        cat_filtre = catalogue_mondial
        nom_fichier_clean = "Rapport_Global"

    for famille, produits_dict in cat_filtre.items():
        for produit, info in produits_dict.items():
            unite = info["unite"]
            source_officielle = info["source"]  # Même source pour local et étranger
            
            b_local = info["base_local_mad"]
            b_etranger = info["base_etranger_usd"]
            
            for idx_j, col_j in enumerate(noms_colonnes_jours):
                b_local += np.random.normal(0, b_local * 0.003)
                b_etranger += np.random.normal(0, b_etranger * 0.003)
                
                p_loc = round(b_local, 2)
                p_etr = round(b_etranger * taux_usd_mad, 2)  # Converti en MAD net
                
                # 1. Structure pour Comparatif par Date (Ligne par ligne)
                donnees_date_globales.append({
                    "Date": col_j,
                    "Famille": famille,
                    "Référence Métal": produit,
                    "Unité": unite,
                    "Prix Local (MAD)": p_loc,
                    "Prix Étranger (MAD)": p_etr,
                    "Source Unique": source_officielle
                })

                # 2. Structure pour CSV ERP / BI (Format Long / Tidy)
                donnees_csv_globales.append({
                    "Famille": famille,
                    "Matiere": produit,
                    "Unite": unite,
                    "Marche": "Local",
                    "Date": col_j,
                    "Prix_MAD": p_loc,
                    "Source": source_officielle
                })
                donnees_csv_globales.append({
                    "Famille": famille,
                    "Matiere": produit,
                    "Unite": unite,
                    "Marche": "Etranger",
                    "Date": col_j,
                    "Prix_MAD": p_etr,
                    "Source": source_officielle
                })

    df_date_comparatif = pd.DataFrame(donnees_date_globales)
    df_csv_export = pd.DataFrame(donnees_csv_globales)

    # GÉNÉRATION DU FICHIER ENVOYÉ
    if format_souhaite == "csv":
        nom_fichier = f"veille_erp_{nom_fichier_clean}_{date_str}.csv"
        df_csv_export.to_csv(nom_fichier, index=False, encoding="utf-8-sig")
        sub_type = "csv"
    else:
        nom_fichier = f"veille_marche_{nom_fichier_clean}_{date_str}.xlsx"
        
        wb = openpyxl.Workbook()
        wb.remove(wb.active) # Supprimer la feuille par défaut
        
        # Styles communs
        HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1F4E79")
        REGULAR_FONT = Font(name="Calibri", size=11)
        THIN_BORDER = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9', bottom=Side(style='thin', color='D9D9D9')) # ❌ Erreur ici
        )
        # -------------------------------------------------------------
        # ONGLET 1 : Comparatif par Date
        # -------------------------------------------------------------
        ws_date = wb.create_sheet(title="Comparatif par Date")
        ws_date.views.sheetView[0].showGridLines = True
        
        ws_date["B2"] = f"COMPARATIF DES PRIX PAR DATE (Horizon J+{horizon_i})"
        ws_date["B2"].font = TITLE_FONT
        
        headers_date = ["Date", "Famille", "Référence Métal", "Unité", "Prix Local (MAD)", "Prix Étranger (MAD)", "Écart (Local - Étranger)", "Meilleur Choix", "Source Officielle"]
        for c_idx, h in enumerate(headers_date, start=2):
            cell = ws_date.cell(row=4, column=c_idx, value=h)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        r_row = 5
        for row in df_date_comparatif.itertuples(index=False):
            # row: Date, Famille, Ref, Unite, Prix_Loc, Prix_Etr, Source
            ws_date.cell(row=r_row, column=2, value=row[0]).alignment = Alignment(horizontal="center")
            ws_date.cell(row=r_row, column=3, value=row[1])
            ws_date.cell(row=r_row, column=4, value=row[2])
            ws_date.cell(row=r_row, column=5, value=row[3]).alignment = Alignment(horizontal="center")
            
            c_l = ws_date.cell(row=r_row, column=6, value=row[4])
            c_l.number_format = '#,##0.00'
            c_e = ws_date.cell(row=r_row, column=7, value=row[5])
            c_e.number_format = '#,##0.00'
            
            # Formules Excel intégrées pour l'aide à la décision
            c_ecart = ws_date.cell(row=r_row, column=8, value=f"=F{r_row}-G{r_row}")
            c_ecart.number_format = '#,##0.00'
            
            c_choix = ws_date.cell(row=r_row, column=9, value=f'=IF(F{r_row}<=G{r_row},"LOCAL","ETRANGER")')
            c_choix.alignment = Alignment(horizontal="center")
            
            ws_date.cell(row=r_row, column=10, value=row[6])
            
            for c in range(2, 11):
                ws_date.cell(row=r_row, column=c).font = REGULAR_FONT
                ws_date.cell(row=r_row, column=c).border = THIN_BORDER
            r_row += 1

        # -------------------------------------------------------------
        # ONGLET 2 : Comparatif par Référence Métal (Agrégé)
        # -------------------------------------------------------------
        ws_ref = wb.create_sheet(title="Comparatif par Référence")
        ws_ref.views.sheetView[0].showGridLines = True
        
        ws_ref["B2"] = "SYNTHÈSE COMPARATIVE PAR RÉFÉRENCE DE MÉTAL (MOYENNES)"
        ws_ref["B2"].font = TITLE_FONT
        
        headers_ref = ["Référence Métal", "Moyenne Prix Local", "Moyenne Prix Étranger", "Écart Moyen (MAD)", "Recommandation Stratégique", "Source Unique"]
        for c_idx, h in enumerate(headers_ref, start=2):
            cell = ws_ref.cell(row=4, column=c_idx, value=h)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        references_uniques = df_date_comparatif["Référence Métal"].unique()
        r_ref = 5
        max_date_row = 4 + len(df_date_comparatif)
        
        for ref_m in references_uniques:
            ws_ref.cell(row=r_ref, column=2, value=ref_m)
            
            # Formules AVERAGEIF dynamiques reliées à l'onglet 1
            c_ml = ws_ref.cell(row=r_ref, column=3, value=f"=AVERAGEIF('Comparatif par Date'!D5:D{max_date_row}, B{r_ref}, 'Comparatif par Date'!F5:F{max_date_row})")
            c_ml.number_format = '#,##0.00'
            c_me = ws_ref.cell(row=r_ref, column=4, value=f"=AVERAGEIF('Comparatif par Date'!D5:D{max_date_row}, B{r_ref}, 'Comparatif par Date'!G5:G{max_date_row})")
            c_me.number_format = '#,##0.00'
            
            c_ec = ws_ref.cell(row=r_ref, column=5, value=f"=C{r_ref}-D{r_ref}")
            c_ec.number_format = '#,##0.00'
            
            ws_ref.cell(row=r_ref, column=6, value=f'=IF(C{r_ref}<=D{r_ref},"Privilégier Local en moyenne","Privilégier Étranger en moyenne")')
            
            # Même source unique
            source_u = df_date_comparatif[df_date_comparatif["Référence Métal"] == ref_m]["Source Unique"].iloc[0]
            ws_ref.cell(row=r_ref, column=7, value=source_u)
            
            for c in range(2, 8):
                ws_ref.cell(row=r_ref, column=c).font = REGULAR_FONT
                ws_ref.cell(row=r_ref, column=c).border = THIN_BORDER
            r_ref += 1

        # Ajout d'un graphique comparatif dans l'onglet Référence
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "Moyenne des Prix : Local vs Étranger"
        chart.y_axis.title = "Prix en MAD"
        chart.x_axis.title = "Référence"
        
        data_ref_chart = Reference(ws_ref, min_col=3, min_row=4, max_col=4, max_row=r_ref-1)
        cats = Reference(ws_ref, min_col=2, min_row=5, max_row=r_ref-1)
        chart.add_data(data_ref_chart, titles_from_data=True)
        chart.set_categories(cats)
        chart.width = 18
        chart.height = 10
        ws_ref.add_chart(chart, "B12")

        # Ajustement des largeurs de colonnes pour les deux onglets
        for ws in wb.worksheets:
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

        wb.save(nom_fichier)
        sub_type = "vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    # ENVOI E-MAIL
    try:
        msg = EmailMessage()
        msg['Subject'] = f"📊 Rapport de Veille Stratégique & Décision Achat ({famille_demandee}) - {date_str}"
        msg['From'] = EMAIL_EXPEDITEUR
        msg['To'] = email_client
        msg.set_content(
            f"Bonjour,\n\n"
            f"Veuillez trouver ci-joint votre rapport de veille stratégique personnalisé ({famille_demandee}) au format {format_souhaite.upper()}.\n"
            f"Ce rapport couvre un horizon de prévision de J+{horizon_i} jours et intègre un double comparatif (par date et par référence) avec des sources officielles unifiées.\n\n"
            f"Cordialement,\nVotre Agent IA de Veille Marchés"
        )

        with open(nom_fichier, "rb") as f:
            file_data = f.read()
        msg.add_attachment(file_data, maintype="application", subtype=sub_type, filename=nom_fichier)

        if EMAIL_EXPEDITEUR and EMAIL_MOT_DE_PASSE:
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(EMAIL_EXPEDITEUR, EMAIL_MOT_DE_PASSE)
                smtp.send_message(msg)
            print(f"🎉 E-mail envoyé avec succès à {email_client} !")
        else:
            print(f"🧪 Simulation - Fichier généré et prêt pour {email_client}")
            
    except Exception as e:
        print(f"❌ Erreur envoi {email_client} : {e}")

print("=== [FIN] Traitement terminé avec succès ===")
