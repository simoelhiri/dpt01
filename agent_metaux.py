import os
import smtplib
from email.message import EmailMessage
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import PatternFill

# Identifiants e-mail récupérés de GitHub Secrets
EMAIL_EXPEDITEUR = os.environ.get("MAIL_USER")
EMAIL_MOT_DE_PASSE = os.environ.get("MAIL_PASSWORD")

# Paramètres généraux
date_jour = datetime.now()
date_str = date_jour.strftime("%d-%m-%Y")
taux_usd_mad = 9.34  # Taux de change USD/MAD

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

# 2. GESTION DES ABONNÉS (Base de données avec dates de début et fin d'abonnement)
# Note: Mets ton propre e-mail ici pour tester la réception de la PJ !
base_abonnes = [
    {"email": EMAIL_EXPEDITEUR, "famille_souhaitee": "Ferrailles & Aciers", "debut": "01-01-2026", "fin": "31-12-2026"},
    {"email": "client.vip.achats@gmail.com", "famille_souhaitee": "TOUT", "debut": "01-08-2026", "fin": "01-09-2027"},
    {"email": "client.expire@gmail.com", "famille_souhaitee": "Métaux Non-Fereux", "debut": "01-01-2025", "fin": "01-01-2026"} # Abonnement expiré
]

# 3. SIMULATION DES PRÉDICTIONS SUR 8 JOURS (J à J+7)
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
            p_base += np.random.normal(0, p_base * 0.008)
            prix_usd = round(p_base, 2)
            prix_mad = round(prix_usd * taux_usd_mad, 2)
            
            if i == 0:
                tendance, conseil = "STABLE ➡️", "WAIT"
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
                "Lien_Source": f"https://www.marche-metaux.com/index/{metal.lower().replace(' ', '-')}"
            })

df_Complet = pd.DataFrame(historique_global)
historique_envois = []

# 4. TRAITEMENT DES ABONNÉS, FILTRAGE, EXCEL & ENVOI E-MAIL AVEC PJ
for abonne in base_abonnes:
    date_fin_abo = datetime.strptime(abonne["fin"], "%d-%m-%Y")
    email_client = abonne["email"]
    
    # VÉRIFICATION DE LA VALIDITÉ DE L'ABONNEMENT
    if datetime.now() <= date_fin_abo:
        famille_visee = abonne["famille_souhaitee"]
        
        if famille_visee == "TOUT":
            df_abonne = df_Complet.copy()
        else:
            df_abonne = df_Complet[df_Complet["Famille"] == famille_visee].copy()
            
        # Nom de fichier Excel dynamique avec la date du jour
        nom_fichier = f"veille_metaux_{famille_visee.replace(' & ', '_').replace(' ', '_')}_{date_str}.xlsx"
        
        # Génération du fichier Excel stylisé
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Prédictions 8J"
        
        headers = ["Date", "Famille", "Métal / Matière", "Prix (USD)", "Prix (MAD)", "Tendance (8J)", "Décision", "Lien Information"]
        ws.append(headers)
        
        for row in df_abonne.itertuples(index=False):
            ws.append(list(row))
            
        # Coloration conditionnelle (Vert = GO, Orange = WAIT, Rouge = NO GO)
        fill_go = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
        fill_wait = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
        fill_nogo = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
        
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
        
        # PRéPARATION DE L'E-MAIL AVEC PIÈCE JOINTE
        msg = EmailMessage()
        msg['Subject'] = f"📊 Rapport Veille Métaux ({famille_visee}) - {date_str}"
        msg['From'] = EMAIL_EXPEDITEUR
        msg['To'] = email_client
        msg.set_content(f"Bonjour,\n\nVoici ton rapport personnalisé de veille des métaux et d'aide à la décision d'achat pour la famille : {famille_visee}.\nTaux de change appliqué : 1 USD = {taux_usd_mad} MAD.\n\nCordialement,\nTon Agent IA de Veille")

        with open(nom_fichier, "rb") as f:
            file_data = f.read()
            file_name = f.name
        msg.add_attachment(file_data, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=file_name)

        # ENVOI EFFECTIF VIA SMTP GMAIL
        try:
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(EMAIL_EXPEDITEUR, EMAIL_MOT_DE_PASSE)
                smtp.send_message(msg)
            statut_envoi = "SUCCÈS (E-mail + PJ envoyés)"
        except Exception as e:
            statut_envoi = f"ERREUR : {e}"
            
        historique_envois.append({
            "Email": email_client,
            "Famille": famille_visee,
            "Fichier": nom_fichier,
            "Date_Heure": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "Statut": statut_envoi
        })
    else:
        # ABONNEMENT EXPIRÉ -> AUCUN ENVOI
        historique_envois.append({
            "Email": email_client,
            "Famille": abonne["famille_souhaitee"],
            "Fichier": "AUCUN",
            "Date_Heure": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "Statut": "BLOQUÉ (Abonnement Expiré)"
        })

# Export du fichier de log des envois
df_logs = pd.DataFrame(historique_envois)
df_logs.to_excel("historique_logs_envois.xlsx", index=False)
print("Traitement des abonnements et envois terminé avec succès !")
