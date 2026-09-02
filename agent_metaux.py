import os
import smtplib
from email.message import EmailMessage
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import PatternFill

# Récupération sécurisée des accès e-mail
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

# 2. LECTURE DE LA BASE DE DONNÉES DES ABONNÉS
fichier_abonnes = "abonnes_db.csv"
if os.path.exists(fichier_abonnes):
    df_abonnes = pd.read_csv(fichier_abonnes)
else:
    df_abonnes = pd.DataFrame([
        {"email": EMAIL_EXPEDITEUR, "famille_souhaitee": "TOUT", "debut": "01-01-2026", "fin": "31-12-2027"}
    ])

# 3. GÉNÉRATION GLOBALE DES PRÉDICTIONS (8 JOURS)
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

# 4. BOUCLE EXPLICITE D'INTERSECTION ET DE TRAITEMENT PAR ABONNÉ
print("=== DÉBUT DU TRAITEMENT DES ABONNÉS ===")
for index, abonne in df_abonnes.iterrows():
    email_client = str(abonne["email"]).strip()
    famille_demandee = str(abonne["famille_souhaitee"]).strip()
    date_fin_abo = datetime.strptime(str(abonne["fin"]).strip(), "%d-%m-%Y")
    
    print(f"Traitement pour : {email_client} | Famille demandée : {famille_demandee}")
    
    # Vérification de l'abonnement
    if datetime.now() > date_fin_abo:
        print(f"-> ABONNEMENT EXPIRÉ pour {email_client}")
        historique_envois.append({
            "Email": email_client, "Famille": famille_demandee, "Fichier": "AUCUN",
            "Date_Heure": datetime.now().strftime("%d/%m/%Y %H:%M:%S"), "Statut": "BLOQUÉ (Expiré)"
        })
        continue # Passe à l'abonné suivant
        
    # Intersection et filtrage selon la famille demandée
    if famille_demandee.upper() == "TOUT":
        df_abonne = df_Complet.copy()
        nom_famille_mail = "Toutes les Familles de Métaux"
        nom_fichier_clean = "Toutes_Familles"
    elif famille_demandee in catalogue_mondial.keys():
        df_abonne = df_Complet[df_Complet["Famille"] == famille_demandee].copy()
        nom_famille_mail = famille_demandee
        nom_fichier_clean = famille_demandee.lower().replace(" & ", "_").replace(" ", "_").replace("(", "").replace(")", "")
    else:
        # Si la famille demandée ne correspond exactement à rien, on prend tout par défaut
        df_abonne = df_Complet.copy()
        nom_famille_mail = "Rapport Global"
        nom_fichier_clean = "Rapport_Global"
        
    # Création du fichier Excel unique pour cet abonné
    nom_fichier = f"veille_metaux_{nom_fichier_clean}_{date_str}.xlsx"
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Prédictions 8J"
    
    headers = ["Date", "Famille", "Métal / Matière", "Prix (USD)", "Prix (MAD)", "Tendance (8J)", "Décision", "Lien Information"]
    ws.append(headers)
    
    for row in df_abonne.itertuples(index=False):
        ws.append(list(row))
        
    # Application des couleurs conditionnelles (Vert = GO, Orange = WAIT, Rouge = NO GO)
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
    
    # Préparation de l'e-mail avec l'objet dynamique contenant la vraie famille
    msg = EmailMessage()
    msg['Subject'] = f"📊 Rapport Veille : {nom_famille_mail} - {date_str}"
    msg['From'] = EMAIL_EXPEDITEUR
    msg['To'] = email_client
    msg.set_content(f"Bonjour,\n\nVoici ton rapport personnalisé de veille des métaux et d'aide à la décision pour la famille : {nom_famille_mail}.\nTaux de change appliqué : 1 USD = {taux_usd_mad} MAD.\n\nCordialement,\nTon Agent IA de Veille")

    # Attachement sécurisé du fichier Excel généré
    with open(nom_fichier, "rb") as f:
        file_data = f.read()
    msg.add_attachment(file_data, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=nom_fichier)

    # Envoi SMTP sécurisé
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_EXPEDITEUR, EMAIL_MOT_DE_PASSE)
            smtp.send_message(msg)
        print(f"-> E-MAIL AVEC PJ ENVOYÉ avec succès à {email_client}")
        statut_envoi = "SUCCÈS (E-mail + PJ)"
    except Exception as e:
        print(f"-> ERREUR lors de l'envoi à {email_client}: {e}")
        statut_envoi = f"ERREUR : {e}"
        
    historique_envois.append({
        "Email": email_client,
        "Famille": nom_famille_mail,
        "Fichier": nom_fichier,
        "Date_Heure": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "Statut": statut_envoi
    })

# Sauvegarde du fichier de log général
df_logs = pd.DataFrame(historique_envois)
df_logs.to_excel("historique_logs_envois.xlsx", index=False)
print("=== FIN DU TRAITEMENT DE TOUS LES ABONNÉS ===")
