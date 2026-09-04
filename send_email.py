import os
import smtplib
from email.message import EmailMessage
import pandas as pd
from datetime import datetime

print("=== [ETAPE 3] Lancement de l'Envoi des E-mails Personnalisés ===")
EMAIL_EXPEDITEUR = os.environ.get("MAIL_USER")
EMAIL_MOT_DE_PASSE = os.environ.get("MAIL_PASSWORD")

fichier_abonnes = "abonnes_db.csv"
if not os.path.exists(fichier_abonnes):
    print("Erreur: abonnes_db.csv introuvable.")
    exit()

df_abonnes = pd.read_csv(fichier_abonnes)
date_str = datetime.now().strftime("%d-%m-%Y")

for index, abonne in df_abonnes.iterrows():
    email_client = str(abonne["email"]).strip()
    statut_abo = str(abonne.get("statut", "ACTIF")).strip().upper()
    
    if statut_abo != "ACTIF":
        continue
        
    # Personnalisation intelligente demandée
    societe = str(abonne.get("nom_societe", "Partenaire")).strip()
    civilite = str(abonne.get("civilite", "")).strip()
    prenom_nom = str(abonne.get("prenom_nom", "").strip())
    
    if prenom_nom and civilite:
        formule_politesse = f"Bonjour {civilite} {prenom_nom},"
    else:
        formule_politesse = f"Bonjour Société {societe},"

    famille_demandee = str(abonne["famille_souhaitee"]).strip()
    format_souhaite = str(abonne.get("format_souhaite", "excel")).strip().lower()
    horizon_i = abonne.get("horizon_jours", 30)

    # Nom du fichier correspondant
    if famille_demandee.upper() == "TOUT":
        nom_f_clean = "Toutes_Familles"
    else:
        nom_f_clean = famille_demandee.lower().replace(" & ", "_").replace(" ", "_")

    if format_souhaite == "csv":
        nom_fichier = f"veille_erp_{nom_f_clean}_{date_str}.csv"
        sub_type = "csv"
    else:
        nom_fichier = f"veille_marche_{nom_f_clean}_{date_str}.xlsx"
        sub_type = "vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    # Construction du mail avec Disclaimer Pro
    try:
        msg = EmailMessage()
        msg['Subject'] = f"📊 [VEILLE STRATÉGIQUE] Cours & Prévisions Métaux ({famille_demandee}) - {date_str}"
        msg['From'] = EMAIL_EXPEDITEUR
        msg['To'] = email_client
        
        corps_texte = f"""{formule_politesse}

Veuillez trouver ci-joint votre rapport de veille stratégique personnalisé pour la famille [{famille_demandee}] couvrant un horizon de J+{horizon_i} jours.

🔍 TRANSPARENCE DE NOTRE MODÈLE :
Pour vous garantir une aide à la décision irréprochable, notre outil distingue désormais clairement :
1. Le "Prix Réel du Marché (Spot J0)" et le "Prix Prévis Fixe" (cours court terme validés).
2. L'"Estimation Calculée selon notre Modèle Propriétaire" (combinant indices de fret, coûts énergétiques et indices de tension géopolitique type détroit d'Hormoz).

⚖️ AVERTISSEMENT LÉGAL (DISCLAIMER) :
Les données portant la mention "Estimation Calculée" sont générées par modélisation mathématique prospective à des fins d'aide à la budgétisation. Elles ne constituent en aucun cas un engagement de prix ferme ou un conseil en investissement direct. Seuls les cours spot constatés engagent les transactions immédiates.

Restant à votre entière disposition pour tout arbitrage stratégique.

Cordialement,
L'Équipe Intelligence Marchés & Veille Industrielle
"""
        msg.set_content(corps_texte)

        if os.path.exists(nom_fichier):
            with open(nom_fichier, "rb") as f:
                file_data = f.read()
            msg.add_attachment(file_data, maintype="application", subtype=sub_type, filename=nom_fichier)

        if EMAIL_EXPEDITEUR and EMAIL_MOT_DE_PASSE:
            import smtplib
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(EMAIL_EXPEDITEUR, EMAIL_MOT_DE_PASSE)
                smtp.send_message(msg)
            print(f"🎉 E-mail envoyé avec succès à {email_client} ({societe}) !")
        else:
            print(f"🧪 Mode Simulation - E-mail prêt et structuré pour {email_client} ({societe})")

    except Exception as e:
        print(f"❌ Erreur lors de l'envoi pour {email_client} : {e}")

print("=== [FIN] Processus d'envoi terminé avec succès ===")
