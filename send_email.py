import os
import smtplib
from email.message import EmailMessage

# Récupération sécurisée de tes secrets configurés sur GitHub
EMAIL_EXPEDITEUR = os.environ.get("MAIL_USER")
EMAIL_MOT_DE_PASSE = os.environ.get("MAIL_PASSWORD")
EMAIL_DESTINATAIRE = os.environ.get("MAIL_USER") # S'envoie à toi-même ou modifie l'adresse

msg = EmailMessage()
msg['Subject'] = "📊 Rapport Automatique : Veille & Décision d'Achat Métaux"
msg['From'] = EMAIL_EXPEDITEUR
msg['To'] = EMAIL_DESTINATAIRE
msg.set_content("Bonjour,\n\nVoici ton rapport automatisé de veille des métaux et d'aide à la décision d'achat du jour.\n\nCordialement,\nTon Agent IA")

# Joindre le fichier Excel généré par le premier script
fichier_joint = "veille_metaux_par_famille.xlsx"
if os.path.exists(fichier_joint):
    with open(fichier_joint, "rb") as f:
        file_data = f.read()
        file_name = f.name
    msg.add_attachment(file_data, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=file_name)

# Connexion au serveur sécurisé de Gmail et envoi
try:
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_EXPEDITEUR, EMAIL_MOT_DE_PASSE)
        smtp.send_message(msg)
    print("E-mail envoyé avec succès !")
except Exception as e:
    print(f"Erreur lors de l'envoi de l'e-mail : {e}")
