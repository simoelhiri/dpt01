import pandas as pd
import numpy as np
from datetime import datetime

# 1. Base de données dynamique structurée par FAMILLES de métaux
catalogue_metaux = {
    "Ferrailles": ["Ferraille Massive", "Ferraille Légère", "Ferraille E40", "Ferraille E3", "Fonte brute"],
    "Métaux Non-Fereux": ["Cuivre Grade A", "Aluminium LME", "Zinc Standard", "Laiton"],
    "Matières Premières & Minéraux": ["Phosphates (Roche BPL 68%)", "Minerai de Fer Standard", "Nickel"]
}

# Simulation d'un historique dynamique (récupération de marché)
np.random.seed(42)
dates = pd.date_range(end=datetime.now(), periods=168, freq="H")
data = []

base_prices = {
    "Ferraille Massive": 315.0, "Ferraille Légère": 260.0, "Ferraille E40": 275.0, "Ferraille E3": 240.0, "Fonte brute": 350.0,
    "Cuivre Grade A": 8900.0, "Aluminium LME": 2400.0, "Zinc Standard": 2700.0, "Laiton": 5800.0,
    "Phosphates (Roche BPL 68%)": 110.0, "Minerai de Fer Standard": 12.0, "Nickel": 16500.0
}

for famille, metaux in catalogue_metaux.items():
    for metal in metaux:
        p = base_prices[metal]
        for d in dates:
            p += np.random.normal(0, p * 0.002)
            data.append({"Famille": famille, "Metal": metal, "Date": d, "Prix": round(p, 2)})

df = pd.DataFrame(data)

print("==========================================================")
print("     AGENT IA DYNAMIQUE : VEILLE & DÉCISION D'ACHAT         ")
print("==========================================================\n")

recommandations = []

# 2. Analyse dynamique par Famille et par Métal
for famille, metaux in catalogue_metaux.items():
    print(f"📁 [FAMILLE : {famille.upper()}]")
    for metal in metaux:
        df_m = df[(df["Famille"] == famille) & (df["Metal"] == metal)].sort_values("Date")
        dernier_prix = df_m.iloc[-1]["Prix"]
        moyenne_24h = df_m.tail(24)["Prix"].mean()
        
        tendance = "HAUSSIÈRE 📈" if dernier_prix > moyenne_24h else "BAISSIÈRE 📉"
        pente_recente = df_m.tail(48)["Prix"].diff().mean()
        prix_prevu_24h = round(dernier_prix + (pente_recente * 24), 2)
        
        if tendance == "BAISSIÈRE 📉":
            constat = f"Repli par rapport à la moyenne 24h ({moyenne_24h:.2f}$). Oportunité potentielle."
            conseil = "⏳ ATTENDRE"
        else:
            constat = f"Pression à la hausse constatée au-dessus de la moyenne."
            conseil = "🛒 ACHETER"

        recommandations.append({
            "Date_Analyse": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "Famille": famille,
            "Metal": metal,
            "Prix_Actuel": dernier_prix,
            "Prix_Prevu_J1": prix_prevu_24h,
            "Tendance": tendance,
            "Conseil_Achat": conseil
        })
        
        print(f"   • {metal} : {dernier_prix} $ (Prévu J+1: {prix_prevu_24h} $) | {tendance} -> {conseil}")
    print("-" * 58)

# 3. Export automatique vers un fichier Excel structuré
df_rec = pd.DataFrame(recommandations)
fichier_sortie = "veille_metaux_par_famille.xlsx"
df_rec.to_excel(fichier_sortie, index=False)
print(f"\n[SUCCÈS] Rapport complet mis à jour et exporté : {fichier_sortie}")

# 4. Simulation d'envoi d'alerte Email (Via Python / smtplib ou intégration Gmail gratuite)
print("\n[NOTIFICATION] Simulation d'envoi du rapport aux abonnés/acheteurs...")
print("-> E-mails cibles : direction-achats@entreprise.com, logistique@entreprise.com")
print("-> Statut : Rapport e-mail envoyé avec succès avec le fichier Excel en pièce jointe !")
