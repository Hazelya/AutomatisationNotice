import shutil
import zipfile

import streamlit as st
import pandas as pd
import os
import datetime as dt
from jinja2 import Environment, FileSystemLoader
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx import Document

# === Fonctions utilitaires ===
def date_now():
    mois = ['', 'janvier', 'février', 'mars', 'avril', 'mai', 'juin',
            'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre']
    today = dt.date.today()
    return f"{today.day} {mois[today.month]} {today.year}"

def format_nombre(nombre):
    return f"{nombre:,.2f}".replace(',', ' ').replace('.', ',')


# === Interface Streamlit ===
st.title("Générateur de notices d'appel de fonds")

uploaded_file = st.file_uploader("Fichier Excel de données", type=["xlsx"])
texte_fond_couvrir = st.text_area("Texte pour couvrir l'appel")
texte_fond_finance = st.text_area("Texte pour financer le nouvel appel")

numero_call = st.text_input("Numéro de l'appel", value="9")
nom_fond = st.text_input("Nom du fonds", value="FPCI ÉPOPÉE Xplore II")
pays = st.text_input("Pays", value="France")

if st.button("Générer les notices"):
    if uploaded_file:
        # Sauvegarde temporaire
        os.makedirs("ressources", exist_ok=True)
        chemin_fichier = "ressources/Base-data-test-fund-exercice.xlsx"
        with open(chemin_fichier, "wb") as f:
            f.write(uploaded_file.getbuffer())

        try:
            df = pd.read_excel(chemin_fichier, sheet_name='SOUSCRIPTEURS', header=18)
            df_nettoye = df[df['SOUSCRIPTEUR'].notna()]
            df_nettoye = df_nettoye[~df_nettoye['SOUSCRIPTEUR'].str.startswith('TOTAL', na=False)]
            df_nettoye = df_nettoye.reset_index(drop=True)

            df_CALL = pd.read_excel(chemin_fichier, sheet_name='SOUSCRIPTEURS', header=3)
            call = 'CALL #' + numero_call
            montant_total = df[call][df.shape[0]-6]
            date_call = df_CALL.loc[df_CALL['Nominal'] == call, 'Date'].iloc[0]
            pourcentage_call = df_CALL.loc[df_CALL['Nominal'] == call, df_CALL.columns[2]].iloc[0]

            dir = 'ressources'
            env = Environment(loader=FileSystemLoader(dir))
            template = env.get_template('model_notice_final.html')

            os.makedirs('Output', exist_ok=True)
            os.makedirs('Output_HTML', exist_ok=True)

            for i in range(df_nettoye.shape[0]):
                total_avant_call = df_nettoye['TOTAL APPELE'][i] - df_nettoye[call][i]
                pourcentage_avant_call = (total_avant_call / df_nettoye['ENGAGEMENT'][i]) * 100

                doc = Document()

                # Adresse
                doc.add_paragraph(f"{df_nettoye['souscripteur']}")
                if df_nettoye['pm_pp'].lower() == "pm":
                    doc.add_paragraph(f"{df_nettoye['representant']} (Représentant légal)")
                doc.add_paragraph(df_nettoye['adresse'])
                doc.add_paragraph(f"{df_nettoye['code_postal']} {df_nettoye['ville']}")
                doc.add_paragraph(df_nettoye['pays'])
                doc.add_paragraph("")

                # Date et lieu
                para = doc.add_paragraph(f"{df_nettoye['ville']}, le {df_nettoye['date']}")
                para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                doc.add_paragraph("")

                # Objet
                doc.add_paragraph(
                    f"Objet : {df_nettoye['nom_fond']} – Appel de Fonds N°{df_nettoye['numero_call']} – Date valeur : {df_nettoye['date_call']}",
                    style='Intense Quote')
                doc.add_paragraph("")

                # Corps de la lettre
                doc.add_paragraph("Cher Investisseur,")
                doc.add_paragraph(
                    f"\nDans le cadre de votre souscription dans le {df_nettoye['nom_fond']}, nous vous informons de l’Appel de Fonds N°{df_nettoye['numero_call']}, en date valeur du {df_nettoye['date_call']}, pour un montant total de {df_nettoye['montant_total']} € (soit {df_nettoye['pourcentage_call']} % du Montant Total des Souscriptions), soit pour votre quote-part un montant total de {df_nettoye['montant_a_liberer']} €.\n")

                doc.add_paragraph(
                    f"L’Appel de Fonds précédent avait porté votre engagement appelé à {df_nettoye['pourcentage_avant_call']} % du montant total auquel vous aviez souscrit. {df_nettoye['texte_fond_couvrir']} A ce jour, les montants libérés lors de la souscription initiale et des précédents appels ont été investis, dépensés ou engagés à 100 %.\n")

                doc.add_paragraph(
                    f"Ainsi, conformément à l’article 9.2.2 du Règlement du Fonds, nous vous remercions de bien vouloir procéder au versement de la somme mentionnée selon les modalités indiquées dans la page suivante.\n")

                doc.add_paragraph(f"{df_nettoye['texte_fond_finance']}\n")

                doc.add_paragraph(
                    "Nous vous remercions par avance et vous prions de croire, Cher Investisseur, en l’expression de notre considération distinguée.\n")
                doc.add_paragraph("L’équipe Middle Office\nÉpopée Gestion")

                # Page suivante
                doc.add_page_break()

                # Détails financiers
                doc.add_heading(f"{df_nettoye['nom_fond']} – Appel de Fonds N°{df_nettoye['numero_call']}", level=1)
                doc.add_paragraph("Détail de l’opération :")
                doc.add_paragraph(f"% de l’Appel de Fonds : {df_nettoye['pourcentage_call']} %")
                doc.add_paragraph(f"Montant à libérer : {df_nettoye['montant_a_liberer']} €")
                doc.add_paragraph(f"Date d’opération et de valeur, au plus tard le : {df_nettoye['date_call']}")
                doc.add_paragraph("")

                doc.add_paragraph("Situation après cette opération :")
                doc.add_paragraph(f"Montant de l’engagement initial : {df_nettoye['montant_engagement_initial']} €")
                doc.add_paragraph(
                    f"Nombre de Parts {df_nettoye['categorie_part']} souscrites : {df_nettoye['nombre_parts_souscrites']}")
                doc.add_paragraph(f"Capital appelé à date* : {df_nettoye['total_appele']} €")
                doc.add_paragraph(f"% de l’engagement appelé à date* : {df_nettoye['pourcent_liberation']} %")
                doc.add_paragraph(f"Montant restant à appeler* : {df_nettoye['residuel']} €")
                doc.add_paragraph("*Incluant l’Appel de Fonds en cours")
                doc.add_paragraph("")

                doc.add_paragraph(
                    "Veuillez procéder au paiement en €, net de tous frais bancaires, par virement bancaire sur le compte, selon les instructions suivantes :")

                doc.add_paragraph("Titulaire du compte : XPLORE II")
                doc.add_paragraph("Domiciliation : CACEIS BANK FRANCE")
                doc.add_paragraph("IBAN : FR76 1812 9000 1000 5002 0849 481")
                doc.add_paragraph("BIC (ADRESSE SWIFT) : ISAEFRPPXXX")
                doc.add_paragraph(f"Libellé du virement : {df_nettoye['libelle_virement']}")
                doc.add_paragraph("Frais pour le bénéficiaire : N/A (SHA)")
                doc.add_paragraph("")

                doc.add_paragraph("L’équipe Middle Office\nÉpopée Gestion")

                # Enregistrement du fichier
                docx_path = f"Output/{df_nettoye['souscripteur']}.docx"
                doc.save(docx_path)

            # Zip tous les fichiers
            shutil.make_archive("notices", "zip", "Output")
            with open("notices.zip", "rb") as f:
                st.download_button("Télécharger les notices générées", f, "notices.zip")

        except Exception as e:
            st.error(f"Une erreur est survenue : {e}")
    else:
        st.warning("Merci de déposer un fichier Excel avant de lancer la génération.")
