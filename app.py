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

                for i in range(len(df_nettoye)):
                    doc = Document()

                    # ======== Haut de page - Coordonnées du souscripteur ========
                    doc.add_paragraph(df_nettoye["SOUSCRIPTEUR"][i])
                    if df_nettoye["TYPE"][i].lower() == "pm":
                        doc.add_paragraph(f"{df_nettoye['Représentant'][i]} (Représentant légal)")
                    doc.add_paragraph(df_nettoye["ADRESSE"][i])
                    doc.add_paragraph(f"{round(df_nettoye['CP'][i])} {df_nettoye['VILLE'][i]}")
                    doc.add_paragraph(pays)

                    # ======== Date et ville ========
                    para = doc.add_paragraph(f"{df_nettoye['VILLE'][i]}, le {date_now()}")
                    para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                    # ======== Objet ========
                    doc.add_paragraph("")
                    doc.add_paragraph(
                        f"Objet : {nom_fond} – Appel de Fonds N°{numero_call} – Date valeur : {date_call.strftime('%d/%m/%Y')}",
                        style="Intense Quote")

                    # ======== Corps de la lettre ========
                    doc.add_paragraph("")
                    doc.add_paragraph("Cher Investisseur,")
                    doc.add_paragraph(f"""
                    Dans le cadre de votre souscription dans le {nom_fond}, nous vous informons de l’Appel de Fonds N°{numero_call}, en date valeur du {date_call.strftime('%d/%m/%Y')}, pour un montant total de {format_nombre(montant_total)} € (soit {pourcentage_call * 100:.2f} % du Montant Total des Souscriptions), soit pour votre quote-part un montant total de {format_nombre(df_nettoye[call][i])} €.
                    """)
                    doc.add_paragraph(f"""
                    L’Appel de Fonds précédent avait porté votre engagement appelé à {format_nombre(pourcentage_avant_call)} % du montant total auquel vous aviez souscrit. {texte_fond_couvrir} A ce jour, les montants libérés lors de la souscription initiale et des précédents appels ont été investis, dépensés ou engagés à 100 %.
                    """)
                    doc.add_paragraph(
                        "Ainsi, conformément à l’article 9.2.2 du Règlement du Fonds, nous vous remercions de bien vouloir procéder au versement de la somme mentionnée selon les modalités indiquées dans la page suivante.")
                    doc.add_paragraph(f"{texte_fond_finance}")
                    doc.add_paragraph("""
                    Nous vous remercions par avance et vous prions de croire, Cher Investisseur, en l’expression de notre considération distinguée.
                    """)
                    doc.add_paragraph("L’équipe Middle Office\nÉpopée Gestion")

                    # Pas de saut de page ici (supprimé volontairement)

                    # ======== TITRE DEUXIÈME PAGE ========
                    doc.add_paragraph("")
                    doc.add_heading(f"{nom_fond}", level=1)
                    doc.add_heading(f"Appel de Fonds N°{numero_call}", level=2)

                    # ======== Tableau : Détail de l’opération ========
                    doc.add_paragraph("Détail de l’opération :")
                    table1 = doc.add_table(rows=3, cols=2)
                    table1.style = 'Table Grid'
                    table1.cell(0, 0).text = "% de l’Appel de Fonds"
                    table1.cell(0, 1).text = f"{pourcentage_call * 100:.2f} %"
                    table1.cell(1, 0).text = "Montant à libérer :"
                    table1.cell(1, 1).text = f"{format_nombre(df_nettoye[call][i])} €"
                    table1.cell(2, 0).text = "Date d’opération et de valeur, au plus tard le :"
                    table1.cell(2, 1).text = date_call.strftime('%d/%m/%Y')

                    doc.add_paragraph("")

                    # ======== Tableau : Situation après cette opération ========
                    doc.add_paragraph("Situation après cette opération :")
                    table2 = doc.add_table(rows=6, cols=2)
                    table2.style = 'Table Grid'
                    table2.cell(0, 0).text = "Montant de l’engagement initial"
                    table2.cell(0, 1).text = f"{format_nombre(df_nettoye['ENGAGEMENT'][i])} €"
                    table2.cell(1, 0).text = f"Nombre de Parts {df_nettoye['PART'][i]} souscrites"
                    table2.cell(1, 1).text = f"{format_nombre(df_nettoye['NBR PARTS'][i])}"
                    table2.cell(2, 0).text = "Capital appelé à date*"
                    table2.cell(2, 1).text = f"{format_nombre(df_nettoye['TOTAL APPELE'][i])} €"
                    table2.cell(3, 0).text = "% de l’engagement appelé à date*"
                    table2.cell(3, 1).text = f"{df_nettoye['%LIBERATION'][i] * 100:.2f} %"
                    table2.cell(4, 0).text = "Montant restant à appeler*"
                    table2.cell(4, 1).text = f"{format_nombre(df_nettoye['RESIDUEL'][i])} €"
                    table2.cell(5, 0).merge(table2.cell(5, 1))
                    table2.cell(5, 0).text = "*Incluant l’Appel de Fonds en cours"

                    doc.add_paragraph("")

                    # ======== Instructions de paiement ========
                    doc.add_paragraph(
                        "Veuillez procéder au paiement en €, net de tous frais bancaires, par virement bancaire sur le compte, selon les instructions suivantes :")
                    doc.add_paragraph("Détail du paiement :")

                    table3 = doc.add_table(rows=6, cols=2)
                    table3.style = 'Table Grid'
                    table3.cell(0, 0).text = "Titulaire du compte"
                    table3.cell(0, 1).text = "XPLORE  II"
                    table3.cell(1, 0).text = "Domiciliation"
                    table3.cell(1, 1).text = "CACEIS BANK FRANCE"
                    table3.cell(2, 0).text = "IBAN"
                    table3.cell(2, 1).text = "FR76 1812 9000 1000 5002 0849 481"
                    table3.cell(3, 0).text = "BIC (ADRESSE SWIFT)"
                    table3.cell(3, 1).text = "ISAEFRPPXXX"
                    table3.cell(4, 0).text = "Libellé du virement"
                    table3.cell(4, 1).text = f"CR {df_nettoye['SOUSCRIPTEUR'][i]} ADF {numero_call}"
                    table3.cell(5, 0).text = "Frais pour le bénéficiaire"
                    table3.cell(5, 1).text = "N/A (SHA)"

                    doc.add_paragraph("")
                    doc.add_paragraph("L’équipe Middle Office\nÉpopée Gestion")

                    # ======== Enregistrement ========
                    docx_path = f"Output/{df_nettoye['SOUSCRIPTEUR'][i]}.docx"
                    doc.save(docx_path)


            # Zip tous les fichiers
            shutil.make_archive("notices", "zip", "Output")
            with open("notices.zip", "rb") as f:
                st.download_button("Télécharger les notices générées", f, "notices.zip")

        except Exception as e:
            st.error(f"Une erreur est survenue : {e}")
    else:
        st.warning("Merci de déposer un fichier Excel avant de lancer la génération.")
