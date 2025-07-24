from pathlib import Path
import pandas as pd
from jinja2 import Environment, FileSystemLoader
import os
from weasyprint import HTML
import datetime as dt

# ============= Fonctions utiles =============
def date_now():
    mois = ['', 'janvier', 'février', 'mars', 'avril', 'mai', 'juin',
            'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre']
    today = dt.date.today()
    return f"{today.day} {mois[today.month]} {today.year}"

def format_nombre(nombre):
    return f"{nombre:,.2f}".replace(',', ' ').replace('.', ',')


# ============= Données Brute (pour l'instant) =============

nom_fond = "FPCI ÉPOPÉE Xplore II"
pays = "France"
numero_call = "9"

# Texte en jaune
texte_fond_couvrir = "Cet Appel était destiné à couvrir notre investissement dans la société RainPath ainsi que le règlement de la commission de gestion du second semestre 2025."
texte_fond_finance = "Cet Appel de Fonds est destiné à financer notre investissement dans la société Cézam et notre second réinvestissement dans la société Deski."



# ============== Ouverture du fichier excel ===============

# Chemin du fichier Excel
chemin_fichier = 'ressources/Base-data-test-fund-exercice.xlsx'

try:
    # Lit le fichier Excel et crée un DataFrame
    df = pd.read_excel(chemin_fichier, sheet_name='SOUSCRIPTEURS', header=18) # Ligne 18 en python = Ligne 19 en Excel

    df_nettoye = df[df['SOUSCRIPTEUR'].notna()]
    df_nettoye = df_nettoye[~df_nettoye['SOUSCRIPTEUR'].str.startswith('TOTAL', na=False)]

    # Réinitialise l'index
    df_nettoye = df_nettoye.reset_index(drop=True)

except FileNotFoundError:
    print(f"Erreur : Le fichier {chemin_fichier} n'a pas été trouvé.")
except Exception as e:
    print(f"Une erreur s'est produite lors de la lecture du fichier : {e}")


try:
    # Lit le fichier Excel et crée un DataFrame
    df_CALL = pd.read_excel(chemin_fichier, sheet_name='SOUSCRIPTEURS', header=3)

except FileNotFoundError:
    print(f"Erreur : Le fichier {chemin_fichier} n'a pas été trouvé.")
except Exception as e:
    print(f"Une erreur s'est produite lors de la lecture du fichier : {e}")


# ============== Ouverture du modele (HTML) ===============

dir = 'ressources'
env = Environment(loader=FileSystemLoader(dir))
template = env.get_template('model_notice_img.html')


# ============== Remplissage des balises ===============

# Taille du dataframe
nb_lignes_nettoye, nb_colonnes_nettoye = df_nettoye.shape
nb_lignes, nb_colonnes = df.shape


# récupération des données du CALL
call = 'CALL #' + numero_call
montant_total = df[call][nb_lignes-6]

date_call = df_CALL.loc[df_CALL['Nominal'] == call, 'Date'].iloc[0]

pourcentage_call = df_CALL.loc[df_CALL['Nominal'] == call, df_CALL.columns[2]].iloc[0]



for i in range(nb_lignes_nettoye):

    total_avant_call = df_nettoye['TOTAL APPELE'][i] - df_nettoye[call][i]
    pourcentage_avant_call = (total_avant_call / df_nettoye['ENGAGEMENT'][i]) * 100

    if pd.isna(df_nettoye["Représentant"][i]):
        representant = ''
    else:
        representant = df_nettoye["Représentant"][i]

    # Les données à injecter
    # 'balise' : 'la donnée',
    data = {
        'souscripteur' : df_nettoye["SOUSCRIPTEUR"][i],
        'pm_pp' : df_nettoye["TYPE"][i],
        'representant' : representant,
        'adresse' : df_nettoye["ADRESSE"][i],
        'code_postal' : round(df_nettoye["CP"][i]),
        'ville' : df_nettoye["VILLE"][i],
        'pays' : pays,
        'date' : date_now(),
        'numero_call' : numero_call,
        'date_call' : date_call.strftime('%d/%m/%Y'),
        'nom_fond' : nom_fond,
        'montant_total' : format_nombre(montant_total),
        'pourcentage_call' : f"{pourcentage_call * 100:.2f}",
        'montant_a_liberer' : format_nombre(df_nettoye[call][i]),
        'pourcentage_avant_call' : format_nombre(pourcentage_avant_call),
        'texte_fond_couvrir' : texte_fond_couvrir,
        'texte_fond_finance' : texte_fond_finance,
        'montant_engagement_initial' : format_nombre(df_nettoye["ENGAGEMENT"][i]),
        'nombre_parts_souscrites' : format_nombre(df_nettoye["NBR PARTS"][i]),
        'categorie_part' : df_nettoye["PART"][i],
        'total_appele' : format_nombre(df_nettoye["TOTAL APPELE"][i]),
        'pourcent_liberation' : f"{df_nettoye['%LIBERATION'][i] * 100:.2f}",
        'residuel' : format_nombre(df_nettoye["RESIDUEL"][i]),
        'libelle_virement' : 'CR '+df_nettoye["SOUSCRIPTEUR"][i]+' ADF '+numero_call
        }

    # Rend le HTML final avec tes vraies données
    html_content = template.render(data)

    # Sauve le résultat dans un fichier
    os.makedirs('Output_HTML', exist_ok=True)
    dir_nom_fichier = 'Output_HTML/' + df_nettoye["SOUSCRIPTEUR"][i] + '_' + df_nettoye["PART"][i] + '.html'
    with open(dir_nom_fichier, 'w', encoding='utf-8') as file:
        file.write(html_content)

    #print('Notice HTML générée avec succès.')

    fichier_html = 'Output_HTML/' + df_nettoye["SOUSCRIPTEUR"][i] + '_' + df_nettoye["PART"][i] + '.html'
    fichier_pdf = 'Output/' + df_nettoye["SOUSCRIPTEUR"][i] + '_' + df_nettoye["PART"][i] + '.pdf'

    base_url = Path('ressources/images').resolve()  # Chemin absolu vers /ressources

    HTML(filename=fichier_html, base_url=base_url.as_uri()).write_pdf(fichier_pdf)


















