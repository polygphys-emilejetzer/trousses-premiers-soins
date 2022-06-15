# -*- coding: utf-8 -*-
"""Transmettre les nouvelles inscriptions au SIMDUT."""

# Bibliothèque standard
from pathlib import Path
import time

# Bibliothèque PIPy
import schedule

# Bibliothèques maison
from polygphys.outils.reseau import OneDrive
from polygphys.outils.reseau.msforms import MSFormConfig, MSForm
from polygphys.outils.reseau.courriel import Courriel


class SSTSIMDUTInscriptionConfig(MSFormConfig):

    def default(self):
        return (Path(__file__).parent / 'premiers_soins.cfg').open().read()

class SSTSIMDUTInscriptionForm(MSForm):

    def nettoyer(self, cadre):
        cadre = self.convertir_champs(cadre)
        colonnes = ['date'] + self.config.getlist('formulaire', 'colonnes')[:-1]
        cadre = cadre.loc[:, colonnes]\
                     .fillna(0)
        return cadre

    def action(self, cadre):
        try:
            if not cadre.empty:
                fichier_temp = Path('trousses.xlsx')
                cadre.to_excel(fichier_temp)
                pièces_jointes = [fichier_temp]

                message = 'Bonjour! Voici des items manquants à certaines trousses de premiers soins. Bonne journée!'
                html = f'<p>{message}</p>{cadre.to_html()}'
            else:
                return
        except Exception as e:
            message = f'L\'erreur {e} s\'est produite.'
            html = f'<p>{message}</p>'

        courriel = Courriel(self.config.get('courriel', 'destinataire'),
                            self.config.get('courriel', 'expéditeur'),
                            self.config.get('courriel', 'objet'),
                            message,
                            html,
                            pièces_jointes=pièces_jointes)
        courriel.envoyer(self.config.get('courriel', 'serveur'))

chemin_config = Path('~').expanduser() / 'premiers_soins.cfg'
config = SSTSIMDUTInscriptionConfig(chemin_config)

dossier = OneDrive('',
                   config.get('onedrive', 'organisation'),
                   config.get('onedrive', 'sous-dossier'),
                   partagé=True)
fichier = dossier / config.get('formulaire', 'nom')
config.set('formulaire', 'chemin', str(fichier))

formulaire = SSTSIMDUTInscriptionForm(config)

schedule.every().friday.at('13:00').do(formulaire.mise_à_jour)

formulaire.mise_à_jour()
try:
    while True:
        schedule.run_pending()
        time.sleep(1)
except KeyboardInterrupt:
    print('Fin.')
