# -*- coding: utf-8 -*-

import os
import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image
from io import BytesIO
from datetime import datetime
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# Configurer Selenium pour Chrome en mode sans tête
options = Options()
options.headless = True  # Mode headless pour ne pas ouvrir de fenêtre de navigateur

# Configurer Selenium pour utiliser ChromeDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# URL fixe pour la recherche de produits
url = 'https://www.aliexpress.com/p/calp-plus/index.html'
driver.get(url)

# Attendre quelques secondes pour que la page soit complètement chargée
time.sleep(5)

# Simuler le défilement pour charger plus de produits
scroll_pause_time = 2  # Temps d'attente après chaque scroll
last_height = driver.execute_script("return document.body.scrollHeight")

produits_collectes = 0  # Compteur de produits collectés
max_produits = 60  # Limiter à 60 produits

while True:
    # Descendre jusqu'en bas de la page
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    # Attendre le chargement des nouveaux produits
    time.sleep(scroll_pause_time)

    # Calculer la nouvelle hauteur après défilement
    new_height = driver.execute_script("return document.body.scrollHeight")

    # Si la hauteur n'a pas changé, cela signifie qu'il n'y a plus de nouveaux produits
    if new_height == last_height:
        print("Fin du défilement, aucun nouveau produit à charger.")
        break

    last_height = new_height

    # Extraire le contenu de la page après chaque scroll pour collecter plus de produits
    page_source = driver.page_source
    soup = BeautifulSoup(page_source, 'html.parser')
    produits = soup.find_all('a', class_='_1UZxx')

    # Liste pour stocker les données
    liste_produits = []

    # Collecter les données des produits
    for produit in produits:
        # Nom du produit
        try:
            nom = produit.find('h3', class_='nXeOv').text.strip()
        except AttributeError:
            nom = 'N/A'

        # Prix du produit
        try:
            prix_elements = produit.find('div', class_='U-S0j').find_all('span')
            prix = ''.join([elem.text for elem in prix_elements]).replace('€', '').strip()
        except AttributeError:
            prix = 'N/A'

        # Image du produit
        try:
            image_url = produit.find('img', class_='_1IH3l product-img')['src']
            if image_url.startswith('//'):
                image_url = 'https:' + image_url
        except (AttributeError, TypeError):
            image_url = 'N/A'

        # Lien du produit
        try:
            lien = produit['href']
            if lien.startswith('//'):
                lien = 'https:' + lien
            else:
                lien = 'https://www.aliexpress.com' + lien
        except KeyError:
            lien = 'N/A'

        # Créer le dictionnaire
        dict_produit = {
            'Nom du Produit': nom,
            'Prix': prix,
            'Image': image_url,
            'Lien': lien
        }

        # Ajouter à la liste
        liste_produits.append(dict_produit)
        produits_collectes += 1

        # Si nous avons collecté 60 produits, nous arrêtons
        if produits_collectes >= max_produits:
            print(f"{max_produits} produits collectés. Fin de la collecte.")
            break

    # Vérifier si nous avons collecté 60 produits, et si oui, arrêter
    if produits_collectes >= max_produits:
        break

# Fermer le navigateur après avoir récupéré le contenu
driver.quit()

# Vérification des données collectées
print(f"Nombre total de produits collectés : {produits_collectes}")
print("Données collectées :")
for produit in liste_produits:
    print(produit)

# Création d'un dossier unique avec la date et l'heure
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
directory_name = f'produits_{timestamp}'
os.makedirs(directory_name)  # Crée un dossier avec la date et l'heure

# Création d'un fichier Excel dans ce dossier
file_name = f'{directory_name}/produits_viraux_{timestamp}.xlsx'

wb = Workbook()
ws = wb.active
ws.title = "Produits Viraux"

# Ajouter les en-têtes
ws.append(['Nom du Produit', 'Prix', 'Image', 'Lien'])
print("En-têtes ajoutées")

# Ajouter les données et les images
for index, produit in enumerate(liste_produits, start=2):  # Start=2 to skip the header row
    print(f"Ajout du produit : {produit['Nom du Produit']}")
    ws.cell(row=index, column=1, value=produit['Nom du Produit'])
    ws.cell(row=index, column=2, value=produit['Prix'])
    ws.cell(row=index, column=4, value=produit['Lien'])

    # Télécharger et ajouter l'image
    if produit['Image'] != 'N/A':
        try:
            image_response = requests.get(produit['Image'])
            img_data = BytesIO(image_response.content)
            img = Image.open(img_data)
            img.thumbnail((100, 100))  # Redimensionner l'image pour qu'elle tienne dans la cellule

            # Sauvegarder l'image dans le dossier créé
            img_path = f"{directory_name}/image_{index}.png"
            img.save(img_path)

            # Insérer l'image dans Excel
            excel_img = ExcelImage(img_path)
            ws.add_image(excel_img, f'C{index}')  # Insérer l'image dans la colonne C
            print(f"Image ajoutée pour {produit['Nom du Produit']}")
        except Exception as e:
            print(f"Impossible de télécharger l'image pour {produit['Nom du Produit']}: {e}")
    else:
        ws.cell(row=index, column=3, value='N/A')

# Enregistrer le fichier Excel avec un nom unique dans le dossier
wb.save(file_name)
print(f"Fichier Excel créé : {file_name}")
