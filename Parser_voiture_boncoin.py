#-------------------------------------------------------------------------------
# Name:        module2
# Purpose:
#
# Author:      Benjamin
#
# Created:     24/09/2020
# Copyright:   (c) Benjamin 2020
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import time
import requests
import datetime
import encodings
import requests
import xlwings as xw
from bs4 import BeautifulSoup
import re

aujourdhui =  datetime.date.today()
aujourdhui = str(aujourdhui)

wb = xw.Book(r'C:\Users\Benjamin\Documents\Projets_Python\annonces_voitures.xlsx')
sht = wb.sheets['Sheet1']


def est_voiture(href):
    return href and re.compile("/voitures").search(href)

header = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:81.0) Gecko/20100101 Firefox/81.0',
    'referer': 'https://www.leboncoin.fr/'
}

tabPagesVoitures = []
##  requetes sur la page résultat de toutes les voitures de 1000€  ##############
r = requests.get("https://www.leboncoin.fr/recherche/?category=2&locations=Castres_81100__43.63007_2.21067_10000&sort=price&order=asc&fuel=2&price=500-500&regdate=1999-max",headers=header)
r.cookies
soup = BeautifulSoup(r.content,'lxml',from_encoding='utf-8')
### récupère tout les <a> contenu "/voitures" dans le "href"
liensVoitures = soup.find_all(href=est_voiture)
##
##
### on parse toutes les pages en liens et insère leur contenu dans le tableau tabPagesVoitures
for liens in liensVoitures:
    r = requests.get('https://www.leboncoin.fr' + liens.attrs["href"],headers=header)
    tabPagesVoitures.append(r)
    print(liens.attrs["href"])
    time.sleep(7)
###########################################################################################

longueurTabPagesVoitures = len(tabPagesVoitures)
for i in range(longueurTabPagesVoitures-1):
    derniereLigne = sht.range('A1').end('down').row

#r = requests.get('https://www.leboncoin.fr/voitures/1846216367.htm',headers=header)
#r.cookies
    soup = BeautifulSoup(tabPagesVoitures[i].content,'lxml',from_encoding='utf-8')
    ### récupère tout les <a> contenu "/voitures" dans le "href"
    zoneMarque = soup.findAll("p",string="Marque")
    itemMarque = str(zoneMarque[0].contents[0].next.contents[0]) #OK

    zoneModele = soup.findAll("p",string="Modèle")
    itemModele = str(zoneModele[0].contents[0].next.contents[0]) #OK

    zoneAnneeModele = soup.findAll("p",string="Année-modèle")
    itemAnneeModele  = str(zoneAnneeModele[0].contents[0].next.contents[0]) #OK

    zoneKilometrage = soup.findAll("p",string="Kilométrage")
    if zoneKilometrage != '':
        itemKilometrage  = str(zoneKilometrage[0].contents[0].next.contents[0]) #OK
        itemKilometrage = itemKilometrage.replace('km','')
        itemKilometrage = itemKilometrage.replace(' ','')
    else :
        itemKilometrage = 'non renseigné'

    zoneCarburant = soup.findAll("p",string="Carburant")
    itemCarburant  = str(zoneCarburant[0].contents[0].next.contents[0]) #OK

    #zoneBoiteVitesse = soup.findAll("p",string="Boîte de vitesse")
    #itemBoiteVitesse  = str(zoneBoiteVitesse[0].contents[0].next.contents[0]) #OK
    zonePrix = soup.findAll(class_=re.compile("price"))
    itemPrix  = str(zonePrix[0].next.next.contents[0].contents[0])
    itemPrix = itemPrix.replace(' ','')

    sht.range(derniereLigne+1,1).value = str(itemMarque)
    sht.range(derniereLigne+1,2).value = str(itemModele)
    sht.range(derniereLigne+1,3).value = str(itemAnneeModele)
    sht.range(derniereLigne+1,5).value = int(itemKilometrage)
    sht.range(derniereLigne+1,6).value = str(itemCarburant)
    sht.range(derniereLigne+1,7).value = int(itemPrix)

wb.save()
wb.close()



