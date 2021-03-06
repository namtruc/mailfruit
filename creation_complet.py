import pandas as pd
import os
import time
import sys
import re

from tkinter import Tk    
from tkinter.filedialog import askopenfilename
from configparser import ConfigParser  
from os import path as os_path
from openpyxl import load_workbook
import datetime

dossier_python = os_path.abspath(os_path.split(__file__)[0])
dossier_usr = dossier_python + '/fichiers_utilisateur'
os.chdir(dossier_python)

parser = ConfigParser() 
parser.read('configuration.ini')


def fct2 (lettre): ### fct verification lettre
    while convert_char(lettre) == 0 :
        print ("erreur")
        lettre = input("Indiquer de nouveau la lettre\n")
    else :
        lettre = convert_char(lettre)
        return lettre-1

def convert_char(old): ### fct convertion lettre en chiffre
    if len(old) != 1:
        return 0
    new = ord(old)
    if 65 <= new <= 90: # Majuscules
        return new - 64
    elif 97 <= new <= 122: # Minuscules   
        return new - 96 
    return 0 # Autres


#######################################################################################################
##### Determiner semaine

date1 = datetime.datetime.today()
date2 = date1.isocalendar()
jour = date1.weekday()
semaine = date2[1]

if jour > 4 :
    semaine += 1

print ('\nLa semaine concernee est la s'+str(semaine))

reponse = input("1. OK\n2. Entrer une autre semaine\n")

while True :

    if reponse=='1':

        break

    elif reponse=='2':

        semaine = input("Entrer le numero de semaine\n")

        break

    else:

        print ("Choix incorrect, recommencer\n")

        reponse = input()

######################################################################################################
##### Determiner le fichier excel 


input("\nAppuyer sur Entree pour choisir le fichier excel comportant la base de donnee clients (ID_groupe.xlsx)\n")

Tk().withdraw() 
contenu1 = askopenfilename()

xl = pd.ExcelFile(contenu1)

v_sheet = 0

if len(xl.sheet_names) > 1 :
    print ("\n---Les differents feuilles presentes---")
    for x in range (len(xl.sheet_names)) :
        print (str(x)+" "+xl.sheet_names[x])

    v_sheet = int(input("Entrer le numero de la feuille excel\n"))



#######################################################################################################
##### Definir les colonnes


vnomf = parser.get('id_groupe', 'nomf')
vdept = parser.get('id_groupe', 'dept')
vgrp = parser.get('id_groupe', 'grp')
vpnom = parser.get('id_groupe', 'pnom')
vresp = parser.get('id_groupe', 'resp')
vmail = parser.get('id_groupe', 'mail')

print ("Lettre colonne departement "+vdept)
print ("Lettre colonne groupe "+vgrp)
print ("Lettre colonne resp "+vresp)
print ("Lettre colonne nom famille "+vnomf)
print ("Lettre colonne prenom "+vpnom)
print ("Lettre colonne mail "+vmail)

print ("\n----Verifier que les informations ci dessus sont correctes----\n")
print ('Attention les lignes avec une case departement vide seront automatiquement supprimees\n')
time.sleep(1)


while 1:

    reponse = input("\nTaper 1 pour modifier manuellement les colonnes prises en compte, sinon taper 2\n")

    if reponse=='1':

        vmail = input("Indiquer la lettre de la colonne comportant les adreses mail\n")
        vmail = fct2 (vmail)
        vdept = input("Indiquer la lettre de la colonne comportant le departement\n")
        vdept = fct2 (vdept)
        vgrp = input("Indiquer la lettre de la colonne comportant le groupe\n")
        vgrp = fct2 (vgrp)
        vpnom = input("Indiquer la lettre de la colonne comportant le prenom\n")
        vpnom = fct2 (vpnom)
        vresp = input("Indiquer la lettre de la colonne comportant la case responsable\n")
        vresp = fct2 (vresp)
        vpnom = input("Indiquer la lettre de la colonne comportant le nom de famille\n")
        vnomf = fct2 (vnomf)
        break
        
    elif reponse == '2':

        vpnom = fct2 (vpnom)
        vmail = fct2 (vmail)
        vdept = fct2 (vdept)
        vgrp = fct2 (vgrp)
        vnomf = fct2 (vnomf)
        vresp = fct2 (vresp)

        break

    else :

        print ('Erreur')


#########################################################################
# Convertion fichier excel
##

read_file = pd.read_excel(contenu1,sheet_name=v_sheet,usecols='A:H')

read_file.to_csv ("Test1.csv",  
                  index = None, 
                  header = ['Dept', 'Grp', 'Resp','','Nom', 'Prenom','Mail',''])

df = pd.read_csv("Test1.csv", usecols=[vdept,vgrp,vresp,vnomf,vpnom,vmail]) 

########################################################################################################### Nettoyage et verif

pd.DataFrame(df.dropna(subset=['Resp'], inplace=True)) ######suppression des non-resp
pd.DataFrame(df.dropna(subset=['Dept'], inplace=True))#Suppr dept vide

df[vresp] = df['Resp'].str.strip()#### suppression espace colonne resp
df = df[(df['Resp'].str.match('X'))|(df['Resp'].str.match('x'))]#### suppression resp sans X

os.remove('Test1.csv')

#print (df)


list2 = [] ### reconstruction de l'index

for x in range (len(df.index)):
 list2.append(x)

df.index = list2


#### Verif mail

EMAIL_REGEX = re.compile("[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+")
        
df2 = df

for x in range (len(df.index)):

    you = (df.at[x,'Mail'])
    you = str(you).strip()

    if not EMAIL_REGEX.match(str(you)):

        print ("\n----Mail manquant ou mal redige groupe "+str(df.at[x,'Grp'])+"----\n")

        reponse_mail = int(input("1 pour continuer sans incorporer ce groupe\n2 pour entrer l'adresse manuellement \n"))

        if reponse_mail == 1 :

            df2 = df2.drop([x])
            
        if reponse_mail == 2 :

            while True :

                reponse_mail2 = input("\n Entrer l'adresse voulue\n")

                if not EMAIL_REGEX.match(reponse_mail2):

                    print ("Erreur redaction mail")

                else :

                    df.at[x,'Mail'] = reponse_mail2.strip()

                    break
                    
            
df = df2

list2 = [] ### reconstruction de l'index

for x in range (len(df.index)):
 list2.append(x)

df.index = list2

#### Verif pnom/grp

def fct3 (mot,var):

    global df2

    if str(df.at[x,var])=='nan' or str(df.at[x,var])=='NaN':

        print ("Erreur case vide colonne ---"+mot+"--- groupe "+str(df.at[x,'Grp'])+" departement "+str(df.at[x,'Dept']) )

        rep = input('Taper entree pour rentrer manuellement un '+mot+'\n')

        nve = input('\nEntrer le '+mot+'\n')
        df.at[x,var] = nve




for x in range (len(df.index)):

    fct3('prenom', 'Prenom')

    
for x in range (len(df.index)):

    fct3('groupe', 'Grp')


    

###############################################################################################
### Enregistrement et rename

os.chdir(dossier_usr)

nom_fichier = 'Fichier_envoi_s'+str(semaine)+'.xlsx'

df.to_excel(nom_fichier, index =False, columns=['Dept', 'Grp', 'Nom', 'Prenom','Mail'])

os.chdir(dossier_python)
print ('\n')
print (nom_fichier+' enregistre dans le dossier\n'+os.path.abspath(os.getcwd())) 

input ('\nAppuyer sur Entree pour quitter')

import main.py
