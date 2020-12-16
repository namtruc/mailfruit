# TODO
# verification si variable presentes dans mail
# API

import smtplib
import sys
import re
import os
import datetime
import pandas as pd
import random
import time
#import getpass

from tkinter import Tk    
from tkinter.filedialog import askopenfilename
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from email.utils import COMMASPACE, formatdate
from email.mime.application import MIMEApplication
from os import path as os_path
from configparser import ConfigParser  
from html2text import html2text


dossier_python = os_path.abspath(os_path.split(__file__)[0])      
dossier_usr = dossier_python + '/fichiers_utilisateur'
os.chdir(dossier_python)

parser = ConfigParser() 
parser.read('configuration.ini')

me = parser.get('settings', 'mail_expe')## parametres presents dans le fichier de config
usr = parser.get('settings', 'username')
srv = parser.get('settings', 'srv_smtp')
prt = parser.get('settings', 'prt_smtp')
psswd = parser.get('settings', 'psswd')

vmail = int(parser.get('fichier_envoi', 'mail'))
vdept = int(parser.get('fichier_envoi', 'dept'))
vgrp = int(parser.get('fichier_envoi', 'grp'))
vpnom = int(parser.get('fichier_envoi', 'pnom'))
vlien = int(parser.get('fichier_envoi', 'lien'))

m=0
n = 1

date = datetime.datetime.today().strftime('%d.%m.%y-%Hh%M')
date1 = datetime.datetime.today()
date2 = date1.isocalendar()
jour = date1.weekday()
semaine = date2[1]


   
#######################################################################################################
#####Determiner fichier HTML

os.chdir(dossier_usr)

reponse = input("Appuyer sur entree pour choisir le fichier mail\n")

Tk().withdraw() 
contenu = askopenfilename()
        
with open(contenu, 'r', encoding="utf8") as file_in :
  file_out = file_in.read()


#######################################################################################################
##### Definir titre

time.sleep(1)
if jour > 4 :
    semaine += 1

titre = 'Catalogue S'+str(semaine)

print ("\n---Titre du mail (exemple)---")
print ("[22_Saint-Brieuc] "+titre+"\n")


reponse = input("1. OK\n2. Entrer un autre titre \n")

while True :
    if reponse=='1':
        break
    elif reponse=='2':
        titre = input("Entrer le titre du mail : (le dep et le groupe seront rajoutes automatiquement)\n")
        print ("\n[22_Saint-Brieuc] "+titre+"\n")
        reponse = input("1. OK\n2.Entrer un autre titre \n")
    else:
        print ("Choix incorrect ! Titre non modifie")
        break




#######################################################################################################
##### Determiner les pieces jointes

os.chdir(dossier_usr)
dossier_pj = dossier_usr

l = 0
e = [] #### liste contenant piece jointe

while 1:
    reponse = input("\nTaper 1 pour choisir des pieces jointes\nTaper 2 si il n'y a pas de pieces jointes\n")
    if reponse=='1':
        break
    elif reponse =='2':
        l = 1
        break
    else:
        print ("Choix incorrect !")

if l == 0 :

    nbr_pj = int(input("Nombre de pieces jointes\n"))
    for x in range (nbr_pj):
        input("\nAppuyer sur Entree pour choisir la piece jointe "+str(x+1))
        time.sleep(1)
        Tk().withdraw() 
        pj = askopenfilename()
        e.append(pj)


#######################################################################################################
##### Determiner le fichier excel 


input("\nAppuyer sur Entree pour choisir le fichier excel comportant la base de donnee avec les mails et les liens catalogue\n")

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
##### Determiner les lignes prises en compte


reponse = input("\nTaper 1 si le script doit prendre en compte l'ensemble des lignes des destinataires.\nTaper 2 pour indiquer quelles lignes doivent etre prises en compte\n")

if reponse=='2':
    vligned = (int(input("Numero de la ligne de debut de la liste des destinataires "))-1)
    vlignef = int(input("Numero de la ligne de fin de la liste des destinataires "))
else:
    vligned = 1
    vlignef = 1
    
    
       
#######################################################################################################
##### Definir les variables


print ('\nLa premiere variable du mail correspond au prenom, le deuxieme au lien one drive')

r = int(input('\nTapper 1 pour continuer\nTapper 2 pour quitter\n'))

if r == 2:
    sys.exit(0)
    
    
#########################################################################
# Convertion fichier excel

vnbrligne = (vlignef-(vligned))

read_file = pd.read_excel(contenu1,sheet_name=v_sheet,skiprows = vligned, header=None)         

read_file.to_csv ("Test.csv",  
                  index = None, 
                  header = True)################uft8 a faire


df = pd.DataFrame(pd.read_csv("Test.csv")) 

if vlignef != 1 :
    df = df[:vnbrligne]

os.remove("Test.csv")



#######################################################################################################
#####Connection

       
print("\nLe mail de l'expediteur est :\n")
print("1.", me)
print("2.", usr)
print("3. Quitter le programme")

while 1:
    reponse = input("Choisir 1,2,3 ou 4:\n")
    if reponse=='1':
        break
    elif reponse =='2':
        me = usr
        break
    elif reponse=='3':
       sys.exit(0)
    
 
mail = smtplib.SMTP(srv, prt)
mail.ehlo()
mail.starttls()
#mail.set_debuglevel(True)
#mdp=getpass.getpass("Entrez le mot de passe pour "+usr+"\n")
mail.login(usr, psswd)

while True:
    print("\nConnection reussie")
    break


################################################################### 
########################## Fct_Envoi


def fct_envoi(essai,Range):

    os.chdir(dossier_python)
    global n
    global m

    for x in range (Range): ### determination adresse mail 

        if essai == True:

            x = random.randint(1,len(df.index)) 
            you = input("L´adresse mail pour l´essai\n")
            print("Essai avec le groupe "+str(df.iat[x,vdept])+" "+str(df.iat[x,vgrp]))

        else :

            you = (df.iat[x,vmail])

        you = you.strip()
        file_out2 = file_out
    
        file_out2 = file_out2.replace('variable1', str(df.iat[x,vpnom]))
        file_out2 = file_out2.replace('variable2', str(df.iat[x,vlien]))
    
        titre2 = ("["+str(df.iat[x,vdept])+"_"+str(df.iat[x,vgrp])+"] "+titre)
        
        msg = MIMEMultipart("alternative")
        msg['Subject'] = titre2
        msg['From'] = me
        msg['To'] = you
        msg['Date'] = formatdate(localtime=True)
        
        html = file_out2
        soup = html2text(file_out2)
        
        html_part = MIMEText(html, "html")
        text_part = MIMEText(soup, "plain")
        
        msg.attach(html_part)
        msg.attach(text_part)
    
        if l == 0 :### pieces jointes
            os.chdir(dossier_pj)

            for z in range (nbr_pj):

                pj = MIMEApplication(open(e[z],'rb').read())
                pj.add_header('Content-Disposition','attachment',filename=os.path.basename(e[z]))
                msg.attach(pj)
    
        os.chdir(dossier_usr)
    

        try:

            mail.sendmail(me, you, msg.as_string())         

            if essai == False :

                with open("sortieincomplete.txt", "a") as myfile:
                    myfile.write(str(df.iat[x-1,vdept])+"/"+str(df.iat[x-1,vpnom])+"/"+str(df.iat[x-1,vgrp])+"/"+you+" mail envoye\n")
                print (str(n)+"/"+str(len(df.index)))
                n=n+1

        except smtplib.SMTPException as f:

            print ("Erreur dans un envoi")
            print (f)

            if essai == False :

                with open("sortieincomplete.txt", "a") as myfile:
                    myfile.write(str(df.iat[x,vdept])+"/"+str(df.iat[x,vpnom])+"/"+str(df.iat[x,vgrp])+"/"+you+" MAIL NON ENVOYE\n")

                m=m+1


####################################################################################
################ Envoi


n = 1
m = 0

essai = int(input("\nVoulez vous faire un essai sur votre adresse mail ?\n1 : oui\n2 : non\n"))

if essai == 1:

    while True:

        fct_envoi(True,1)
            
        essai = int(input("\n1 pour envoyer un autre essai, 2 pour arreter le programme, 3 pour continuer les envois\n"))

        if essai == 2 :

            sys.exit(0)

        if essai == 3 :

            break


essai2 = int(input("\n"+str(len(df.index))+" mails seront envoyes \n1 : oui\n2 : arreter le programme\n"))

if essai2 == 2 :

    sys.exit(0)

else :

    fct_envoi(False, len(df.index))


#######################################################################################################

mail.quit()

print ("\n"+str(n-1)+" mails envoyes,"+str(m)+" erreurs")

os.chdir(dossier_usr)
os.rename("sortieincomplete.txt", date+"_envoi.txt" )

print ("\n"+"fichier "+date+"_envoi.txt cree dans le dossier fichier utilisateur")

fin = input("Envois termines, appuyer sur Entree pour quitter le programme")
