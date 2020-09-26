import smtplib
import sys
import re
import os
import datetime
import pandas as pd
import random

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from email.utils import COMMASPACE, formatdate
from email.mime.application import MIMEApplication
from os import path as os_path

def convert_char(old): ### fct convertion lettre en chiffre
	if len(old) != 1:
		return 0
	new = ord(old)
	if 65 <= new <= 90: # Majuscules
           return new - 64
	elif 97 <= new <= 122: # Minuscules   
		return new - 96 
	return 0 # Autres
       
def fct2 (lettre): ### fct verification lettre
	while convert_char(lettre) == 0 :
		print ("erreur")
		lettre = input("Indiquer de nouveau la lettre\n")

	else :
		lettre = convert_char(lettre)
		return lettre-1
		
#me = 'anne_cecile.laveder@fruitstock.eu'
me = 'catalogue@fruitstock.eu'

m=0
n = 1
date = datetime.datetime.today().strftime('%d.%m.%y-%Hh%M')
dossier_python = os_path.abspath(os_path.split(__file__)[0])



#######################################################################################################
#####Connection

mail = smtplib.SMTP('smtp.gmail.com', 587)
mail.ehlo()
mail.starttls()

usr = me#input("Entrez le mail utilisateur\n")
mdp = input("Entrez le mot de passe\n")
mail.login(usr, mdp)

while True:
     print("Connection reussie")
     break
          
print("Le mail de l'expediteur est :")
print("1.", me)
print("2.", usr)
print("3. Entrer une autre adresse")
print("4. Quitter le programme")

while 1:
    reponse = input("Choisir 1,2,3 ou 4:\n")
    if reponse=='1':
        break
    elif reponse =='2':
        me = usr
        break
    elif reponse=='3':
        me = input("Rentrer l'adresse voulue\n")
        print("-----")
        print("Le mail de l'expediteur est :")
        print("1.", me)
        print("2.", usr)
        print("3. Entrer une autre adresse")
        print("4. Quitter le programme")
    elif reponse=='4':
        sys.exit(0)
    else:
        print ("Choix incorrect !")
    
    
#######################################################################################################
#####Determiner fichier HTML

dossier = dossier_python

while 1:
    reponse = input("Taper 1 pour choisir le fichier HTML dans le repertoire actuel (recommande)\nTaper 2 pour choisir un autre repertoire\n")
    if reponse=='1':
        break
    elif reponse =='2':
        dossier = input("Entrer le chemin du dossier\n")
        break
    else:
        print ("Choix incorrect !")   
        
items = os.listdir(dossier)
newlist = []
for names in items:
    if names.endswith(".html"):
        newlist.append(names)
print ("Contenu du dossier\n", newlist)
contenu = input("Entrer le nom complet du fichier HTML contenant le texte brut du mail\n")

os.chdir(dossier)

with open(contenu, 'r', encoding="utf8") as file_in :
  file_out = file_in.read()

os.chdir(dossier_python)

while True :
  titre = input("Entrer le titre du mail :\n")
  reponse = input("1. OK\n2.Recommencer\nChoisir 1 ou 2\n")
  if reponse=='1':
      break
  elif reponse=='2':
      print("...\n")
  else:
      print ("Choix incorrect !")
      

#######################################################################################################
##### Definir les variables

d = []

nombre_variables = input("Entrer le nombre de variables presentes dans le texte du mail(nom, lien, etc...)\n")

vmail = input("Indiquer la lettre de la colonne comportant les adreses mail\n")#-1
vmail = fct2 (vmail)

vdept = input("Indiquer la lettre de la colonne comportant le departement\n")#-1
vdept = fct2 (vdept)

vpnom = input("Indiquer la lettre de la colonne comportant le prenom\n")#-1
vpnom = fct2 (vpnom)

for x in range(1, (int(nombre_variables)+1)):
	vvar = input("Indiquer la lettre de la colonne comportant la variable "+str(x)+"\n")#)-1)
	vvar = fct2 (vvar)
	d.append(int(vvar))
	
vligned = int(input("Numero de la ligne de debut de la liste des destinataires "))-1

vlignef = int(input("Numero de la ligne de fin de la liste des destinataires "))-1


#######################################################################################################
##### Determiner les pieces jointes

os.chdir(dossier_python)
dossier_pj = dossier_python

l = 0
e = []

while 1:
	reponse = input("Taper 1 pour choisir les pieces jointes dans le repertoire actuel (recommande)\nTaper 2 pour choisir un autre repertoire\nTaper 3 si il n´y a pas de pieces jointes\n")
	if reponse=='1':
		break
	elif reponse =='2':
		dossier_pj = input("Entrer le chemin du dossier contenant TOUTES les pieces jointes\n")
		break
	elif reponse =='3':
		l = l+1
		break
	else:
		print ("Choix incorrect !")

if l == 0 :
	
	nbr_pj = int(input("Nombre de pieces jointes\n"))
	
	items = os.listdir(dossier_pj)
	newlist3 = []
	for names in items:
		newlist3.append(names)
	print ("Contenu du dossier\n", newlist3)
	
	for x in range (nbr_pj):
		e.append(str(input("Entrer le nom complet de la piece jointe "+str(x+1)+"\n")))


#######################################################################################################
##### Determiner le fichier excel

dossier = dossier_python

while 1:
    reponse = input("Taper 1 pour choisir le fichier excel ou libreoffice dans le repertoire actuel (recommande)\nTaper 2 pour choisir un autre repertoire\n")
    if reponse=='1':
        break
    elif reponse =='2':
        dossier = input("Entrer le chemin du dossier\n")
        break
    else:
        print ("Choix incorrect !")   
        
items = os.listdir(dossier)
newlist2 = []
for names in items: 
    if names.endswith(".ods"):
        newlist2.append(names)
    elif names.endswith(".xlsx"):
        newlist2.append(names)
    elif names.endswith(".xls"):
        newlist2.append(names)
print ("Contenu du dossier\n", newlist2)

contenu = input("Entrer le nom complet du fichier contenant la base de donnee\n")

os.chdir(dossier) 

df = pd.DataFrame(pd.read_excel(contenu))
	

#######################################################################################################
##### Verification
  
  
EMAIL_REGEX = re.compile("[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+")
i = 0

for x in range (vligned, vlignef+1): ### verification validite mails
	you = (df.iat[x-1,vmail])
	if not EMAIL_REGEX.match(str(you)):
		print ("erreur mail ligne "+str(x))      
		i = 1

if i == 1 :
	sys.exit(0)

###################################################################	
########################## Essai

essai = int(input("Voulez vous faire un essai sur votre adresse mail ?\n1 : oui\n2 : non\n"))

if essai == 1 :
	you = input("L´adresse mail pour l´essai\n")
	x = random.randint(vligned,vlignef)
	file_out2 = file_out
	print("Essai avec la ligne "+str(x+1))
	for y in range (1, (int(nombre_variables)+1)): ### remplacement variable
		file_out2 = file_out2.replace('variable'+str(y), str(df.iat[x-1,d[y-1]]))
		
	msg = MIMEMultipart()
	msg['Subject'] = titre
	msg['From'] = me
	msg['To'] = you
	msg['Date'] = formatdate(localtime=True)
	
	html = file_out2
	
	html_part = MIMEText(html, "html")
	msg.attach(html_part)

	if l == 0 :### pieces jointes
		os.chdir(dossier_pj)
		for z in range (nbr_pj):
			pj = MIMEApplication(open(e[z],'rb').read())
			pj.add_header('Content-Disposition','attachment',filename=str(e[z]))
			msg.attach(pj)

	try:
		mail.sendmail(me, you, msg.as_string())         
		print ("Mail Envoyé")
		
	except smtplib.SMTPException as f:
		print ("Erreur dans l´envoi")
		print (f)
		sys.exit(0)

	essai = int(input("1 pour continuer, 2 pour arreter\n"))
	if essai == 2 :
		sys.exit(0)
		
#######################################################################################################
##### Envoi du mail

k= int((vlignef+1)-(vligned)) ### nbr mails a envoyer

for x in range (vligned, vlignef+1): ### determination adresse mail
	you = (df.iat[x-1,vmail])
	file_out2 = file_out
	for y in range (1, (int(nombre_variables)+1)): ### remplacement variable
		file_out2 = file_out2.replace('variable'+str(y), str(df.iat[x-1,d[y-1]]))
		
	msg = MIMEMultipart()
	msg['Subject'] = titre
	msg['From'] = me
	msg['To'] = you
	msg['Date'] = formatdate(localtime=True)
	
	html = file_out2
	
	html_part = MIMEText(html, "html")
	msg.attach(html_part)

	if l == 0 :### pieces jointes
		os.chdir(dossier_pj)
		for z in range (nbr_pj):
			pj = MIMEApplication(open(e[z],'rb').read())
			pj.add_header('Content-Disposition','attachment',filename=str(e[z]))
			msg.attach(pj)

	try:
		mail.sendmail(me, you, msg.as_string())         
		with open("sortie.txt", "a") as myfile:
			myfile.write(str(df.iat[x-1,vdept])+"/"+str(df.iat[x-1,vpnom])+"/"+you+" mail envoye\n")
		print (str(n)+"/"+str(k))
		n=n+1
	except smtplib.SMTPException as f:
		print ("Erreur dans un envoi")
		print (f)
		with open("sortie.txt", "a") as myfile:
			myfile.write(str(df.iat[x-1,vdept])+"/"+str(df.iat[x-1,vpnom])+"/"+you+" MAIL NON ENVOYE\n")
		m=m+1


#######################################################################################################

mail.quit()

print (n-1, "mails envoyes,",m,"erreurs")

os.rename("sortie.txt", date+".txt" )

# essai avec + 10 variables
#Essai avec tableau complete
#Connection ?
#Lien Hypertexte ?
# tester erreur lignes vide
##### Faire fichier parametres
##### Essayer avec differant types de connection
##### Tester avec 0 parametres
