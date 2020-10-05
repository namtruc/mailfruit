import smtplib
import sys
import re
import os
import datetime
import pandas as pd
import random
import time

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from email.utils import COMMASPACE, formatdate
from email.mime.application import MIMEApplication
from os import path as os_path
from configparser import ConfigParser  

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
		
def fct3 (vide, mot): ###fct verif case vide
	vide = df.columns.values[vide]
	check = pd.isna(df[vide])
	for x in range ((len(df.index))):
		if check[x+1] == True :
			print ("Erreur case vide colonne "+mot+" groupe "+str(df.iat[x,vgrp]) )
			i = 1

def fct4 (vide, mot): ###fct 3 pour vgrp
	vide = df.columns.values[vide]
	check = pd.isna(df[vide])
	for x in range ((len(df.index))):
		if check[x+1] == True :
			print ("Erreur case vide colonne "+mot+" groupe "+str(df.iat[x,vdept]) )
			i = 1	
					
parser = ConfigParser() 
parser.read('configuration.ini')

me = parser.get('settings', 'mail_expe')## parametres presents dans le fichier de config
usr = parser.get('settings', 'username')
srv = parser.get('settings', 'srv_smtp')
prt = parser.get('settings', 'prt_smtp')

m=0
n = 1
date = datetime.datetime.today().strftime('%d.%m.%y-%Hh%M')
dossier_python = os_path.abspath(os_path.split(__file__)[0])



#######################################################################################################
#####Connection

mail = smtplib.SMTP(srv, prt)
mail.ehlo()
mail.starttls()

#input("Entrez l'identifiant de connection (mail utilisateur)\n")
mdp = input("Entrez le mot de passe pour "+usr+"\n")
mail.login(usr, mdp)

while True:
	print("Connection reussie")
	break
          
print("Le mail de l'expediteur est :")
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
  print ("-------------------------------------")
  print ("[00_perpette les oies]"+titre)
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

print ("-----Attention, une erreur de frappe peut entrainer le crash du programme-----")
time.sleep(3)

nombre_variables = input("Entrer le nombre de variables presentes dans le texte du mail, max 10\n")

vmail = input("Indiquer la lettre de la colonne comportant les adreses mail\n")#-1
vmail = fct2 (vmail)

vdept = input("Indiquer la lettre de la colonne comportant le departement\n")#-1
vdept = fct2 (vdept)

vgrp = input("Indiquer la lettre de la colonne comportant le groupe\n")#-1
vgrp = fct2 (vgrp)

vpnom = input("Indiquer la lettre de la colonne comportant le prenom\n")#-1
vpnom = fct2 (vpnom)

for x in range(1, (int(nombre_variables)+1)):
	vvar = input("Indiquer la lettre de la colonne comportant la variable "+str(x)+"\n")#)-1)
	vvar = fct2 (vvar)
	d.append(int(vvar))


	
#################################


print ("-----Attention, les mails seront envoyes uniquement aux responsables avec case cochee avec x ou X sur le tableur-----")
print ("-----Tout autre lettre dans la colonne responsable empechera l'envoi du mail-----")
time.sleep(3)

vresp = input("Indiquer la lettre de la colonne cochee indiquant les responsables de groupe\n")
vresp = fct2 (vresp)

vligned = int(input("Numero de la ligne de debut de la liste des destinataires "))

vlignef = int(input("Numero de la ligne de fin de la liste des destinataires "))


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
##### Determiner le fichier excel Faire attention sheet!!!!! 



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

v_sheet = int(input("Entrer le numero de la feuille excel\n"))-1

os.chdir(dossier) 


#########################################################################
# Convertion fichier excel

vnbrligne = ((vlignef+1)-vligned)

read_file = pd.read_excel(contenu,sheet_name=v_sheet,skiprows = vligned-1, header=None)         

read_file.to_csv ("Test.csv",  
                  index = None, 
                  header = True)################uft8 a faire

#print (read_file)##########

df = pd.DataFrame(pd.read_csv("Test.csv")) 
df = df[:vnbrligne]

#print (df)############

vresp = df.columns.values[vresp] 

pd.DataFrame(df.dropna(subset=[vresp], inplace=True)) ######suppression des non-resp
df[vresp] = df[vresp].str.strip()#### suppression espace colonne resp
df = df[(df[vresp].str.match('X'))|(df[vresp].str.match('x'))]#### suppression resp sans X

list = []

for x in range (1,len(df.index)+1):
 list.append(x)

df.index = list
########print (df)

#######################################################################################################
##### Verification

EMAIL_REGEX = re.compile("[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+")
i = 0
		
df2 = df

for x in range (1,len(df.index)+1):
	you = (df.iat[x-1,vmail])
	if str(you) == "nan": ## verif case vide
		print ("\n----Mail manquant groupe "+str(df.iat[x-1,vgrp])+"----\n")
		reponse_mail = int(input("1 pour continuer sans envoyer ce mail\n2 pour arreter le programme apres verification des autres mail\n3 pour entrer l'adresse manuellement \n"))
		if reponse_mail == 1 :
			df2 = df2.drop([x])
			you = 'aaa'
			i = i-1
		if reponse_mail == 2 :
			you = 'aaa'
			i=1
		if reponse_mail == 3 :
			vmail2 = df.columns.values[vmail] 
			reponse_mail2 = input("\n Entrer l'adresse voulue\n")
			df.at[x,vmail2] = reponse_mail2
			you = (df.iat[x-1,vmail])
			
	you = you.strip()
	if not EMAIL_REGEX.match(str(you)):
		print ("----Erreur mail groupe "+str(df.iat[x-1,vgrp])+"----\n")    
		i = i+1

df = df2

list2 = [] ### reconstruction de l'index

for x in range (1,len(df.index)+1):
 list2.append(x)

df.index = list2

fct3 (vdept, 'departement')
fct3 (vpnom, 'prenom')
fct4 (vgrp, 'groupe')

if i == 1 :##########################################################nbr erreur
	sys.exit(0)

###################################################################	
########################## Essai 

essai = int(input("Voulez vous faire un essai sur votre adresse mail ?\n1 : oui\n2 : non\n"))

if essai == 1 :
	you = input("L´adresse mail pour l´essai\n")
	x = random.randint(1,len(df.index)) 
	file_out2 = file_out
	print("Essai avec le groupe "+str(df.iat[x,vgrp]))
	
	for y in range (1, (int(nombre_variables)+1)): ### remplacement variable
		file_out2 = file_out2.replace('variable'+str(y), str(df.iat[x,d[y-1]]))############# a verif
	
	titre2 = ("["+str(df.iat[x,vdept])+"_"+str(df.iat[x,vgrp])+"] "+titre)

	msg = MIMEMultipart()
	msg['Subject'] = titre2
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

essai3 = int(input("Afficher la base de donnee modifiee qui sera traitee par le programme ? \n1 : oui\n2 : non\n"))
if essai3 == 1 :
	print (df) #^^^^^^^^^^^^^^^^^^ affiche en entier
		
essai2 = int(input(str(len(df.index))+" mails seront envoyes \n1 : oui\n2 : non\n"))
if essai2 == 2 :
	sys.exit(0)
	

#######################################################################################################
##### Envoi du mail


for x in range (0,len(df.index)): ### determination adresse mail 
	you = (df.iat[x,vmail])
	you = you.strip()
	file_out2 = file_out
	for y in range (1, (int(nombre_variables)+1)): ### remplacement variable
		file_out2 = file_out2.replace('variable'+str(y), str(df.iat[x,d[y-1]]))
	
	titre2 = ("["+str(df.iat[x,vdept])+"_"+str(df.iat[x,vgrp])+"] "+titre)
	
	msg = MIMEMultipart()
	msg['Subject'] = titre2
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
			myfile.write(str(df.iat[x-1,vdept])+"/"+str(df.iat[x-1,vpnom])+"/"+str(df.iat[x-1,vgrp])+"/"+you+" mail envoye\n")
		print (str(n)+"/"+str(len(df.index)))
		n=n+1
	except smtplib.SMTPException as f:
		print ("Erreur dans un envoi")
		print (f)
		with open("sortie.txt", "a") as myfile:
			myfile.write(str(df.iat[x-1,vdept])+"/"+str(df.iat[x-1,vpnom])+"/"+str(df.iat[x-1,vgrp])+"/"+you+" MAIL NON ENVOYE\n")
		m=m+1


#######################################################################################################

mail.quit()

print (n-1, "mails envoyes,",m,"erreurs")

os.rename("sortie.txt", date+".txt" )
os.remove("Test.csv")

fin = input("Envois termines, appuyer sur une touche pour quitter le programme")
