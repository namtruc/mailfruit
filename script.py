import smtplib
import sys
import re
import os
import datetime
import pandas as pd
import random
import time
import getpass

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
			return (1)

def fct4 (vide, mot): ###fct 3 pour vgrp
	vide = df.columns.values[vide]
	check = pd.isna(df[vide])
	for x in range ((len(df.index))):
		if check[x+1] == True :
			print ("Erreur case vide colonne "+mot+" departement "+str(df.iat[x,vdept]) )
			return (1)	
					
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
dossier_usr = dossier_python + '/fichiers_utilisateur'
os.chdir(dossier_usr)

date1 = datetime.datetime.today()
date2 = date1.isocalendar()
jour = date1.weekday()
semaine = date2[1]


#######################################################################################################
#####Merger excel

while 1:
	reponse = input("Taper 1 pour copier automatiquement les liens du tableur suivi-envoi avec la base de donnees clients (un nouveau fichier sera cree)\nTaper 2 pour passer cette etape\n")
	if reponse=='1':
		import script2
		break
	elif reponse =='2':
		break
	else:
		print ("Choix incorrect !")  


#######################################################################################################
#####Connection

mail = smtplib.SMTP(srv, prt)
mail.ehlo()
mail.starttls()

mdp=getpass.getpass("Entrez le mot de passe pour "+usr+"\n")
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

input("\nAppuyer sur Entree pour choisir le fichier mail au format html\n")

Tk().withdraw() 
contenu = askopenfilename()

with open(contenu, 'r', encoding="utf8") as file_in :
  file_out = file_in.read()


#######################################################################################################
##### Definir titre

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
		input("\n Appuyer sur Entree pour choisir la piece jointe "+str(x+1))
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

print ("\n--Attention les lignes avec la case departement vide seront suprimees automatiquement")
time.sleep (1)

reponse = input("\nTaper 1 si le script doit prendre en compte l'ensemble des lignes des destinataires.\nTaper 2 pour indiquer quelles lignes doivent etre prises en compte\n")

if reponse=='2':
	vligned = int(input("Numero de la ligne de debut de la liste des destinataires "))
	vlignef = int(input("Numero de la ligne de fin de la liste des destinataires "))
else:
	vligned = 1
	vlignef = 1
	
	
#######################################################################################################
##### Definir les colonnes


vmail = parser.get('colonnes', 'mail')
vdept = parser.get('colonnes', 'dept')
vgrp = parser.get('colonnes', 'grp')
vpnom = parser.get('colonnes', 'pnom')

print ("-----Attention, une erreur de frappe peut entrainer le crash du programme-----")
time.sleep(3)

print ("Lettre colonne mail "+vmail)
print ("Lettre colonne departement "+vdept)
print ("Lettre colonne groupe "+vgrp)
print ("Lettre colonne prenom "+vpnom)

print ("\n----Verifier que les informations ci dessus sont correctes----")
time.sleep(3)

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
		break
		
	elif reponse == '2' :
		
		vpnom = fct2 (vpnom)
		vmail = fct2 (vmail)
		vdept = fct2 (vdept)
		vgrp = fct2 (vgrp)
		break

	else :

		print ("Choix incorrect !")
		break
		
#######################################################################################################
##### Definir les variables


d = []

nombre_variables = input("Entrer le nombre de variables presentes dans le texte du mail, max 9\n")

for x in range(1, (int(nombre_variables)+1)):
	vvar = input("Indiquer la lettre de la colonne comportant la variable "+str(x)+"\n")
	vvar = fct2 (vvar)
	d.append(int(vvar))


	
#######################################################################################################
##### Definir les responsables de groupe


print ("-----Attention, les mails seront envoyes uniquement aux responsables avec case cochee avec x ou X sur le tableur-----")
print ("-----Tout autre lettre dans la colonne responsable empechera l'envoi du mail-----")
time.sleep(3)

vresp = input("Indiquer la lettre de la colonne cochee indiquant les responsables de groupe\n")
vresp = fct2 (vresp)

	
#########################################################################
# Convertion fichier excel

vnbrligne = ((vlignef+1)-vligned)

read_file = pd.read_excel(contenu1,sheet_name=v_sheet,skiprows = vligned-1, header=None)         

read_file.to_csv ("Test.csv",  
                  index = None, 
                  header = True)################uft8 a faire

#print (read_file)##########

df = pd.DataFrame(pd.read_csv("Test.csv")) 
vrespa = df.columns.values[vresp] 
vdepta = df.columns.values[vdept] 

if vlignef != 1 :
	df = df[:vnbrligne]
	pd.DataFrame(df.dropna(subset=[vdepta], inplace=True))
else :
	pd.DataFrame(df.dropna(subset=[vdepta], inplace=True))

#print (df)############

pd.DataFrame(df.dropna(subset=[vrespa], inplace=True)) ######suppression des non-resp
df[vresp] = df[vrespa].str.strip()#### suppression espace colonne resp
df = df[(df[vrespa].str.match('X'))|(df[vrespa].str.match('x'))]#### suppression resp sans X

list = []

for x in range (1,len(df.index)+1):
 list.append(x)

df.index = list

os.remove("Test.csv")


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
			#you = 'erreur@gmail.com'
			#i = i-1
		if reponse_mail == 2 :
			#you = 'erreur@gmail.com'
			i= i+1
		
		if reponse_mail == 3 :
			vmail2 = df.columns.values[vmail] 
			
			while True :
				reponse_mail2 = input("\n Entrer l'adresse voulue\n")
				if not EMAIL_REGEX.match(reponse_mail2):
					print ("Erreur redaction mail")
				else :
					break
					
			df.at[x,vmail2] = reponse_mail2
			you = (df.iat[x-1,vmail])
	
	else :	
		if not EMAIL_REGEX.match(str(you)):
			print ("----Erreur redaction mail groupe "+str(df.iat[x-1,vgrp])+"----\n")    
			i = i+1
	you = you.strip()

df = df2

list2 = [] ### reconstruction de l'index

for x in range (1,len(df.index)+1):
 list2.append(x)

df.index = list2



if fct3 (vdept, 'departement') == 1 :
	i = i+1

if fct3 (vpnom, 'prenom') == 1 :
	i = i+1

if fct4 (vgrp, 'groupe') == 1 :
	i = i+1

if i > 0 :
	print (str(i)+" erreurs")
	time.sleep (3)
	sys.exit(0)


###################################################################	
########################## Essai 

essai = int(input("Voulez vous faire un essai sur votre adresse mail ?\n1 : oui\n2 : non\n"))

if essai == 1 :
	you = input("L´adresse mail pour l´essai\n")
	x = random.randint(1,len(df.index)) 
	file_out2 = file_out
	print("Essai avec le groupe "+str(df.iat[x,vdept]+" "+str(df.iat[x,vgrp])))
	
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
	print (df) 


essai2 = int(input(str(len(df.index))+" mails seront envoyes \n1 : oui\n2 : arreter le programme\n"))
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

	os.chdir(dossier_usr)

	try:
		mail.sendmail(me, you, msg.as_string())         
		with open("sortieincomplete.txt", "a") as myfile:
			myfile.write(str(df.iat[x-1,vdept])+"/"+str(df.iat[x-1,vpnom])+"/"+str(df.iat[x-1,vgrp])+"/"+you+" mail envoye\n")
		print (str(n)+"/"+str(len(df.index)))
		n=n+1
	except smtplib.SMTPException as f:
		print ("Erreur dans un envoi")
		print (f)
		with open("sortieincomplete.txt", "a") as myfile:
			myfile.write(str(df.iat[x-1,vdept])+"/"+str(df.iat[x-1,vpnom])+"/"+str(df.iat[x-1,vgrp])+"/"+you+" MAIL NON ENVOYE\n")
		m=m+1


#######################################################################################################

mail.quit()

print (n-1, "mails envoyes,",m,"erreurs")

os.chdir(dossier_usr)
os.rename("sortieincomplete.txt", date+".txt" )

print ("fichier"+date+".txt cree dans le dossier fichier utilisateur")

fin = input("Envois termines, appuyer sur une touche pour quitter le programme")
