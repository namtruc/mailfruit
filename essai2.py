import smtplib
import sys
import re
import os
import datetime
import pandas as pd

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# me == my email address
# you == recipient's email address

me = 'kildek@caramail.fr'
#you = 'vincentd@gmx.us'
n = 0
date = datetime.datetime.today().strftime('%d.%m.%y-%Hh%M')
dossier_python = os.getcwd()



#######################################################################################################
#####Connection

mail = smtplib.SMTP('mail.gmx.com', 587)
mail.ehlo()
mail.starttls()
mail.login('vincentd@gmx.us', 'eyno8smj1g')#############################################
usr = 'vincentd@gmx.us'#################################################################
#usr = input("Entrez le mail utilisateur\n")
#mdp = input("Entrez le mot de passe\n"))
#mail.login(usr, mdp)

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
##### Operations fichier HTML

dossier = '.'

while 1:
    reponse = input("Taper 1 pour choisir le fichier HTML dans le repertoire actuel (recommande)\nTaper 2 pour choisir un autre repertoire\n")
    if reponse=='1':
        break
    elif reponse =='2':
        dossier = input("Entrer le chemin du dossier\n")
        break
    else:
        print ("Choix incorrect !")   
        
items = os.listdir(dossier)############## tester avec windows
newlist = []
for names in items:
    if names.endswith(".html"):
        newlist.append(names)
print ("Contenu du dossier\n", newlist)
contenu = input("Entrer le nom complet du fichier HTML contenant le texte brut du mail\n")

os.chdir(dossier)

with open(contenu, 'r') as file_in :
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

dossier = '.'

d = []

nombre_variables = input("Entrer le nombre de variables (nom, lien, etc...)\n")
vmail = int(input("Indiquer le numero de la colonne comportant les adreses mail\n"))-1

for x in range(1, (int(nombre_variables)+1)):
	d.append(int(input("Numero de la colonne avec la variable "+str(x)+"\n"))-1)
	
vligned = int(input("Numero de la ligne de debut de la liste des destinataires "))-1
vlignef = int(input("Numero de la ligne de fin de la liste des destinataires "))-1


#######################################################################################################
##### Determiner le fichier excel

while 1:
    reponse = input("Taper 1 pour choisir le fichier excel ou libreoffice dans le repertoire actuel (recommande)\nTaper 2 pour choisir un autre repertoire\n")
    if reponse=='1':
        break
    elif reponse =='2':
        dossier = input("Entrer le chemin du dossier\n")
        break
    else:
        print ("Choix incorrect !")   
        
items = os.listdir(dossier)############## tester avec windows
newlist2 = []
for names in items: ###### tester les noms
    if names.endswith(".ods"):
        newlist2.append(names)
    elif names.endswith(".xlsx"):
        newlist2.append(names)
    elif names.endswith(".xls"):
        newlist2.append(names)
print ("Contenu du dossier\n", newlist2)

contenu = input("Entrer le nom complet du fichier contenant la base de donnee\n")

os.chdir(dossier) 

#read_file = pd.read_excel(contenu)
df = pd.DataFrame(pd.read_excel(contenu))
	

#######################################################################################################
##### Creation du mail
  
  
EMAIL_REGEX = re.compile("[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+")
i = 0

for x in range (vligned, vlignef+1):
	you = (df.iat[x-1,vmail])
	if not EMAIL_REGEX.match(str(you)):
		print ("erreur mail ligne "+str(x))      
		i = 1

if i == 1 :
	sys.exit(0)


for x in range (vligned, vlignef+1):
	you = (df.iat[x-1,vmail])
	file_out2 = file_out
	for y in range (1, (int(nombre_variables)+1)):
		file_out2 = file_out2.replace('variable'+str(y), str(df.iat[x-1,d[y-1]]))
		
	msg = MIMEMultipart('alternative')
	msg['Subject'] = titre
	msg['From'] = me
	msg['To'] = you

	html = file_out2
	
	html_part = MIMEText(html, "html")
	msg.attach(html_part)

	mail.sendmail(me, you, msg.as_string())         
	with open("sortie.txt", "a") as myfile:
		myfile.write(you+" Successfully sent email\n")
	n=n+1
	#except mail.Exception:#######erreur
		#print ("Error: unable to send email")
		#n=n+1


#######################################################################################################

mail.quit()

print (n, "mails envoyes")

os.rename("sortie.txt", date+".txt" )


#Essai avec tableau complete
#Piece jointe ?
#Connection ?
#Lien Hypertexte ?
#essai wind


##### Convertir lettre en chiffre
##### Renvoyer le resultat dans un fichier
##### Message x envoye, x non envoye
##### Proposer essai avec mail utilisateur
##### Faire fichier parametres (avertissement gmail)
##### Essayer avec differant types de connection
##### Tester avec 0 parametres
