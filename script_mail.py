# -*- coding: utf-8 -*-

import email
import imaplib
import re
import os
import getpass

from os import path as os_path
from configparser import ConfigParser  

dossier_python = os_path.abspath(os_path.split(__file__)[0])
dossier_usr = dossier_python + '/fichiers_utilisateur'

os.chdir(dossier_python)

parser = ConfigParser() 
parser.read('configuration.ini')

srv = parser.get('settings', 'srv_imap')
usr = parser.get('settings', 'username')

os.chdir(dossier_usr)

list_sender = []
list_title = []
list_id = []

mdp=getpass.getpass("Entrez le mot de passe pour "+usr+"\n")

mail = imaplib.IMAP4_SSL(srv)
mail.login(usr, mdp)
mail.select('inbox')


status, data = mail.search(None, 'ALL')
mail_ids = []

for block in data:
    mail_ids += block.split()


for i in mail_ids:
    status, data = mail.fetch(i, '(RFC822)')
    for response_part in data:
        
        if isinstance(response_part, tuple):
            message = email.message_from_bytes(response_part[1])
            mail_from = message['from']
            mail_subject = message['subject'] 
            list_sender.append(mail_from)
            list_title.append(mail_subject)
            list_id.append(i)
           
l = len(list_sender)
m=len(list_sender)


if m == 0:
	print ("Boite mail vide")

if m > 2:
  for x in range (1,4):
    print (str(x)+" "+list_sender[m-x])
    print (' ',list_title[m-x])

if m == 2:
  for x in range (1,3):
    print (str(x)+" "+list_sender[m-x])
    print (' ',list_title[m-x])
    
if m == 1:
  for x in range (1,2):
    print (str(x)+" "+list_sender[m-x])
    print (' ',list_title[m-x])
    
y = int(input("Entrer le numero du mail\n"))
y = m-y

j= mail_ids[y]

status, data = mail.fetch(j, '(RFC822)')
for response_part in data:
   if isinstance(response_part, tuple):
            message = email.message_from_bytes(response_part[1])
            mail_from = message['from']
            mail_subject = message['subject']
            
            if message.is_multipart():
                mail_content = ''
                for part in message.get_payload():          
					     
                    if part.get_content_type() == 'text/html':
                        body = part.get_payload(decode=True).decode() 
                        with open('sortie_mail.html', "w") as myfile:
	                        myfile.write(body)
            else:
                body = message.get_payload(decode=True).decode()
                with open("sortie_mail.html", "w") as myfile:
	                myfile.write(body)



