### TODO
# enregistrer les fichiers dans un rep temp
# Proposer creation dossier et recup ID
# verifier list-error
# Creation Dossier archive ?
# Fct remplissage a revoir
# determiner num semaine
# erreur pre-remplissage
# gitignore + base 64
# ajout config.py a config.ini
# si erreur upload ne pas supprimer et copier dans fichier utilisateur
# affucher progression upload

import pyperclip
import config
import json
import time
import os
import webbrowser
import pandas as pd
import requests

from microsoftgraph.client import Client
from adal import AuthenticationContext
from os import path as os_path 
from os import listdir
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from shutil import copyfile

dossier_python = os_path.abspath(os_path.split(__file__)[0])
dossier_usr = dossier_python + '/fichiers_utilisateur'

####################dossier_usr = dossier_python + '/s49'

#AUTHORITY_URL = 'https://login.microsoftonline.com/common'
#RESOURCE = 'https://graph.microsoft.com'


def device_flow_session(client_id, auto=True):
    
    ctx = AuthenticationContext(config.AUTHORITY_URL, api_version=None)
    device_code = ctx.acquire_user_code(config.RESOURCE,
                                        client_id)

    # display user instructions
    if auto:

        print(f'Le code {device_code["user_code"]} a ete copie dans le presse-papier '
              f' et votre navigateur est en train d´ouvrir la page {device_code["verification_url"]}. '
              'Coller le code pour se connecter.')
        pyperclip.copy(device_code['user_code']) # copy user code to clipboard
        webbrowser.open(device_code['verification_url']) # open browser
        
    else:
        print(device_code['message'])

    token_response = ctx.acquire_token_with_device_code(config.RESOURCE,
                                                        device_code,
                                                        client_id)
    if not token_response.get('accessToken', None):

        return None

    session = requests.Session()
    session.headers.update({'Authorization': f'Bearer {token_response["accessToken"]}'})

    return session

########################
######################
#####################


def creation_nom(sav):

    dico = dict()
    #dico2 = dict()
    
    for x in range (len(df_envoi.index)):

        if sav == False :

            dico[x] = ('cmd_'+str(df_envoi.iat[x,vdept])+'_'+str(df_envoi.iat[x,vgrp])+'_S'+semaine+'.xlsx')	

        if sav == True :

            dico[x] = ('sav_'+str(df_envoi.iat[x,vdept])+'_'+str(df_envoi.iat[x,vgrp])+'_S'+semaine+'.xlsx')	

    return dico
 
###

def remplissage (index, lien,sav_lien):

    #print('rempl '+str(sav_lien))
    if sav_lien == True :

        df_envoi.at[index, 'sav'] = lien

    else :

        df_envoi.at[index, 'lien'] = lien

###

def fct_copy(dst,src):

    copyfile(src, dst)


#################
################
###############

def fct_upload (name_cat,name_sav, repeat, nom_dossier,index,SAV):
    
    print ("\n")
    print (name_cat)
    headers = {'Content-Type' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}
    data = open(name_cat, 'rb')
    r = session.put('https://graph.microsoft.com/v1.0/me/drive/root:/'+nom_dossier+'/'+name_cat+':/content', data=(data), headers=headers)

    fct_retour (r, 'Upload', False, repeat,name_cat,name_sav,False)


    if SAV == True: 
        
        print (name_sav)
        #headers = {'Content-Type' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}
        data = open(name_sav, 'rb')

        r = session.put('https://graph.microsoft.com/v1.0/me/drive/root:/'+nom_dossier+'/SAV/'+name_sav+':/content', data=(data), headers=headers)
        
        fct_retour (r, 'Upload', False, repeat,name_cat,name_sav,True)


###

def fct_share(name_cat, name_sav, repeat,nom_dossier,index,sav_lien):

    m = 0
    #input (str(SAV))

    if sav_lien == True: 

        #print (name+'%%%%%%%%')
        #print ('ok')
        r = session.post('https://graph.microsoft.com/v1.0/me/drive/root:/'+nom_dossier+'/SAV/'+name_sav+':/createlink',json = {"type": "edit", "scope": "anonymous"})

        #print (str(r))
        fct_retour (r, 'Partage', True, repeat,name_cat,name_sav, True)

    else :

        #print (name+'%%%%%%%%')
        r = session.post('https://graph.microsoft.com/v1.0/me/drive/root:/'+nom_dossier+'/'+name_cat+':/createlink',json = {"type": "edit", "scope": "anonymous"})

        fct_retour (r, 'Partage', True, repeat,name_cat,name_sav, False)

###

def fct_retour(r, text, share, repeat,name_cat,name_sav,sav_lien):

    n = 0
    if sav_lien == True :
        name = name_sav
    else :
        name = name_cat

    while True :
    
           if str(r) == '<Response [200]>' or str(r) == '<Response [201]>':

               if share == True :

                    print (text +' ok')
                    d = (r.json().get('link'))
                    link = d.get('webUrl')
                    remplissage (index,link,sav_lien)

                    break
    
               else :

                    print (text +' ok')
                    fct_share(name_cat, name_sav, repeat,nom_dossier,index,sav_lien)
                    os.remove (name)
    
                    break
    
           elif n < 2  and repeat == True:
    
               print (str(r))
               print ('Erreur, attendre 10s')
               time.sleep(10)
               n = n+1
    
           elif n >= 2  and repeat == True:
    
               print (str(r))
               print ('ERREUR '+text+name)
               list_error.append('Erreur '+text+name)
    
               break
    
           else :

               print (str(r)) 
               print ('Erreur non defenitive '+text+name) 
               dico_nom_error[index] = name_cat
    
               break



###################################################################################################
##### Determiner le fichier excel envoi


os.chdir (dossier_python)

input("\nAppuyer sur Entree pour choisir le Fichier_envoi \n")

Tk().withdraw() 
contenu_envoi = askopenfilename()
#contenu_envoi = 'ex_envoi.xlsx'

read_file = pd.read_excel(contenu_envoi)         

read_file.to_csv ("Test.csv",  
                  index = None, 
                  header = True)################uft8 a faire

df_envoi = pd.read_csv('Test.csv')
os.remove('Test.csv')

df_envoi['lien']=''
df_envoi['sav']=''

###################################################################################################
##### Determiner le fichier excel catalogue et sav


input("\nAppuyer sur Entree pour choisir le Fichier catalogue\n")

Tk().withdraw() 
contenu_catalogue = askopenfilename()


r = int(input('1. Uploader en meme temps le fichier SAV\n2. Uploader uniquement le fichier catalogue\n'))

if r == 1:

    input("\nAppuyer sur Entree pour choisir le Fichier Sav\n")

    Tk().withdraw() 
    fichier_sav = askopenfilename()

    SAV = True

elif r == 2:

    SAV = False

else :
            
    print('Erreur dans le choix')



#contenu_catalogue = 'ex_catalogue.xlsx'
#contenu_catalogue = 'ex_toto.xlsx'


#############################################################################
##################


vgrp = 1
vdept = 0
semaine = input('Entrer le numero de semaine\n')

session = device_flow_session(config.CLIENT_ID)

list_error = []
#list_error_share = []

dico_cat = dict()
dico_sav = dict()
dico_nom_error= dict()
#dico_nom_error_sav = dict()

dico_cat = creation_nom(False)
dico_sav = creation_nom(SAV)

nom_dossier = input('Entrer un nom de dossier pour l´upload (pre-existant ou non)\n')

#print (dico_nom)
#input ('TTT')

for index, nom in dico_cat.items():

    fct_copy(nom, contenu_catalogue)

    if SAV == True:

        fct_copy(dico_sav[index], fichier_sav)

    fct_upload (nom,dico_sav[index],False,nom_dossier,index,SAV)

for index, nom in dico_nom_error.items():

    fct_upload (nom,dico_sav[index],True,nom_dossier,index,SAV)


###########################################################################
############ Enregistrement

#print (list_error)

os.chdir (dossier_usr)
df_envoi.to_excel(contenu_envoi, index = False, header = True)
print ('\nFichier_envoi mise a jour')

if list_error :

    os.chdir(dossier_usr)
    myfile = open('Erreurs_OneDrive_s'+str(semaine)+'.txt', 'w')

    print ('\nErreurs au cours de l´upload, fichier erreur enregistre dans le dossier Fichiers_utilisateur\n')
    
    print (('Verifier les fichiers cmd ET sav\n'), file = myfile)
    input ('Appuyer sur Entree pour continuer')

    for element in list_error :

        print (element)
        print (element, file = myfile)
        

input = ('Appuyer sur Entree pour revenir au menu')



#os.chdir(dossier_usr)
#liste = os.listdir()###################
#list_error = ()
#
#df = pd.DataFrame(columns = ['dpt','grp','lien'])
#
#nom_dossier = 's00'##################
#
##liste = fct_creation_liste_nom(fichier_complet)
#for name in liste:
#    fct_upload (name,True,nom_dossier)

################## ajout fct verif




#os.chdir (dossier_python)
#df.to_excel('Fichier_envoi.xlsx', index =False)




#r = session.post('https://graph.microsoft.com/v1.0/me/drive/root/children', json = {"name": "New Folder", "folder": { }, "@microsoft.graph.conflictBehavior": "rename"})
#
#json_object = json.loads(r.text) 
#print(json.dumps(json_object, indent = 1))
#input()

#r = session.get('https://graph.microsoft.com/v1.0/me/drive/root:/Test:/children')
#d = r.json()

#print (r.json(id))
#d = r.json()
#for key, value in r.json().items():
#    print (key)
#    print (value)

#items = r['values']
#print (r)
#json_object = json.loads(r.string) 
#print (json_object)
#print(json.dumps(json_object, indent = 1))

#input ()



