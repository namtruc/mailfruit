### TODO
# Proposer creation dossier et recup ID
# verifier list-error
# Creation Dossier archive ?
# Fct remplissage a revoir
# determiner num semaine
# erreur pre-remplissage
# gitignore + base 64
# ajout config.py a config.inimport requests 
# si erreur upload ne pas supprimer et copier dans fichier utilisateur

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


def creation_nom():

    dico = dict()
    #dico2 = dict()
    
    for x in range (len(df_envoi.index)):

        dico[x] = ('cmd_'+str(df_envoi.iat[x,vdept])+'_'+str(df_envoi.iat[x,vgrp])+'_S'+semaine+'.xlsx')	
    	#dico.update( {df_envoi.iat[x,vgrp] : df_envoi.iat[x,vdept]} )
        #dico2[x] = ''

    #print (dico)	
    #print (len(df.index))
    
        #for key,value in dico.items() : 

    return dico
 

def fct_upload (name, repeat, nom_dossier,index):
    
    n = 0
    headers = {'Content-Type' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}
    data = open(name, 'rb')
    r = session.put('https://graph.microsoft.com/v1.0/me/drive/root:/'+nom_dossier+'/'+name+':/content', data=(data), headers=headers)

    while True :

        if str(r) == '<Response [200]>' or str(r) == '<Response [201]>':

            print (name + ' Upload ok')
            ### fct_remplissage_cat
            fct_share(index, nom_dossier,name)
            #list_error.append('Ca marche ')###################
            os.remove (name)

            break

        elif n < 2  and repeat == True:

            print (str(r))
            print ('Erreur, attendre 10s')
            time.sleep(10)
            n = n+1

        elif n >= 2  and repeat == True:

            print (str(r))
            print ('ERREUR UPLOAD '+name)
            list_error.append('Erreur Upload '+name)

            break

        else :

            print ('Erreur non defenitive '+name) 
            dic_nom_error[index] = name

            break


def fct_share(index,nom_dossier,name):
    m = 0

    while True:

        r = session.post('https://graph.microsoft.com/v1.0/me/drive/root:/'+nom_dossier+'/'+name+':/createlink',json = {"type": "edit", "scope": "anonymous"})

        if str(r) == '<Response [200]>' or str(r) == '<Response [201]>':

           print (name + ' Partage ok')
           d = (r.json().get('link'))
           link = d.get('webUrl')
           remplissage (index,link)
           #list_error_share.append('ca marche')###################

           break 

        elif m>= 2 :

            print (str(r))
            print ('ERREUR PARTAGE '+name)
            list_error.append('Erreur partage '+name)

            break

        else :

            print (str(r))
            print ('Erreur attendre 10s')
            m = m+1
            time.sleep (10)


def remplissage (index, lien):

    #input ('lien : '+lien)
    df_envoi.at[index, 'lien'] = lien
    #print (df_envoi.loc[index].at['lien'] )
    #input ()



def fct_copy(dst,src):

    copyfile(src, dst)


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


###################################################################################################
##### Determiner le fichier excel catalogue 


input("\nAppuyer sur Entree pour choisir le Fichier catalogue\n")

Tk().withdraw() 
contenu_catalogue = askopenfilename()
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
dico_nom = dict()
dico_nom_error = dict()

dico_nom = creation_nom()

nom_dossier = input('Entrer un nom de dossier pour l´upload (pre-existant ou non)\n')

#print (dico_nom)
#input ('TTT')

for index, nom in dico_nom.items():

    #print (nom+ str(contenu_catalogue))
    #input ('verif')
    fct_copy(nom, contenu_catalogue)
    fct_upload (nom,False,nom_dossier,index)

for index, nom in dico_nom_error.items():
    fct_upload (nom,True,nom_dossier,index)


###########################################################################
############ Enregistrement

#print (list_error)

os.chdir (dossier_python)
df_envoi.to_excel('Fichier_envoi_s'+str(semaine)+'.xlsx', index = False, header = True)

if list_error :

    os.chdir(dossier_usr)
    myfile = open('Erreurs_OneDrive_s'+str(semaine)+'.txt', 'w')

    print ('\nErreurs au cours de l´upload, fichier erreur enregistre dans le dossier Fichiers_utilisateur\n')
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



