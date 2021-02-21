import sys

print ('\n1. Creer le fichier_envoi a partir de la base de donnee client \n'
        '2. Uploader sur OneDrive le catalogue et le fichier SAV \n'
        '3. Envoyer les mails \n'
        '4. Quitter le programme\n')

r = input('Entrer 1, 2 ou 3\n')

if r == '1':
    import creation_complet.py

elif r == '2' :
    import od_flow.py

elif r == '3':
    import script.py

elif r == '4':
    sys.exit(0)

else :
    print ('Erreur dans le choix')
