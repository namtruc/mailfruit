## akoiçassaire


*   Ce programme sert à envoyer les liens catalogues chaque semaine par mail
* Il lit les bases de données et modifie automatiquement chaque mail avec les données de son choix (prénom et lien par exemple)
* Il peut prendre en compte 2 tableurs différents si les liens sont sur un autre fichier que sur celui des ID clients

## Utilisateurs Windows 10

* Installer python : [Sur le store de microsoft](http://www.microsoft.com/fr-fr/p/python-37/9nj46sx7x90p?rtc=1&amp;activetab=pivot:overviewtab)
* Dans la barre de recherche du menu démarrer  taper `cmd` et lancer l´application *invite de commande*
* Dans l’invite de commande taper `pip install pandas xlrd openpyxl html2text`
* Fermer cette invite de commande
* Télécharger le code et l´extraire dans le dossier de son choix
* Ouvrir le dossier mailfruit, précédemment téléchargé, cliquer sur fichier, puis ouvrir Windows PowerShell (cela permet de facilement se placer dans le bon répertoire sans avoir à le faire via une ligne de commande)
* Taper `python script.py` pour lancer le programme

## Mail et spam

* A remplir, voir [https://www.mail-tester.com/](https://www.mail-tester.com/)

## Utilisation du programme

* Aucun fichier n’est modifié ou supprimé par ce programme
* Le fichier associant les groupes au lien onedrive peut etre à part du fichier client, le programme proposera de les merger
* Nécessite de passer par un serveur gmail, vous devez avoir les identifiants pour utiliser ce programme
* Le mail peut être récupérer soit ~~en l´envoyant à l´adresse gmail (il sera récupéré par la suite)~~  (à éviter, source d´erreur) soit en l´enregistrant sous format html (certains clients de messagerie proposent cette option, me contacter si besoin)
* Le mail doit être rédigé avec les termes variable1, variable2, etc… à la place des mots ou des liens (attention à la redirection) qui devront être remplacés par les valeurs présente dans les fichiers excels. L´ordre importe peu, 9 variables max
* Si possible mettre tous les fichiers (html, piece jointe, tableur) dans le dossier utilisateur


## Catalogue

 - En cours
 - Le script 'commande' permet, à partir de la base de données client et du fichier commande de générer les fichier à mettre dans OneDrive avec le nom correspondant au groupe et en pré-remplissant la case département pour débloquer les prix

## Contact

 - Retwiin via le serveur discord du groupe ou par mail vincent_fds@gmx.com
