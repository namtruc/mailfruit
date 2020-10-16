## akoiçassaire

 - Ce programme sert à envoyer les liens catalogues chaque semaine par mail
 - Il lit les bases de données et modifie automatiquement chaque mail avec les données de son choix (prénom et lien par exemple)
 - Il peut prendre en compte 2 tableurs différents si les liens sont sur un autre fichier que sur celui des ID clients

## Utilisateurs Windows 10

* Installer python : www.microsoft.com/fr-fr/p/python-37/9nj46sx7x90p?rtc=1&activetab=pivot:overviewtab
* Dans la barre de recherche du menu démarrer  taper `cmd` et lancer l´application *invite de commande*
* Dans l'invite de commande taper ` pip install pandas xlrd openpyxl html2text`
* Fermer cette invite de commande
* Télécharger le code et l´extraire dans le dossier de son choix
* Ouvrir le dossier mailfruit, précédemment téléchargé, cliquer sur fichier, puis ouvrir Windows PowerShell (cela permet de facilement se placer dans le bon répertoire sans avoir à le faire via une ligne de commande)
* Taper ` python script.py ` pour lancer le programme


## Utilisation du programme

* Aucun fichier n'est modifié ou supprimé par ce programme
* Nécessite de passer par un serveur gmail, vous devez avoir les identifiants pour utiliser ce programme
* Le mail peut être récupérer soit en l´envoyant à l´adresse gmail (il sera récupéré par la suite) soit en l´enregistrant sous format html (certains clients de messagerie proposent cette option, me contacter si besoin)
* Le mail doit être rédigé avec les termes variable1, variable2, etc... à la place des mots ou des liens (attention à la redirection) qui devront être remplacés par les valeurs présente dans les fichiers excels. L´ordre importe peu, 9 variables max
* Si possible mettre tous les fichiers (html, piece jointe, tableur) dans le dossier fichiers_usr present dans le dossier mailfruit
* Mettre les tableurs en xls ou xlsx 
* Pas d´inquiétude, il y a la possibilité d´envoyer un mail d´essai à l´adresse de son choix avant de finaliser la procédure 
* Des fichiers tests sons dispo dans le dossier *test* avec des bases de donnée fictive et des exemples de mail
* **Attention à bien lire ce que le programme affiche pour éviter d´envoyer 80 mauvais mails**

## Contact

 - Retwiin via le serveur discord du groupe ou par mail vincent_fds@gmx.com
