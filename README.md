

<h2 id="akoiçassaire">akoiçassaire</h2>
<ul>
<li>Ce programme sert à envoyer les liens catalogues chaque semaine par mail</li>
<li>Il lit les bases de données et modifie automatiquement chaque mail avec les données de son choix (prénom et lien par exemple)</li>
<li>Il peut prendre en compte 2 tableurs différents si les liens sont sur un autre fichier que sur celui des ID clients</li>
</ul>
<h2 id="utilisateurs-windows-10">Utilisateurs Windows 10</h2>
<ul>
<li>Installer python : <a href="http://www.microsoft.com/fr-fr/p/python-37/9nj46sx7x90p?rtc=1&amp;activetab=pivot:overviewtab">www.microsoft.com/fr-fr/p/python-37/9nj46sx7x90p?rtc=1&amp;activetab=pivot:overviewtab</a></li>
<li>Dans la barre de recherche du menu démarrer  taper <code>cmd</code> et lancer l´application <em>invite de commande</em></li>
<li>Dans l’invite de commande taper <code>pip install pandas xlrd openpyxl html2text</code></li>
<li>Fermer cette invite de commande</li>
<li>Télécharger le code et l´extraire dans le dossier de son choix</li>
<li>Ouvrir le dossier mailfruit, précédemment téléchargé, cliquer sur fichier, puis ouvrir Windows PowerShell (cela permet de facilement se placer dans le bon répertoire sans avoir à le faire via une ligne de commande)</li>
<li>Taper <code>python script.py</code> pour lancer le programme</li>
</ul>
<h2 id="mail-et-spam">Mail et spam</h2>
<ul>
<li>A remplir, voir <a href="https://www.mail-tester.com/">https://www.mail-tester.com/</a></li>
</ul>
<h2 id="utilisation-du-programme">Utilisation du programme</h2>
<ul>
<li>Aucun fichier n’est modifié ou supprimé par ce programme</li>
<li>Le fichier associant les groupes au lien onedrive peut etre à part du fichier client, le programme proposera de les merger</li>
<li>Nécessite de passer par un serveur gmail, vous devez avoir les identifiants pour utiliser ce programme</li>
<li>Le mail <s>peut être récupérer soit en l´envoyant à l´adresse gmail (il sera récupéré par la suite)</s>  (à éviter, source d´erreur) soit en l´enregistrant sous format html (certains clients de messagerie proposent cette option, me contacter si besoin)</li>
<li>Le mail doit être rédigé avec les termes variable1, variable2, etc… à la place des mots ou des liens (attention à la redirection) qui devront être remplacés par les valeurs présente dans les fichiers excels. L´ordre importe peu, 9 variables max</li>
<li>Si possible mettre tous les fichiers (html, piece jointe, tableur) dans le d## akoiçassaire

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

## Mail et spam

 - A remplir, voir https://www.mail-tester.com/

## Utilisation du programme

* Aucun fichier n'est modifié ou supprimé par ce programme
* Le fichier associant les groupes au lien onedrive peut etre à part du fichier client, le programme proposera de les merger
* Nécessite de passer par un serveur gmail, vous devez avoir les identifiants pour utiliser ce programme
* Le mail ~~peut être récupérer soit en l´envoyant à l´adresse gmail (il sera récupéré par la suite)~~  (à éviter, source d´erreur) soit en l´enregistrant sous format html (certains clients de messagerie proposent cette option, me contacter si besoin)
* Le mail doit être rédigé avec les termes variable1, variable2, etc... à la place des mots ou des liens (attention à la redirection) qui devront être remplacés par les valeurs présente dans les fichiers excels. L´ordre importe peu, 9 variables max
* Si possible mettre tous les fichiers (html, piece jointe, tableur) dans le dossier fichiers_usr present dans le dossier mailfruit
* Mettre les tableurs en xls ou xlsx 
* Pas d´inquiétude, il y a la possibilité d´envoyer un mail d´essai à l´adresse de son choix avant de finaliser la procédure 
* Des fichiers tests sons dispo dans le dossier *test* avec des bases de donnée fictive et des exemples de mail
* **Attention à bien lire ce que le programme affiche pour éviter d´envoyer 80 mauvais mails**

## Catalogue

 - En cours
 - Le script 'commande' permet, à partir de la base de données client et du fichier commande de générer les fichier à mettre dans OneDrive avec le nom correspondant au groupe et en pré-remplissant la case département pour débloquer les prix

## Contact

 - Retwiin via le serveur discord du groupe ou par mail vincent_fds@gmx.com

<!--stackedit_data:
eyJoaXN0b3J5IjpbMTc4NDk5NTE1MywtNTQ5MzA4OTIyXX0=
-->