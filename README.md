# poubelle-projet

Développeur : Orie-san

#requirements

Ce programme a été codé sous la version  3.8 de Python et utilise les modules pandas, numpy,  openpyxl, xlsxwriter, datetime et random.



Avant d'exécuter les différentes commandes dans le fichier principal main.py veuillez vous assurez que les fichiers Gestion.xlsx, En_tete_Excel.py, Personne.py,
Poubelle.py, Camion.py, Employee.py, Fonctions_utiles.py et Fonctions_ramassage.py se trouvent dans le même répertoire.

#détails sur les fichiers

Gestion.xlsx est un fichier excel constitué de 4 feuilles de calcul à savoir:

	Employés : qui contient toutes les informations relatives aux Employés

	Poubelles : qui contient toutes les informations relatives aux Poubelles

	Camions : qui contient toutes les informations relatives aux Camions

	Ramassage : qui contient l'historique des missions de ramassage (DATE,HEURE,ID POUBELLE,Localisation,Type de poubelle,Immatriculation du camion en charge, 
	Chauffeur du camion, liste des ramasseurs convoqués pour le ramassage)

En_tête_Excel.py: qui permet de créer les en-têtes du fichier Excel Gestion.xlsx

Personne.py: qui definit les attributs de la classe personne

Employee.py : qui définit les attributs et les méthodes de la classe Employee

Poubelle.py: qui définit les attributs de la classe Poubelle

Camion.py: qui définit les attributs de la classe Camion

Fonctions_utiles.py : qui contient une fonction nécessaire au traitement des informations du fichier Excel Gestion.xlsx

Fonctions_ramassage.py: qui définit les méthodes nécessaires à la simulation d'une opération de ramassage

main.py : permet de tester différentes fonctions pour établir leur bon fonctionnement.
