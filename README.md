# ExcelExport
Used to export excel sheets to pdf files

- Récupération d'un listening des feuilles

- Prévisualisation des feuilles pour l'utilisateur :
	- Navigation flèche droite-gauche sur le coté
	- Affichage du Nom de la feuille(liste déroulante) pour la sélection -> Sommet
	- Petite case dans un coin permettant de sélectionner la feuille -> Ajout à la liste prévisualisation
	- Affichage des feuilles sélectionnées en bas de la fenêtre (Slider prévisualisation), sur le clic d'une des feuilles -> affichage dans la grande prévisualisation

- Export de chaque feuille dans un fichier pdf :
	- Boucle sur la liste des feuilles
	- Génération d'un fichier pdf pour chacune des feuilles
	- Convention de nommage incluant le nom de la feuille : 
		- {year}.{month}.{day}_{SheetName}_Custom.pdf
		- Menu options -> Description avec explications des paramètres possibles
		- Bloquage de modification d'extension (Extension dans un label juste après la textbox)


- Menu :
	- File :
		- Open
	- Settings :
		- File name

- Ouverture de l'application :
	- Affichage d'un gros "bouton" open file