# concours
Permet de générer un excel avec les parties qualificatives d'un concours de palet

Le programme genere autant de fichier excel que d'équipe potentielle.
Ce nombre d'équipe est en dur dans le main (à externaliser) : Nombre minimum et nombre maximum.

Le fichier excel généré est une copie du fichier modele (src/main/resources/com/elbedd/palet/excel/modele.xls)

Le fichier est compléter par les lignes correspondant aux équipes, en dupliquant l'onglet partie 5 fois (Main#main : nbPartieQualificative)

La 1ère feuille contient le tableau syntétiques des résultats par équipe avec le nombre de partie gagnée (Les parties se jouent en 11 points, voir excel modèle)


