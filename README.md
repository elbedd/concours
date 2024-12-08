Permet de générer un excel avec les parties qualificatives d'un concours de palet

Le programme génère autant de fichier excel que d'équipes potentielles.
Ce nombre d'équipe est en dur dans le main (à externaliser) : Nombre minimum et nombre maximum.

Le fichier excel généré est une copie du fichier modèle (src/main/resources/com/elbedd/palet/excel/modele.xls)

Le fichier est complété par les lignes correspondant aux équipes, en dupliquant l'onglet partie 5 fois (Main#main : nbPartieQualificative).

La 1ère feuille contient le tableau syntétique des résultats par équipe avec le nombre de partie gagnée (Les parties se jouent en 11 points, voir excel modèle). Le progarmme permet de générer des formules. Par exemple pour que le nom des joueurs ne soit modifié qu'une fois dans l'excel.

L'algorithme fait en sorte que les équipes ne se rencontrent pas 2 fois.
C'est basé sur un scénario optimiste (on a généralement suffisament d'équipe par rapport au nombre de parties)
L'algorithme détermine également l'équipe exempt (Si nombre impair), et évite qu'une équipe soit 2 fois exemptes et une autre ne le soit pas.


Licence :
The MIT License (MIT)
Copyright (c) 2016-2019 Laurent BRAUD

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
