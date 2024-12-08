// The MIT License (MIT)
// Copyright (c) 2016-2018 Laurent BRAUD
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
package com.elbedd.palet;

import com.elbedd.palet.excel.Generator;
import com.elbedd.palet.model.Concours;
import com.elbedd.palet.model.Equipe;

public class Main {

	public static void main(String[] arg) {
		int nbPartieQualificative = 6;
		// Le nombre de partie aux hasard : mettre le maxumim. Peut etre généré dans l'excel
		int nbPartieHazard = nbPartieQualificative;
	
		int nbEquipeMin = 8;
		int nbEquipeMax = 80;
		// int scoreToWin = 11;
		
		boolean withClassementIntermediaire = true;
		
		for (int nbEquipe = nbEquipeMin; nbEquipe <= nbEquipeMax; nbEquipe++) {
			Concours concours = new Concours(nbPartieQualificative);

			for (int numEquipe = 0; numEquipe < nbEquipe; numEquipe++) {
				Equipe equipe = new Equipe(numEquipe + 1);
				concours.addEquipe(equipe);
			}

			concours.effectueTirageQualification(nbPartieHazard);
			// concours.gerePrincipale();
			//concours.display();
			Generator generator = new Generator(concours, false);
			
			String numExcel = "000" + nbEquipe;
			numExcel = numExcel.substring(numExcel.length() - 3, numExcel.length());
			// enregistrer dans un XLS
			try {
				generator.generateExcel("d:/temp/tirage/tirage" + numExcel + ".xls", withClassementIntermediaire, nbPartieHazard);
			} catch(Exception e) {
				e.printStackTrace();
			}
		}
		
		
		

	}

}
