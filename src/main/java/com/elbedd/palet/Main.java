// The MIT License (MIT)
// Copyright (c) 2016 Laurent BRAUD
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
		int nbPartieQualificative = 5;

		Concours concours = new Concours(nbPartieQualificative);

		int nbEquipe = 41;
		for (int numEquipe = 0; numEquipe < nbEquipe; numEquipe++) {
			Equipe equipe = new Equipe(numEquipe + 1);
			concours.addEquipe(equipe);
		}

		concours.effectueTirageQualification();
		//concours.display();
		Generator generator = new Generator(concours);
		
		// enregistrer dans un XLS
		try {
			generator.generateExcel("d:/temp/myExcel.xls");	
		} catch(Exception e) {
			e.printStackTrace();
		}
		

	}

}
