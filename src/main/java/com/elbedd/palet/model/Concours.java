// The MIT License (MIT)
// Copyright (c) 2016 Laurent BRAUD
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
package com.elbedd.palet.model;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Concours {

	private final int nbPartieQualificative;

	private List<Partie> parties;

	private Map<Integer, Equipe> equipes;

	public Concours(int nbPartieQualificative) {
		this.nbPartieQualificative = nbPartieQualificative;
		equipes = new HashMap<Integer, Equipe>();
	}
	
	public List<Partie> getParties() {
		return parties;
	}
	
	public Map<Integer, Equipe> getEquipes() {
		return equipes;
	}

	
	
	public void effectueTirageQualification() {
		effectueTirageQualification(nbPartieQualificative);
	}

	public void effectueTirageQualification(int nbPartie) {
		parties = new ArrayList<Partie>();
		for (int i = 1; i <= nbPartieQualificative; i++) {
			Partie tirage = Partie.effectueTirage(i, equipes, parties, i <= nbPartie );
			tirage.display();
			parties.add(tirage);
		}
//		for (int i = nbPartie + 1; i <= this.nbPartieQualificative; i++) {
//			Partie tirage = new Partie(i);
//			parties.add(tirage);
//			nbPartie++;
//		}
	}

	public void display() {
		for (Partie partie : parties) {
			partie.display();
		}

	}

	public boolean addEquipe(Equipe equipe) {
		boolean ret = false;
		if (equipe != null) {
			Equipe a = equipes.put(new Integer(equipe.getNumero()), equipe);
			ret = a != null;
		}
		return ret;
	}

	public void gerePrincipale() {

		
	}
	
	
	public void calcNbTableau() {
		int nbTableau;
		int[] nbEquipeTableau;
		if (equipes.size() % 32 == 0) {
			nbTableau = equipes.size() / 32;
			nbEquipeTableau = new int[1];
			nbEquipeTableau[0]=32;
		} else {
			
			
		}
	}

}
