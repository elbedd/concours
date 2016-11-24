// The MIT License (MIT)
// Copyright (c) 2016 Laurent BRAUD
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
package com.elbedd.palet;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;

public class Partie {
	private int numero;

	private List<Match> matchs;

	private Map<Equipe, Match> matchsByEquipe;

	public Partie() {
		matchs = new ArrayList<Match>();
		matchsByEquipe = new HashMap<Equipe, Match>();
	}

	public int getNumero() {
		return numero;
	}

	public void setNumero(int numero) {
		this.numero = numero;
	}

	public List<Match> getMatchs() {
		return matchs;
	}

	public Match getMatchOfEquipe(Equipe equipe) {
		Match match = matchsByEquipe.get(equipe);
		return match;
	}

	public boolean isMatchExmpt(Equipe equipe) {
		Match match = matchsByEquipe.get(equipe);
		return match != null && match.getEquipeB() == null;
	}

	public void addMatchExmpt(Equipe equipe) {
		Match match = new Match();
		match.setEquipeA(equipe);
		matchsByEquipe.put(equipe, match);
		matchs.add(match);

	}

	public Match createMatch(Equipe equipeA, Equipe equipeB) {
		Match match = new Match();
		match.setNumeroPlaque(matchs.size() + 1);
		match.setEquipeA(equipeA);
		match.setEquipeB(equipeB);
		matchsByEquipe.put(equipeA, match);
		matchsByEquipe.put(equipeB, match);
		matchs.add(match);

		return match;
	}

	public void display() {
		System.out.println("Partie " + numero);
		for (Match match : matchs) {
			if (match.getEquipeB() != null) {
				System.out.println("Equipe " + match.getEquipeA().getNumero() + " vs " + match.getEquipeB().getNumero()
						+ " sur la planche " + match.getNumeroPlaque());
			} else {
				System.out.println("Equipe " + match.getEquipeA().getNumero() + " exempt");
			}

		}

	}

	public static Partie effectueTirage(int numeroPartie, Map<Integer, Equipe> equipesByNum,
			List<Partie> tiragePrecedant) {
		Partie ret = new Partie();
		ret.setNumero(numeroPartie);
		List<Equipe> equipeSansMatch = new ArrayList<Equipe>(equipesByNum.values());

		// On gère les exempts en 1er pour éviter d'avoir 2 fois le même exempt
		// à la fin.
		Equipe equipeExempte = findExempt(equipesByNum, tiragePrecedant);
		if (equipeExempte != null) {
			equipeSansMatch.remove(equipeExempte);
		}

		// Faire des matchs aux hasards
		Random rand = new Random();

		while (equipeSansMatch.size() > 0) {
			int indiceEquipeA = rand.nextInt(equipeSansMatch.size());
			int indiceEquipeB = rand.nextInt(equipeSansMatch.size());

			if (indiceEquipeA != indiceEquipeB) {
				Equipe equipeA = equipeSansMatch.get(indiceEquipeA);
				Equipe equipeB = equipeSansMatch.get(indiceEquipeB);
				ret.createMatch(equipeA, equipeB);
				equipeSansMatch.remove(equipeA);
				equipeSansMatch.remove(equipeB);
			}
		}

		// Recherche si match déjà joué dans une précédante partie
		if (tiragePrecedant.size() > 0) {
			List<Match> matchAChanger = new ArrayList<Match>();
			for (Match match : ret.getMatchs()) {
				if (matchDejaJoueAvant(tiragePrecedant, match)) {
					matchAChanger.add(match);
				}
			}

			// Revoir les tirages
			while (matchAChanger.size() > 0) {

				Match matchAModifier = matchAChanger.get(matchAChanger.size() - 1);
				// Trouve un match non joue [Rq : ne foncttionnera que si
				// nbPartie pas elevé et Nb equipe >]
				Equipe equipeA = matchAModifier.getEquipeA();
				Equipe equipeB = matchAModifier.getEquipeB();
				for (Match match : ret.getMatchs()) {
					if (match != matchAModifier) {
						Equipe equipeC = match.getEquipeA();
						Equipe equipeD = match.getEquipeB();

						// L'équipe A e telle joue contre C et equipe B contre D
						// ?
						// Sinon : Intervertir le match
						Match matchTest1 = new Match();
						matchTest1.setEquipeA(equipeA);
						matchTest1.setEquipeB(equipeC);

						Match matchTest2 = new Match();
						matchTest2.setEquipeA(equipeB);
						matchTest2.setEquipeB(equipeD);
						boolean dejaJoue = matchDejaJoueAvant(tiragePrecedant, matchTest1)
								|| matchDejaJoueAvant(tiragePrecedant, matchTest2);

						if (dejaJoue) {
							// Match A-D et B-C ?
							matchTest1 = new Match();
							matchTest1.setEquipeA(equipeA);
							matchTest1.setEquipeB(equipeD);

							matchTest2 = new Match();
							matchTest2.setEquipeA(equipeB);
							matchTest2.setEquipeB(equipeC);
							dejaJoue = matchDejaJoueAvant(tiragePrecedant, matchTest1)
									|| matchDejaJoueAvant(tiragePrecedant, matchTest2);
						}

						if (!dejaJoue) {
							// "Match " + matchAModifier.getNumeroPlaque() +"
							// <-> " + match.getNumeroPlaque()
							// intervertir match.
							match.setEquipeA(matchTest1.getEquipeA());
							match.setEquipeB(matchTest1.getEquipeB());
							matchAChanger.remove(match); // May be true when 2
															// matchs to
															// reverse.

							matchAModifier.setEquipeA(matchTest2.getEquipeA());
							matchAModifier.setEquipeB(matchTest2.getEquipeB());
							matchAChanger.remove(matchAModifier);

							break;
						}

					}
				}
			}

		}

		// On ajoute le match exempt à la dernière plaque.
		if (equipeExempte != null) {
			ret.addMatchExmpt(equipeExempte);
		}

		return ret;
	}

	protected static boolean matchDejaJoueAvant(List<Partie> tiragePrecedant, Match match) {
		boolean ret = false;
		Equipe equipe = match.getEquipeA();

		for (Partie partie : tiragePrecedant) {
			Match matchPrecedant = partie.getMatchOfEquipe(equipe);

			if (match.hasSameEquipe(matchPrecedant)) {
				// == match de la plaque " + match.getNumeroPlaque() + " deja
				// joue partie " + partie.getNumero() + " a la plaque " +
				// matchPrecedant.getNumeroPlaque()
				ret = true;
				break;
			}
		}
		return ret;
	}

	/**
	 * Recherche EquipeExempte (0 ou 1 par Partie) Cette équipe ne doit pas
	 * avoir été exempte auparavant.
	 * 
	 * @param equipesByNum
	 * @param tiragePrecedant
	 * @return l'équipe exempte le cas échéant
	 */
	private static Equipe findExempt(Map<Integer, Equipe> equipesByNum, List<Partie> tiragePrecedant) {
		Equipe equipeExempte = null;
		// nombre d'équipe impaire => Trouver un exempt
		int nbEquipe = equipesByNum.size();

		if (nbEquipe % 2 > 0) {
			int nbMaxExemptAvant = tiragePrecedant.size() / nbEquipe;// Serait
																		// différent
																		// si
																		// nombre_partie_jouée>
																		// nbEquipe

			do {
				Random rand = new Random();
				int numEquipeAleatoire = rand.nextInt(nbEquipe) + 1;
				// déjà été exempt ?
				int nombreFoisExempt = 0;

				for (Partie partie : tiragePrecedant) {
					Equipe equipe = equipesByNum.get(new Integer(numEquipeAleatoire));
					if (partie.isMatchExmpt(equipe)) {
						nombreFoisExempt++;
					}
				}

				if (nombreFoisExempt <= nbMaxExemptAvant) {
					equipeExempte = equipesByNum.get(new Integer(numEquipeAleatoire));
					;
				}

			} while (equipeExempte == null);

			System.out.println(equipeExempte.getNumero());
		}
		return equipeExempte;
	}

}
