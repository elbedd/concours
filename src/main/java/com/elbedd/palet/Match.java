// The MIT License (MIT)
// Copyright (c) 2016 Laurent BRAUD
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
package com.elbedd.palet;

public class Match {

	private Equipe equipeA;

	/**
	 * null : A exempt
	 */
	private Equipe equipeB;

	private int numeroPlaque;

	private int resultatA;

	private int resultatB;

	public Equipe getEquipeA() {
		return equipeA;
	}

	public Equipe getEquipeB() {
		return equipeB;
	}

	public int getNumeroPlaque() {
		return numeroPlaque;
	}

	public void setNumeroPlaque(int numeroPlaque) {
		this.numeroPlaque = numeroPlaque;
	}

	public int getResultatA() {
		return resultatA;
	}

	public void setResultatA(int resultatA) {
		this.resultatA = resultatA;
	}

	public int getResultatB() {
		return resultatB;
	}

	public void setResultatB(int resultatB) {
		this.resultatB = resultatB;
	}

	public void setEquipeA(Equipe equipeA) {
		this.equipeA = equipeA;
	}

	public void setEquipeB(Equipe equipeB) {
		this.equipeB = equipeB;
	}

	public boolean hasSameEquipe(Match match) {
		return (equipeA == match.getEquipeA() || equipeA == match.getEquipeB())
				&& (equipeB == match.getEquipeA() || equipeB == match.getEquipeB());

	}

	public boolean hasEquipe(Equipe equipe) {
		return (equipeA == equipe || equipeB == equipe);

	}

}
