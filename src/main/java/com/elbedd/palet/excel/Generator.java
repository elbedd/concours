// The MIT License (MIT)
// Copyright (c) 2016-2024 Laurent BRAUD
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

package com.elbedd.palet.excel;

import java.io.FileOutputStream;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import com.elbedd.palet.model.Concours;
import com.elbedd.palet.model.Equipe;
import com.elbedd.palet.model.Match;
import com.elbedd.palet.model.Partie;

public class Generator {
	
	private final static String REF_SHEET = "!";// In apachePOI ?
	private final static String REF_RANGE = ":";// In apachePOI ?
	
	private final static String SHEET_LIST_NAME = "Liste";
	private final static int SHEET_LIST_FIRSTLINE = 3;
	private final static int INDEX_SHEET_LISTE_COLUMN_PALET = 7;
	
	private final static String SHEET_TEAM_NAME = "Equipes";
	private final static String SHEET_CLASSEMENT = "Clt";
	private final static int SHEET_TEAM_NUMBER = 1;
	private final static int SHEET_TEAM_FIRSTLINE = 4;
	
	
	private final static String SHEET_PARTIE_NAME = "Partie";
	private final static int SHEET_PARTIE_NUMBER = 2;
	private final static int SHEET_PARTIE_FIRSTLINE = 3;
	
	
	private String CELL_SCORE_WIN = SHEET_TEAM_NAME + "!G1";
	// private int SCORE_WIN = 11;
	
	private int firstColumnPartieInSheetTeam = 3;
	
	Concours concours;
	boolean tableauFinale = true;

	public Generator(Concours concours) {
		this(concours, false);
	}
	
	public Generator(Concours concours, boolean tableauFinale) {
		this.concours = concours;
		if (tableauFinale) {
			firstColumnPartieInSheetTeam = 4;
		}
		
		this.tableauFinale = tableauFinale;
	}
	
	private Workbook loadModeleExcel() {
		Workbook ret = null;
		try {
			// src/main/resources/com/elbedd/palet/excel
			InputStream is = null;
			if (tableauFinale) {
				is = getClass().getResourceAsStream("modeleFinale.xls");
			} else {
				is = getClass().getResourceAsStream("modele.xls");
			}
			
			ret = new HSSFWorkbook(is);
			
		} catch(Exception e) {
			// Exception : Fichier mod�le non charg� 
			ret = new HSSFWorkbook();
		} finally {
			
		}
		return ret;
	}
	/**
	 * TODO : review Exception
	 * 
	 * @param fileName
	 * @throws Exception
	 */
	public void generateExcel(String fileName) throws Exception {
		generateExcel(fileName, false, this.concours.getParties().size());
	 }
	public void generateExcel(String fileName, boolean withClassementIntermediaire, int nbPartieHazard) throws Exception {
		Workbook wb = loadModeleExcel();
		
		//setGrayStyleSheetFromWb(wb);
		//setVerticalStyleSheetFromWb(wb);
		writeSheetEquipes(wb);
		String plageEquipeOrdre = null;
		String plageEquipeName = null;
		String plageEquipeAdverse = null;
		
		if (tableauFinale) {
			plageEquipeOrdre = computeTeamRange("A", "B", concours.getEquipes().size());
			plageEquipeName = computeTeamRange("B", "D", concours.getEquipes().size());
		} else {
			plageEquipeName = computeTeamRange("A", "C", concours.getEquipes().size());
			// W : 6 match (Colonne 6 + 6*2 match +4)
			plageEquipeAdverse = computeTeamRange("A", "X", concours.getEquipes().size());
		}
		
		
		for (Partie partie : concours.getParties()) {
			//wb.createSheet("Partie" + partie.getNumero());
			Sheet sheet = wb.cloneSheet(SHEET_PARTIE_NUMBER);
			wb.setSheetName(wb.getSheetIndex(sheet), SHEET_PARTIE_NAME + partie.getNumero());
			writeSheetPartieHeader(sheet, partie);

			for (Match match : partie.getMatchs()) {
				writeSheetPartieMatch(sheet, partie.getNumero(), concours.getParties().size(), match, plageEquipeOrdre, plageEquipeName, plageEquipeAdverse, nbPartieHazard <= partie.getNumero());
			}
			
			
			
		}
		wb.removeSheetAt(SHEET_PARTIE_NUMBER);
		
		
		writeMoreInSheetEquipes(wb);

		if (withClassementIntermediaire) {
			for (Partie partie : concours.getParties()) {
				Sheet classementI = wb.cloneSheet(SHEET_TEAM_NUMBER);
				replaceConstantByRef(classementI);
				String sheetClassement = SHEET_CLASSEMENT + partie.getNumero();
				wb.setSheetName(wb.getSheetIndex(classementI), sheetClassement);
				// pose probl�me car change formule de Sheet 1.
				//wb.setSheetOrder(classementI.getSheetName(), partie.getNumero()*2);
				// 7 : numero, joueur1, joueur2, NbWin, PP, PC, DIFF, 
				// 3* : PP, PC, Num Adversaire
				final int numCellInfoTeam = 7 + 3*concours.getParties().size();
				if (partie.getNumero() == concours.getParties().size()) {
					int lastRow = SHEET_TEAM_FIRSTLINE + concours.getEquipes().size();
					for(int iRow = SHEET_TEAM_FIRSTLINE; iRow < lastRow;iRow++) {
						Row row = Util.getOrCreateRow(classementI, iRow);
						Cell cell = Util.getOrCreateCell(row, numCellInfoTeam);
						String formula = buildFormulaIndexInListe("A" + (iRow + 1), INDEX_SHEET_LISTE_COLUMN_PALET).toString();
						cell.setCellFormula(formula);
					}
				}
				//
				if (partie.getNumero() > 1) {
					Sheet sheetPartie = wb.getSheet(SHEET_PARTIE_NAME + partie.getNumero());
					//
					int rowClt = SHEET_TEAM_FIRSTLINE;
					for (Match match : partie.getMatchs()) {
						int numRow = match.getNumeroPlaque() + 1;
						Row r= sheetPartie.getRow(numRow);
						Cell TeamA = r.getCell(1);
						rowClt++;
						String sheetnameRef = SHEET_CLASSEMENT + (partie.getNumero()-1) + "!A" + rowClt;
						// La formule regarde le nombre de partie al�atoire indiqu� dans la feuille Equipe
						// Si nombre d�pass�, alors on va r�cup�r� l'info dans le classement.
						String formulaCondition = partie.getNumero() + ">"+ SHEET_TEAM_NAME + "!E1";
						String formulaA = "IF(" + formulaCondition + "," + sheetnameRef + "," + match.getEquipeA().getNumero() + ")";
						TeamA.setCellFormula(formulaA);

						if (match.getEquipeB()!=null) {
							Cell TeamB = r.getCell(3);
							rowClt++;
							sheetnameRef = SHEET_CLASSEMENT + (partie.getNumero()-1) + "!A" + rowClt;
							String formulaB = "IF(" + formulaCondition + "," + sheetnameRef + "," + match.getEquipeB().getNumero() + ")";
							TeamB.setCellFormula(formulaB);
						}
						
					}
				}
				
					
				//}
				
			}
			
		}
		
		wb.setActiveSheet(1);

		try (FileOutputStream out = new FileOutputStream(fileName)) {
			wb.write(out);
			out.close();
		} catch (Exception e) {
			throw e;
		}
		wb.close();
	}

	/**
	 * Remplace l'information dupliqu� de la feuille Equipe
	 * Par une formule qui r�cup�re la valeur de la cellule (devient dynamique)
	 */
	private void replaceConstantByRef(Sheet classementI) {
		Row headerRow = Util.getOrCreateRow(classementI, 0);
		for (int cellnum = 0; cellnum < 10;) {
			Cell cell = Util.getOrCreateCell(headerRow, cellnum);
			String columnLetter = Util.getColumnLetter(++cellnum);
			cell.setCellFormula(SHEET_TEAM_NAME + REF_SHEET + columnLetter + "1");
		}
	}
	private String computeTeamRange(String firstColumn, String lastComumn, int nbTeam) {
		String referenceStart = SHEET_TEAM_NAME + REF_SHEET + firstColumn+ (SHEET_TEAM_FIRSTLINE + 1);
		int lastLine = SHEET_TEAM_FIRSTLINE + nbTeam;
		String referenceEnd = lastComumn + lastLine;
		String plageEquipe = referenceStart + REF_RANGE + referenceEnd;// Exemple "Equipes!A3:C43";
		return plageEquipe;
	}

	protected void writeSheetPartieHeader(Sheet sheet, Partie partie) {
		Row row = Util.getOrCreateRow(sheet, 0);
		
		Cell cell = Util.getOrCreateCell(row, 2);
		if(tableauFinale) {
			cell.setCellValue("Tableau");
		} else {
			cell.setCellValue("Partie n�" + partie.getNumero());
		}
		
		Row rowHeader = Util.getOrCreateRow(sheet, 1);
		Cell cellE1 = Util.getOrCreateCell(rowHeader, 7);
		cellE1.setCellValue("Rappel Equipe1");
		sheet.setColumnWidth(7, 0);
		
	}
	
	protected void writeSheetPartieMatch(Sheet sheet, int numeroPartie, int nbPartie, Match match, String rangeTeamOrder, String rangeTeamName, String teamRange, boolean TeamByClt) {
		int numRow = match.getNumeroPlaque() + 1;
		String excelRowNumber = "" + (numRow + 1);
		
		int TeamNumberA = match.getEquipeA().getNumero();
		int TeamNumberB = -1;
		if (match.getEquipeB() != null) {
			TeamNumberB = match.getEquipeB().getNumero();
		}
		
		Row row = Util.getOrCreateRow(sheet, numRow);
		 
		Row rowModel = null;
		//CellStyle styleRow = null;
		if (match.getNumeroPlaque() % 2 == 1) {
			rowModel = Util.getOrCreateRow(sheet, 2);
		} else {
			rowModel = Util.getOrCreateRow(sheet, 3);
		}
		
		//row.setRowStyle(rowModel.getRowStyle());
		row.setHeight(rowModel.getHeight());
		
		int i = 0;
		
		Cell cell = Util.getOrCreateCell(row, i);
		cell.setCellValue(match.getNumeroPlaque());
		cell.setCellStyle(rowModel.getCell(i).getCellStyle());
		
		i++;
		cell = Util.getOrCreateCell(row, i);
		if (TeamByClt) {
			// Gerer apr�s avoir ecrit la feuille de classement
			cell.setCellValue(TeamNumberA);
		} else if (tableauFinale) {
			// Dans ce cas, il ne faut pas afficher le nu�mro de l'�quipe mais le num�ro d'ordre
			searchFormulaTeamByOrdre(TeamNumberA, rangeTeamOrder, cell, "B");
		} else {
			// num�ro d'�quipe = numero d'ordre
			cell.setCellValue(TeamNumberA);
		}
		
		
		cell.setCellStyle(rowModel.getCell(i).getCellStyle());
		
		i++;
		cell = Util.getOrCreateCell(row, i);//row.createCell(i++, CellType.FORMULA);
		//cell.setCellFormula("VLOOKUP(" + match.getEquipeA().getNumero() + ",A3:C43,3,FALSE)");
		//cell.setCellType(CellType.FORMULA);
		searchFormulaTeamNames(rangeTeamName, cell, "B");
		cell.setCellStyle(rowModel.getCell(i).getCellStyle());
		
		i++;
		cell = Util.getOrCreateCell(row, i);
		cell.setCellStyle(rowModel.getCell(i).getCellStyle());
		if (TeamNumberB > 0) {
			// Non exempt
			if (TeamByClt) {
				//int TeamNumberB = (match.getNumeroPlaque()-1) * 2 +1;
				cell.setCellValue(TeamNumberB);
			} else if (tableauFinale) {
				searchFormulaTeamByOrdre(TeamNumberB, rangeTeamOrder, cell, "B");
			} else {
				cell.setCellValue(TeamNumberB);
			}
		}
		
		i++;
		cell = Util.getOrCreateCell(row, i);
		cell.setCellStyle(rowModel.getCell(i).getCellStyle());
		if (match.getEquipeB() != null) {
			// Non exempt
			searchFormulaTeamNames(rangeTeamName, cell, "D");
		}
		
		// Score (vide)
		i++;
		cell = Util.getOrCreateCell(row, i);
		cell.setCellStyle(rowModel.getCell(i).getCellStyle());
//		if (match.getEquipeB() == null && !TeamByClt) {
//			cell.setCellValue(SCORE_WIN);
//		}
		i++;
		cell = Util.getOrCreateCell(row, i);
		cell.setCellStyle(rowModel.getCell(i).getCellStyle());
//		if (match.getEquipeB() == null) {
//			cell.setCellValue(0);
//		}
		// Rappel numero Equipe 1
		i++;
		cell = Util.getOrCreateCell(row, i);
		cell.setCellFormula("B" + excelRowNumber);
		
		if (numeroPartie > 1) {
			int indexWin = nbPartie * 2 + 4;
			// Nb victoire A
			i++;
			cell = Util.getOrCreateCell(row, i);
			applyFormulaIndexColumnByTeam(cell, teamRange, "B", excelRowNumber, indexWin);
			// Nb victoire B
			i++;
			cell = Util.getOrCreateCell(row, i);
			applyFormulaIndexColumnByTeam(cell, teamRange, "D", excelRowNumber, indexWin);
			
			// Rappel Adversaire Equipe A
			i++;
			cell = Util.getOrCreateCell(row, i);
			
			searchFormulaTeamAdv(cell, teamRange, numeroPartie, numRow, "B", nbPartie);
			i++;
			if (match.getEquipeB() != null) {
				// Rappel Adversaire Equipe B (sauf si A exempt)
				cell = Util.getOrCreateCell(row, i);
				searchFormulaTeamAdv(cell, teamRange, numeroPartie, numRow, "D", nbPartie);
			}
			
			i++;
			cell = Util.getOrCreateCell(row, i);
			cell.setCellFormula("IF(ISERROR(FIND(CONCATENATE(\" \",B" + excelRowNumber + ", \"-\"),L" + excelRowNumber + ")),\"\",\"DEJA JOUE\")");
			
		}
		
		
		
		
		
		
	}

	protected void searchFormulaTeamNames(String rangeEquipe, Cell cell, String columnNumEquipe) {
		StringBuilder formula = new StringBuilder("CONCATENATE(");
		int rowFormula = cell.getRow().getRowNum() + 1;
		int posJoueurDansTableau = 2;
		
		formula.append("VLOOKUP(" + columnNumEquipe + rowFormula + ", " + rangeEquipe + "," + posJoueurDansTableau + ",FALSE)");// First Player
		formula.append(",");//SEP CONCAT
		formula.append("CHAR(10)");//Retour Chariot
		formula.append(",");//SEP CONCAT
		posJoueurDansTableau++;
		formula.append("VLOOKUP(" + columnNumEquipe + rowFormula + ", " + rangeEquipe + "," + posJoueurDansTableau + ",FALSE)");// 2nd Player (colonne 3)
		formula.append(")");//END CONCAT
		
		cell.setCellFormula(formula.toString());
	}
	
	protected void searchFormulaTeamByOrdre(int teamNumber, String rangeEquipe, Cell cell, String columnNumEquipe) {
		StringBuilder formula = new StringBuilder();
		formula.append("VLOOKUP(" + teamNumber + ", " + rangeEquipe + ",2,FALSE)");
		cell.setCellFormula(formula.toString());
	}
	
	protected StringBuilder buildFormulaTeamAdv(String rangeEquipe, String cellTeamNumber, int indexPartie) {
		StringBuilder formula = new StringBuilder();
		formula.append("VLOOKUP(" + cellTeamNumber + ", " + rangeEquipe + "," + indexPartie + ",FALSE)");
		return formula;
	}
	
	protected StringBuilder buildFormulaIndexInListe(String cellTeamNumber, int index) {
		StringBuilder formula = new StringBuilder();
		int endLine = SHEET_LIST_FIRSTLINE + concours.getEquipes().size();
		final String rangeEquipe = SHEET_LIST_NAME + "!A" + (SHEET_LIST_FIRSTLINE + 1) + ":G" + endLine;
		formula.append("VLOOKUP(" + cellTeamNumber + ", " + rangeEquipe + "," + index + ",FALSE)");
		return formula;
	}
	
	protected void searchFormulaTeamAdv(Cell cell, String plageEquipeAdverse, int numeroPartie, int numRow, String columnEquipe, int nbPartie) {
		StringBuilder formula = new StringBuilder("CONCATENATE(");
		for (int iPartie = 1; iPartie < numeroPartie; iPartie++) {
			if (iPartie > 0) {
				formula.append(", \"  \", ");
			} else {
				formula.append("\"  \", ");
			}
			// 7: colonnes :Equipe, Joueurs (2), PG, PP, PC, Diff
			formula.append(buildFormulaTeamAdv(plageEquipeAdverse, columnEquipe + (numRow + 1) , nbPartie*2 + 7 + iPartie));
			formula.append(", \"-\"");
		}
		formula.append(")");
		cell.setCellFormula(formula.toString());
	}

	/**
	 * Set ce
	 * @param rangeEquipe
	 * @param cell
	 * @param columnNumEquipe
	 * @param rowFormula
	 * @param index
	 */
	protected void applyFormulaIndexColumnByTeam(Cell cell, String rangeEquipe, String columnNumEquipe, String rowFormula, int index) {
		String formula = "VLOOKUP(" + columnNumEquipe + rowFormula + ", " + rangeEquipe + "," + index + ",FALSE)";
		cell.setCellFormula(formula.toString());
	}
	
	protected void writeSheetEquipes(Workbook wb) {
		Sheet sheet = wb.getSheet(SHEET_TEAM_NAME);
		Row row = Util.getOrCreateRow(sheet, 0);
		
		Cell cell = Util.getOrCreateCell(row, 2);
		cell.setCellValue(concours.getParties().size());
		
		for (Equipe equipe : concours.getEquipes().values()) {
			writeEquipe(sheet, equipe);
		}

	}
	
	
	
	protected void writeEquipe(Sheet sheet, Equipe equipe) {
		// -1 =  -2 + 1 (+1 => On ajoute la ligne modele)
		Row row = Util.getOrCreateRow(sheet, SHEET_TEAM_FIRSTLINE + equipe.getNumero() - 1);
		
		int numCell = 0;
		Cell cell = Util.getOrCreateCell(row, numCell++);
		
		if (tableauFinale) {
			// Numero d'ordre
			cell.setCellValue(equipe.getNumero());
			cell = Util.getOrCreateCell(row, numCell++);
		}
		
		// Num�ro d'�quipe
		cell.setCellValue(equipe.getNumero());
		
		cell = Util.getOrCreateCell(row, numCell++);
		cell.setCellValue("Player1 Team" + equipe.getNumero());// 1st player
		
		cell = Util.getOrCreateCell(row, numCell++);
		cell.setCellValue("Player2 Team" + equipe.getNumero());// 2nd player
		
	}
	
	protected void writeClassementEquipe(Sheet sheet, Equipe equipe, String plageEquipe) {
		Row row = sheet.getRow(SHEET_TEAM_FIRSTLINE + equipe.getNumero() - 2);
		row.setHeight((short) (row.getHeight() * 2)); 
		int numCell = 1;
		// Dans cette cellule, remplacer le nom du joueur 1 par le nom de l'�quipe
		Cell cell = row.getCell(numCell++);
		cell.setCellValue(equipe.getNumero());
		// Remplacer le nom du joueur 2 par jouer1+2
		cell = row.getCell(numCell);
		row.removeCell(cell);
		cell = row.createCell(numCell, CellType.FORMULA);
		searchFormulaTeamNames(plageEquipe, cell, "A");
		//cell.setCellStyle(verticalCellStyleSheet);
		
		
	}
	
	private void writeMoreInSheetEquipes(Workbook wb) {
		Sheet sheet = wb.getSheet(SHEET_TEAM_NAME);
		
		// Nombre de partie aux hasards g�r�s
		Row row0 = sheet.getRow(0);
		Cell cellRandom = Util.getOrCreateCell(row0, 4);
		cellRandom.setCellValue(3);	// Par defaut
		cellRandom = Util.getOrCreateCell(row0, 5);
		cellRandom.setCellValue("au hasard");
		
		cellRandom = Util.getOrCreateCell(row0, 6);
		cellRandom.setCellValue(11);	// Par defaut
		cellRandom = Util.getOrCreateCell(row0, 7);
		cellRandom.setCellValue("Points Gagnant");
		// On recopie la ligne 3 de l'excel en ajustant de 6 � nbPartie.
		
		Row rowModel = Util.getOrCreateRow(sheet, 2);// Ligne 3
		Row rowHeader = writeHeaderInSheetEquipe(sheet, rowModel);
		sheet.removeRow(rowModel);
		
		for (Equipe equipe : concours.getEquipes().values()) {
			int numLigne = SHEET_TEAM_FIRSTLINE + equipe.getNumero();	
			Row row = sheet.getRow(numLigne - 1);
			
			int numCell = firstColumnPartieInSheetTeam;
			// ajouter le r�sultat des parties
			for (int iPartie = 0; iPartie < concours.getParties().size(); iPartie++) {
				// Score �quipe
				Cell cell = row.createCell(numCell++, CellType.FORMULA);
				cell.setCellType(CellType.FORMULA);
				String formula = searchFormulaScore(equipe, iPartie + 1);
				cell.setCellFormula(formula.toString()); 
				cell.setCellStyle(rowHeader.getCell(cell.getColumnIndex()).getCellStyle());
   
				// Score adverse
			
				Cell cell2 = row.createCell(numCell++, CellType.FORMULA);
				String formula2 = searchFormulaScoreAdverse(equipe, iPartie + 1);
				cell2.setCellType(CellType.FORMULA);
				cell2.setCellFormula(formula2.toString());
				cell2.setCellStyle(rowHeader.getCell(cell2.getColumnIndex()).getCellStyle());
				
			}
			// Faire le nombre de victoire 
			// Si 5 partie : 
			//  =SI(D4<11;0;1)+SI(F4<11;0;1)+SI(H4<11;0;1)+SI(J4<11;0;1)+SI(L4<11;0;1)
			char[] tabLettre = {'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P','Q','R' ,'S','T','U','V','W','X','Y','Z'};
			int indexLettre = 3;
			if (tableauFinale) {
				indexLettre++;
			}
			
			StringBuilder sbWin = new StringBuilder("");
			StringBuilder sbPour = new StringBuilder("");
			StringBuilder sbContre = new StringBuilder("");
			
			for (int iPartie = 0; iPartie < concours.getParties().size(); iPartie++) {
				sbWin.append("+IF(" + tabLettre[indexLettre] + numLigne + "<" + CELL_SCORE_WIN +",0,1)");
				sbPour.append("+" +tabLettre[indexLettre] + numLigne);
				indexLettre ++;
				sbContre.append("+" +tabLettre[indexLettre] + numLigne);
				indexLettre ++;
			}
			
			
			Cell cellWin = row.createCell(numCell++, CellType.FORMULA);
			String formula = sbWin.substring(1);// Remove the first +
			cellWin.setCellFormula(formula);
			cellWin.setCellStyle(rowHeader.getCell(cellWin.getColumnIndex()).getCellStyle());
			indexLettre ++;
			
			Cell cellPP = row.createCell(numCell++, CellType.FORMULA);
			formula = sbPour.substring(1);// Remove the first +
			cellPP.setCellFormula(formula);
			cellPP.setCellStyle(rowHeader.getCell(cellPP.getColumnIndex()).getCellStyle());
			indexLettre ++;
		
			Cell cellPC = row.createCell(numCell++, CellType.FORMULA);
			formula = sbContre.substring(1);// Remove the first +
			cellPC.setCellFormula(formula);
			cellPC.setCellStyle(rowHeader.getCell(cellPC.getColumnIndex()).getCellStyle());
			indexLettre ++;
			
			Cell cellDiff = row.createCell(numCell++, CellType.FORMULA);
			cellDiff.setCellFormula("" + tabLettre[indexLettre - 2]+ numLigne + "-" + tabLettre[indexLettre - 1]+ numLigne);
			cellDiff.setCellStyle(rowHeader.getCell(cellPC.getColumnIndex()).getCellStyle());
			
			// 1 colonne par Adversaire
			for (int iPartie = 0; iPartie < concours.getParties().size(); iPartie++) {
				Cell Adv = row.createCell(numCell++, CellType.FORMULA);
				formula = searchFormulaAdversaire(equipe, iPartie + 1);
				Adv.setCellFormula(formula);
			}
			
			
		}
		
	}

	/**
	 *
	 * @param sheet
	 * @return
	 */
	private Row writeHeaderInSheetEquipe(Sheet sheet, Row rowModel) {
		
		Row row = sheet.createRow(rowModel.getRowNum()+1);
		row.setHeight(rowModel.getHeight());
		for (int numCell = 0; numCell < firstColumnPartieInSheetTeam; numCell++) {
			Cell cellEntete = row.createCell(numCell);
			cellEntete.setCellValue(Util.getOrCreateCell(rowModel, numCell).getStringCellValue());
			cellEntete.setCellStyle(rowModel.getCell(numCell).getCellStyle());
		}
		
		// Les parties (cellules fusionn�es)
		// 
		for (int numPartie = 0; numPartie < concours.getParties().size(); numPartie++) {
			int index = firstColumnPartieInSheetTeam + numPartie * 2;
			sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(),row.getRowNum(), index, index + 1));
			Cell cellEntete1 = row.createCell(index);
			Cell cellEntete2 = row.createCell(cellEntete1.getColumnIndex() + 1);
			cellEntete1.setCellValue("Partie " + (numPartie + 1));
			if (numPartie % 2 == 0) {
				cellEntete1.setCellStyle(rowModel.getCell(firstColumnPartieInSheetTeam).getCellStyle());
				cellEntete2.setCellStyle(rowModel.getCell(firstColumnPartieInSheetTeam + 1).getCellStyle());	
			} else {
				cellEntete1.setCellStyle(rowModel.getCell(firstColumnPartieInSheetTeam + 2).getCellStyle());
				cellEntete2.setCellStyle(rowModel.getCell(firstColumnPartieInSheetTeam + 3).getCellStyle());	
			}
		}
		
		// Les totaux : // 3 ou 4 colonnes colonnes ({ordre}/equipe/nom1/Nom2)  + 2 colonnes par partie
		// la feuille modele � 6 partie
		int firstIndexCellInRowModel = firstColumnPartieInSheetTeam + 6 * 2;
		int numCellSum = firstColumnPartieInSheetTeam + concours.getParties().size() * 2;
		// Parties Gagn�es, PP, PC, DIFF
		for (int numCell = 0; numCell < 4; numCell++) {
			Cell cellEntete = row.createCell(numCellSum + numCell);
			Cell cellModel = Util.getOrCreateCell(rowModel, firstIndexCellInRowModel + numCell);
			cellEntete.setCellValue(cellModel.getStringCellValue());
			cellEntete.setCellStyle(cellModel.getCellStyle());
		}
		// Adversaire
		numCellSum = numCellSum + 3;
		for(int iPartie = 1;iPartie <= concours.getParties().size();iPartie++) {
			Cell cellEntete = row.createCell(numCellSum + iPartie);
			Cell cellModel = Util.getOrCreateCell(rowModel, firstIndexCellInRowModel + 3);
			cellEntete.setCellValue(cellModel.getStringCellValue());
			cellEntete.setCellStyle(cellModel.getCellStyle());
			cellEntete.setCellValue("Adv" + iPartie);
		}
		
		
		return row;
	}
	
	
	
	protected String searchFormulaScore(Equipe equipe, int numPartie) {
		return searchFormulaScore(equipe, numPartie, 5, 4);
	}
	
	protected String searchFormulaScoreAdverse(Equipe equipe, int numPartie) {
		return searchFormulaScore(equipe, numPartie, 6, 3);
	}
	
	protected String searchFormulaScore(Equipe equipe, int numPartie, int colonneScore1, int colonneScore2) {
		StringBuilder formula = new StringBuilder("");
		// SHEET_TEAM_FIRSTLINE + 1 => Duplication de la ligne d'entete
		String columnNumberInTeam = "A";
		if (tableauFinale) {
			columnNumberInTeam = "B";
		}
		String cellSearched = columnNumberInTeam + (SHEET_TEAM_FIRSTLINE + equipe.getNumero()); 
		// =SI(ESTNA(RECHERCHEV($A4;Partie1!$B$2:$G$22;5;FAUX));RECHERCHEV($A4;Partie1!$D$2:$G$22;4;FAUX);RECHERCHEV($A4;Partie1!$B$2:$G$22;5;FAUX))
		// 
		int nbMatch = concours.getEquipes().size() / 2;
		String rechercheEquipe1 = "VLOOKUP(" + cellSearched + "," + SHEET_PARTIE_NAME + numPartie + "!$B$"+SHEET_PARTIE_FIRSTLINE+":$G$"+(SHEET_PARTIE_FIRSTLINE + nbMatch)+","+colonneScore1+",FALSE)";
		String rechercheEquipe2 = "VLOOKUP(" + cellSearched + "," + SHEET_PARTIE_NAME + numPartie + "!$D$"+SHEET_PARTIE_FIRSTLINE+":$G$"+(SHEET_PARTIE_FIRSTLINE + nbMatch)+","+colonneScore2+",FALSE)";
		formula.append("IF(");
		// COND
		formula.append("ISNA(" + rechercheEquipe1 + "),");
		// THEN
		formula.append(rechercheEquipe2 + ",");
		// ELSE
		formula.append(rechercheEquipe1);
		formula.append(")"); // END IF
		
		return formula.toString();
	}
	
	/**
	 * Construit formule qui renvoie le code de l'adversaire � la partie {numPartie}, (0 pour Exempt)
	 * @param equipe Num�ro de l'�quipe dont on cherche l'adversaire
	 * @param numPartie Num�ro de partie � laquelle on recherche
	 * @return
	 */
	protected String searchFormulaAdversaire(Equipe equipe, int numPartie) {
		StringBuilder formula = new StringBuilder("");
		
		String columnNumberInTeam = "A";
		if (tableauFinale) {
			columnNumberInTeam = "B";
		}
		String cellSearched = columnNumberInTeam + (SHEET_TEAM_FIRSTLINE + equipe.getNumero()); 
		// =SI(ESTNA(RECHERCHEV($A4;Partie1!$B$2:$G$22;5;FAUX));RECHERCHEV($A4;Partie1!$D$2:$G$22;4;FAUX);RECHERCHEV($A4;Partie1!$B$2:$G$22;5;FAUX))
		// 
		int nbMatch = concours.getEquipes().size() / 2;
		
		int colonneAdv1 = 3;
		int colonneAdv2 = 5;
		
		String rechercheEquipe1 = "VLOOKUP(" + cellSearched + "," + SHEET_PARTIE_NAME + numPartie + "!$B$"+SHEET_PARTIE_FIRSTLINE+":$G$"+(SHEET_PARTIE_FIRSTLINE + nbMatch)+","+colonneAdv1+",FALSE)";
		// L'�quipe 1 a �t� recopi� dans la colonne H
		String rechercheEquipe2 = "VLOOKUP(" + cellSearched + "," + SHEET_PARTIE_NAME + numPartie + "!$D$"+SHEET_PARTIE_FIRSTLINE+":$H$"+(SHEET_PARTIE_FIRSTLINE + nbMatch)+","+colonneAdv2+",FALSE)";
		
		formula.append("IF(");
		// COND
		formula.append("ISNA(" + rechercheEquipe1 + "),");
		// THEN
		formula.append(rechercheEquipe2 + ",");
		// ELSE
		formula.append(rechercheEquipe1);
		formula.append(")"); // END IF
		
		return formula.toString();
	}
	

}
