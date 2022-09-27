// The MIT License (MIT)
// Copyright (c) 2016-2018 Laurent BRAUD
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
	
	private final static String SHEET_TEAM_NAME = "Equipes";
	private final static int SHEET_TEAM_FIRSTLINE = 4;
	
	private final static int SHEET_PARTIE_FIRSTLINE = 3;
	
	
	private final static int SCORE_WIN = 11; 
	
	Concours concours;

	public Generator(Concours concours) {
		this.concours = concours;
	}
	
	private Workbook loadModeleExcel() {
		Workbook ret = null;
		try {
			// src/main/resources/com/elbedd/palet/excel
			InputStream is = getClass().getResourceAsStream("modele.xls");
			ret = new HSSFWorkbook(is);
			
		} catch(Exception e) {
			// Exception : Fichier modèle non chargé 
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
		Workbook wb = loadModeleExcel();
		
		//setGrayStyleSheetFromWb(wb);
		//setVerticalStyleSheetFromWb(wb);
		writeSheetEquipes(wb);
		
		String referenceStart = SHEET_TEAM_NAME + REF_SHEET + "A" + (SHEET_TEAM_FIRSTLINE + 1);
		int lastLine = SHEET_TEAM_FIRSTLINE + concours.getEquipes().size();
		String referenceEnd = "C" + lastLine;
		String plageEquipe = referenceStart + REF_RANGE + referenceEnd;//"Equipes!A3:C43";
		
		for (Partie partie : concours.getParties()) {
			//wb.createSheet("Partie" + partie.getNumero());
			Sheet sheet = wb.cloneSheet(1);
			wb.setSheetName(wb.getSheetIndex(sheet), "Partie" + partie.getNumero());
			
			
			writeSheetPartieHeader(sheet, partie);

			for (Match match : partie.getMatchs()) {
				writeSheetPartieMatch(sheet, match, plageEquipe);
			}
		}
		wb.removeSheetAt(1);
		wb.setSheetOrder("ClassementQualif", wb.getNumberOfSheets()-1);
		writeMoreInSheetEquipes(wb);
		wb.setActiveSheet(0);

		try (FileOutputStream out = new FileOutputStream(fileName)) {
			wb.write(out);
			out.close();
		} catch (Exception e) {
			throw e;
		}
		wb.close();
	}

	protected void writeSheetPartieHeader(Sheet sheet, Partie partie) {
		Row row = Util.getOrCreateRow(sheet, 0);
		
		Cell cell = Util.getOrCreateCell(row, 2);
		cell.setCellValue("Partie n°" + partie.getNumero());
		
		
	}
	
	protected void writeSheetPartieMatch(Sheet sheet, Match match, String rangeEquipe) {
		int numRow = match.getNumeroPlaque() + 1;
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
		cell.setCellValue(match.getEquipeA().getNumero());
		cell.setCellStyle(rowModel.getCell(i).getCellStyle());
		
		i++;
		cell = Util.getOrCreateCell(row, i);//row.createCell(i++, CellType.FORMULA);
		//cell.setCellFormula("VLOOKUP(" + match.getEquipeA().getNumero() + ",A3:C43,3,FALSE)");
		cell.setCellType(CellType.FORMULA);
		searchFormulaTeamNames(match.getEquipeA(), rangeEquipe, cell, "B");
		cell.setCellStyle(rowModel.getCell(i).getCellStyle());
		
		i++;
		cell = Util.getOrCreateCell(row, i);
		cell.setCellStyle(rowModel.getCell(i).getCellStyle());
		if (match.getEquipeB() != null) {
			// Non exempt
			cell.setCellValue(match.getEquipeB().getNumero());
		}
		
		i++;
		cell = Util.getOrCreateCell(row, i);
		cell.setCellStyle(rowModel.getCell(i).getCellStyle());
		if (match.getEquipeB() != null) {
			// Non exempt
			searchFormulaTeamNames(match.getEquipeB(), rangeEquipe, cell, "D");
		}
		
		// Score (vide)
		i++;
		cell = Util.getOrCreateCell(row, i);
		cell.setCellStyle(rowModel.getCell(i).getCellStyle());
		i++;
		cell = Util.getOrCreateCell(row, i);
		cell.setCellStyle(rowModel.getCell(i).getCellStyle());
		
		
	}

	protected void searchFormulaTeamNames(Equipe equipe, String rangeEquipe, Cell cell, String columnNumEquipe) {
		StringBuilder formula = new StringBuilder("CONCATENATE(");
		int rowFormula = cell.getRow().getRowNum() + 1;
		formula.append("VLOOKUP(" + columnNumEquipe + rowFormula + ", "+ rangeEquipe +",2,FALSE)");//First Player
		formula.append(",");//SEP CONCAT
		formula.append("CHAR(10)");//Retour Chariot
		formula.append(",");//SEP CONCAT
		formula.append("VLOOKUP(" + columnNumEquipe + rowFormula  + ", "+ rangeEquipe +",3,FALSE)");//2nd Player (colonne 3)
		formula.append(")");//END CONCAT
		
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
		// Dans cette cellule, remplacer le nom du joueur 1 par le nom de l'équipe
		Cell cell = row.getCell(numCell++);
		cell.setCellValue(equipe.getNumero());
		// Remplacer le nom du joueur 2 par jouer1+2
		cell = row.getCell(numCell);
		row.removeCell(cell);
		cell = row.createCell(numCell, CellType.FORMULA);
		searchFormulaTeamNames(equipe, plageEquipe, cell, "A");
		//cell.setCellStyle(verticalCellStyleSheet);
		
		
	}
	
	private void writeMoreInSheetEquipes(Workbook wb) {
		Sheet sheet = wb.getSheet(SHEET_TEAM_NAME);
		
		// On recopie la ligne 3 de l'excel en ajustant de 6 à nbPartie.
		
		Row rowModel = Util.getOrCreateRow(sheet, 2);// Ligne 3
		Row rowHeader = writeHeaderInSheetEquipe(sheet, rowModel);
		sheet.removeRow(rowModel);
		
		for (Equipe equipe : concours.getEquipes().values()) {
			int numLigne = SHEET_TEAM_FIRSTLINE + equipe.getNumero();	
			Row row = sheet.getRow(numLigne - 1);
			
			int numCell = 3;
			// ajouter le résultat des parties
			for (int iPartie = 0; iPartie < concours.getParties().size(); iPartie++) {
				// Score équipe
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
			char[] tabLettre = {'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P','Q','R' };
			int indexLettre = 3;
			
			StringBuilder sbWin = new StringBuilder("");
			StringBuilder sbPour = new StringBuilder("");
			StringBuilder sbContre = new StringBuilder("");
			
			for (int iPartie = 0; iPartie < concours.getParties().size(); iPartie++) {
				sbWin.append("+IF(" + tabLettre[indexLettre] + numLigne + "<" + SCORE_WIN +",0,1)");
				sbPour.append("+" +tabLettre[indexLettre] + numLigne);
				indexLettre ++;
				sbContre.append("+" +tabLettre[indexLettre] + numLigne);
				indexLettre ++;
			}
			
			
			Cell cellWin = row.createCell(numCell++, CellType.FORMULA);
			String formula = sbWin.substring(1);// Remove the first +
			cellWin.setCellFormula(formula);
			cellWin.setCellStyle(rowHeader.getCell(cellWin.getColumnIndex()).getCellStyle());
			
			Cell cellPP = row.createCell(numCell++, CellType.FORMULA);
			formula = sbPour.substring(1);// Remove the first +
			cellPP.setCellFormula(formula);
			cellPP.setCellStyle(rowHeader.getCell(cellPP.getColumnIndex()).getCellStyle());
			
			Cell cellPC = row.createCell(numCell++, CellType.FORMULA);
			formula = sbContre.substring(1);// Remove the first +
			cellPC.setCellFormula(formula);
			cellPC.setCellStyle(rowHeader.getCell(cellPC.getColumnIndex()).getCellStyle());
			
			
		}
		
	}

	/**
	 *
	 * @param sheet
	 * @return
	 */
	private Row writeHeaderInSheetEquipe(Sheet sheet, Row rowModel) {
		// 3 colonnes (equipe/nom1/Nom2)  + 2 colonnes par partie
		int numCellSum = 3 + concours.getParties().size() * 2;
		
		Row row = sheet.createRow(rowModel.getRowNum()+1);
		row.setHeight(rowModel.getHeight());
		for (int numCell = 0; numCell < 3; numCell++) {
			Cell cellEntete = row.createCell(numCell);
			cellEntete.setCellValue(Util.getOrCreateCell(rowModel, numCell).getStringCellValue());
			cellEntete.setCellStyle(rowModel.getCell(numCell).getCellStyle());
		}
		
		// Les parties (cellules fusionnées)
		// 
		for (int numPartie = 0; numPartie < concours.getParties().size(); numPartie++) {
			sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(),row.getRowNum(),3 + numPartie * 2 ,3 + numPartie * 2 + 1));
			Cell cellEntete1 = row.createCell(3 + numPartie * 2);
			Cell cellEntete2 = row.createCell(cellEntete1.getColumnIndex() + 1);
			cellEntete1.setCellValue("Partie " + (numPartie + 1));
			if (numPartie % 2 == 0) {
				cellEntete1.setCellStyle(rowModel.getCell(3).getCellStyle());
				cellEntete2.setCellStyle(rowModel.getCell(4).getCellStyle());	
			} else {
				cellEntete1.setCellStyle(rowModel.getCell(5).getCellStyle());
				cellEntete2.setCellStyle(rowModel.getCell(6).getCellStyle());	
			}
		}
		
		// Les totaux
		// la feuille modele à 6 partie
		int firstIndexCellInRowModel = 3 + 6 * 2;;
		for (int numCell = 0; numCell < 3; numCell++) {
			Cell cellEntete = row.createCell(numCellSum + numCell);
			Cell cellModel = Util.getOrCreateCell(rowModel, firstIndexCellInRowModel + numCell);
			cellEntete.setCellValue(cellModel.getStringCellValue());
			cellEntete.setCellStyle(cellModel.getCellStyle());
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
		String cellSearched = "A" + (SHEET_TEAM_FIRSTLINE + equipe.getNumero()); 
		// =SI(ESTNA(RECHERCHEV($A4;Partie1!$B$2:$G$22;5;FAUX));RECHERCHEV($A4;Partie1!$D$2:$G$22;4;FAUX);RECHERCHEV($A4;Partie1!$B$2:$G$22;5;FAUX))
		// 
		int nbMatch = concours.getEquipes().size() / 2;
		String rechercheEquipe1 = "VLOOKUP(" + cellSearched + ",Partie" + numPartie + "!$B$"+SHEET_PARTIE_FIRSTLINE+":$G$"+(SHEET_PARTIE_FIRSTLINE + nbMatch)+","+colonneScore1+",FALSE)";
		String rechercheEquipe2 = "VLOOKUP(" + cellSearched + ",Partie" + numPartie + "!$D$"+SHEET_PARTIE_FIRSTLINE+":$G$"+(SHEET_PARTIE_FIRSTLINE + nbMatch)+","+colonneScore2+",FALSE)";
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
