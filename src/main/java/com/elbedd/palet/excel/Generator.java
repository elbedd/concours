package com.elbedd.palet.excel;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

import com.elbedd.palet.model.Concours;
import com.elbedd.palet.model.Equipe;
import com.elbedd.palet.model.Match;
import com.elbedd.palet.model.Partie;

public class Generator {
	
	private final static String REF_SHEET = "!";// In apachePOI ?
	private final static String REF_RANGE = ":";// In apachePOI ?
	
	private final static String SHEET_TEAM_NAME = "Equipes";
	private final static int SHEET_TEAM_FIRSTLINE = 3; 
	
	
	
	Concours concours;

	public Generator(Concours concours) {
		this.concours = concours;
	}

	/**
	 * TODO : review Exception
	 * 
	 * @param fileName
	 * @throws Exception
	 */
	public void generateExcel(String fileName) throws Exception {
		Workbook wb = new HSSFWorkbook();
		writeSheetEquipes(wb);
		
		
		String referenceStart = SHEET_TEAM_NAME + REF_SHEET + "A" + SHEET_TEAM_FIRSTLINE;
		int lastLine = SHEET_TEAM_FIRSTLINE + concours.getEquipes().size() - 1;
		String referenceEnd = "C" + lastLine;
		String plageEquipe = referenceStart + REF_RANGE + referenceEnd;//"Equipes!A3:C43";
		
		CellStyle styleJoueurEquipe = wb.createCellStyle();
		styleJoueurEquipe.setVerticalAlignment(VerticalAlignment.JUSTIFY);
		
		for (Partie partie : concours.getParties()) {
			Sheet sheet = wb.createSheet("Partie" + partie.getNumero());
			writeSheetPartieHeader(sheet);

			for (Match match : partie.getMatchs()) {
				writeSheetPartieMatch(sheet, match, plageEquipe, styleJoueurEquipe);
			}
		}
		

		try (FileOutputStream out = new FileOutputStream(fileName)) {
			wb.write(out);
			out.close();
		} catch (Exception e) {
			throw e;
		}
		wb.close();
	}

	protected void writeSheetPartieHeader(Sheet sheet) {
		Row row = sheet.createRow(0);
		int i = -1;
		Cell cell = row.createCell(++i, CellType.STRING);
		cell.setCellValue("Plaque");
		
		cell = row.createCell(++i, CellType.STRING);
		cell.setCellValue("Equipe 1");
		
		cell = row.createCell(++i, CellType.STRING);
		cell.setCellValue("Noms Equipe 1");
		sheet.setColumnWidth(i, (short) (sheet.getColumnWidth(i) * 3));
		
		cell = row.createCell(++i, CellType.STRING);
		cell.setCellValue("Equipe 2");
		
		cell = row.createCell(++i, CellType.STRING);
		cell.setCellValue("Noms Equipe 2");
		sheet.setColumnWidth(i, (short) (sheet.getColumnWidth(i) * 3));
	}
	
	protected void writeSheetPartieMatch(Sheet sheet, Match match, String rangeEquipe, CellStyle styleJoueurEquipe) {
		Row row = sheet.createRow(match.getNumeroPlaque());
		row.setHeight((short) (row.getHeight() * 2)); 
		
		int i = 0;
		Cell cell = row.createCell(i++, CellType.STRING);
		cell.setCellValue(match.getNumeroPlaque());
		
		cell = row.createCell(i++, CellType.NUMERIC);
		cell.setCellValue(match.getEquipeA().getNumero());

		cell = row.createCell(i++, CellType.FORMULA);
		//cell.setCellFormula("VLOOKUP(" + match.getEquipeA().getNumero() + ",A3:C43,3,FALSE)");
		searchFormulaTeamNames(match.getEquipeA(), rangeEquipe, cell);
		
		cell.setCellStyle(styleJoueurEquipe);
		

		if (match.getEquipeB() != null) {
			// Non exempt
			cell = row.createCell(i++, CellType.NUMERIC);
			cell.setCellValue(match.getEquipeB().getNumero());

			cell = row.createCell(i++, CellType.STRING);
			cell.setCellValue("Noms Equipe 2");
			searchFormulaTeamNames(match.getEquipeB(), rangeEquipe, cell);
			cell.setCellStyle(styleJoueurEquipe);
			
		}
		
	}

	protected void searchFormulaTeamNames(Equipe equipe, String rangeEquipe, Cell cell) {
		StringBuilder formula = new StringBuilder("CONCATENATE(");
		formula.append("VLOOKUP(" + equipe.getNumero() + ", "+ rangeEquipe +",2,FALSE)");//First Player
		formula.append(",");//SEP CONCAT
		formula.append("CHAR(10)");//Retour Chariot
		formula.append(",");//SEP CONCAT
		formula.append("VLOOKUP(" + equipe.getNumero() + ", "+ rangeEquipe +",3,FALSE)");//2nd Player (colonne 3)
		formula.append(")");//END CONCAT
			
		
		cell.setCellFormula(formula.toString());
	}

	protected void writeSheetEquipes(Workbook wb) {
		Sheet sheet = wb.createSheet(SHEET_TEAM_NAME);
		Row row = sheet.createRow(0);
		
		Cell cell = row.createCell(0, CellType.STRING);
		cell.setCellValue("Concours");
		
		cell = row.createCell(2, CellType.NUMERIC);
		cell.setCellValue(concours.getParties().size());
		
		cell = row.createCell(3, CellType.STRING);
		cell.setCellValue("parties qualificatives");
		
		for (Equipe equipe : concours.getEquipes().values()) {
			writeEquipe(sheet, equipe);
		}
		
		/*Name namedCell = wb.createName();
		namedCell.setNameName("EquipeRange");
		String referenceStart = SHEET_TEAM_NAME + REF_SHEET + "A" + SHEET_TEAM_FIRSTLINE;
		int lastLine = SHEET_TEAM_FIRSTLINE + concours.getEquipes().size() - 1;
		String referenceEnd = "C" + lastLine;
		System.out.println(referenceStart + REF_RANGE + referenceEnd);
		namedCell.setRefersToFormula(referenceStart + REF_RANGE + referenceEnd);*/
		
	}
	
	protected void writeEquipe(Sheet sheet, Equipe equipe) {
		Row row = sheet.createRow(SHEET_TEAM_FIRSTLINE + equipe.getNumero() - 2);
		
		Cell cell = row.createCell(0, CellType.NUMERIC);
		cell.setCellValue(equipe.getNumero());
		
		cell = row.createCell(1, CellType.STRING);
		cell.setCellValue("Player1 Team" + equipe.getNumero());// 1st player
		
		cell = row.createCell(2, CellType.STRING);
		cell.setCellValue("Player2 Team" + equipe.getNumero());// 2nd player
	}

}
