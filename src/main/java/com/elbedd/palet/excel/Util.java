package com.elbedd.palet.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Methodes utilitaires pour excel
 * @author Laurent BRAUD
 *
 */
public class Util {

	
	protected static Row getOrCreateRow(Sheet sheet, int rownum) {
		Row ret = sheet.getRow(rownum);
		if (ret == null) {
			ret = sheet.createRow(rownum);
		}
		return ret;
		
	}
	
	protected static Cell getOrCreateCell(Row row, int cellnum) {
		Cell ret = row.getCell(cellnum);
		if (ret == null) {
			ret = row.createCell(cellnum);
		}
		return ret;
		
	}
}
