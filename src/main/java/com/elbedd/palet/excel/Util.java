// The MIT License (MIT)
// Copyright (c) 2016-2018 Laurent BRAUD
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

package com.elbedd.palet.excel;

import org.apache.poi.ss.usermodel.Cell;
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
	
	public static String getColumnLetter(int columnIndex) {
        StringBuilder columnLetter = new StringBuilder();
        while (columnIndex >= 0) {
            int remainder = columnIndex % 26;
            columnLetter.insert(0, (char) (remainder + 'A'));
            columnIndex = (columnIndex / 26) - 1;
        }
        return columnLetter.toString();
    }
}
