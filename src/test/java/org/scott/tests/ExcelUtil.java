package org.scott.tests;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {

	private static List<String> readRow(Row row) {
		List<String> values = new ArrayList<>();
		
		Iterator<Cell> cellIterator = row.iterator();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            String value = "";

            if (cell.getCellTypeEnum() == CellType.STRING) {
                value = cell.getStringCellValue();
            } else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                DataFormatter dataFormatter = new DataFormatter();
                value = dataFormatter.formatCellValue(cell);
            }
            
            values.add(value);
        }
        
        return values;
	}
	
	public static Object[][] readSheet(String filename) {
		List<Object[]> sheet = new ArrayList<>(); 
		
		try {
            FileInputStream excelFile = new FileInputStream(new File(filename));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();

            // Skip header (first row of spreadsheet)
            iterator.next();

            // Read each row of data into a list
            while (iterator.hasNext()) {
                Row row = iterator.next();
                List<String> columns = readRow(row);
                
                Object[] rowObj = new Object[] {columns}; 
                sheet.add(rowObj);
            }
            
    		workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
		
		// Convert our list of row objects into an object array which is what TestNG requires
		Object[][] sheetObj = new Object[sheet.size()][];
        sheetObj = sheet.toArray(sheetObj);
        
        return sheetObj;
	}
}
