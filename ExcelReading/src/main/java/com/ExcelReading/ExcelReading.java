package com.ExcelReading;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReading {
	public ArrayList testData(String filePath, int sheetNumber) {
		ArrayList data = new ArrayList();
		try {
			FileInputStream fis = new FileInputStream(filePath);
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheetAt = wb.getSheetAt(sheetNumber);
			Iterator<Row> itr = sheetAt.iterator();
			while(itr.hasNext()) {
				Row row = itr.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				while(cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					if(cell.getCellType()==CellType.STRING) {
						data.add(cell.getStringCellValue());
					}
					}
			}
			return data;
		} catch (Exception e) {
			
		}
		return null;
		
	}
}