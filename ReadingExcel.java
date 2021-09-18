package exceloperations;

import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcel {

	public static void main(String[] args) {

		try {
			String filePath = ".\\datafiles\\countries.xlsx";
			FileInputStream fis = new FileInputStream(filePath);

			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheet("Sheet1");

			// using for loop:
			/*int rows = sheet.getLastRowNum();
			int cols = sheet.getRow(1).getLastCellNum();

			for (int r = 0; r <= rows; r++) {
				
				XSSFRow row = sheet.getRow(r);

				for (int c = 0; c < cols; c++) {
					
					XSSFCell cell = row.getCell(c);
					
					switch(cell.getCellType())
					{
					case STRING: System.out.print(cell.getStringCellValue());break;
					case NUMERIC:System.out.print(cell.getNumericCellValue());break;
					case BOOLEAN:System.out.print(cell.getStringCellValue());break;
					}
					System.out.print(" | ");
				}
				System.out.println();
			}
			*/
			
			//2. using iterator method
			Iterator<Row> iterator = sheet.iterator();
			
			while(iterator.hasNext()) {
				
				Row row = iterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				
				while(cellIterator.hasNext()) {
					
					Cell cell = cellIterator.next();
					
					switch(cell.getCellType())
					{
					case STRING: System.out.print(cell.getStringCellValue());break;
					case NUMERIC:System.out.print(cell.getNumericCellValue());break;
					case BOOLEAN:System.out.print(cell.getBooleanCellValue());break;
					}
					System.out.println(" | ");
				}
				System.out.println();
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
