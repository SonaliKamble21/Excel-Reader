package exceloperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//workbook-->sheet-->rows-->cells
public class WritingExcel {

	public static void main(String[] args) throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook();

		XSSFSheet sheet = workbook.createSheet("Emp Info");

		Object empdata[][] = { 
				{ "EmpID", "Name", "Job" },
				{ "101", "John", "Tester" },
				{ "102", "Smith", "Engineer" }, 
				};

		// using for loop
//		int rows = empdata.length;
//		int cols = empdata[0].length;
//
//		System.out.println(rows);
//		System.out.println(cols);
//
//		for (int r = 0; r < rows; r++) {
//
//			XSSFRow row = sheet.createRow(r);
//
//			for (int c = 0; c < cols; c++) {
//
//				XSSFCell cell = row.createCell(c);
//				Object value = empdata[r][c];
//
//				if (value instanceof String)
//					cell.setCellValue((String) value);
//				if (value instanceof Boolean)
//					cell.setCellValue((Integer) value);
//				if (value instanceof Boolean)
//					cell.setCellValue((Boolean) value);
//			}
//		}
		
		//2. using for each loop:
		
		int rowCount=0;
		
		for(Object emp[]:empdata) {
			
			XSSFRow row = sheet.createRow(rowCount++);
			int colCount=0;
				for(Object value:emp) {
					
				 XSSFCell cell = row.createCell(colCount++);
				 
				 if (value instanceof String)
						cell.setCellValue((String) value);
				 if (value instanceof Boolean)
						cell.setCellValue((Integer) value);
				 if (value instanceof Boolean)
						cell.setCellValue((Boolean) value);
				 
			}
		}
		
		
		
		String filePath = ".\\datafiles\\employee.xlsx";
		FileOutputStream fout = new FileOutputStream(filePath);
		workbook.write(fout);
		
		fout.close();
		System.out.println("employee.xlsx file written successfully");

	}
}
