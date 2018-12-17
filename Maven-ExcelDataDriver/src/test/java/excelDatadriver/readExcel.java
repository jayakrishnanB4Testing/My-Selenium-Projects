package excelDatadriver;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class readExcel {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fileIn = new FileInputStream("C:\\Users\\j.a.balakrishnan\\Google Drive\\JK Automation\\Selenium Automation\\TestData.xlsx");
		XSSFWorkbook eworkBook = new XSSFWorkbook(fileIn);
		
		XSSFSheet sheet = eworkBook.getSheet("sample");
		
		Iterator<Row> rows = sheet.iterator();
		Row firstRow = rows.next();
		
		Iterator<Cell> cells = firstRow.cellIterator();
		int colNum = 0;
		while(cells.hasNext()) {
			Cell cell = cells.next();
			if(cell.getStringCellValue().equalsIgnoreCase("user")){
				colNum = cell.getColumnIndex();
			}
			
		}
		
//		System.out.println(colNum);
		
		while(rows.hasNext()) {
			Row row = rows.next();
			if(row.getCell(colNum).getStringCellValue().equalsIgnoreCase("trip")) {
				Iterator<Cell> Cells = row.cellIterator();
				while(Cells.hasNext()) {
					Cell Cell = Cells.next();
					System.out.println(Cell.getStringCellValue());
				}
			}
			
		}

	}

}
