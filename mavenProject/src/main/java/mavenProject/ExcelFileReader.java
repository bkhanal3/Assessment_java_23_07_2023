package mavenProject;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileReader {
	public static void main(String[] args) throws IOException {

		File file = new File("/Users/bisasonkhanal/Downloads/EC_newWorkspace/mavenProject/src/main/resources/employee.xlsx");
		FileInputStream fileInputStream = new FileInputStream(file);

		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		int noOfSheet = workbook.getNumberOfSheets();
		System.out.println("Number of Sheet present in Xls file is : " + noOfSheet);

		XSSFSheet xssfSheet = workbook.getSheetAt(0);

		Iterator<Row> rowIterator = xssfSheet.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();

			Iterator<Cell> columnIterator = row.cellIterator();
			while (columnIterator.hasNext()) {
				Cell cell = columnIterator.next();
				switch (cell.getCellType()) {
				case STRING:
					System.out.println(cell.getStringCellValue() + "\t");
					break;
				case NUMERIC:
					System.out.println(cell.getNumericCellValue() + "\t");
					break;
				case BOOLEAN:
					System.out.println(cell.getBooleanCellValue() + "\t");
					break;
				}
			}
		}

	}

}
