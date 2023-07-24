package com.fileReadWrite;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

public class ReadExcelFile {
	private static final int NUMERIC = 0;
	private static final int BOOLEAN = 0;

	public static void main(String[] args) {
		try {

			// Read the .xlsx / .xls file
			File file = new File(
					"/Users/bisasonkhanal/Downloads/EC_newWorkspace/mavenProject/src/src/test/resources/data/students.xlsx");
			FileInputStream fileInputStream = new FileInputStream(file);

			// Create workbook instance and capture the data from excel file
			XSSFWorkbook workbook = new XSSFWorkbook();
			int noOfSheet = workbook.getNumberOfSheets();
			System.out.println("Number of Sheet present in Xls file is : " + noOfSheet);

			// Read the first sheet
			XSSFSheet xssfSheet = workbook.getSheetAt(0);

			// Read the rows one by one
			Iterator<Row> rowIterator = xssfSheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				// fetch all the column / cell
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

		} catch (Exception e) {
			System.err.println("Error is :" + e.getMessage());
		}
	}
}
