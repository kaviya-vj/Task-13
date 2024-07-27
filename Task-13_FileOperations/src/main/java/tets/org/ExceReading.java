package tets.org;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExceReading {

	public static void main(String[] args) throws IOException {
		// specified the location of excel file
		File src = new File("D:\\ExcelReader\\Demo.xlsx");
		// Load file
		FileInputStream fls = new FileInputStream(src);
		// Load workbook
		XSSFWorkbook wb = new XSSFWorkbook(fls);
		// Load worksheet
		XSSFSheet sheet = wb.getSheet("Sheet1");
		// Print the name of the sheet
		System.out.println(sheet.getSheetName());
		// Print username from excelsheet
		System.out.println(sheet.getRow(0).getCell(0).getStringCellValue());
		// Print P2 from excelsheet
		System.out.println(sheet.getRow(2).getCell(1).getStringCellValue());
		// Print total number of rows
		System.out.println(sheet.getPhysicalNumberOfRows());
		// Print total number of coloumn
		System.out.println(sheet.getRow(0).getPhysicalNumberOfCells());
		// to print all the element
		int rows = (sheet.getLastRowNum() - sheet.getFirstRowNum()) + 1;
		System.out.println(rows);
		int columns = sheet.getRow(0).getLastCellNum();
		System.out.println(columns);
		for (int i = 0; i < rows; i++) {
			for (int j = 0; j < columns; j++) {
				System.out.println(sheet.getRow(i).getCell(j).getStringCellValue());

			}

		}

	}

}
