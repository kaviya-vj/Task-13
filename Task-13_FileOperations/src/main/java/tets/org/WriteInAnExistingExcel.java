package tets.org;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteInAnExistingExcel {

	public static void main(String[] args) throws IOException {

		// Load workbook
		XSSFWorkbook wb = new XSSFWorkbook();
		// Create a worksheet
		XSSFSheet sh = wb.createSheet("Sheet1");

		// Print the name of loaded sheet
		// System.out.println(sheet.getSheetName());

		ArrayList<Object[]> ofz = new ArrayList();
		ofz.add(new Object[] { "Unit", "Floor", "Password" });
		ofz.add(new Object[] { "Unit1", 30, "P1" });
		ofz.add(new Object[] { "Unit2", 28, "P2" });
		ofz.add(new Object[] { "Unit3", 35, "P3" });
		ofz.add(new Object[] { "Unit4", 37, "P4" });

		int rownum = 0;
		for (Object[] emp : ofz) {

			XSSFRow row = sh.createRow(rownum++);
			int cellnum = 0;
			for (Object value : emp) {
				XSSFCell cell = row.createCell(cellnum++);
				if (value instanceof String)
					cell.setCellValue((String) value);
				if (value instanceof Integer)
					cell.setCellValue((Integer) value);
				if (value instanceof Boolean)
					cell.setCellValue((Boolean) value);

			}

		}
		// specified the location of excel file
		File src = new File("D:\\ExcelReader\\WriteAnExcel.xlsx");
		// Load file
		OutputStream fls = new FileOutputStream(src);
		wb.write(fls);
		fls.close();
		System.out.println("WriteAnExcel.xlsx file written successfully");
	}

}
