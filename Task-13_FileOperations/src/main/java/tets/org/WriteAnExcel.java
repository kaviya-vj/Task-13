package tets.org;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteAnExcel {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook wb = new XSSFWorkbook();

		XSSFSheet sh = wb.createSheet("Sheet1");
		ArrayList<Object[]> empData = new ArrayList();
		empData.add(new Object[] { "Name", "Age", "Email" });

		empData.add(new Object[] { "John Doe", 30, "john@tets.com" });
		empData.add(new Object[] { "Jane Doe", 28, "john@tets.com" });
		empData.add(new Object[] { "Bob Smith", 35, "jacky@example.com" });
		empData.add(new Object[] { "Swapnil", 37, "swapnil@example.com" });

		int rownum = 0;
		for (Object[] emp : empData) {

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
		String filepath = "C:\\Users\\Admin\\eclipse-workspace\\Task-13_FileOperations\\datafiles\\employees.xlsx";
		FileOutputStream file = new FileOutputStream(filepath);
		wb.write(file);

		file.close();
		System.out.println("Employees.xlsx file written successfully");
	}


	}


