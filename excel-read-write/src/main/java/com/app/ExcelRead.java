package com.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.app.model.Employee;
import com.sun.org.apache.xerces.internal.impl.xpath.regex.ParseException;

public class ExcelRead {

	private static final String FILE_NAME = "MyFirstExcel.xlsx";
	XSSFWorkbook workbook = null;

	private void readSheet() throws ParseException, IOException {
		FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
		Workbook workbook = new XSSFWorkbook(excelFile);
		Sheet sheet = workbook.getSheetAt(0);

		Iterator<Row> rowItr = sheet.iterator();
		List<Employee> empList = new ArrayList<Employee>();
		// Iterate through rows
		while (rowItr.hasNext()) {
			Employee emp = new Employee();
			Row row = rowItr.next();
			// skip header (First row)
			if (row.getRowNum() == 0) {
				continue;
			}
			Iterator<Cell> cellItr = row.cellIterator();
			// Iterate each cell in a row
			while (cellItr.hasNext()) {

				Cell cell = cellItr.next();
				int index = cell.getColumnIndex();
				switch (index) {
				case 0:
					emp.setName((String) getValueFromCell(cell));
					break;
				case 1:
					emp.setEmail((String) getValueFromCell(cell));
					break;
				case 2:
					emp.setDateOfBirth((Date) getValueFromCell(cell));
					break;
				case 3:
					emp.setSalary((double) getValueFromCell(cell));
					break;
				}
			}
			empList.add(emp);
		}
		for (Employee emp : empList) {
			System.out.println(emp.toString());
		}
		workbook.close();
	}

	// Utility method to get cell value based on cell type
	@SuppressWarnings("deprecation")
	private Object getValueFromCell(Cell cell) {
		switch (cell.getCellTypeEnum()) {
		case STRING:
			return cell.getStringCellValue();
		case BOOLEAN:
			return cell.getBooleanCellValue();
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				return cell.getDateCellValue();
			}
			return cell.getNumericCellValue();
		case FORMULA:
			return cell.getCellFormula();
		case BLANK:
			return "";
		default:
			return "";
		}
	}

	public static void main(String[] args) throws ParseException, IOException {
		ExcelRead excelRead = new ExcelRead();
		excelRead.readSheet();
	}

}
