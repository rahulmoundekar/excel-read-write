package com.app;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.app.model.Employee;

public class ExcelWrite {
	private static String[] columns = { "Name", "Email", "Date Of Birth", "Salary" };

	private static final String FILE_NAME = "MyFirstExcel.xlsx";
	XSSFWorkbook workbook = null;

	public List<Employee> getEmployees() {
		List<Employee> employees = new ArrayList<Employee>();
		Calendar dateOfBirth = Calendar.getInstance();
		dateOfBirth.set(1992, 7, 21);
		employees.add(new Employee("Rajeev Singh", "rajeev@example.com", dateOfBirth.getTime(), 1200000.0));

		dateOfBirth.set(1965, 10, 15);
		employees.add(new Employee("Thomas cook", "thomas@example.com", dateOfBirth.getTime(), 1500000.0));

		dateOfBirth.set(1987, 4, 18);
		employees.add(new Employee("Steve Maiden", "steve@example.com", dateOfBirth.getTime(), 1800000.0));
		return employees;
	}

	public void writeToExcel() throws IOException {
		workbook = new XSSFWorkbook();
		// Create a Sheet
		Sheet sheet = workbook.createSheet("Employee");

		// Create a Row
		Row headerRow = sheet.createRow(0);
		// Create cells
		for (int i = 0; i < columns.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(columns[i]);
		}

		/*
		 * CreationHelper helps us create instances of various things like DataFormat,
		 * Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way
		 */
		CreationHelper createHelper = workbook.getCreationHelper();

		// Create Cell Style for formatting Date
		CellStyle dateCellStyle = workbook.createCellStyle();
		dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

		// Create Other rows and cells with employees data
		int rowNum = 1;
		for (Employee employee : getEmployees()) {
			Row row = sheet.createRow(rowNum++);
			row.createCell(0).setCellValue(employee.getName());
			row.createCell(1).setCellValue(employee.getEmail());

			Cell dateOfBirthCell = row.createCell(2);
			dateOfBirthCell.setCellValue(employee.getDateOfBirth());
			dateOfBirthCell.setCellStyle(dateCellStyle);

			row.createCell(3).setCellValue(employee.getSalary());
		}

		// Resize all columns to fit the content size
		for (int i = 0; i < columns.length; i++) {
			sheet.autoSizeColumn(i);
		}

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream(new File(FILE_NAME));
		workbook.write(fileOut);
		fileOut.close();

		// Closing the workbook
		workbook.close();
	}

	public static void main(String[] args) {
		ExcelWrite write = new ExcelWrite();
		try {
			write.writeToExcel();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
