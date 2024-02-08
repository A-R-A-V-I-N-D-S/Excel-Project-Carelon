package com.demo.ExcelProject;

import java.io.FileOutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateInvoices {

	public static void main(String[] args) {
		try {
			// for .xlsx format
			Workbook workbook = new XSSFWorkbook();
			// for .xsl format
			// Workbook workbook = new HSSFWorkbook();
			Sheet sh = workbook.createSheet("Invoices");
			String[] colHeadings = { "Item ID", "Item Name", "Qty", "Item Price", "Sold Date" };
			Font headerFont = workbook.createFont();
			// headerFont.setBoldweight(true);
			headerFont.setFontHeightInPoints((short) 12);
			headerFont.setColor(IndexedColors.BLACK.index);
			CellStyle headerStyle = workbook.createCellStyle();
			headerStyle.setFont(headerFont);
			// headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			// headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT);
			Row headerRow = sh.createRow(0);
			for (int i = 0; i < colHeadings.length; i++) {
				Cell cell = headerRow.createCell(i);
				cell.setCellValue(colHeadings[i]);
				cell.setCellStyle(headerStyle);
			}
			// FillData
			ArrayList<Invoices> a = createData();
			CreationHelper creationHelper = workbook.getCreationHelper();
			CellStyle dateStyle = workbook.createCellStyle();
			dateStyle.setDataFormat(creationHelper.createDataFormat().getFormat("MM/dd/yyyy"));
			int rowNum = 1;
			for (Invoices i : a) {
				Row row = sh.createRow(rowNum++);
				row.createCell(0).setCellValue(i.getItemId());
				row.createCell(1).setCellValue(i.getItemName());
				row.createCell(2).setCellValue(i.getItemQty());
				row.createCell(3).setCellValue(i.getTotalPrice());
				Cell dateCell = row.createCell(4);
				dateCell.setCellValue(i.getItemSoldDate());
				dateCell.setCellStyle(dateStyle);
			}
			// AutoSize Columns
			for (int i = 0; i < colHeadings.length; i++) {
				sh.autoSizeColumn(i);
			}
			Sheet sh2 = workbook.createSheet("Second");
			FileOutputStream fileOut = new FileOutputStream(
					"C:\\Users\\AL04040\\OneDrive - Elevance Health\\Documents\\VA_SBE\\Invoices.xlsx");
			workbook.write(fileOut);
			fileOut.close();
			System.out.println("Excel generated successfully...!");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static ArrayList<Invoices> createData() throws ParseException {
		ArrayList<Invoices> a = new ArrayList<Invoices>();
		a.add(new Invoices(1, "Book", 2, 10.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(2, "Table", 1, 50.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/02/2020")));
		a.add(new Invoices(3, "Lamp", 5, 100.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(4, "Pen", 100, 20.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/02/2020")));
		a.add(new Invoices(5, "Book", 2, 10.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(6, "Table", 1, 50.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/02/2020")));
		a.add(new Invoices(7, "Lamp", 5, 100.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(8, "Pen", 100, 20.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/02/2020")));
		a.add(new Invoices(9, "Book", 2, 10.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(10, "Table", 1, 50.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/02/2020")));
		a.add(new Invoices(11, "Lamp", 5, 100.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(12, "Pen", 100, 20.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/02/2020")));
		a.add(new Invoices(13, "Book", 2, 10.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(14, "Table", 1, 50.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/02/2020")));
		a.add(new Invoices(15, "Lamp", 5, 100.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		return a;
	}

}
