package com.demo.ExcelProject;

import java.io.FileInputStream;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Locale;
import java.util.TimeZone;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DateChecker {

	public static void main(String[] args) {
		try {
			FileInputStream file = new FileInputStream(
					"C:\\Users\\AL04040\\OneDrive - Elevance Health\\Documents\\VA_SBE\\temp\\Daily Sheet_TEST.xlsx");
			XSSFWorkbook wb = new XSSFWorkbook(file);
			XSSFSheet sh1 = wb.getSheetAt(0);
			DataFormatter dtfm = new DataFormatter();
			DateFormat df = new SimpleDateFormat("dd/MM/yy");
			String orderDate = dtfm.formatCellValue(sh1.getRow(1).getCell(20));
			Date dt = df.parse(orderDate);
			Calendar cal = Calendar.getInstance();
			cal.setTime(dt);
			String day = cal.getDisplayName(cal.DAY_OF_WEEK, cal.LONG, Locale.US);
			System.out.println(df.format(dt));
			System.out.println(day);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
