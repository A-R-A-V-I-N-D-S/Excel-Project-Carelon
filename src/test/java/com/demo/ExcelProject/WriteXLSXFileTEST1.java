package com.demo.ExcelProject;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Scanner;
import java.util.TimeZone;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteXLSXFileTEST1 {

	public static void main(String[] args) {
		try {
			FileInputStream file = new FileInputStream(
					"C:\\Users\\AL04040\\OneDrive - Elevance Health\\Documents\\VA_SBE\\temp\\Daily Sheet1 - Copy.xlsx");
			XSSFWorkbook wb = new XSSFWorkbook(file);
			XSSFSheet sh1 = wb.getSheet("CTM Details");
			XSSFSheet sh2 = wb.getSheet("Checklist");
			DataFormatter fm = new DataFormatter();
			int sh2Len = 490;
			Cell scheduleCell = null;
			Cell serverCell = null;
			String schedule = "";
			String server = "";
			for (int i = 0; i < sh2Len; i++) {
				scheduleCell = sh2.getRow(i).getCell(6);
				serverCell = sh2.getRow(i).getCell(5);
				schedule = scheduleCell.getStringCellValue();
				server = serverCell.getStringCellValue();
				System.out.printf("%d) %s => %s\n", (i+1), server, isJobNewScan(server)==1?"New scan":"No new scan");
			}			
			/*FileOutputStream fileOut = new FileOutputStream(
					"C:\\Users\\AL04040\\OneDrive - Elevance Health\\Documents\\VA_SBE\\temp\\Daily Sheet_TEST.xlsx");
			wb.write(fileOut);
			fileOut.close();
			System.out.println("File successfully written and timings marked for completed jobs.");*/
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public static int isMonthlyJobTodayPresent(String date){
		DateFormat dateNTimeFormat = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
		DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
		dateFormat.setTimeZone(TimeZone.getTimeZone("EST"));
		dateNTimeFormat.setTimeZone(TimeZone.getTimeZone("EST"));
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 1);
		Date curDatenTime = new Date();
		Date curDate = new Date();
		String tdyDate = dateFormat.format(curDate);
		String tmrDate = dateNTimeFormat.format(cal.getTime());
		Date tdy3pmEST = null, tmr3pmEST = null;
		Date tdy6amEST = null, tmr6amEST = null;
		return 0;
	}
	public static int isJobNewScan(String srvr){
		DateFormat dateNTimeFormat = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
		DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
		dateFormat.setTimeZone(TimeZone.getTimeZone("EST"));
		dateNTimeFormat.setTimeZone(TimeZone.getTimeZone("EST"));
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 1);
		Date curDatenTime = new Date();
		Date curDate = new Date();
		String tdyDate = dateFormat.format(curDate);
		String tmrDate = dateNTimeFormat.format(cal.getTime());
		Date tdy3pmEST = null, tmr3pmEST = null;
		Date tdy6amEST = null, tmr6amEST = null;
		try {
			tdy3pmEST = dateNTimeFormat.parse(tdyDate+" 15:00:00");
			tdy6amEST = dateNTimeFormat.parse(tdyDate+" 06:00:00");
			tmr3pmEST = dateNTimeFormat.parse(tmrDate+" 15:00:00");
			tmr6amEST = dateNTimeFormat.parse(tmrDate+" 06:00:00");
//			System.out.println(dateNTimeFormat.format(curDatenTime));
//			System.out.println(tmrDate);
		} catch (ParseException e) {
			e.printStackTrace();
		}
		if(srvr.equals("CTM200")){
			if(curDatenTime.after(tdy3pmEST) && curDatenTime.before(tmr3pmEST))
				return 1;
			else
				return 0;
		}else{
			if(curDatenTime.after(tdy6amEST) && curDatenTime.before(tmr6amEST))
				return 1;
			else
				return 0;
		}
	}

}
