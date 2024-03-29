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
					"C:\\Users\\AL04040\\OneDrive - Elevance Health\\Documents\\VA_SBE\\temp\\Daily Sheet_TEST.xlsx");
			XSSFWorkbook wb = new XSSFWorkbook(file);
			XSSFSheet sh1 = wb.getSheet("CTM Details");
			XSSFSheet sh2 = wb.getSheet("Checklist");
			DataFormatter fm = new DataFormatter();
			DateFormat df = new SimpleDateFormat("hh:mm:ss a");
			int sh1Len = sh1.getLastRowNum();
			int sh2Len = sh2.getLastRowNum();
			System.out.printf("%d , %d\n", sh1Len, sh2Len);
			Cell scheduleCell = null;
			String schedule = "";
			int[] flag = new int[sh2Len + 1];
			for (int i = 0; i <= sh2Len; i++) {
				flag[i] = 0;
				scheduleCell = sh2.getRow(i).getCell(6);
				schedule = scheduleCell.getStringCellValue();
				if (schedule.contains("On Demand") || schedule.contains("OnDemand")) {
					Cell naCell = sh2.getRow(i).createCell(11);
					naCell.setCellValue("NA");
				} else if (schedule.contains("&"))
					continue;
				else if (schedule.matches(".*[0-9].*") && schedule.length() <= 12) {

				}
			}
			for (int i = 0; i <= sh1Len; i++) {
				String jobStatus = fm.formatCellValue(sh1.getRow(i).getCell(2));
				for (int j = 0; j <= sh2Len; j++) {
					String sh1JobName = fm.formatCellValue(sh1.getRow(i).getCell(1));
					String sh1FldrName = fm.formatCellValue(sh1.getRow(i).getCell(28));
					String sh2JobName = fm.formatCellValue(sh2.getRow(j).getCell(3));
					String sh2FldrName = fm.formatCellValue(sh2.getRow(j).getCell(4));
					if (sh1JobName.equals(sh2JobName) && sh1FldrName.equals(sh2FldrName)) {
						XSSFCell cellStartDateINP = sh1.getRow(i).getCell(14);
						XSSFCell cellEndDateINP = sh1.getRow(i).getCell(15);
						String startDate = "";
						String endDate = "";
						Cell cellJobStatus = sh2.getRow(j).getCell(11);
						String sh2JobStatus = fm.formatCellValue(sh2.getRow(j).getCell(11));
						if ((!sh2JobStatus.equalsIgnoreCase("ended ok")) && flag[j] == 0) {
							flag[j] = 1;
							Cell cellStDtOUT = sh2.getRow(j).createCell(9);
							Cell cellEnDtOUT = sh2.getRow(j).createCell(10);
							cellJobStatus = sh2.getRow(j).createCell(11);
							if (cellStartDateINP != null) {// ignoring
															// NullPointerException
								if (cellStartDateINP.getCellType() != 3) {
									if (cellStartDateINP.getCellType() != 1) {
										Date date1 = cellStartDateINP.getDateCellValue();
										startDate = df.format(date1);
									} else {
										startDate = cellStartDateINP.getStringCellValue();
									}
								}
							}
							if (cellEndDateINP != null) {// ignoring
															// NullPointerException
								if (cellEndDateINP.getCellType() != 3) {
									if (cellEndDateINP.getCellType() != 1) {
										Date date2 = cellEndDateINP.getDateCellValue();
										endDate = df.format(date2);
									} else {
										endDate = cellStartDateINP.getStringCellValue();
									}
								}
							}
							cellStDtOUT.setCellValue(startDate);
							cellEnDtOUT.setCellValue(endDate);
							switch (jobStatus) {
							case "Ended OK":
								cellJobStatus.setCellValue("Ended OK");
								break;
							case "Wait Condition":
								cellJobStatus.setCellValue("Executing");
								break;
							case "Executing":
								cellJobStatus.setCellValue("Executing");
								break;
							default:
								break;
							}
							String srvr = fm.formatCellValue(sh2.getRow(j).getCell(5));
							if (startDate.equals("") && endDate.equals(""))
								cellJobStatus.setCellValue("Yet to start");
							else if (isJobNewScan(srvr)==1)
								cellJobStatus.setCellValue("Ended OK");
							System.out.printf("%d)%s - %s\n", (i + 1), sh2JobName, sh2FldrName);
							break;
						}
					}
				}
			}
			sh2.autoSizeColumn(9);
			sh2.autoSizeColumn(10);
			sh2.autoSizeColumn(11);
			FileOutputStream fileOut = new FileOutputStream(
					"C:\\Users\\AL04040\\OneDrive - Elevance Health\\Documents\\VA_SBE\\temp\\Daily Sheet_TEST.xlsx");
			wb.write(fileOut);
			fileOut.close();
			System.out.println("File successfully written and timings marked for completed jobs.");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static int isMonthlyJobTodayPresent(String date) {
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

	public static int isJobNewScan(String srvr) {
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
			tdy3pmEST = dateNTimeFormat.parse(tdyDate + " 15:00:00");
			tdy6amEST = dateNTimeFormat.parse(tdyDate + " 06:00:00");
			tmr3pmEST = dateNTimeFormat.parse(tmrDate + " 15:00:00");
			tmr6amEST = dateNTimeFormat.parse(tmrDate + " 06:00:00");
		} catch (ParseException e) {
			e.printStackTrace();
		}
		if (srvr.equals("CTM200")) {
			if (curDatenTime.after(tdy3pmEST) && curDatenTime.before(tmr3pmEST))
				return 1;
			else
				return 0;
		} else {
			if (curDatenTime.after(tdy6amEST) && curDatenTime.before(tmr6amEST))
				return 1;
			else
				return 0;
		}
	}

}
