package com.demo.ExcelProject;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteXLSXFileTEST {

	public static void main(String[] args) {
		try {
			FileInputStream file = new FileInputStream(
					"C:\\Users\\AL04040\\OneDrive - Elevance Health\\Documents\\VA_SBE\\temp\\Daily Sheet1 - Copy.xlsx");
			XSSFWorkbook wb = new XSSFWorkbook(file);
			XSSFSheet sh1 = wb.getSheet("CTM Details");
			XSSFSheet sh2 = wb.getSheet("Checklist");
			DataFormatter fm = new DataFormatter();
			DateFormat df = new SimpleDateFormat("hh:mm:ss a");
			Scanner scan = new Scanner(System.in);
			System.out.print("Enter number of jobs: ");
			int sh1Len = scan.nextInt();
			scan.close();
			int sh2Len = 490;
			Cell scheduleCell = null;
			String schedule = "";
			for (int i = 0; i < sh2Len; i++) {
				scheduleCell = sh2.getRow(i).getCell(6);
				schedule = scheduleCell.getStringCellValue();
				if (schedule.contains("On Demand") || schedule.contains("OnDemand")) {
					Cell naCell = sh2.getRow(i).createCell(11);
					naCell.setCellValue("NA");
				}else if(schedule.contains("&"))
					continue;
				else if(schedule.matches(".*[0-9].*") && schedule.length()<=12){
					
				}
			}
			for (int i = 0; i < sh1Len; i++) {
				String jobStatus = fm.formatCellValue(sh1.getRow(i).getCell(2));
				// if (jobStatus.equals("Ended OK")) {
				for (int j = 0; j < sh2Len; j++) {
					String sh1JobName = fm.formatCellValue(sh1.getRow(i).getCell(1));
					String sh1FldrName = fm.formatCellValue(sh1.getRow(i).getCell(28));
					String sh2JobName = fm.formatCellValue(sh2.getRow(j).getCell(3));
					String sh2FldrName = fm.formatCellValue(sh2.getRow(j).getCell(4));
					if (sh1JobName.equals(sh2JobName) && sh1FldrName.equals(sh2FldrName)) {
						System.out.printf("%d)%s - %s\n", (i + 1), sh2JobName, sh2FldrName);
						XSSFCell cellStartDateINP = sh1.getRow(i).getCell(14);
						XSSFCell cellEndDateINP = sh1.getRow(i).getCell(15);
						String startDate = "";
						String endDate = "";
						Cell cellJobStatus = sh2.getRow(j).getCell(11);
						if (cellJobStatus == null || cellJobStatus.getCellType() == 3) {
							Cell cellStDtOUT = sh2.getRow(j).createCell(9);
							Cell cellEnDtOUT = sh2.getRow(j).createCell(10);
							cellJobStatus = sh2.getRow(j).createCell(11);
							// cellStartDateINP = sh1.getRow(i).getCell(14);
							// cellEndDateINP = sh1.getRow(i).getCell(15);
							// System.out.println(sh1.getRow(i).getCell(14).getCellType());
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
									if (cellStartDateINP.getCellType() != 1) {
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
							if (startDate.equals("") && endDate.equals(""))
								cellJobStatus.setCellValue("Yet to start");
							break;
						}
					}
				}
				// }
			}
			FileOutputStream fileOut = new FileOutputStream(
					"C:\\Users\\AL04040\\OneDrive - Elevance Health\\Documents\\VA_SBE\\temp\\Daily Sheet_TEST.xlsx");
			wb.write(fileOut);
			fileOut.close();
			System.out.println("File successfully written and timings marked for completed jobs.");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
