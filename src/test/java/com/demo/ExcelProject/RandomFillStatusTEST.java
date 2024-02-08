package com.demo.ExcelProject;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class RandomFillStatusTEST {

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
			String schedule = "";
			final String[] allSchedules = new String[40];
			int j = 0;
			for (int i = 0; i < sh2Len; i++) {
				scheduleCell = sh2.getRow(i).getCell(6);
				schedule = scheduleCell.getStringCellValue();
				if (!Arrays.asList(allSchedules).contains(schedule)){
					allSchedules[j++] = schedule;System.out.printf("%02d) %s\n",(j-1),schedule);}
				
			}/*
		    Map<String, List<String>> exactSchedule = new HashMap<String, List<String>>(){{
		    	put(allSchedules[1], new ArrayList<String>(){{add("On Demand");}});
		    	put(allSchedules[2], new ArrayList<String>(){{add("MON");add("TUE");add("WED");add("THU");add("FRI");add("SAT");add("SUN");}});
		    	put(allSchedules[3], new ArrayList<String>(){{add("MON");add("TUE");add("WED");add("THU");add("FRI");}});
		    	put(allSchedules[4], new ArrayList<String>(){{add("MON");add("TUE");add("WED");add("THU");add("FRI");add("SAT");}});
		    	put(allSchedules[5], new ArrayList<String>(){{add("SUN");add("MON");add("TUE");add("WED");add("THU");}});
		    	put(allSchedules[6], new ArrayList<String>(){{add("");}});
		    	put(allSchedules[7], new ArrayList<String>(){{add("On Demand");}});
		    	put(allSchedules[8], new ArrayList<String>(){{add("MON");}});
		    	put(allSchedules[9], new ArrayList<String>(){{add("");}});
		    }};
		    System.out.println(exactSchedule.get(allSchedules[5]).get(0));
			// // }
			/*
			 * FileOutputStream fileOut = new FileOutputStream(
			 * "C:\\Users\\AL04040\\OneDrive - Elevance Health\\Documents\\VA_SBE\\temp\\Daily Sheet_TEST.xlsx"
			 * ); wb.write(fileOut); fileOut.close(); System.out.
			 * println("File successfully written and timings marked for completed jobs."
			 * );
			 */
//			String result = allSchedules[17].matches(".*[0-9].*");
			System.out.println(allSchedules[17].matches(".*[0-9].*"));
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
