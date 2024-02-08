package com.demo.ExcelProject;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.TimeZone;

public class DateChecker {

	public static void main(String[] args) {
		DateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
		String t = "01-06-2005";
		Date date = null;
		try {
			date = dateFormat.parse(t);
			System.out.println("Date format is correct - "+dateFormat.format(date));
		} catch (ParseException e) {
			System.out.println("Error occured.");
			e.printStackTrace();
		}
	}

}
