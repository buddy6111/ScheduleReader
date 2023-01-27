package com.demo.ExcelProject;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
* This applet requests a file to be uploaded by the user.
* It can detect if this is a proper Excel file containing a schedule.
* 
* It then accesses Google Calendar and retrieves the user's Calendars
* while also retrieving the employees found on the schedule. It then
* allows the user to select an employee and a Calendar, and will then
* upload the selected schedule to the selected Calendar.
* 
* Once completed, it will display a list of errors that the applet
* encountered while uploading the schedule.
*
*
* @author  Brennan Kerkstra
* @version 1.0
* @since   2023-01-27 
*/

/**
 * ExcelFileReader.java
 * 
 * This file handles the data extraction from the given Excel document,
 * given that the file is valid.
 * 
 * @param fileLocation : String containing the directory for the file selected by user
 * @return fileValid   : Boolean value that checks if the file is a valid schedule
 * @return allShifts   : ArrayList that contains a List for each employee that contains each shift
 * 		in order with their corresponding dates
 * @return dates       : List of Dates that go with each shift
 * @return names	   : List of Strings that hold the names of all the employees
 * 
 * @author bkerk
 */

public class ExcelFileReader
{
	static boolean fileValid;
	private static List<Date> dates;
	private static ArrayList<List<String>> allShifts;
	static List<String> names;
	
	/**
	 * Returns if the file is valid
	 * 
	 */
	
	public static boolean isFileValid() {
		return fileValid;
	}
	
	/**
	 * Returns all the shifts
	 * 
	 */
	
	public static ArrayList<List<String>> getAllShifts() {
		return allShifts;
	}
	/**
	 * Returns the Array of Dates
	 * 
	 */
	
	public static List<Date> getDates() {
		return dates;
	}
	/**
	 * Returns all of the employees listed on the schedule
	 * 
	 */
	
	public static List<String> getNames() {
		return names;
	}
	
	
	/**
	 * This function takes the Microsoft Excel file and filters through looking for only the employees schedules
	 * as well as the dates for each scheduled shift. Filters out the unneccessary information
	 * 
	 * @param fileLocation
	 * @return Nothing
	 * 
	 */
	public static void ReadSchedule(String fileLocation)
	{
		fileValid = false;
		
		List<String> shift = new ArrayList<String>();
		//ROW BY COLUMN : ROW_LENGTH is the number of columns
		allShifts = new ArrayList<List<String>>();
		
		dates = new ArrayList<Date>();
		names = new ArrayList<String>();
	
		try 
		{
			FileInputStream file = new FileInputStream(new File(fileLocation));
			Workbook workbook = new XSSFWorkbook(file);
			DataFormatter dataFormatter = new DataFormatter();
			Iterator<Sheet> sheets = workbook.sheetIterator();
	
			while(sheets.hasNext())
			{
				Sheet sh = sheets.next();
				Iterator<Row> iterator = sh.rowIterator();
	
				while(iterator.hasNext())
				{
					Row row = iterator.next();
					boolean rowValid = false;
					Iterator<Cell> cellIterator = row.iterator();
					Cell firstCell = cellIterator.next();
					String name = dataFormatter.formatCellValue(firstCell);
					
					if(name.contains("DAY OF WEEK"))
						fileValid = true;
					
					if(name.contains("DAY OF MONTH"))
					{
						dates.add(new Date());
						while(cellIterator.hasNext())
						{
							
							Cell cell = cellIterator.next();
							Date cellValue = cell.getDateCellValue();
							dates.add(cellValue);
						}
						
					}
					else if(!name.equals("") && name.contains("(") && name.contains(")") 
							&& !name.contains("OPEN"))
					{
						names.add(name);
						shift.add(name);
						rowValid = true;

						while(cellIterator.hasNext())
						{
							Cell cell = cellIterator.next();
							String cellValue = dataFormatter.formatCellValue(cell);

							if((cell.getCellType() == CellType.STRING || cellValue.equals("") || cellValue.indexOf("*") != -1 && rowValid))
							{
								if(cellValue == "" || cellValue.indexOf("*") != -1)
									shift.add("DO");
								else
									shift.add(cellValue);
							}
						}
						allShifts.add(shift);
						shift = new ArrayList<String>();
					}
				}
			}
			workbook.close();
		} 
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}
}
