package com.demo.ExcelProject;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.TimeZone;

import com.google.api.client.util.DateTime;
import com.google.api.services.calendar.model.Event;
import com.google.api.services.calendar.model.EventDateTime;

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
 * EventCreator.java
 * 
 * 
 * This file handles the interpretation of the uploaded Excel file.
 * It then uses the information to create Events with the appropriate start
 * and end times for that shift.
 * 
 * @param shiftType  : String containing the shift title to create an event for
 * @param date       : Date of the corresponding shift
 * @return event     : Event containing the details of the shift, including the shift, 
 * 		start and end times, and location
 * @return alerts    : Returns a list of strings containing any issues that occurred in this
 * 		step, which is mainly if there are missing definitions for shifts
 * 
 * @author bkerk
 */
public class EventCreator
{

	private static Event event = new Event()
			.setLocation("1521 Gull Road, Kalamazoo, MI, 49048");
	private static DateTime start, end;
	private static Date startDate;
	private static List<String> alerts = new ArrayList<String>();
	
	public Event getEvent()
	{
		return event;
	}

	public List<String> getAlerts()
	{
		return alerts;
	}
	
	public static Event createEvent(String shiftType, Date date)
	{
		
		startDate = date;
		System.out.println(startDate);
		event.setSummary("Work - " + shiftType);
		
		boolean sendEvent = true;
		
		if(shiftType.equals("IV1D") || shiftType.equals("IV2D" )
				|| shiftType.equals("U1D") || shiftType.equals("MC")
				|| shiftType.equals("PX1D"))
		{
			shiftTime6();
		}
		else if(shiftType.equals("OR"))
			shiftTime530();
		else if(shiftType.equals("IVC") || shiftType.equals("IVT"))
			shiftTime630();
		else if(shiftType.equals("PACK"))
			shiftTime7();
		else if(shiftType.equals("PX2D") || shiftType.equals("PX3D"))
			shiftTime8();
		else if(shiftType.equals("IVN"))
			shiftTime14();
		else if(shiftType.equals("U2N")|| shiftType.equals("PXN"))
			shiftTime1430();
		else if(shiftType == "M")
			shiftTime1930();
		else
		{
			sendEvent = false;
			alerts.add(date.getMonth()+1 + "/" + date.getDate() + " : No definition for shift - " + shiftType);
		}

		if(sendEvent)
		{
			event.setStart(new EventDateTime().setDateTime(start));
			event.setEnd(new EventDateTime().setDateTime(end));
			return event;
		}

		return null;
	}
	
	
	
	private static void shiftTime1930()
	{
		
	}
	private static void shiftTime1430()
	{
		startDate = new Date(startDate.getTime() + 52200000);
		Date endDate = new Date(startDate.getTime() + 30600000);
		start = new DateTime(startDate, TimeZone.getTimeZone("Detroit"));
		end = new DateTime(endDate, TimeZone.getTimeZone("Detroit"));
	}
	private static void shiftTime14()
	{
		startDate = new Date(startDate.getTime() + 50400000);
		Date endDate = new Date(startDate.getTime() + 30600000);
		start = new DateTime(startDate, TimeZone.getTimeZone("Detroit"));
		end = new DateTime(endDate, TimeZone.getTimeZone("Detroit"));
	}
	private static void shiftTime8()
	{
		startDate = new Date(startDate.getTime() + 28800000);
		Date endDate = new Date(startDate.getTime() + 30600000);
		start = new DateTime(startDate, TimeZone.getTimeZone("Detroit"));
		end = new DateTime(endDate, TimeZone.getTimeZone("Detroit"));
	}
	private static void shiftTime7()
	{
		startDate = new Date(startDate.getTime() + 25200000);
		Date endDate = new Date(startDate.getTime() + 30600000);
		start = new DateTime(startDate, TimeZone.getTimeZone("Detroit"));
		end = new DateTime(endDate, TimeZone.getTimeZone("Detroit"));
	}
	private static void shiftTime630() 
	{
		startDate = new Date(startDate.getTime() + 23400000);
		Date endDate = new Date(startDate.getTime() + 30600000);
		start = new DateTime(startDate, TimeZone.getTimeZone("Detroit"));
		end = new DateTime(endDate, TimeZone.getTimeZone("Detroit"));
	}
	private static void shiftTime530()
	{
		startDate = new Date(startDate.getTime() + 19800000);
		Date endDate = new Date(startDate.getTime() + 30600000);
		start = new DateTime(startDate, TimeZone.getTimeZone("Detroit"));
		end = new DateTime(endDate, TimeZone.getTimeZone("Detroit"));
	}
	private static void shiftTime6() 
	{
		startDate = new Date(startDate.getTime() + 21600000);
		Date endDate = new Date(startDate.getTime() + 30600000);
		start = new DateTime(startDate, TimeZone.getTimeZone("Detroit"));
		end = new DateTime(endDate, TimeZone.getTimeZone("Detroit"));
	}
}
