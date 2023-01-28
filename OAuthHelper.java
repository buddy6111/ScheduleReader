package com.ScheduleUploader;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.List;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.calendar.Calendar;
import com.google.api.services.calendar.CalendarScopes;
import com.google.api.services.calendar.model.CalendarList;
import com.google.api.services.calendar.model.CalendarListEntry;
import com.google.api.services.calendar.model.Event;


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
 * This file handles any and all transactions with the Google Calendar API
 * 
 * @param Users login information as well as permissions to allow edits to Google Calendar
 * 
 * @author bkerk
 */

public class OAuthHelper
{

	/**
	 * Application name.
	 * Directory to store authorization tokens for this application.
	 */
	private static final String APPLICATION_NAME = "ScheduleReader",
				TOKENS_DIRECTORY_PATH = "tokens",
				CREDENTIALS_FILE_PATH = "credentials.json";
	/**
	 * Global instance of the JSON factory.
	 */
	private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
	/**
	 * Global instance of the scopes required by this quickstart.
	 * If modifying these scopes, delete your previously saved tokens/ folder.
	 */
	private static final List<String> SCOPES =
			Collections.singletonList(CalendarScopes.CALENDAR);

	private static Credential credentials;
	static NetHttpTransport HTTP_TRANSPORT;
	private static Calendar service;
	private static List<CalendarListEntry> calendars;

	public static Credential returnCredential()
	{
		return credentials;
	}

	public static JsonFactory getJsonFactory()
	{
		return JSON_FACTORY;
	}

	/**
	 * Creates an authorized Credential object.
	 *
	 * @param HTTP_TRANSPORT The network HTTP Transport.
	 * @return An authorized Credential object.
	 * @throws IOException If the credentials.json file cannot be found.
	 */
	private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT1) throws IOException
	{
		HTTP_TRANSPORT = HTTP_TRANSPORT1;
		// Load client secrets.
		InputStream in = ScheduleUploader.class.getResourceAsStream(CREDENTIALS_FILE_PATH);

		if (in == null) 
		{
			throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
		}

		GoogleClientSecrets clientSecrets =
				GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
				HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
				.setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
				.setAccessType("offline")
				.setApprovalPrompt("force")
				.build();

		LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8880).build();
		Credential credential = new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");

		//returns an authorized Credential object.
		return credential;
	}

	/**
	 * This is the initial access of the Google Calendar. It is mainly used to retrieve the Calendar Names
	 * so the user can select which Calendar they want to upload the schedule to
	 */
	public static List<CalendarListEntry> AccessGoogleCalendar() throws GeneralSecurityException, IOException
	{
		final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
		credentials = getCredentials(HTTP_TRANSPORT);

		// Build a new authorized API client service.
		Calendar service = getCalendars();
		String pageToken = null;
		do {

			CalendarList calendarList = service.calendarList().list().setPageToken(pageToken).execute();
			calendars = calendarList.getItems();
			pageToken = calendarList.getNextPageToken();

		} while (pageToken != null);
		
		return calendars;
	}
	
	
	  /**
	   * First accesses the Google API, then if loops through the shifts until the correct employee is found.
	   * When the correct shifts are found, it creates Events for the shift and sends it to Google to be added
	   * 
	   */
	public static List<String> addScheduleToCalendar(String selectedName, String selectedCalendar, List<List<String>> allShifts, List<Date> dates)
	{	
		try {
			List<String> alerts = new ArrayList<String>();
			EventCreator e = new EventCreator();
			
			CalendarListEntry currentCalendar = new CalendarListEntry();
			for(CalendarListEntry cal : calendars)
				if(cal.getSummary() == selectedCalendar)
					currentCalendar = cal;
			
			for(List<String> shifts : allShifts)
			{
				if(shifts.get(0) == selectedName)
				{
					//For now, I just want to do the first 3 events for testing purposes so its easier to delete later
					for(int i = 1; i < 4; i++)//dates.size(); i++)
					{
						System.out.println(shifts.get(i));
						if(!shifts.get(i).equals("DO"))
						{
							Event event = e.createEvent(shifts.get(i), dates.get(i));

							if(event != null) {
								service.events().insert(currentCalendar.getId(), event).execute();
							}
						}
					}
				}
			}
			alerts = e.getAlerts();
			return alerts;

		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	
	public static Calendar getCalendars()
	{
		Credential credentials = returnCredential();

		service = new Calendar.Builder(HTTP_TRANSPORT, OAuthHelper.getJsonFactory(), credentials)
				.setApplicationName("Schedule Reader").build();

		return service;
	}

}
