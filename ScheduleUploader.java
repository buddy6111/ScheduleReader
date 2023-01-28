package com.ScheduleUploader;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;
import java.awt.BorderLayout;
import java.awt.GridLayout;
import java.awt.event.*;    
import java.io.*;   
import java.util.*;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.services.calendar.Calendar;
import com.google.api.services.calendar.model.CalendarListEntry;
import java.security.GeneralSecurityException;


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
 * This particular file is the starting point. It constructs the GUI
 * and adds functionality to all the buttons that allow the applet to work.
 * 
 * References 3 different files:
 * 	1. OAuthHelper      : handles all transactions with the Google Calendar API
 *  2. ExcelFileReader  : handles all data retrieval from the schedule
 *  3. EventCreator     : handles the creation of events to be uploaded using the OAuthHelper
 * 
 * @author bkerk
 */

public class ScheduleUploader extends JFrame implements ActionListener 
{
	//Defining required variables
	
	//GUI Elements
	JTextArea ta;
	static boolean fileValid;
	static JLabel l, firstName, calendar, newOrUpdate;
	static JButton submit, cancel, sync, currentButton;
	String[] array = {"<Select>"};
	static JComboBox<String> cbNames, cbCalendars;
	private DefaultListCellRenderer listRenderer;
	private static JPanel information;
	
	//OAuth Information
	Credential credentials;
	static Calendar service;
	
	//Data holding variables
	private static String selectedName, selectedCalendar, fileLocation;
	private static List<CalendarListEntry> calendars;
	private static List<String> names, alerts;
	private static List<Date> dates;
	private static ArrayList<List<String>> allShifts;
	
	
	public static void main(String args[])
	{
		ScheduleUploader reader = new ScheduleUploader();
	}
	
	
	/**
	 * This is the Constructor method for the class. It mainly only handles the construction of the GUI
	 * and the button actions with the appropriate listeners.
	 * 
	 * The GUI functions as the main interaction with the user. It gets all the necessary information,
	 * so that's why this acts as the main class as well as the GUI construction.
	 * 
	 *  @param args Unused
	 *  @return Nothing
	 *  
	 */

	ScheduleUploader()
	{    
		//Creating the Frame
        final JFrame frame = new JFrame("Schedule Reader");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(380, 350);
        frame.setLocation(300, 300);

        //Sets the upload Button
        JPanel panel = new JPanel(); // the panel is not visible in output
        JLabel label = new JLabel("Select File: ");
        JButton open = new JButton("Select");
        open.addActionListener(this);
        
        l = new JLabel("");
        
        panel.add(label);
        panel.add(open);
        panel.add(l);

        // Text Area at the Bottom for information about whats happening
        ta = new JTextArea(8,1);
        
        ta.append("To Begin:\nUpload the Schedule\n\n"
        		+ "Fill in necessary information\n"
        		+ "Choose \"New\" to sync for the first time\n"
        		+ "Choose \"Update\" to resync the schedule with any updates\n\n"
        		+ "Click Submit and this text field will update with any messages");
        
        //-------------------------------------------------------------------------------------------------------
        
        information = new JPanel();
        
        GridLayout grid = new GridLayout(0,2);
        
        grid.setVgap(7);
        grid.setHgap(20);
        information.setLayout(grid);
        
        firstName = new JLabel("First Name: ");
        firstName.setHorizontalAlignment(JLabel.CENTER);  
        
        information.add(firstName);
        
        cbNames = new JComboBox<String>(array);
        listRenderer = new DefaultListCellRenderer();
        listRenderer.setHorizontalAlignment(DefaultListCellRenderer.CENTER); // center-aligned items
        cbNames.setRenderer(listRenderer);
        information.add(cbNames);
        
        calendar = new JLabel("Calendar Name: ");
        calendar.setHorizontalAlignment(JLabel.CENTER);
        
        information.add(calendar);
        
        cbCalendars = new JComboBox<String>(array);
        listRenderer = new DefaultListCellRenderer();
        listRenderer.setHorizontalAlignment(DefaultListCellRenderer.CENTER); // center-aligned items
        cbCalendars.setRenderer(listRenderer);
        information.add(cbCalendars);
        
        submit = new JButton("Read File");
        submit.addActionListener( new ActionListener() {
        	public void actionPerformed(ActionEvent e)
            {
            	ta.selectAll();
            	ta.replaceSelection("");
            	ta.append("Select your name from the list, and select which calendar\n"
            			+ "you'd like to sync to.\n\nPress \'Sync to Calendar\' to begin upload.");
        		
        		try {
					calendars = OAuthHelper.AccessGoogleCalendar();
					service = OAuthHelper.getCalendars();
					switchButtons();
					update();
				} catch (GeneralSecurityException e1) {
					e1.printStackTrace();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
        		setComponentsActive();
            }
        	});
            
        sync = new JButton("Sync to Calendar");
        sync.addActionListener( new ActionListener() {
        	public void actionPerformed(ActionEvent e)
            {
        		selectedName = cbNames.getSelectedItem().toString();
        		selectedCalendar = cbCalendars.getSelectedItem().toString();
        		
        		String[] name = selectedName.split(" ");
            	alerts = OAuthHelper.addScheduleToCalendar(selectedName, selectedCalendar, allShifts, dates);
            	
            	ta.selectAll();
            	ta.replaceSelection("");
        		ta.append(name[0] + "'s Schedule Upload Completed.\nGoogle Calendar may take a few moments to update.\nCheck below for any alerts.\n\n");
        		
        		for(String s : alerts)
        		{
        			ta.append(s + "\n");
        		}
        		
            }
        	});
        currentButton = submit;
        information.add(currentButton);
        
        cancel = new JButton("Cancel"); 
        cancel.addActionListener( new ActionListener() {
        	public void actionPerformed(ActionEvent e)
            {
        		frame.dispose();
            }
        	});   
        information.add(cancel);
        
        //Adding Components to the frame.
        frame.getContentPane().add(BorderLayout.NORTH, panel);
        frame.getContentPane().add(BorderLayout.CENTER, information);
        frame.getContentPane().add(BorderLayout.SOUTH, ta);
        frame.pack();
        setComponentsInactive();
        frame.setVisible(true);
	}

	  /**
	   * Updates the GUI with the new variables to be selected
	   * @param args Unused.
	   * @return Nothing.
	   */

	private void update()
	{
		cbNames.removeAllItems();
		for(String s : names)
			cbNames.addItem(s);

		cbCalendars.removeAllItems();
		for(CalendarListEntry cal : calendars)
		{
			//Removes known calendars that are uneditable, so there won't be an error if the user selects them
			if(!cal.getSummary().equals("Birthdays") && !cal.getSummary().equals("Holidays in United States"))
			{
				cbCalendars.addItem(cal.getSummary());
			}
		}
		
		this.invalidate();
		this.validate();
		this.repaint();
	}
	  /**
	   * Sets the components of the GUI active once the required files are attached and variables can be retrieved
	   * @param args Unused.
	   * @return Nothing.
	   */
	static void setComponentsActive()
	{
		cbNames.setEnabled(true);
		cbCalendars.setEnabled(true);
        calendar.setEnabled(true);
        firstName.setEnabled(true);
	}
	
	  /**
	   * Sets the components of the GUI inactive to prevent selection before variables are assigned
	   * @param args Unused.
	   * @return Nothing.
	   */
	static void setComponentsInactive()
	{
		submit.setEnabled(false);
		cbNames.setEnabled(false);
		cbCalendars.setEnabled(false);
        calendar.setEnabled(false);
        firstName.setEnabled(false);
	}

	  /**
	   * Switches the Submit button with the Sync button
	   * @param args Unused.
	   * @return Nothing.
	   */
	static void switchButtons()
	{
		if(currentButton == sync)
		{
			currentButton = submit;
			information.remove(sync);
			information.remove(cancel);
			information.add(submit);
			information.add(cancel);
		}
		else if(currentButton == submit)
		{
			currentButton = sync;
			information.remove(submit);
			information.remove(cancel);
			information.add(sync);
			information.add(cancel);
		}
	}
	
	  /**
	   * Checks for actions performed for the File Chooser function
	   * @param ActionEvent e : When the Select File button is clicked, it triggers this action event
	   * @return Nothing.
	   */
	@Override
	public void actionPerformed(ActionEvent e)
	{
		  
		JFileChooser j = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
		 
		//Restricts the file type to Excel files only
        j.setAcceptAllFileFilterUsed(false);
        j.setDialogTitle("Select a .xlsx (Microsoft Excel) file");
        FileNameExtensionFilter restrict = new FileNameExtensionFilter("Only .xlsx files", "xlsx");
        j.addChoosableFileFilter(restrict);

        // invoke the showsOpenDialog function to show the save dialog
        int r = j.showOpenDialog(null);

        // if the user selects a file
        if (r == JFileChooser.APPROVE_OPTION)
        {
            // set the label to the path of the selected file
            fileLocation = j.getSelectedFile().getAbsolutePath();
            String[] temp = fileLocation.split("\\\\");
            l.setText(temp[temp.length - 1]);
            
            //This references the ExcelFileReader class to read the inputed file and save the data in the 
            //necessary data structures for later use
            ExcelFileReader.ReadSchedule(fileLocation);
            names = ExcelFileReader.getNames();
    		allShifts = ExcelFileReader.getAllShifts();
    		dates = ExcelFileReader.getDates();
            
    		//Determines if the file is valid or not
            if(ExcelFileReader.isFileValid())
            {
            	ta.selectAll();
            	ta.replaceSelection("");
            	ta.append(temp[temp.length - 1] + " - File is valid. Press 'Read File' to begin."
            			+ "\n\nFirst time users, Google should ask for permission to \naccess your calendar."
            			+ "\n\nPlease note that it is only asking for access to the calendar.");
            	submit.setEnabled(true);
            	//update();
            }
            else
            {
            	ta.selectAll();
            	ta.replaceSelection("");
            	ta.append(temp[temp.length - 1] + " - File is not valid, select another file.");
            	if(currentButton == sync)
            	{
            		switchButtons();
            		submit.setEnabled(false);
            	}
            	setComponentsInactive();
            }
        }
        // if the user cancelled the operation
        else
            l.setText("the user cancelled the operation");
	}
}