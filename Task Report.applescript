(*
OmniFocus Task and Project Completion Report
Version 1.0
Original by By: Ben Waldie <https://about.me/benwaldie>
Updated By: Ben Mason <https://ben.the-collective.net/>

This script generates a list of completed Projects and Completed tasks for a given time frame in Omnifocus. 

*)

-- Prepare a name for the new note
set theNoteName to "OmniFocus Completed Task Report"

-- Prompt the user to choose a scope for the report
activate
set theReportScope to choose from list {"Today", "Yesterday", "This Week", "Last Week", "Last Month", "This Month", "This Quarter", "Last Quarter"} default items {"Today"} with prompt "Generate a report for:" with title "OmniFocus Completed Task Report"

if theReportScope = false then return
set theReportScope to item 1 of theReportScope

-- Calculate the task start and end dates, based on the specified scope
set theStartDate to current date
set hours of theStartDate to 0
set minutes of theStartDate to 0
set seconds of theStartDate to 0
set theEndDate to theStartDate + (23 * hours) + (59 * minutes) + 59
set timetrackrange to false
set icalrange to false
set startWeek to Monday
set endWeek to Sunday
set theProgressDetail to ""

if theReportScope = "Today" then
	set theDateRange to date string of theStartDate
	set icalrange to "eventsToday"
	set timetrackrange to ""
	
else if theReportScope = "Yesterday" then
	set theStartDate to theStartDate - 1 * days
	set theEndDate to theEndDate - 1 * days
	set theDateRange to date string of theStartDate
	set icalrange to "eventsFrom:yesterday to:yesterday"
	set timetrackrange to ":yesterday"
	
else if theReportScope contains "Week" then
	
	if theReportScope = "This Week" then
		set timetrackrange to ":week"
		set timetrackrange to "eventsFrom:monday to:friday"
	else if theReportScope = "Last Week" then
		set theStartDate to theStartDate - 7 * days
		set theEndDate to theEndDate - 7 * days
		set timetrackrange to ":lastweek"
		
	end if
	
	repeat until (weekday of theStartDate) = startWeek
		set theStartDate to theStartDate - 1 * days
	end repeat
	
	repeat until (weekday of theEndDate) = endWeek
		set theEndDate to theEndDate + 1 * days
	end repeat
	
else if theReportScope contains "Month" then
	set day of theStartDate to 1
	
	if theReportScope = "This Month" then
		set timetrackrange to ":month"
		
	else if theReportScope = "Last Month" then
		set month of theStartDate to (month of theStartDate) - 1
		set month of theEndDate to (month of theEndDate) - 1
		set timetrackrange to ":lastmonth"
		
	end if
	
	repeat until (month of theEndDate) is not equal to (month of theStartDate)
		set theEndDate to theEndDate + 1 * days
	end repeat
	set theEndDate to theEndDate - 1 * days
	
	
	
else if theReportScope contains "Quarter" then
	
	set currentMonth to month of (current date) as integer
	set currentQuarter to (((currentMonth - 1) div 3 + 1) as text)
	
	if theReportScope = "This Quarter" then
		set timetrackrange to ":quarter"
		
	else if theReportScope = "Last Quarter" then
		set currentQuarter to currentQuarter - 1
		-- Set date to previous year if in Q1
		if currentQuarter = 0 then
			set currentQuarter to "4"
			set the year of theStartDate to (year of theStartDate) - 1
			set the year of theEndDate to (year of theEndDate) - 1
		end if
		set timetrackrange to ":lastquarter"
		
	end if
	
	if currentQuarter = "1" then
		set the month of theStartDate to 1
		set the day of theStartDate to 1
		set the month of theEndDate to 3
		set the day of theEndDate to 31
	else if currentQuarter = "2" then
		set the month of theStartDate to 4
		set the day of theStartDate to 1
		set the month of theEndDate to 6
		set the day of theEndDate to 30
	else if currentQuarter = "3" then
		set the month of theStartDate to 7
		set the day of theStartDate to 1
		set the month of theEndDate to 9
		set the day of theEndDate to 31
	else if currentQuarter = "4" then
		set the month of theStartDate to 10
		set the day of theStartDate to 1
		set the month of theEndDate to 12
		set the day of theEndDate to 31
	end if
	
	
end if

set theDateRange to (date string of theStartDate) & " through " & (date string of theEndDate)

set {year:y, month:m, day:d} to (current date)
set theDate to y & "/" & (m * 1) & "/" & (d) as string

-- Begin preparing the task list as HTML.
set theProgressDetail to theProgressDetail & "Journal - " & theDate & return & return


-- Completed Projects
set theProjects to ""
set theProjects to "Completed Projects - " & theDateRange & return
set theProjects to theProjects & "-------------------------------" & return

tell application "OmniFocus"
	tell front document
		set theCompletedProjects to every flattened project where its completed = true and modification date is greater than theStartDate and modification date is less than theEndDate
		
		if theCompletedProjects is not equal to {} then
			repeat with a from 1 to length of theCompletedProjects
				set theCurrentProject to item a of theCompletedProjects
				-- Append the project name to the task list
				set theProjects to theProjects & "* " & name of theCurrentProject & return
			end repeat
		else
			set theProjects to "No Completed Projects" & return
			
		end if
	end tell
end tell
set theProgressDetail to theProgressDetail & theProjects & return & return

-- Task Collection
set theTasks to ""
set theTasks to theTasks & "Completed Tasks - " & theDateRange & return
set theTasks to theTasks & "-------------------------------" & return

-- Retrieve a list of projects modified within the specified scope
set modifiedTasksDetected to false
tell application "OmniFocus"
	tell front document
		set theModifiedProjects to every flattened project where its modification date is greater than theStartDate and modification date is less than theEndDate
		
		-- Loop through any detected projects
		repeat with a from 1 to length of theModifiedProjects
			set theCurrentProject to item a of theModifiedProjects
			
			-- Retrieve any project tasks modified within the specified scope
			set theCompletedTasks to (every flattened task of theCurrentProject where its completed = true and modification date is greater than theStartDate and modification date is less than theEndDate and number of tasks = 0)
			
			-- Loop through any detected tasks
			if theCompletedTasks is not equal to {} then
				set modifiedTasksDetected to true
				
				-- Append the project name to the task list
				set theTasks to theTasks & name of theCurrentProject & return
				
				repeat with b from 1 to length of theCompletedTasks
					set theCurrentTask to item b of theCompletedTasks
					
					-- Append the tasks's name to the task list
					set theTasks to theTasks & "* " & name of theCurrentTask & " (" & completion date of theCurrentTask & ")" & return
					
				end repeat
				set theTasks to theTasks & return
			end if
		end repeat
	end tell
end tell

set theProgressDetail to theProgressDetail & theTasks

if timetrackrange is not equal to false then
	-- doing tracked time
	set theProgressDetail to theProgressDetail & return & "Time Tracking today data" & return
	set timetrackdata to (do shell script "/usr/local/bin/timew summary " & timetrackrange)
	set theProgressDetail to theProgressDetail & timetrackdata & return
end if

if icalrange is not equal to false then
	-- Add ICal (Calendar) Data
	set theProgressDetail to theProgressDetail & return & "Calendar Data" & return
	set icaldata to (do shell script "/usr/local/bin/icalbuddy -npn -nc -ps \"/ - /\" -eep \"url\",notes,attendees -ec \"Daily Routine\",\"Birthdays\",\"Facebook Events\",\"IFTTT\" " & icalrange & " | uniq")
	set theProgressDetail to theProgressDetail & icaldata & return
end if

-- Create the note in TextEdit.
tell application "TextEdit"
	activate
	make new document at the front
	set the text of the front document to theProgressDetail
end tell
