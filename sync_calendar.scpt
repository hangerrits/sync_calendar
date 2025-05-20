tell application "Calendar"
	-- Define date range
	set currentDate to current date
	set startDateRange to currentDate - (30 * days)
	set endDateRange to currentDate + (180 * days)
	
	-- Get the calendars
	set sourceCalendar to calendar "Original calendar on MS"
	set destinationCalendar to calendar "Your iCloud calendar to copy the events into"
	
	log "Starting intelligent calendar sync from " & (name of sourceCalendar) & " to " & (name of destinationCalendar)
	
	-- STEP 1: GET ALL EVENTS FROM BOTH CALENDARS
	log "Collecting events from both calendars..."
	
	set sourceEvents to (every event of sourceCalendar whose start date ≥ startDateRange and start date ≤ endDateRange)
	set destEvents to (every event of destinationCalendar whose start date ≥ startDateRange and start date ≤ endDateRange)
	
	log "Found " & (count of sourceEvents) & " events in source and " & (count of destEvents) & " events in destination"
	
	-- STEP 2: CREATE SIGNATURES FOR EVENTS
	-- These signatures will help us compare events between calendars
	
	-- Create signatures for source events
	set sourceSignatures to {}
	repeat with anEvent in sourceEvents
		-- Extract event properties
		set eventSummary to summary of anEvent
		set eventStart to start date of anEvent
		
		-- Get hour and minute
		set eventHour to hours of eventStart as string
		set eventMinute to minutes of eventStart as string
		
		-- Get date components
		set eventDay to day of eventStart as string
		set eventMonth to month of eventStart as string
		set eventYear to year of eventStart as string
		
		-- Check if recurring
		set isRecurring to false
		set recurrenceInfo to ""
		try
			set recurrenceValue to recurrence of anEvent
			if recurrenceValue is not missing value then
				set isRecurring to true
				set recurrenceInfo to recurrenceValue
			end if
		end try
		
		-- Create a structured signature record
		set eventSignature to {|summary|:eventSummary, hour:eventHour, minute:eventMinute, |day|:eventDay, |month|:eventMonth, |year|:eventYear, recurring:isRecurring, |recurrence|:recurrenceInfo, |sourceEvent|:anEvent}
		
		set end of sourceSignatures to eventSignature
	end repeat
	
	-- Create signatures for destination events
	set destSignatures to {}
	repeat with anEvent in destEvents
		-- Extract event properties
		set eventSummary to summary of anEvent
		set eventStart to start date of anEvent
		
		-- Get hour and minute
		set eventHour to hours of eventStart as string
		set eventMinute to minutes of eventStart as string
		
		-- Get date components
		set eventDay to day of eventStart as string
		set eventMonth to month of eventStart as string
		set eventYear to year of eventStart as string
		
		-- Check if recurring
		set isRecurring to false
		set recurrenceInfo to ""
		try
			set recurrenceValue to recurrence of anEvent
			if recurrenceValue is not missing value then
				set isRecurring to true
				set recurrenceInfo to recurrenceValue
			end if
		end try
		
		-- Create a structured signature record
		set eventSignature to {|summary|:eventSummary, hour:eventHour, minute:eventMinute, |day|:eventDay, |month|:eventMonth, |year|:eventYear, recurring:isRecurring, |recurrence|:recurrenceInfo, |destEvent|:anEvent}
		
		set end of destSignatures to eventSignature
	end repeat
	
	-- STEP 3: IDENTIFY EVENTS TO DELETE FROM DESTINATION
	log "Identifying events to delete from destination..."
	
	set eventsToDelete to {}
	
	-- For each destination event, check if it exists in source
	repeat with destSig in destSignatures
		set destSummary to |summary| of destSig
		set destHour to hour of destSig
		set destMinute to minute of destSig
		set destDay to |day| of destSig
		set destMonth to |month| of destSig
		set destYear to |year| of destSig
		set destIsRecurring to recurring of destSig
		
		-- Look for a matching event in source
		set foundMatch to false
		
		repeat with sourceSig in sourceSignatures
			-- Compare essential properties
			if (|summary| of sourceSig is equal to destSummary) and ¬
				(hour of sourceSig is equal to destHour) and ¬
				(minute of sourceSig is equal to destMinute) and ¬
				(|day| of sourceSig is equal to destDay) and ¬
				(|month| of sourceSig is equal to destMonth) and ¬
				(|year| of sourceSig is equal to destYear) and ¬
				(recurring of sourceSig is equal to destIsRecurring) then
				
				-- Found a match!
				set foundMatch to true
				exit repeat
			end if
		end repeat
		
		if not foundMatch then
			-- This destination event doesn't exist in source - mark for deletion
			set end of eventsToDelete to |destEvent| of destSig
		end if
	end repeat
	
	log "Found " & (count of eventsToDelete) & " events to delete from destination"
	
	-- STEP 4: IDENTIFY EVENTS TO ADD TO DESTINATION
	log "Identifying events to add to destination..."
	
	set eventsToAdd to {}
	
	-- For each source event, check if it exists in destination
	repeat with sourceSig in sourceSignatures
		set sourceSummary to |summary| of sourceSig
		set sourceHour to hour of sourceSig
		set sourceMinute to minute of sourceSig
		set sourceDay to |day| of sourceSig
		set sourceMonth to |month| of sourceSig
		set sourceYear to |year| of sourceSig
		set sourceIsRecurring to recurring of sourceSig
		
		-- Look for a matching event in destination
		set foundMatch to false
		
		repeat with destSig in destSignatures
			-- Compare essential properties
			if (|summary| of destSig is equal to sourceSummary) and ¬
				(hour of destSig is equal to sourceHour) and ¬
				(minute of destSig is equal to sourceMinute) and ¬
				(|day| of destSig is equal to sourceDay) and ¬
				(|month| of destSig is equal to sourceMonth) and ¬
				(|year| of destSig is equal to sourceYear) and ¬
				(recurring of destSig is equal to sourceIsRecurring) then
				
				-- Found a match!
				set foundMatch to true
				exit repeat
			end if
		end repeat
		
		if not foundMatch then
			-- This source event doesn't exist in destination - mark for addition
			set end of eventsToAdd to |sourceEvent| of sourceSig
		end if
	end repeat
	
	log "Found " & (count of eventsToAdd) & " events to add to destination"
	
	-- STEP 5: PERFORM DELETIONS
	-- We'll do recurring events first, then non-recurring
	
	log "Deleting events from destination..."
	
	-- First delete recurring events
	set recurringDeleted to 0
	set nonRecurringDeleted to 0
	
	repeat with eventToDelete in eventsToDelete
		try
			set isRecurring to false
			try
				if recurrence of eventToDelete is not missing value then
					set isRecurring to true
				end if
			end try
			
			if isRecurring then
				log "Deleting recurring event: " & (summary of eventToDelete)
				delete eventToDelete
				set recurringDeleted to recurringDeleted + 1
				delay 0.5 -- Give extra time for recurring event deletion
			end if
		on error errMsg
			log "Error deleting recurring event: " & errMsg
		end try
	end repeat
	
	-- Allow time for Calendar to process recurring deletions
	if recurringDeleted > 0 then
		delay 3
	end if
	
	-- Now delete non-recurring events
	repeat with eventToDelete in eventsToDelete
		try
			set isRecurring to false
			try
				if recurrence of eventToDelete is not missing value then
					set isRecurring to true
				end if
			end try
			
			if not isRecurring then
				log "Deleting non-recurring event: " & (summary of eventToDelete)
				delete eventToDelete
				set nonRecurringDeleted to nonRecurringDeleted + 1
			end if
		on error errMsg
			log "Error deleting non-recurring event: " & errMsg
		end try
	end repeat
	
	log "Deleted " & recurringDeleted & " recurring events and " & nonRecurringDeleted & " non-recurring events"
	
	-- Allow time for Calendar to process all deletions
	if (recurringDeleted + nonRecurringDeleted) > 0 then
		delay 2
	end if
	
	-- STEP 6: PERFORM ADDITIONS
	-- We'll do recurring events first, then non-recurring
	
	log "Adding events to destination..."
	
	-- First add recurring events
	set recurringAdded to 0
	set nonRecurringAdded to 0
	
	repeat with eventToAdd in eventsToAdd
		try
			set isRecurring to false
			try
				if recurrence of eventToAdd is not missing value then
					set isRecurring to true
				end if
			end try
			
			if isRecurring then
				set eventSummary to summary of eventToAdd
				set eventStartDate to start date of eventToAdd
				set eventEndDate to end date of eventToAdd
				
				if eventEndDate is missing value then
					set eventEndDate to eventStartDate + (60 * minutes)
				end if
				
				log "Adding recurring event: " & eventSummary
				
				-- Create the event
				set newEvent to make new event at end of destinationCalendar with properties {summary:eventSummary, start date:eventStartDate, end date:eventEndDate}
				
				-- Set the recurrence pattern
				set recurrencePattern to recurrence of eventToAdd
				set recurrence of newEvent to recurrencePattern
				
				set recurringAdded to recurringAdded + 1
				delay 0.5 -- Give extra time for recurring event creation
			end if
		on error errMsg
			log "Error adding recurring event: " & errMsg
		end try
	end repeat
	
	-- Allow time for Calendar to process recurring additions
	if recurringAdded > 0 then
		delay 3
	end if
	
	-- Now add non-recurring events
	repeat with eventToAdd in eventsToAdd
		try
			set isRecurring to false
			try
				if recurrence of eventToAdd is not missing value then
					set isRecurring to true
				end if
			end try
			
			if not isRecurring then
				set eventSummary to summary of eventToAdd
				set eventStartDate to start date of eventToAdd
				set eventEndDate to end date of eventToAdd
				
				if eventEndDate is missing value then
					set eventEndDate to eventStartDate + (60 * minutes)
				end if
				
				log "Adding non-recurring event: " & eventSummary
				
				-- Create the event
				make new event at end of destinationCalendar with properties {summary:eventSummary, start date:eventStartDate, end date:eventEndDate}
				
				set nonRecurringAdded to nonRecurringAdded + 1
			end if
		on error errMsg
			log "Error adding non-recurring event: " & errMsg
		end try
	end repeat
	
	log "Added " & recurringAdded & " recurring events and " & nonRecurringAdded & " non-recurring events"
	
	-- Completion message
	log "Calendar sync completed successfully"
	log "Summary: " & recurringDeleted & " recurring and " & nonRecurringDeleted & " regular events deleted, " & recurringAdded & " recurring and " & nonRecurringAdded & " regular events added"
end tell
