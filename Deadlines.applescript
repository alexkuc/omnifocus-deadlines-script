use application "OmniFocus"

use scripting additions

tell application "OmniOutliner"
	
	if version > 5 then
		
		set template to ((path to application support from user domain as string) & "The Omni Group:OmniOutliner:Pro Templates:" & "Blank.otemplate:") --- template path for OO5
		
	else
		
		set template to ((path to application support from user domain as string) & "The Omni Group:OmniOutliner:Templates:" & "Blank.oo3template:") --- template for OO3 or OO4
		
	end if
	
	set currentDate to current date
	
	activate
	
	open template
	
	tell front document
		
		set title of second column to "Project Name"
		
		make new column with properties {title:"Days Left", column type:numeric, column format:{id:"no-thousands-no-decimal"}, sort order:ascending, alignment:center}
		
		repeat with thisProject in (flattened projects of default document whose status is active and due date is not missing value)
			
			if defer date of thisProject ² currentDate then
				
				if text of second cell of first row is equal to "" then
					
					set newRow to first row
					
				else
					
					set newRow to make new row at end of (parent of last row)
					
				end if
				
				set text of second cell of newRow to name of thisProject
				
				set value of attribute named "link" of style of text of second cell of newRow to "omnifocus:///task/" & id of thisProject
								
				set timeLeft to ((due date of thisProject) - currentDate) / 3600 / 24 as integer
				
				set value of third cell of newRow to timeLeft
				
			end if
			
		end repeat
		
	end tell
	
end tell
