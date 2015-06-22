-- DocumentTool for InDesign
-- version 1.2

-- created by medul6, Michael Heck, 2014
-- NOT open sourced YET on September 7th, 2012 on Github > check the LICENSE.txt and README.md in the repository for detailed information
-- https://github.com/medul6/...

-- еееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееее

-- global variables
global activeDocument
global openDocuments
global otherDocuments
--global activeWindow
global openWindows
--global pdfPresetsOnComputer
--global preservedPageRange
global stopBool
global splittedRange
--global pageNumberInsertionpoint
--global inputRange
global splittedMagic
--global splittedRangeReverse
global splittedRangeMagic
--global splittedRangeMagicLoop
--global incrementValue
--global repeatNumber
--global textOverflows

--test variables!!!
--global xxx
--global filePath
--global chosenPresetText
--global docName
--global newFilePath
--global pathItems
--global pageRange
--global newdocName
--global failedLinks
--global textOverflows
--global modifiedLinks
--global missingLinks
--global exportPreset

--properties!
property functionChoice : {"Vorschau einschalten"}
--property chosenPreset : {"sk-Screen"}
--property pageRange : "all pages"

-- еееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееее

tell application id "com.adobe.InDesign"
	
	set stopBool to false
	-- set up some informations from the current state as variables
	set activeDocument to active document
	set activeWindow to active window
	set openWindows to every window
	set openDocuments to every document
	set otherDocuments to every document whose id is not activeDocument's id
	-- only pdf presets are captured that are not build in. we have our own! remove the whose clause to show all of them, or modify the whose clause to show only them.
	--set pdfPresetsOnComputer to name of every PDF export preset whose name does not contain "["
	
	-- initialize some lists (to be filled in the next two repeat loops)
	--set splittedMagic to {}
	
	
	--my linkCheck()
	--my textOverflowCheck()
	my pageCountCheck()
	
	my functionChooser()
	
	
	if stopBool is true then
		displayTheEnd() of me
	end if
	
end tell

-- еееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееее

on functionChooser()
	set functionChoice to choose from list {"Bezugspunkt setzen", "Vorschau einschalten", "Vorschau ausschalten", "Alle KapitelanfКnge lЪschen", "Seiten lЪschen ...", "Seiten einfЯgen ...", "Seiten verschieben ..."} default items functionChoice with prompt "Funktion wКhlen:" OK button name "Weiter!"
	
	if the functionChoice = {"Bezugspunkt setzen"} then
		my setOrigin()
	else if the functionChoice = {"Vorschau einschalten"} then
		my previewOn()
	else if the functionChoice = {"Vorschau ausschalten"} then
		my previewOff()
	else if the functionChoice = {"Alle KapitelanfКnge lЪschen"} then
		my deleteEverySection()
	else if the functionChoice = {"Seiten lЪschen ..."} then
		my deletePages()
	else if the functionChoice = {"Seiten einfЯgen ..."} then
		my insertPages()
	else if the functionChoice = {"Seiten verschieben ..."} then
		my movePages()
	end if
	
end functionChooser

-- еееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееее

on pageCountCheck()
	tell application id "com.adobe.InDesign"
		
		set pageCountBool to true
		set pageCount to count pages of activeDocument
		set pageCountRepeat to pageCount
		
		repeat with x from 1 to count otherDocuments
			set pageCountRepeat to count pages of otherDocuments's item x
			if pageCountRepeat is not equal to pageCount then
				set pageCountBool to false
			end if
			if pageCountBool is false then
				display dialog "Dokumente benЪtigen die gleiche Seitenanzahl! " & return & "-----------------------------------------" & return & ((name of otherDocuments's item x) as string) & return & "-----------------------------------------" & return & "hat eine unterschiedliche Seitenanzahl!" buttons "OK" default button "OK"
			end if
		end repeat
	end tell
end pageCountCheck

-- еееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееее

on deleteEverySection()
	tell application id "com.adobe.InDesign"
		repeat with x from 1 to count openDocuments -- this iterates through all open documents
			
			set sectionsOfActiveDocument to every section of openDocuments's item x
			
			repeat with y from 2 to count sectionsOfActiveDocument
				delete item y of sectionsOfActiveDocument
			end repeat
			
		end repeat
		
	end tell
	set stopBool to true
end deleteEverySection

-- еееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееее

on deletePages()
	tell application id "com.adobe.InDesign"
		
		set buttonName to functionChoice & " !" as string
		
		display dialog "Welche Seiten sollen gelЪscht werden?" & return & "Seiten mЯssen nicht zusammenhКngen, z.B. '2-3,8-19'" & return & "Aber Reihenfolge einhalten(!): '2-3,8-19' nicht '8-19,2-3'" default answer "" buttons {"Abbrechen!", (buttonName as string)} default button (buttonName as string)
		if button returned of result is "Abbrechen!" then
			return
		else
			set inputRange to (text returned of result)
		end if
		
		my inputRangeSplitter(inputRange)
		my MagicSplitter(splittedRange)
		
		--set splittedRangeReverse to reverse of splittedRange
		set splittedRangeReverse to reverse of splittedMagic
		
		repeat with x from 1 to count openDocuments -- this iterates through all open documents
			repeat with y from 1 to count splittedRangeReverse -- this iterates through all pages
				delete page (splittedRangeReverse's item y) of openDocuments's item x
			end repeat
		end repeat
		
	end tell
	set stopBool to true
end deletePages

-- еееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееее

on insertPages()
	tell application id "com.adobe.InDesign"
		
		set buttonName to functionChoice & " !" as string
		
		display dialog "Wieviele Seiten sollen eingefЯgt werden?" & return & "NUR ganze Zahlen, gerade oder ungerade! z.B. '2' oder '7'" default answer "" buttons {"Abbrechen!", (buttonName as string)} default button (buttonName as string)
		if button returned of result is "Abbrechen!" then
			return
		else
			set numerOfPagesToInsert to (text returned of result)
		end if
		
		
		
		if numerOfPagesToInsert contains "," then
			display dialog "NUR ganze Zahlen!!!" buttons {"Hab's verstanden!"} default button "Hab's verstanden!"
			return
		else if numerOfPagesToInsert contains "-" then
			display dialog "NUR ganze Zahlen!!!" buttons {"Hab's verstanden!"} default button "Hab's verstanden!"
			return
		end if
		
		
		display dialog "Nach welcher Seite sollen die Seiten eingefЯgt werden?" & return & "NUR ganze Zahlen, keine Bereiche!" default answer "" buttons {"Abbrechen!", (buttonName as string)} default button (buttonName as string)
		if button returned of result is "Abbrechen!" then
			return
		else
			set pageNumberInsertionpoint to (text returned of result)
		end if
		
		
		if pageNumberInsertionpoint contains "," then
			display dialog "NUR ganze Zahlen!!!" buttons {"Hab's verstanden!"} default button "Hab's verstanden!"
			return
		else if pageNumberInsertionpoint contains "-" then
			display dialog "NUR ganze Zahlen!!!" buttons {"Hab's verstanden!"} default button "Hab's verstanden!"
			return
		end if
		
		
		repeat with x from 1 to count openDocuments -- this iterates through all open documents
			repeat with y from 1 to (numerOfPagesToInsert as integer)
				make page at after page pageNumberInsertionpoint of openDocuments's item x
			end repeat
		end repeat
		
	end tell
	set stopBool to true
end insertPages

-- еееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееее

on movePages()
	tell application id "com.adobe.InDesign"
		
		set buttonName to functionChoice & " !" as string
		
		display dialog "Welche Seiten sollen verschoben werden?" & return & "Immer nur zusammenhКngende DruckbЪgen!" & return & "Mit Divis oder Kommagetrennt! z.B. '4-5' oder '4,5'" & return & "Es kЪnnen auch mehrere BЪgen sein! z.B. '4-9' oder '12-23'" default answer "" buttons {"Abbrechen!", (buttonName as string)} default button (buttonName as string)
		if button returned of result is "Abbrechen!" then
			return
		else
			set inputRange to (text returned of result)
		end if
		
		--if inputRange contains "" then
		--	return
		--end if
		
		if inputRange contains "," then
			my inputRangeSplitter(inputRange)
		else if inputRange contains "-" then
			my inputRangeSplitterFromTo(inputRange)
		end if
		
		--my inputRangeSplitter(inputRange)
		
		display dialog "Nach welcher Seite sollen die Seiten verschoben werden?" & return & "NUR ganze Zahlen, keine Bereiche!" default answer "" buttons {"Abbrechen!", (buttonName as string)} default button (buttonName as string)
		if button returned of result is "Abbrechen!" then
			return
		else
			set pageNumberInsertionpoint to (text returned of result)
		end if
		
		if inputRange contains pageNumberInsertionpoint then
			return
		end if
		
		
		--set splittedRangeReverse to reverse of splittedRange
		if (count splittedRange) > "2" then
			display dialog "Immer Doppelseite fЯr Doppelseite verschieben!" & return & "Es wurden mehr als zwei Seiten angegeben!"
			set splittedRange to {"", ""}
		else if (count splittedRange) < "2" then
			display dialog "Immer Doppelseite fЯr Doppelseite verschieben!" & return & "Es wurde nur eine Seite angegeben!"
			set splittedRange to {"", ""}
		end if
		
		repeat with x from 1 to count openDocuments -- this iterates through all open documents
			--repeat with y from 1 to count splittedRange -- this iterates through all pages
			--delete page (splittedRange's item y) of openDocuments's item x
			tell openDocuments's item x
				move (pages (splittedRange's item 1 as integer) thru (splittedRange's item 2 as integer)) to after (page (pageNumberInsertionpoint as integer))
			end tell
			--end repeat
		end repeat
		
		--tell the active document
		--	move (pages 2 thru 3) to before page 1
		--end tell
		
		
		
	end tell
	set stopBool to true
end movePages

-- еееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееее

on inputRangeSplitter(inputRange)
	set oldDelimiters to AppleScript's text item delimiters -- always preserve original delimiters
	set AppleScript's text item delimiters to {","}
	set splittedRange to text items of inputRange
	set AppleScript's text item delimiters to oldDelimiters -- always restore original delimiters
	return splittedRange
end inputRangeSplitter

-- еееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееее

on inputRangeSplitterFromToMagic(inputRange)
	set oldDelimiters to AppleScript's text item delimiters -- always preserve original delimiters
	set AppleScript's text item delimiters to {"-"}
	
	set splittedRangeMagic to text items of inputRange
	
	if (splittedRangeMagic's item 2 as integer) is not ((splittedRangeMagic's item 1 as integer) + 1) then
		set incrementValue to (splittedRangeMagic's item 1 as integer) + 1
		set splittedRangeMagicLoop to {(splittedRangeMagic's item 1 as integer)}
		
		--repeat with x from (splittedRangeMagic's item 1 as integer) to (splittedRangeMagic's item 2 as integer) -- this iterates through
		--	set splittedRangeMagicLoop to splittedRangeMagicLoop & (incrementValue + 1)
		--end repeat
		
		set repeatNumber to (splittedRangeMagic's item 2 as integer) - (splittedRangeMagic's item 1 as integer)
		repeat repeatNumber times -- this iterates through
			set splittedRangeMagicLoop to splittedRangeMagicLoop & incrementValue
			set incrementValue to incrementValue + 1
		end repeat
		
		
		set splittedRangeMagic to splittedRangeMagicLoop
	end if
	
	set AppleScript's text item delimiters to oldDelimiters -- always restore original delimiters
	return splittedRangeMagic
end inputRangeSplitterFromToMagic

-- еееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееее

on inputRangeSplitterFromTo(inputRange)
	set oldDelimiters to AppleScript's text item delimiters -- always preserve original delimiters
	set AppleScript's text item delimiters to {"-"}
	set splittedRange to text items of inputRange
	set AppleScript's text item delimiters to oldDelimiters -- always restore original delimiters
	return splittedRange
end inputRangeSplitterFromTo

-- еееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееее

on MagicSplitter(splittedRange)
	set oldDelimiters to AppleScript's text item delimiters -- always preserve original delimiters
	
	set AppleScript's text item delimiters to {"-"}
	set splittedMagic to {}
	
	--set splittedMagic to text items of splittedRange
	
	repeat with x from 1 to count splittedRange -- this iterates through
		if splittedRange's item x does not contain "-" then
			set splittedMagic to splittedMagic & splittedRange's item x
			--set splittedRange's item x to end of splittedMagic
		else if splittedRange's item x contains "-" then
			inputRangeSplitterFromToMagic(splittedRange's item x)
			--set xxx to splittedRange
			set splittedMagic to splittedMagic & splittedRangeMagic's items
			--set splittedRange's items to end of splittedMagic
		end if
		
	end repeat
	
	set AppleScript's text item delimiters to oldDelimiters -- always restore original delimiters
	return splittedMagic
end MagicSplitter

-- еееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееее

on previewOn()
	tell application id "com.adobe.InDesign"
		
		repeat with x from 1 to count openWindows -- this iterates through all open documents
			set screen mode of openWindows's item x to preview to page
		end repeat
		
	end tell
end previewOn

-- еееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееее

on previewOff()
	tell application id "com.adobe.InDesign"
		
		repeat with x from 1 to count openWindows -- this iterates through all open documents
			set screen mode of openWindows's item x to preview off
		end repeat
		
	end tell
end previewOff

-- еееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееее

on setOrigin()
	tell application id "com.adobe.InDesign"
		
		set myOriginDialog to make dialog with properties {name:"Bezugspunkt setzen:"}
		
		tell myOriginDialog
			make dialog column
			tell the result
				make border panel
				tell the result
					make dialog column
					tell the result
						set myTopLeftCheckbox to make checkbox control with properties {static label:""}
						set myLeftCenterCheckbox to make checkbox control with properties {static label:""}
						set myBottomLeftCheckbox to make checkbox control with properties {static label:""}
					end tell
					--end tell
					
					make dialog column
					tell the result
						set myTopCenterCheckbox to make checkbox control with properties {static label:""}
						set myCenterCheckbox to make checkbox control with properties {static label:""}
						set myBottomCenterCheckbox to make checkbox control with properties {static label:""}
					end tell
					
					make dialog column
					tell the result
						set myTopRightCheckbox to make checkbox control with properties {static label:""}
						set myRightCenterCheckbox to make checkbox control with properties {static label:""}
						set myBottomRightCheckbox to make checkbox control with properties {static label:""}
					end tell
					
				end tell
				
			end tell
			
		end tell
		
		set myResult to show myOriginDialog
		
		if myResult is true then
			--Get the control settings from the dialog box.
			set myTopLeft to checked state of myTopLeftCheckbox
			set myLeftCenter to checked state of myLeftCenterCheckbox
			set myBottomLeft to checked state of myBottomLeftCheckbox
			set myTopCenter to checked state of myTopCenterCheckbox
			set myCenter to checked state of myCenterCheckbox
			set myBottomCenter to checked state of myBottomCenterCheckbox
			set myTopRight to checked state of myTopRightCheckbox
			set myRightCenter to checked state of myRightCenterCheckbox
			set myBottomRight to checked state of myBottomRightCheckbox
			
			destroy myOriginDialog
			
			if (myTopLeft or myLeftCenter or myBottomLeft or myTopCenter or myCenter or myBottomCenter or myTopRight or myRightCenter or myBottomRight) then
				repeat with x from 1 to count openWindows -- this iterates through all open documents
					if myTopLeft then
						set transform reference point of openWindows's item x to top left anchor
					else if myLeftCenter then
						set transform reference point of openWindows's item x to left center anchor
					else if myBottomLeft then
						set transform reference point of openWindows's item x to bottom left anchor
					else if myTopCenter then
						set transform reference point of openWindows's item x to top center anchor
					else if myCenter then
						set transform reference point of openWindows's item x to center anchor
					else if myBottomCenter then
						set transform reference point of openWindows's item x to bottom center anchor
					else if myTopRight then
						set transform reference point of openWindows's item x to top right anchor
					else if myRightCenter then
						set transform reference point of openWindows's item x to right center anchor
					else if myBottomRight then
						set transform reference point of openWindows's item x to bottom right anchor
					end if
				end repeat
			else
				display dialog "Es wurde keine Position ausgewКhlt"
			end if
			
		else
			destroy myOriginDialog
		end if
	end tell
end setOrigin

-- еееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееееее

on displayTheEnd()
	--display dialog "Fertig!" buttons "OK" default button "OK" giving up after 1
	say "OK!" using "Zarvox" --"Zarvox"
end displayTheEnd
