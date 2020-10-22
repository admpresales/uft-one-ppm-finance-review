'===========================================================
'20200929 - DJ: Added .sync statements after .click statements
'20201001 - DJ: Added the ClickLoop and the PPMProposalSearch functions, removed duplicative code
'20201006 - DJ: Updated PPMProposalSearch function to look for the text Saved Searches to know Search click worked and commented .sync after that to prevent PPM from auto closing the menu popup
'			Corrected inaccurate comment
'			Removed unused function
'			Added function comments
'			Updated OR to have more logical names
'20201006 - DJ: Added steps to add projected financial costs into the Financial Summary
'20201006 - DJ: Updated steps that weren't working on lower resolutions
'20201008 - DJ: Updated for missing step of saving the changes to the financial details before closing and associate .sync
'20201013 - DJ: Modified the ClickLoop retry counter to be 3 instead of 90
'20201020 - DJ: Updated to handle changes coming in UFT One 15.0.2
'				Commented out the msgbox, which can cause UFT One to be in a locked state when executed from Jenkins
'20201022 - DJ: Updated ClickLoop to gracefully abort if failure number reached
'20201022 - DJ: Disabled smart identification on Browser("Create a Blank Staffing").Page("Edit Costs_2").Frame("CopyCostsDialog").WebButton("CopyButton")
'				Updated the click on the add button statement to use ClickLoop
'===========================================================

'===========================================================
'Function to retry action if next step doesn't show up
'===========================================================
Function ClickLoop (AppContext, ClickStatement, SuccessStatement)
	
	Dim Counter
	
	Counter = 0
	Do
		ClickStatement.Click
		AppContext.Sync																				'Wait for the browser to stop spinning
		Counter = Counter + 1
		wait(1)
		If Counter >=3 Then
			Reporter.ReportEvent micFail, "Click Statement", "The Success Statement didn't display within " & Counter & " attempts.  Aborting run"
			'===========================================================================================
			'BP:  Logout
			'===========================================================================================
			AIUtil.SetContext AppContext																'Tell the AI engine to point at the application
			Browser("Search Requests").Page("Req Details").WebElement("menuUserIcon").Click
			AppContext.Sync																				'Wait for the browser to stop spinning
			AIUtil.FindText("Sign Out (").Click
			AppContext.Sync																				'Wait for the browser to stop spinning
			While Browser("CreationTime:=0").Exist(0)   												'Loop to close all open browsers
				Browser("CreationTime:=0").Close 
			Wend
			ExitAction
		End If
	Loop Until SuccessStatement.Exist(10)
	AppContext.Sync																				'Wait for the browser to stop spinning

End Function

'===========================================================
'Function to search for the PPM proposal in the appropriate status
'===========================================================
Function PPMProposalSearch (CurrentStatus, NextAction)
	'===========================================================================================
	'BP:  Click the Search menu item
	'===========================================================================================
	Set ClickStatement = AIUtil.FindText("SEARCH", micFromTop, 1)
	Set SuccessStatement = AIUtil.FindText("Saved Searches")
	ClickLoop AppContext, ClickStatement, SuccessStatement
	'AppContext.Sync																				'Wait for the browser to stop spinning
	
	'===========================================================================================
	'BP:  Click the Requests text
	'===========================================================================================
	Set ClickStatement = AIUtil.FindTextBlock("Requests", micFromTop, 1)
	Set SuccessStatement = AIUtil("text_box", "Request Type:")
	ClickLoop AppContext, ClickStatement, SuccessStatement
	AppContext.Sync																				'Wait for the browser to stop spinning
	
	'===========================================================================================
	'BP:  Enter PFM - Proposal into the Request Type field
	'===========================================================================================
	AIUtil("text_box", "Request Type:").Type "PFM - Proposal"
	AIUtil.FindText("Status").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	
	'===========================================================================================
	'BP:  Enter a status of "New" into the Status field
	'===========================================================================================
	AIUtil("text_box", "Status").Type CurrentStatus
	
	'===========================================================================================
	'BP:  Click the Search button (OCR not seeing text, use traditional OR)
	'===========================================================================================
	Browser("Search Requests").Page("Search Requests").Link("Search").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	
	'===========================================================================================
	'BP:  Click the first record returned in the search results
	'===========================================================================================
	DataTable.Value("dtFirstReqID") = Browser("Search Requests").Page("Request Search Results").WebTable("Req #").GetCellData(2,2)
	Set ClickStatement = AIUtil.FindTextBlock(DataTable.Value("dtFirstReqID"))
	Set SuccessStatement = AIUtil.FindText(NextAction)
	ClickLoop AppContext, ClickStatement, SuccessStatement
	AppContext.Sync																				'Wait for the browser to stop spinning
	
End Function

Dim BrowserExecutable, Counter, mySendKeys, rc

While Browser("CreationTime:=0").Exist(0)   												'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend
BrowserExecutable = DataTable.Value("BrowserName") & ".exe"
SystemUtil.Run BrowserExecutable,"","","",3													'launch the browser specified in the data table
Set AppContext=Browser("CreationTime:=0")													'Set the variable for what application (in this case the browser) we are acting upon
Set AppContext2=Browser("CreationTime:=1")													'Set the variable for what application (in this case the browser) we are acting upon

'===========================================================================================
'BP:  Navigate to the PPM Launch Pages
'===========================================================================================

AppContext.ClearCache																		'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.Navigate DataTable.Value("URL")													'Navigate to the application URL
AppContext.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil.SetContext AppContext																'Tell the AI engine to point at the application

'===========================================================================================
'BP:  Click the Strategic Portfolio link
'===========================================================================================
AIUtil.FindText("Strategic Portfolio").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Andy Stein (IT Financial Manager) link to log in as Andy Stein
'===========================================================================================
AIUtil.FindTextBlock("Andy Stein").Click
AppContext.Sync																				'Wait for the browser to stop spinning
rc = AIUtil.FindTextBlock("Proposals Eligible for My Action (Financial Review)").Exist

'===========================================================================================
'BP:  Search for propsals currently in a status of "Finance Review"
'===========================================================================================
PPMProposalSearch "Finance Review", "Approved"

'===========================================================================================
'BP:  Click the Business Case Sta link to move down the form
'===========================================================================================
AIUtil.FindTextBlock("Business Case Sta").Click

'===========================================================================================
'BP:  Click the link for the Financial Summary
'===========================================================================================
AIUtil.FindText("Proposal Name ", micFromBottom, 1).Click

'===========================================================================================
'BP:  Maximize the popup window
'===========================================================================================
AppContext2.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext2.Sync																				'Wait for the browser to stop spinning
AIUtil.SetContext AppContext2																'Tell the AI engine to point at the application

'===========================================================================================
'BP:  Click the Add Costs link, use traditional OR as it isn't visible on the screen, but is on the page
'===========================================================================================
Browser("Create a Blank Staffing").Page("Financial Summary").Link("Add Costs").Click
AppContext2.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the copy costs button
'===========================================================================================
Browser("Create a Blank Staffing").Page("Edit Costs").WebElement("Copy Costs Button").Click

'===========================================================================================
'BP:  Click the Copy from Another Request text 
'===========================================================================================
AIUtil.FindTextBlock("Copy from Another Request").Click

'===========================================================================================
'BP:  Click the Include Project radio button
'===========================================================================================
AIUtil("radio_button", "0 Include Project:").SetState "on"

'===========================================================================================
'BP:  Type Web for One World into the Include Project text bos
'===========================================================================================
AIUtil("text_box", "@ Include Project:").Type "Web for One World"

'===========================================================================================
'BP:  Click the Copy Cost Lines text to get the application to run the value entry validation
'===========================================================================================
AIUtil.FindTextBlock("Copy Cost Lines").Click
Set ClickStatement = AIUtil.FindTextBlock("Copy Cost Lines")
Set SuccessStatement = AIUtil("button", "Add")
ClickLoop AppContext, ClickStatement, SuccessStatement

'===========================================================================================
'BP:  Click the Add button
'===========================================================================================
Set ClickStatement = AIUtil("button", "Add")
Set SuccessStatement = AIUtil.FindTextBlock("Are you sure you want to copy cost lines from the source request?")
ClickLoop AppContext, ClickStatement, SuccessStatement

'===========================================================================================
'BP:  Click the Copy Forecast Values check box
'===========================================================================================
AIUtil("check_box", "C1 Copy Forecast Values").SetState "On"

'===========================================================================================
'BP:  Click the Copy Copy button, detection improvement submitted.
'===========================================================================================
'AIUtil("button", "", micFromBottom, 1).Click
Browser("Create a Blank Staffing").Page("Edit Costs_2").Frame("CopyCostsDialog").WebButton("CopyButton").Click

'===========================================================================================
'BP:  Click the first 0.00 field and type 100
'===========================================================================================
AIUtil.FindTextBlock("0.000", micFromTop, 3).Click
Window("Edit Costs").Type "100" @@ hightlight id_;_1771790_;_script infofile_;_ZIP::ssf2.xml_;_
AIUtil.FindTextBlock("Contractor").Click
Browser("Create a Blank Staffing").Page("Edit Costs_3").WebButton("Save").Click
AppContext2.Sync																			'Close the application at the end of your script

'===========================================================================================
'BP:  Click the Done button, detection improvement submitted.
'===========================================================================================
'AIUtil("button", "", micFromRight, 2).Click
Browser("Create a Blank Staffing").Page("Edit Costs_2").WebButton("Done").Click
AppContext2.Close																			'Close the application at the end of your script

'===========================================================================================
'BP:  Close the popup window
'===========================================================================================
AIUtil.SetContext AppContext																'Tell the AI engine to point at the application

'===========================================================================================
'BP:  Click the Save text
'===========================================================================================
Set ClickStatement = AIUtil.FindText("Save", micFromLeft, 1)
Set SuccessStatement = AIUtil.FindText("Approved")
ClickLoop AppContext, ClickStatement, SuccessStatement

'===========================================================================================
'BP:  Click the Approved button
'===========================================================================================
AIUtil.FindText("Approved").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Logout
'===========================================================================================
Browser("Search Requests").Page("Req Details").WebElement("menuUserIcon").Click
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil.FindTextBlock("Sign Out (Andy Stein)").Click
AppContext.Sync																				'Wait for the browser to stop spinning

AppContext.Close																			'Close the application at the end of your script

