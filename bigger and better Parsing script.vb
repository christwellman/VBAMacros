'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This procedure exports parsed data from an email and exports the  '
' items to an excel workbook/sheet                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ExportToExcel()

On Error GoTo ErrHandler

Dim appExcel As Excel.Application
Dim wkb As Excel.Workbook
Dim wks As Excel.Worksheet
Dim rng As Excel.Range
Dim strSheet As String
Dim strPath As String
Dim intRowCounter As Integer
Dim intColumnCounter As Integer
Dim msg As Outlook.MailItem
Dim nms As Outlook.NameSpace
Dim fld As Outlook.MAPIFolder
Dim itm As Object
Dim NextRow As Integer
Dim projType As String


'Text Parsing Variables
Dim ParseText As String
Dim ParseDate As Double
Dim ParseNumber As Integer
Dim intLocLabel As Integer
Dim intLocCRLF As Integer
Dim intLenLabel As Integer
Dim strText As String

'Define location of document to add new records to
'File name
strSheet = "DataCenterPracticeNewMetricsDatasheet.xlsm"
'File Path
strPath = "C:\Users\ctwellma\Desktop\"
strSheet = strPath & strSheet

Debug.Print strSheet
  'Select export folder
Set nms = Application.GetNamespace("MAPI")

Set fld = nms.PickFolder
  'Handle potential errors with Select Folder dialog box.
If fld Is Nothing Then
    MsgBox "There are no mail messages to export", vbOKOnly, "Error"

Exit Sub

ElseIf fld.DefaultItemType <> olMailItem Then

    MsgBox "There are no mail messages to export", vbOKOnly, "Error"

Exit Sub

ElseIf fld.Items.Count = 0 Then

    MsgBox "There are no mail messages to export", vbOKOnly, "Error"

Exit Sub

End If
  'Open and activate Excel workbook.
Set appExcel = CreateObject("Excel.Application")

appExcel.Workbooks.Open (strSheet)

Set wkb = appExcel.ActiveWorkbook

Set wks = wkb.Sheets(1)

wks.Activate

appExcel.Application.Visible = True

'Where to start Populating Data:
'NextRow = 3
NextRow = LastRow(wks.Range("A1:AS1"))

'***Debug Find LastRow?
'MsgBox "Last Row is: " & NextRow

intRowCounter = NextRow


  'Copy field items in mail folder.
For Each itm In fld.Items
 
        intColumnCounter = 1
        Set msg = itm
        intRowCounter = intRowCounter + 1
        PID = ParseTextLinePair(msg.body, "PID:")
        
        'Parse Submit time
        Set rng = wks.Cells(intRowCounter, intColumnCounter)
        rng.Value = ParseTextLinePair(msg.body, "Date and Time of Submission:")
        'rng.Value = CDate(SumbissionDate)
        intColumnCounter = intColumnCounter + 1
           
        'Parse Requestor
        Set rng = wks.Cells(intRowCounter, intColumnCounter)
        rng.Value = msg.SenderName
        intColumnCounter = intColumnCounter + 1
		
        'Parse PID
        Set rng = wks.Cells(intRowCounter, intColumnCounter)
        rng.Value = ParseTextLinePair(msg.body, "PID:")
        intColumnCounter = intColumnCounter + 1
		
		'PID Status
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
        rng.Value = InputBox(Prompt:="What is the project Status in OP for PID: " & PID, Title:="Project Type?", Default:="")
        intColumnCounter = intColumnCounter + 1
		
		'Customer Name
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
        rng.Value = ParseTextLinePair(msg.body, "Customer Name:")
        intColumnCounter = intColumnCounter + 1
		
		'Eng Desk Notes
		intColumnCounter = intColumnCounter + 1
		
		'Service Type(s)
		intColumnCounter = intColumnCounter + 1
		
		'Delivery Location City
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
		locCity = ParseTextLinePair(msg.body, "Customer Site Location:")
        rng.Value = InputBox(Prompt:="What is the delivery City for PID: " & PID, Title:="Delivery City?", Default:=locCity)
        intColumnCounter = intColumnCounter + 1
		
		'Delivery Location State
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
		locState = ParseTextLinePair(msg.body, "Customer Site Location:")
        rng.Value = InputBox(Prompt:="What is the delivery State for PID: " & PID, Title:="Delivery State?", Default:=locState)
        intColumnCounter = intColumnCounter + 1
		
		'Request Type
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
        rng.Value = ParseTextLinePair(msg.body, "Request Type:")
        intColumnCounter = intColumnCounter + 1
		
		'Project Type
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
        projType = ParseTextLinePair(msg.body, "Project Type:")
        rng.Value = InputBox(Prompt:="What is the project type in OP for PID: " & PID, Title:="Project Type?", Default:=projType)
        intColumnCounter = intColumnCounter + 1
		
		'Customer Primary Contact
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
        rng.Value = ParseTextLinePair(msg.body, "Customer Primary Contact:")
        intColumnCounter = intColumnCounter + 1
		
		'Services Description
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
        rng.Value = ParseTextLinePair(msg.body, "Service Description:")
        intColumnCounter = intColumnCounter + 1
		
		'Services Revenue
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
        ServRev = ParseTextLinePair(msg.body, "Services Revenue:")
        rng.Value = InputBox(Prompt:="How much service revenue is generated by PID: " & PID, Title:="Services Revenue?", Default:=ServRev)
        intColumnCounter = intColumnCounter + 1
		
		'Funding
		intColumnCounter = intColumnCounter + 1
		
		'Oracle Project Name
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
        rng.Value = InputBox(Prompt:="What's the OP Project Name for PID: " & PID, Title:="OP Project Name?", Default:="")
        intColumnCounter = intColumnCounter + 1
		
		'Theater
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
        Segment = ParseTextLinePair(msg.body, "Theather/Market: Mkt Seg -")
        rng.Value = InputBox(Prompt:="What is the project segment in OP for PID: " & (ParseTextLinePair(msg.body, "PID:")), Title:="Project Type?", Default:=Segment)
        'rng.Value = ParseTextLinePair(msg.body, "Theather/Market: Mkt Seg - ")
        intColumnCounter = intColumnCounter + 1
		
		'SO#
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
        rng.Value = ParseTextLinePair(msg.body, "Sales Order Nbr:")
        intColumnCounter = intColumnCounter + 1
		
		'Deal ID
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
        rng.Value = ParseTextLinePair(msg.body, "DID:")
        intColumnCounter = intColumnCounter + 1
		
		'Project Start Date
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
        rng.Value = ParseTextLinePair(msg.body, "Project Start Date:")
        rng.Value = CDate(rng.Value)
		'need to add if statement to only CDATE If numeric
        intColumnCounter = intColumnCounter + 1
		
		'Project Kick Off Date
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
        rng.Value = ParseTextLinePair(msg.body, "Project Kick-Off Meeting:")
		rng.Value = CDate(rng.Value)
        intColumnCounter = intColumnCounter + 1
		
		'Project End Date
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
        rng.Value = ParseTextLinePair(msg.body, "End Date:")
        rng.Value = CDate(rng.Value)
        intColumnCounter = intColumnCounter + 1
		
		'Margin Analysis 
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
        rng.Value = ParseTextLinePair(msg.body, "Latest version of proposal or SOW:")
        intColumnCounter = intColumnCounter + 1
		
		'SOW/ASPT Quote
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
        rng.Value = ParseTextLinePair(msg.body, "Margin analysis spreadsheet:")
        intColumnCounter = intColumnCounter + 1
		
		'Market Segment
		''''Segment Vs Theater??
		intColumnCounter = intColumnCounter + 1
		
		'Project Status
		Set rng = wks.Cells(intRowCounter, intColumnCounter)
        rng.Value = "New"
        intColumnCounter = intColumnCounter + 1
		
		'Delivery Close Date
        intColumnCounter = intColumnCounter + 1

		'Project Close Date
		intColumnCounter = intColumnCounter + 1
		
		'Past Due Date
		intColumnCounter = intColumnCounter + 1	
			
		'Walker Survey Sent Date
		intColumnCounter = intColumnCounter + 1
		
		'Sales rep
	    intColumnCounter = intColumnCounter + 1
		
		'DC PM Assigned
		'Set rng = wks.Cells(intRowCounter, intColumnCounter)
        'rng.Value = "Needs DCPM"
        intColumnCounter = intColumnCounter + 1
		
		'Work Manager Assigned
		'Set rng = wks.Cells(intRowCounter, intColumnCounter)
        'rng.Value = "Needs DCPM"
        intColumnCounter = intColumnCounter + 1
		
		'Technical resourcing Status
		intColumnCounter = intColumnCounter + 1
		
		'Initial Follow up Sent
		intColumnCounter = intColumnCounter + 1
		
		'DCPM Assigned Date
		intColumnCounter = intColumnCounter + 1
		
		'Technical Resource Assigned Date
		intColumnCounter = intColumnCounter + 1
		
		'WM has PID
		intColumnCounter = intColumnCounter + 1
		
		'PM Assigned to PID
		intColumnCounter = intColumnCounter + 1
		
		'Last Cost Forecast Date
		intColumnCounter = intColumnCounter + 1
		
		'Days since last Cost Forecast
		intColumnCounter = intColumnCounter + 1
		
		'Workplan Chargeable
		intColumnCounter = intColumnCounter + 1
		
		'Revenue Recognized to date
		intColumnCounter = intColumnCounter + 1
		
		'Costs to Date
		intColumnCounter = intColumnCounter + 1
		
		'Margin
		intColumnCounter = intColumnCounter + 1
       
        
        
Next itm

Set appExcel = Nothing
Set wkb = Nothing
Set wks = Nothing
Set rng = Nothing
Set msg = Nothing
Set nms = Nothing
Set fld = Nothing
Set itm = Nothing
Exit Sub

ErrHandler: If Err.Number = 1004 Then

    MsgBox strSheet & " doesn't exist", vbOKOnly, "Error"

Else

    MsgBox Err.Number & "; Description: ", vbOKOnly, "Error"

End If

'Zero Variables
Set appExcel = Nothing
Set wkb = Nothing
Set wks = Nothing
Set rng = Nothing
Set msg = Nothing
Set nms = Nothing
Set fld = Nothing
Set itm = Nothing
End Sub

Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application
         
    Set objApp = CreateObject("Outlook.Application")
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
        Case Else
            ' anything else will result in an error, which is
            ' why we have the error handler above
    End Select
     
    Set objApp = Nothing
End Function