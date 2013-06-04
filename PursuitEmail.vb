

Function RangetoHTML(rng As Range)
' Works in Excel 2000, Excel 2002, Excel 2003, Excel 2007, Excel 2010, Outlook 2000, Outlook 2002, Outlook 2003, Outlook 2007, and Outlook 2010.
Dim fso As Object
Dim ts As Object
Dim TempFile As String
Dim TempWB As Workbook

TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
' Copy the range and create a workbook to receive the data.
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        '.DrawingObjects.Delete
        On Error GoTo 0
    End With
    ' Publish the sheet to an .htm file.
    With TempWB.PublishObjects.Add( _
        SourceType:=xlSourceRange, _
        Filename:=TempFile, _
        Sheet:=TempWB.Sheets(1).name, _
        Source:=TempWB.Sheets(1).UsedRange.Address, _
        HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
    
    ' Read all data from the .htm file into the RangetoHTML subroutine.
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", "align=left x:publishsource=")
   
    ' Close TempWB.
    TempWB.Close savechanges:=False

    ' Delete the htm file.
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

Sub SendTableToOutlook()
'Copy's Specified text from excel sheet to a temporary sheet and then pastes table into outlook message
Dim OutApp As Object
Dim OutMail As Object
Dim recip As String
Dim orgsheet As Worksheet
Dim Src As Range
Dim rng As Range
Dim val As Variant
Dim buf_in() As Variant
Dim cl As Collection

'Set Object libraries
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

'Set Orgin Sheet
Set orgsheet = ActiveWorkbook.Worksheets("Pursuit Tracker")

'Define where to get list of emails
Set Src = ActiveWorkbook.Worksheets("Pursuit Tracker").Range("S2:S1000") 'need to make this size dynamically

'Hide unneccessary Columns
Set rng = Nothing
orgsheet.Range("A:E,H:T,V:W,Y:Y").EntireColumn.Hidden = True
Set rng = orgsheet.UsedRange.SpecialCells(xlCellTypeVisible)

'Add Title to column Z
Cells(1, "Z").Value = "Last Updated"

'Turn off features to optimize
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    ' Catch issue if there is only one valid email
    If Src.Cells.Count = 1 Then
        dst.Value = Src.Value   ' ...which is not an array for a single cell
        MsgBox ("Theres only one email in the list -- come on you can do that manually")
        Exit Sub
    End If
    ' read all values at once
    buf_in = Src.Value
    Set cl = New Collection
    ' Skip all already-present or invalid values
    On Error Resume Next
    For Each val In buf_in
        cl.Add val, CStr(val)
    Next
    On Error GoTo 0
    
    For Each val In cl
        'check for valid email address
        If val Like "?*@?*.?*" Then 'And _ ' check for secondary validation
            'If email address exists and looks valid then filter Sheet by it
            orgsheet.ListObjects("Table_owssvr_1").Range.AutoFilter Field:=19, Criteria1:=val
            
            'Filter to just new pursuits
            orgsheet.ListObjects("Table_owssvr_1").Range.AutoFilter Field:=21, Criteria1:="Not Started", Operator:=xlOr, Criteria2:="="
            
            'Build email for filtered recipient
            Set OutMail = OutApp.CreateItem(0)
            On Error Resume Next
            With OutMail
                .To = val 'address mail to validated email
                .Subject = "Pursuits requiring your input"
                .htmlbody = "<p>Hi <br></p>" & _
                    "<p>The following oppertunities from the Sales Pipeline are listed in the Pursuit tracker under you name and are either incomplete or haven’t been updated in 14 or more days.</p>" & _
                    "<h3>New Pursuits</h3>" & _
                    "<p>Please begin the qualification process and provide a Qualification Status in the <a href=" & _
                    "http://ecm-link.cisco.com/ecm/view/objectId/0b0dcae183ece46a/app/ciscodocs" & _
                    ">Pursuit Tracker</a> by updating column L (Qualification Status).</ol>" & _
                    RangetoHTML(rng) & "<br>"
            End With ' Close top half of mail
            'Remove Filter to just new pursuits
            orgsheet.ListObjects("Table_owssvr_1").Range.AutoFilter Field:=21
            
            'Add Filter for oppertunities not update in 14 or more Days
            orgsheet.ListObjects("Table_owssvr_1").Range.AutoFilter Field:=26, Criteria1:="<" & (Now() - 15), Operator:=xlAnd
           ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           ' Seccond Table starts below this                                                              '
           ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            With OutMail 'open bottom part of mail
                    .htmlbody = .htmlbody & "<h3>Aging Pursuits</h3>" & _
                    "<p>These oppertunities haven't been updated in teh tracker in over 14 days. Please review" & _
                    "Qualification Status in column L and update if necessary. If no updates are necessary please simply update the date of review. Statuses are expected to be reviewed/updated weekly.</p>" & _
                    RangetoHTML(rng) & "<br>" & _
                    "<p><a href=" & _
                    "http://ecm-link.cisco.com/ecm/view/objectId/090dcae183ed3974/versionLabel/CURRENT" & _
                    ">More Information, instructions Qualification and Risk Assesment Templates</a> are available on Cisco Docs.<br><br>" & _
                    "Please refer to <a href=" & _
                    "http://iwe.cisco.com/web/view-post/post/-/posts?postId=223600176" & _
                    ">Checking out and checking in library docs</a> if you need more information on how to properly reserve and update thePursuit Tracker.</p>"
                    'You can also add files like this:
                    '.Attachments.Add ("C:\test.txt")
                
                .Display ' can also use .Send
            End With
            'Remove Filter for oppertunities not update in 14 or more Days
            orgsheet.ListObjects("Table_owssvr_1").Range.AutoFilter Field:=26
            On Error GoTo 0
            Set OutMail = Nothing
            
        End If
        orgsheet.ListObjects("Table_owssvr_1").Range.AutoFilter Field:=19
    Next val


cleanup:
    Set OutApp = Nothing
    orgsheet.Cells.EntireColumn.Hidden = False
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
End Sub
