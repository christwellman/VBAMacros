Sub CombineSheetsForMetrics()
    Dim sh As Worksheet
    Dim DestSh As Worksheet

    Dim Last As Long
    Dim shLast As Long
    Dim CopyRng As Range
    Dim StartRow As Long

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    

    ' Delete unneccssary sheets.
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets("LOVs").Delete
    'ActiveWorkbook.Worksheets("Complete - Presales - Scoping").Delete
    'ActiveWorkbook.Worksheets("Cold Projects").Delete
    'ActiveWorkbook.Worksheets("Closed Projects").Delete
    ActiveWorkbook.Worksheets("Project Tracking - GPAGE").Delete
    ActiveWorkbook.Worksheets("Archive Desk Complete").Delete
    ActiveWorkbook.Worksheets("Archive Cold Projects").Delete
    ActiveWorkbook.Worksheets("Archive Closed Projects").Delete
    ActiveWorkbook.Worksheets("DMColWidths").Delete
    ActiveWorkbook.Worksheets("New WM Mapping").Delete
    
    On Error GoTo 0
    

    ' Add a new summary worksheet.
    Set DestSh = ActiveWorkbook.Worksheets.Add
    DestSh.Name = "MasterCombinedData"

    ' Fill in the start row.
    StartRow = 2
    

    ' Loop through all worksheets and copy the data to the
    ' summary worksheet.
    For Each sh In ActiveWorkbook.Worksheets
        If sh.Name <> DestSh.Name Then

            ' Find the last row with data on the summary
            ' and source worksheets.
            Last = Lastrow(DestSh)
            shLast = Lastrow(sh)

            ' If source worksheet is not empty and if the last
            ' row >= StartRow, copy the range.
            If shLast > 0 And shLast >= StartRow Then
                'Set the range that you want to copy
                Set CopyRng = sh.Range(sh.Rows(StartRow), sh.Rows(shLast))

               ' Test to see whether there are enough rows in the summary
               ' worksheet to copy all the data.
                If Last + CopyRng.Rows.Count > DestSh.Rows.Count Then
                   MsgBox "There are not enough rows in the " & _
                   "summary worksheet to place the data."
                   GoTo ExitTheSub
                End If

                ' This statement copies values and formats.
                CopyRng.Copy
                With DestSh.Cells(Last + 1, "A")
                    .PasteSpecial xlPasteValues
                    .PasteSpecial xlPasteFormats
                    Application.CutCopyMode = False
                End With

            End If

        End If
    Next

ExitTheSub:




'Get Rid of Extra Columns
'Sheets("MasterCombinedData").Select
'[I:J].Delete
'[T:V].Delete
'[U:AC].Delete
'[V:V].Delete
'[AB:AH].Delete
'[U:XFD].Delete

'Confirm Dates are in Date Format
'Last = Lastrow(DestSh)
'Firstrow = ActiveSheet.UsedRange.Cells(1).Row
'    Lrow = Last + Firstrow - 1
    
'    With DestSh
'        .DisplayPageBreaks = False
            
'            For Lrow = Last To Firstrow Step -1
        'Delete requests received before desk
'                If IsDate(.Cells(Lrow, "D").Value) And .Cells(Lrow, "D").Value > CDate("8/1/2011") = True Then


 '               Else
 '                   .Cells(Lrow, "D").EntireRow.Delete

 '              End If
 '           Next
 '   End With

'Add Column Headers Back
    Application.Goto DestSh.Cells(1)
    ActiveCell.EntireRow.Select
    Selection.Insert Shift:=xlDown
    Sheets("Project Pipeline").Select
    Range("A1:AZ1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("MasterCombinedData").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Sheets("MasterCombinedData").Cells.Select
    Selection.clearformats
    
'Delete Source Worksheets now that we're done with them
ActiveWorkbook.Worksheets("Project Tracking - BBODEN").Delete
ActiveWorkbook.Worksheets("Project Tracking - SMILESNI").Delete
ActiveWorkbook.Worksheets("Project Tracking - EVOGEL").Delete
ActiveWorkbook.Worksheets("Complete - Presales - Scoping").Delete
ActiveWorkbook.Worksheets("Cold Projects").Delete
ActiveWorkbook.Worksheets("Closed Projects").Delete
ActiveWorkbook.Worksheets("Project Pipeline").Delete

'
'    DestSh.Cells(1, 1).Value = "Flow Type"
'    DestSh.Cells(1, 2).Value = "Firm Start Date?"
'    DestSh.Cells(1, 3).Value = "Desk Owner"
'    DestSh.Cells(1, 4).Value = "Request Date"
'    DestSh.Cells(1, 5).Value = "Requestor"
'    DestSh.Cells(1, 6).Value = "Project ID (PID):"
'    DestSh.Cells(1, 7).Value = "PID Status:"
'    DestSh.Cells(1, 8).Value = "Customer Name"
'    DestSh.Cells(1, 9).Value = "Technology"
'    DestSh.Cells(1, 10).Value = "Service Type(s)"
'    DestSh.Cells(1, 11).Value = "Location"
'    DestSh.Cells(1, 12).Value = "Start Date:"
'    DestSh.Cells(1, 13).Value = "Kick-Off Date:"
'    DestSh.Cells(1, 14).Value = "End Date:"
'    DestSh.Cells(1, 15).Value = "Request Type:"
'    DestSh.Cells(1, 16).Value = "Project Type:"
'    DestSh.Cells(1, 17).Value = "Work Manager Assigned"
'    DestSh.Cells(1, 18).Value = "DCPM Assigned"
'    DestSh.Cells(1, 19).Value = "DM Assigned"
'    DestSh.Cells(1, 20).Value = "Services Revenue:"
'    DestSh.Cells(1, 21).Value = "DCPM Project Status:"
'    DestSh.Cells(1, 22).Value = "Initial Follow Up Request Sent"
'    DestSh.Cells(1, 23).Value = "DCPM Assigned Date"
'    DestSh.Cells(1, 24).Value = "Technical Resource(s) Assigned Date"
'    DestSh.Cells(1, 25).Value = "Walker Survey Sent Date?"
'    DestSh.Cells(1, 26).Value = "Delivery Close Date:"
'    DestSh.Cells(1, 27).Value = "Project Close Date:"
'    DestSh.Cells(1, 28).Value = "Date Ready for Staffing"
'    DestSh.Cells(1, 29).Value = "Last Update"


    ' AutoFit the column width in the summary sheet.
'    DestSh.Columns.AutoFit
'    DestSh.Rows.AutoFit

    With Application
        .Calculation = CalcMode
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With
End Sub