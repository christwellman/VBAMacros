Sub CIACProjectReport()
'This Process:
' 1) Creates a copy of an existing file
' 2) Saves the copy with today's date
' 3) Deletes the unused Sheets
' 4) Combines all the other sheets into 1 master
' 5) Deletes row in the master which doesn't pertain to "CIAC"
' 6) Pulls in data from existing master PID health Report file
' 7) Joins the information from the 2 workbooks
' 8) Creates a consolidated summary worksheet with the KPI information
' Known issue 1 - Formula is not getting written correctly to L2 on "Raw Data" Sheet
'   Formula should be: "=INDEX(PIDHealth,MATCH($A2,PIDHealthData!$A:$A,0)-1,MATCH('Raw Data'!L$1,PIDHealthData!$1:$1,0))"
' Known issue 2 - CombinedRawData and PIDHEALTH Data tables are not getting dynamically sized correctly
    
    Dim Sh As Worksheet
    Dim DestSh As Worksheet

    Dim Last As Long
    Dim shLast As Long
    Dim CopyRng As Range
    Dim bottomRow As Long
    Dim StartRow As Long

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    
''''''''''''''''''''''''''''''''''''''''''''''
'Save Report with a new file name            '
''''''''''''''''''''''''''''''''''''''''''''''
    
    'Create Timestamp
     vNow = Now()
     vMthStr = CStr(Month(vNow))
     vDayStr = CStr(Day(vNow))
    'Add leading zeroes to month, day, hour, minutes
     If Len(vMthStr) = 1 Then
        vMthStr = "0" & vMthStr
     End If
     If Len(vDayStr) = 1 Then
        vDayStr = "0" & vDayStr
     End If

    'Get date string in yyyymmddhhnn format.
     vDateStr = Year(vNow) & vMthStr & vDayStr

    SheetPrefix = "PMO Tracked CIAC Projects - "
    SheetName = SheetPrefix & vDateStr & ".xlsx"
    
    'File name
    strSheet = SheetName
    'File Path
    strPath = "C:\Users\ctwellma\Documents\AS\Reports\CIAC Project Report\"
    
    'File Name
    strSheet = strPath & strSheet
    
    'Save As
    ActiveWorkbook.saveas Filename:=strSheet, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

    ' Delete unneccssary sheets.
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets("LOVs").Delete
    ActiveWorkbook.Worksheets("Project Tracking - GPAGE").Delete
    ActiveWorkbook.Worksheets("Archive Desk Complete").Delete
    ActiveWorkbook.Worksheets("Archive Cold Projects").Delete
    ActiveWorkbook.Worksheets("Archive Closed Projects").Delete
    ActiveWorkbook.Worksheets("New WM Mapping").Delete
    ActiveWorkbook.Worksheets("DMColWidths").Delete
    ActiveWorkbook.Worksheets("Project Tracking - AMALAN").Delete
    ActiveWorkbook.Worksheets("Complete - Presales - Scoping").Delete
    ActiveWorkbook.Worksheets("Cold Projects").Delete
    ActiveWorkbook.Worksheets("Closed Projects").Delete
    ActiveWorkbook.Worksheets("GPPR").Delete
    ActiveWorkbook.Worksheets("VlookAllSheets").Delete
    ActiveWorkbook.Worksheets("Cloud Workshops").Delete
    On Error GoTo 0
    

    ' Add a Combine all sheets to a new one
    Set DestSh = ActiveWorkbook.Worksheets.Add
    DestSh.Name = "PMO Tracked CIAC Projects"

    ' Fill in the start row.
    StartRow = 2
    

    ' Loop through all worksheets and copy the data to the
    ' summary worksheet.
    For Each Sh In ActiveWorkbook.Worksheets
        If Sh.Name <> DestSh.Name Then

            ' Find the last row with data on the summary
            ' and source worksheets.
            Last = Lastrow(DestSh)
            shLast = Lastrow(Sh)

            ' If source worksheet is not empty and if the last
            ' row >= StartRow, copy the range.
            If shLast > 0 And shLast >= StartRow Then
                'Set the range that you want to copy
                Set CopyRng = Sh.Range(Sh.Rows(StartRow), Sh.Rows(shLast))

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

'Add Column Headers Back
Application.Goto DestSh.Cells(1)
ActiveCell.EntireRow.Select
Selection.Insert Shift:=xlDown
Sheets("Project Pipeline").Select
Range("A1:AZ1").Select
Application.CutCopyMode = False
Selection.Copy
Sheets("PMO Tracked CIAC Projects").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Sheets("PMO Tracked CIAC Projects").Cells.Select
Selection.clearformats
    
'Delete Source Worksheets now that we're done with them
ActiveWorkbook.Worksheets("Project Tracking - SMILESNI").Delete
ActiveWorkbook.Worksheets("Project Tracking - EVOGEL").Delete
ActiveWorkbook.Worksheets("Project Tracking - MILASKIN").Delete
ActiveWorkbook.Worksheets("Project Tracking - TINGRAM").Delete
ActiveWorkbook.Worksheets("Project Tracking - JSMESSAE").Delete
ActiveWorkbook.Worksheets("Project Tracking - SHAHANG").Delete



'Loop through to eliminate everything but CIAC Projects
Last = Lastrow(DestSh)
Firstrow = ActiveSheet.UsedRange.Cells(2).Row
    Lrow = Last + Firstrow - 1
    
    With DestSh
        .DisplayPageBreaks = False
            
            For Lrow = Last To Firstrow Step -1
                'If IsError(.Cells(Lrow, "K").Value) Then


                If .Cells(Lrow, "K").Value <> "CIAC" Then
                    .Rows(Lrow).EntireRow.Delete


                End If
            Next
    End With

'Cleanup
'Add Column Headers Back
Application.Goto DestSh.Cells(1)
ActiveCell.EntireRow.Select
Selection.Insert Shift:=xlDown
Sheets("Project Pipeline").Select
Range("A1:AZ1").Select
Application.CutCopyMode = False
Selection.Copy
Sheets("PMO Tracked CIAC Projects").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Sheets("PMO Tracked CIAC Projects").Cells.Select
Selection.clearformats
ActiveWorkbook.Worksheets("Project Pipeline").Delete

'Get Rid of Extra Columns
Sheets("PMO Tracked CIAC Projects").Select
[A:E].Delete
[E:E].Delete
[I:I].Delete
[O:Q].Delete
[S:U].Delete
[U:XFD].Delete
    
'Clean Formats / Fit Columns
Sheets("PMO Tracked CIAC Projects").Cells.Select
Selection.clearformats
DestSh.Columns.AutoFit
Columns("A:A").Select
Selection.NumberFormat = "General"
Columns("D:D").Select
Selection.ColumnWidth = 30
Columns("F:F").Select
Selection.ColumnWidth = 40
DestSh.Rows.AutoFit
    

''''''''''''''''''''''''''''''''''''''''''''''''''
'Add New Sheet and import PID Health report data '
''''''''''''''''''''''''''''''''''''''''''''''''''
    Set NewSh = ActiveWorkbook.Worksheets.Add
    NewSh.Name = "PIDHealthData"

'Open PID Health Report
    'Workbooks.Open Filename:= _
    '    "C:\Users\ctwellma\Documents\AS\Reports\PID Health Report\MergedPIDReportFY12.xlsx"
    'Workbooks.Open Filename:= _
    '    "C:\Users\ctwellma\Documents\AS\Reports\PID Health Report\CombinedReports\MergedPIDReportFY12.xlsx"
    Workbooks.Open Filename:= _
        "C:\Users\ctwellma\Documents\AS\Reports\PID Health Report\PreReport\lastRun.xls"

'Copy PID Health Detail
    Sheets("Detail").Select

'*********** Dynamically select all data from the Merged PID health Report
    Range("A1").Resize(Cells.Find(What:="*", SearchOrder:=xlRows, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Row, _
      Cells.Find(What:="*", SearchOrder:=xlByColumns, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Column).Select
    Selection.Copy
    Application.Goto NewSh.Cells(1)

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.NumberFormat = "General"

'Move PID to the head of the class (First Column)
    Rows("1:1").Delete
    Columns("G:G").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    'Name "PIDHealth" as a Range
    '****************************************
    'Dynamic Selection
    bottomRow = Lastrow(Worksheets("PIDHealthData"))
    endOfRange = "BS" & bottomRow
    Range("A1", endOfRange).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1", endOfRange), , xlYes).Name = "PIDHealth"
    
    'Format As Table
    Range("PIDHealth[#All]").Select
    ActiveSheet.ListObjects("PIDHealth").TableStyle = "TableStyleLight1"
    
    Windows("lastRun.xls").Activate
    ActiveWindow.Close

''''''''''''''''''''''''''''''''''''''''''''''''''
'Add New Sheet to join PID health and PMO Data   '
''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets("Sheet4").Select
    Sheets("Sheet4").Name = "Raw Data"
    Sheets("PMO Tracked CIAC Projects").Select
    'Select All PMO Data
    Range("A1").Resize(Cells.Find(What:="*", SearchOrder:=xlRows, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Row, _
      Cells.Find(What:="*", SearchOrder:=xlByColumns, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Column).Select
    Selection.Copy
    
    'Paste into "Raw Data"
    Sheets("Raw Data").Select
    Range("A1").Select
    ActiveSheet.Paste
    'Range("L1").Select
    Sheets("PIDHealthData").Select
    Range("PIDHealth[[#Headers],[Project Number]]").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Sheets("Raw Data").Select
    
    'Uncommented below 6 lines
    Range("N1").Select
    ActiveSheet.Paste
    Range("N2").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX(PIDHealth,MATCH(RC1,PIDHealthData!C1,0)-1,MATCH('Raw Data'!R1C,PIDHealthData!R1,0))"
    Range("N2").Select
    
    'Range("L2").FormulaR1C1 = "=INDEX(PIDHealth,MATCH(RC1,PIDHealthData!C1,0)-1,MATCH('Raw Data'!R1C,PIDHealthData!R1,0))"
    'bottomRow = Lastrow(Worksheets("Raw Data"))
    'With Worksheets("Raw Data").Range("L2")
    '    .AutoFill Destination:=Range("L2:L" & bottomRow&)
    'End With
    
    'Selection.AutoFill Destination:=Range("L2:L67")
    'Range("L2:L67").Select
    'Selection.AutoFill Destination:=Range("L2:CC67"), Type:=xlFillDefault
    'Range("L2:CC67").Select
    '''''''''''''''''''''''''''''''''''''' New Try Here
    'Range("M2").FormulaR1C1 =
    '    "=INDEX(PIDHealth,MATCH(RC1,PIDHealthData!C1,0)-1,MATCH('Raw Data'!R1C,PIDHealthData!R1,0))"
    'Range("M2").Select
    Selection.AutoFill Destination:=Range("N2:N76"), Type:=xlFillDefault
    Range("N2:N76").Select
    Selection.AutoFill Destination:=Range("N2:CC76"), Type:=xlFillDefault
    Range("N2:CC76").Select
    Selection.NumberFormat = "General"
    
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    
    'Name Range
    bottomRow = Lastrow(Worksheets("Raw Data"))
    endOfRange = "CC" & bottomRow
    Range("A1", endOfRange).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1", endOfRange), , xlYes).Name = "CombinedRawData"
    Range("CombinedRawData[#All]").Select
    ActiveSheet.ListObjects("CombinedRawData").TableStyle = "TableStyleLight1"
    
    'Filter out extra data from PID Health
    '*** NEED TO DEVELOP THIS PROCESS *****
    ''''''''''''''''''''''''''''''''''''''
    
'Add Dynamic Summary
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets("Sheet5").Select
    Sheets("Sheet5").Name = "Dynamic Summary"
    Range("A1:Z1").Select
    
    'Add Default Headers and format widths
    Range("A1").Value = "Project ID (PID)"
    Columns("A:A").Select
    Selection.ColumnWidth = 10

    Range("B1").Value = "PID Status"
    Columns("B:B").Select
    Selection.ColumnWidth = 13.5

    Range("C1").Value = "Customer Name"
    Columns("C:C").Select
    Selection.ColumnWidth = 40

    Range("D1").Value = "DCPM Assigned"
    Columns("D:D").Select
    Selection.ColumnWidth = 18

    Range("E1").Value = "AS Approved Cost Budget"
    Columns("E:E").Select
    Selection.ColumnWidth = 17

    Range("F1").Value = "Actual Costs"
    Columns("F:F").Select
    Selection.ColumnWidth = 17

    Range("G1").Value = "% Over / Under  Budget"
    Columns("G:G").Select
    Selection.ColumnWidth = 13
    Selection.NumberFormat = "%"
    Range("H1").Value = "As Approved Budgeted Hours"
    Columns("H:H").Select
    Selection.ColumnWidth = 15

    Range("I1").Value = "Total hours"
    Columns("I:I").Select
    Selection.ColumnWidth = 10
    Selection.NumberFormat = "0.00"
    Range("J1").Value = "Delivery Manager"
    Columns("J:J").Select
    Selection.ColumnWidth = 18
    'Selection.NumberFormat = "@"
    Range("K1").Value = "Actual PreSales Cost % of Actual Total Cost"
    Columns("K:K").Select
    Selection.ColumnWidth = 10
    Selection.NumberFormat = "%"
    
    'Format Headers
    Range("A1:K1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984740745262
        .PatternTintAndShade = 0
    End With
    Rows("1:1").RowHeight = 30
    Rows("1:1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .Bold = True
    End With
          
    
    
    'Input Formula for Dynamic table
    Range("A2").Select
    ActiveCell.Formula = _
        "=INDEX('Raw Data'!$A$1:$CC$67,ROWS($A$2:A2)+1,MATCH(A$1,'Raw Data'!$1:$1,0))"
    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A67")
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((ISERROR(INDEX('Raw Data'!R1C1:R67C81,MATCH(RC1,'Raw Data'!C1,0),MATCH('Dynamic Summary'!R1C,'Raw Data'!R1,0)))),""-"",INDEX('Raw Data'!R1C1:R67C81,MATCH(RC1,'Raw Data'!C1,0),MATCH('Dynamic Summary'!R1C,'Raw Data'!R1,0)))"
    'ActiveCell.FormulaR1C1 = _
    '    "=INDEX('Raw Data'!R1C1:R67C81,MATCH(RC1,'Raw Data'!C1,0),MATCH('Dynamic Summary'!R1C,'Raw Data'!R1,0))"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B67"), Type:=xlFillDefault
    Range("B2:B67").Select
    Selection.AutoFill Destination:=Range("B2:K67"), Type:=xlFillDefault
    Range("A2:K67").Select
    Selection.NumberFormat = "General"
    Columns("E:F").Select
    'Selection.Style = "Currency"
    Columns("G:G").Select
    Selection.Style = "Percent"
    Columns("H:I").Select
    Selection.Style = "Comma"
    Columns("K:K").Select
    Selection.Style = "Percent"

    Sheets("Raw Data").Select
    Range("CombinedRawData[[#Headers],[Project ID (PID)]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CombinedRawData[[#Headers],[Project ID (PID)]]").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("PMO Tracked CIAC Projects").Select
    Application.CutCopyMode = False
    'ActiveWindow.SelectedSheets.Delete
    'Sheets("PIDHealthData").Select
    'ActiveWindow.SelectedSheets.Delete
    
    Sheets("Dynamic Summary").Select
    
    Columns("E:E").Select
    Selection.NumberFormat = _
        "_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* ""-""??_);_(@_)"
    ActiveWindow.SmallScroll ToRight:=3
    Columns("F:F").Select
    Selection.NumberFormat = _
        "_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* ""-""??_);_(@_)"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR((RC[-1]-RC[-2])/RC[-2],""-"")"
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G66")
    Range("G2:G66").Select
    'Selection.FormatConditions(1).StopIfTrue = False
    'Conditional Formatting $$
    Columns("F:F").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=F2>E2"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    'Conditional Formatting Budget %
    Columns("G:G").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="> 0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    'Conditional Fomatting Hours
    Columns("I:I").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=I2>H2"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Columns("K:K").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
'Turn on Updates
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    ActiveWorkbook.Save

End Sub