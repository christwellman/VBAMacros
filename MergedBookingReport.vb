Sub UpdateMergedBookingReport()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This macro copies detailed data from a Booking Report and places the information into a      '
' consolidated table saving only the latest record for each PID                                '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim FillRange As Range
Dim ShName As Worksheet
Dim DestSh As Worksheet

'Turn off Alerts and Screen updating
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

'Confirm user is ready to Run Report
    Answer = MsgBox("Do you have the latest version of the Booking  health report (from SharePoint) open?", vbQuestion + vbYesNo, "PID Health Open?")

    If Answer = vbNo Then
        'Code for No button Press
        Exit Sub
    Else
        'Code for Yes button Press

''''''Select PID health Data ''''''''''''''''
    Application.Goto Reference:=Worksheets("DATA").Range("A1")
    
    Range("A2").Resize(Cells.Find(What:="*", SearchOrder:=xlRows, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Row - 3, _
      Cells.Find(What:="*", SearchOrder:=xlByColumns, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Column).Select

''''''Copy PID Health data as values to new sheet ''''''''''''
    Selection.Copy
    Set ShName = ActiveWorkbook.Worksheets.Add
    ShName.Name = "Data Extract"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("CD1").Select
    Application.CutCopyMode = False
    
''''''Fill Cells with Report Week etc
    ActiveCell.Formula = _
        "=MID(CELL(""filename"",A1),FIND(""["",CELL(""filename"",A1))+1,10)"
    Range("CD1").Select
    'Selection.AutoFill Destination:=Range("BS1:BS264")
    'AutoFill the range down
    Last = Lastrow(ShName)
    Selection.AutoFill Destination:=Range(Cells(1, "CD"), Cells(Last, "CD"))

''''''Copy data to move to merged file
    Range("A1").Select
    Range("A1").Resize(Cells.Find(What:="*", SearchOrder:=xlRows, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Row, _
      Cells.Find(What:="*", SearchOrder:=xlByColumns, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Column).Select
    'Range("A1:BS264").Select
    Selection.Copy
    'Open the Combined PID Report XLSX
    Workbooks.Open Filename:= _
        "C:\Users\ctwellma\Documents\AS\Reports\Booking Report\CombinedReports\CombinedBookings.xlsx"
    'Copy PID Health Detail
    'Sheets("Merged Reports").Select
    'Windows("MergedPIDReportFY12.xlsx").Activate
    Set ShName = ActiveWorkbook.Worksheets("Merged Reports")
    Last = Lastrow(ShName) + 1
       
    Range(Cells(Last, "A"), Cells(Last, "A")).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'Variables:


'Define Sheet we're using
Set DestSh = ActiveWorkbook.Worksheets("Merged Reports")
Last = Lastrow(DestSh)

'Select Full Active Range to be sorted
    Range("A1").Resize(Cells.Find(What:="*", SearchOrder:=xlRows, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Row, _
      Cells.Find(What:="*", SearchOrder:=xlByColumns, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Column).Select
      
    'Sort Selection FIrst on PID (G) and then on report Quarter and Week (BS)
      'Selection.Sort _
      '      Key1:=Worksheets("Merged Reports").Range("G1"), Order1:=xlAscending, Header:=xlGuess, _
      '      key2:=Worksheets("Merged Reports").Range("BS1"), Order1:=xlAscending, Header:=xlGuess


'Loop through to eliminate everything but latest info for each PID

'Last = Lastrow(DestSh)
'Firstrow = ActiveSheet.UsedRange.Cells(2).Row
'    Lrow = Last + Firstrow - 1

'   With DestSh
'            For Lrow = Last To Firstrow Step -1
'                If IsError(.Cells(Lrow, "G").Value) Then
'
'
'                ElseIf .Cells(Lrow, "G").Value = .Cells(Lrow + 1, "G").Value Then
'                    If .Cells(Lrow, "BS").Value < .Cells(Lrow + 1, "BS").Value Then
'                        .Rows(Lrow).EntireRow.Delete
'                    End If
'
'                End If
'            Next
'    End With
'    End If
'Turn off Alerts and Screen updating
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub