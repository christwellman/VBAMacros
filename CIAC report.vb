
Sub CIACProjectReport()
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
    ActiveWorkbook.Worksheets("Project Tracking - GPAGE").Delete
    ActiveWorkbook.Worksheets("Archive Desk Complete").Delete
    ActiveWorkbook.Worksheets("Archive Cold Projects").Delete
    ActiveWorkbook.Worksheets("Archive Closed Projects").Delete
    ActiveWorkbook.Worksheets("New WM Mapping").Delete
    
    On Error GoTo 0
    

    ' Add a new summary worksheet.
    Set DestSh = ActiveWorkbook.Worksheets.Add
    DestSh.Name = "CIAC Projects"

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

'Add Column Headers Back
Application.Goto DestSh.Cells(1)
ActiveCell.EntireRow.Select
Selection.Insert Shift:=xlDown
Sheets("Project Pipeline").Select
Range("A1:AZ1").Select
Application.CutCopyMode = False
Selection.Copy
Sheets("CIAC Projects").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Sheets("CIAC Projects").Cells.Select
Selection.clearformats
    
'Delete Source Worksheets now that we're done with them
ActiveWorkbook.Worksheets("Project Tracking - BBODEN").Delete
ActiveWorkbook.Worksheets("Project Tracking - SMILESNI").Delete
ActiveWorkbook.Worksheets("Project Tracking - EVOGEL").Delete
ActiveWorkbook.Worksheets("Project Tracking - MILASKIN").Delete
ActiveWorkbook.Worksheets("Project Tracking - AMALAN").Delete
ActiveWorkbook.Worksheets("Complete - Presales - Scoping").Delete
ActiveWorkbook.Worksheets("Cold Projects").Delete
ActiveWorkbook.Worksheets("Closed Projects").Delete



Last = Lastrow(DestSh)
Firstrow = ActiveSheet.UsedRange.Cells(2).Row
    Lrow = Last + Firstrow - 1
    
    With DestSh
        .DisplayPageBreaks = False
            
            For Lrow = Last To Firstrow Step -1
                If IsError(.Cells(Lrow, "K").Value) Then


                ElseIf .Cells(Lrow, "K").Value <> "CIAC" Then
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
Sheets("CIAC Projects").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Sheets("CIAC Projects").Cells.Select
Selection.clearformats
ActiveWorkbook.Worksheets("Project Pipeline").Delete

'Get Rid of Extra Columns
Sheets("CIAC Projects").Select
[A:E].Delete
[E:E].Delete
[I:I].Delete
[O:Q].Delete
[S:U].Delete
[U:XFD].Delete
    
'Clear Formats
Sheets("CIAC Projects").Cells.Select
Selection.clearformats
DestSh.Rows.AutoFit
DestSh.Columns.AutoFit
Columns("A:A").Select
Selection.NumberFormat = "General"
Columns("D:D").Select
Selection.ColumnWidth = 30
Columns("F:F").Select
Selection.ColumnWidth = 40

    


DestSh.Rows.AutoFit

    With Application
        .Calculation = CalcMode
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    'Save New File for Report                    '
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

    SheetPrefix = "CIAC Projects List - "
    Sheetname = SheetPrefix & vDateStr & ".xlsx"
    
    'File name
    strSheet = Sheetname
    'File Path
    strPath = "C:\Users\ctwellma\Documents\AS\Reports\CIAC Project Report\"
    
    'File Name
    strSheet = strPath & strSheet
    
    'Save As
    ActiveWorkbook.saveas Filename:=strSheet, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

End Sub