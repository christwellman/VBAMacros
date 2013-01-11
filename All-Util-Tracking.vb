Sub UpdateMaster()
    Dim Sh As Worksheet
    Dim DestSh As Worksheet
    Dim FillRange As Range
    Dim last As Long
    Dim shLast As Long
    Dim CopyRng As Range
    Dim StartRow As Long
    Dim PMName As String
    

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    

    ' Delete unneccssary sheets.
    Application.DisplayAlerts = False
    On Error Resume Next
    'Delet the old Master Sheet
    ActiveWorkbook.Worksheets("Master").Delete
    ActiveWorkbook.Worksheets("ALL").Visible = xlSheetHidden
    
    On Error GoTo 0
    

    ' Add a new summary worksheet.
    Set DestSh = ActiveWorkbook.Worksheets.Add
    DestSh.Name = "Master"

    ' Fill in the start row.
    StartRow = 2
    

    ' Loop through all worksheets and copy the data to the
    ' summary worksheet.
    For Each Sh In ActiveWorkbook.Worksheets
        If Sh.Name <> DestSh.Name And Sh.Visible <> xlSheetHidden Then
        
        PMName = Sh.Name

            ' Find the last row with data on the summary
            ' and source worksheets.
            last = Lastrow(DestSh)
            RowCount = last + 1
            shLast = Lastrow(Sh)

            ' If source worksheet is not empty and if the last
            ' row >= StartRow, copy the range.
            If shLast > 0 And shLast >= StartRow Then
                'Set the range that you want to copy
                Set CopyRng = Sh.Range(Sh.Rows(StartRow), Sh.Rows(shLast))

               ' Test to see whether there are enough rows in the summary
               ' worksheet to copy all the data.
                If last + CopyRng.Rows.Count > DestSh.Rows.Count Then
                   MsgBox "There are not enough rows in the " & _
                   "summary worksheet to place the data."
                   GoTo ExitTheSub
                End If

                ' This statement copies values and formats.
                CopyRng.Copy
                With DestSh.Cells(last + 1, "A")
                    .PasteSpecial xlPasteValues
                    .PasteSpecial xlPasteFormats
                    Application.CutCopyMode = False
                End With
                
                'Add EM/PM Name From Sheet
                Range(DestSh.Cells(RowCount, 2), DestSh.Cells(shLast + RowCount, 2)).Select
                'MsgBox ((DestSh.Cells(RowCount, 2) & DestSh.Cells(shLast, 2)))
                'Range(DestSh.Rows(RowCount), DestSh.Rows(shLast)).Select
                Selection.Insert shift:=xlToRight
                DestSh.Cells(RowCount, 2).Value = PMName
                
                'FillDown Range
                'With DestSh
                '    '.Range(.Cells(RowCount, 2), .Cells(RowCount, shLast)).FillDown
                '    MsgBox (RowCount)
                '    MsgBox (shLast)
                'End With
                
            End If

        End If
    Next
ExitTheSub:

'Add Column Headers Back
    Application.Goto DestSh.Cells(1)
    ActiveCell.EntireRow.Select
    Selection.Insert shift:=xlDown

    'DestSh.Cells("A1:A200").Value = ""
    DestSh.Cells(1, 1).Value = "No."
    DestSh.Cells(1, 2).Value = "EM/PM"
    DestSh.Cells(1, 3).Value = "Customer"
    DestSh.Cells(1, 4).Value = "Project"
    DestSh.Cells(1, 5).Value = "Percent"
    DestSh.Cells(1, 6).Value = "Project ID (PID):"
    'DestSh.Cells(1, 7).Value = "Start Date"
    'DestSh.Cells(1, 8).Value = "End Date"
    'DestSh.Cells(1, 9).Value = "Comments"
    DestSh.Cells(1, 10).Value = "=(TODAY()+(7-WEEKDAY(TODAY(),2)+1))-7"
    DestSh.Cells(1, 11).Value = "=J1+7"
    DestSh.Range("K1:BK1").FillRight
    
    'Remove Blanks
    last = Lastrow(DestSh)
    firstrow = ActiveSheet.UsedRange.Cells(1).Row
    Lrow = last + firstrow - 1
    
    With DestSh
        .DisplayPageBreaks = False
            
            For Lrow = last To firstrow Step -1
                If IsError(.Cells(Lrow, "D").Value) Then
                    MsgBox (Lrow)
                
                'This Method Leaves the total row for each individual
                ElseIf .Cells(Lrow, "B").Value = "" And .Cells(Lrow, "C").Value = "" And .Cells(Lrow, "D").Value = "" Then
                'This Method Removes the total row for each individual
                'ElseIf .Cells(Lrow, "B").Value = "" And .Cells(Lrow, "C").Value = "" Then
                    'MsgBox (lrow)
                    .Rows(Lrow).EntireRow.Delete
                    '.Rows(Lrow).EntireRow.Interior.ColorIndex = 43
                    

                End If
            Next
            Counter = 1
            last = Lastrow(DestSh)
            For firstrow = 2 To last
                If .Cells(firstrow, "C").Value <> "" And .Cells(firstrow, "D").Value <> "" Then
                DestSh.Cells(firstrow, 1).Value = Counter
                Counter = Counter + 1
                Else
                DestSh.Cells(firstrow, 1).Value = ""
                End If
            Next
            
            
    End With
    

    ' AutoFit the column width in the summary sheet.
   DestSh.Columns.AutoFit
'    DestSh.Rows.AutoFit

    With Application
        .Calculation = CalcMode
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With
End Sub


Sub Utilized()
Dim CurrentSheet As Worksheet
Dim rCell As Range
Dim last As Integer ' Last Row on sheet
Dim CheckRng As Range ' Range to Update
Dim Comments As String



Set CurrentSheet = ActiveSheet
last = Lastrow(CurrentSheet)
'Debug: Last Range on Sheet
'MsgBox (last)

Set CheckRng = Range(Cells(2, 6), Cells(last, "BR"))
'Deubg: Check Range
'MsgBox (CheckRng.Address)

'Loop Through Each Cell in the Range
For Each rCell In CheckRng.Cells
    'MsgBox (rCell.Address)
'Next Cell
        
        cellcolor = rCell.Interior.ColorIndex
        'Debug: Cell Color Index
        'MsgBox (cellcolor)
        RowNum = rCell.Row
        'Debug: Row Number
        
        'MsgBox (RowNum)
        'Debug: Utilziation %
        'MsgBox (ActiveSheet.Cells(RowNum, 4).Value)
        
        If cellcolor = 44 Then
        'MsgBox (rCell.Address)
        'Check for Onhold COlor?
        'ElseIf cellcolor = 42 Then
           rCell.Value = "-"
        ElseIf cellcolor = -4142 Then
            'Do Nothign for Cells with no fill
        Else
            rCell.Value = CurrentSheet.Cells(RowNum, 4).Value
        End If
Next rCell
End Sub

Sub UpgradeTeamUtilToChrisFormat()
Dim Sh As Worksheet
Dim CurrentSheet As Worksheet
Dim rCell As Range
Dim last As Integer ' Last Row on sheet
Dim CheckRng As Range ' Range to Update
Dim Comments As String

With Application
    .ScreenUpdating = False
    .EnableEvents = False
End With


' Delete unneccssary sheets.
Application.DisplayAlerts = False
On Error Resume Next

ActiveWorkbook.Worksheets("Master").Delete
ActiveWorkbook.Worksheets("ALL").Visible = xlSheetHidden
    
    On Error GoTo 0
    

 

' Loop through all worksheets and copy the data to the
' summary worksheet.
    For Each Sh In ActiveWorkbook.Worksheets
        If Sh.Name <> DestSh.Name And Sh.Visible <> xlSheetHidden Then
        
        MsgBox (Sh.Name)

        Set CurrentSheet = Sh
        last = Lastrow(CurrentSheet)
        'Debug: Last Range on Sheet
        'MsgBox (last)
        
        Set CheckRng = Range(Cells(2, 6), Cells(last, "BR"))
        
        Range("F:G").Insert shift:=xlToRight
        Range("F1").Value = "Start Date"
        Range("G1").Value = "End Date"
        
        'Deubg: Check Range
        'MsgBox (CheckRng.Address)
        
        'Loop Through Each Cell in the Range
        For Each rCell In CheckRng.Cells
            'MsgBox (rCell.Address)
        'Next Cell
                
                cellcolor = rCell.Interior.ColorIndex
                ColNum = rCell.Column
                RowNum = rCell.Row
                ColumnDate = Cells(1, ColNum).Value
                'Debug: Column Date
                'MsgBox (ColumnDate)
                'Debug: Utilziation %
                'MsgBox (ActiveSheet.Cells(RowNum, 4).Value)
                
                If cellcolor = 44 Then
                'MsgBox (rCell.Address)
                'Check for Onhold COlor?
                'ElseIf cellcolor = 42 Then
                   'rCell.Value = "-"
                ElseIf cellcolor = -4142 Then
                    'Do Nothign for Cells with no fill
                Else
                    'What to do for valid dates
                    'rCell.Value = CurrentSheet.Cells(RowNum, 4).Value ' Fills in util % in date column
                    CurrentSheet.Cells(RowNum, 7).Value = ColumnDate
                End If
        Next rCell
        
        For r = 2 To last
            Comments = ""
            Cells(r, 3).ClearComments
            For c = 65 To 8 Step -1
                    cellcolor = Cells(r, c).Interior.ColorIndex
                    ColumnDate = Cells(1, c).Value
                    'Debug: Column Date
                    'MsgBox (ColumnDate)
                    'Debug: Utilziation %
                    'MsgBox (ActiveSheet.Cells(RowNum, 4).Value)
                    'Get Comments
                    If Cells(r, c).Value <> "" Then
                        Comments = Comments & ColumnDate & " " & Cells(r, c).Value
                    End If
                                
                    If cellcolor = 44 Then
                    'MsgBox (rCell.Address)
                    'Check for Onhold COlor?
                    'ElseIf cellcolor = 42 Then
                       'rCell.Value = "-"
                    ElseIf cellcolor = -4142 Then
                        'Do Nothign for Cells with no fill
                    Else
                        'What to do for valid dates
                        'rCell.Value = CurrentSheet.Cells(RowNum, 4).Value ' Fills in util % in date column
                        CurrentSheet.Cells(r, 6).Value = ColumnDate
                    End If
            Next c
            If Comments <> "" Then
                Cells(r, 3).AddComment (Comments)
            End If
        Next r
        End If
    Next Sh

End Sub

Function Lastrow(Sh As Worksheet)
'This Fucntion takes a worksheet as an input and returns the last used row in the sheet

    On Error Resume Next
    Lastrow = Sh.Cells.Find(What:="*", _
                            After:=Sh.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    On Error GoTo 0
End Function

