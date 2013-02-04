Sub UpdateMaster()
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
	'Delet the old Master Sheet
    ActiveWorkbook.Worksheets("Master").Delete
	ActiveWorkbook.Worksheets("ALL").Visible = xlsheethidden
    
    On Error GoTo 0
    

    ' Add a new summary worksheet.
    Set DestSh = ActiveWorkbook.Worksheets.Add
    DestSh.Name = "Master"

    ' Fill in the start row.
    StartRow = 2
    

    ' Loop through all worksheets and copy the data to the
    ' summary worksheet.
    For Each sh In ActiveWorkbook.Worksheets
        If Sh.Name <> DestSh.Name And Sh.Visible <> xlSheetHidden Then

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
                With DestSh.Cells(Last + 1, "C")
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


    DestSh.Cells(1, 1).Value = "No."
    DestSh.Cells(1, 2).Value = "EM/PM"
    DestSh.Cells(1, 3).Value = "Customer"
    DestSh.Cells(1, 4).Value = "Project"
    DestSh.Cells(1, 5).Value = "Percent"
    DestSh.Cells(1, 6).Value = "Project ID (PID):"
    DestSh.Cells(1, 7).Value = "Start Date"
    DestSh.Cells(1, 8).Value = "End Date"
    DestSh.Cells(1, 9).Value = "Comments"


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