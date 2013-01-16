Attribute VB_Name = "Module1"
Sub UpdateMaster()
    Dim Sh As Worksheet
    Dim DestSh As Worksheet
    Dim FillRange As Range
    Dim last As Long
    Dim shLast As Long
    Dim CopyRng As Range
    Dim mastertable As Range
    Dim StartRow As Long
    Dim PMName As String
    

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        
    End With
     ActiveWindow.FreezePanes = False
    On Error Resume Next
    'Clear Data from old Master Sheet
    ActiveWorkbook.Worksheets("Master").Range("A2:BT600").Clear
    'Previous Version was deleting this sheet
    'ActiveWorkbook.Worksheets("Master").Delete
    ActiveWorkbook.Worksheets("ALL").Visible = xlSheetHidden
    
    On Error GoTo 0
    

    ' Add a new summary worksheet.
    Set DestSh = ActiveWorkbook.Worksheets("Master")

    ' Fill in the start row.
    StartRow = 3
    

    ' Loop through all worksheets and copy the data to the
    ' summary worksheet.
    For Each Sh In ActiveWorkbook.Worksheets
        If Sh.Name <> DestSh.Name And Sh.Visible <> xlSheetVeryHidden Then
        
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
                Selection.Insert Shift:=xlToRight
                DestSh.Cells(RowCount, 2).Value = PMName
                
                
            End If

        End If
    Next
ExitTheSub:

'Add Column Headers Back
    DestSh.Rows(1).Insert Shift:=xlDown
    
    'Insert Hyperlink for Update Macro
    With Worksheets("Master")
     .Hyperlinks.Add Anchor:=.Range("C1"), _
     Address:="", _
     ScreenTip:="Refresh worksheet Data", _
     TextToDisplay:="Update Sheet"
    End With
    
    'Insert Hyperlink for Add PM/EM
    With Worksheets("Master")
     .Hyperlinks.Add Anchor:=.Range("B1"), _
     Address:="", _
     ScreenTip:="Add a new Tab for a new EM/PM", _
     TextToDisplay:="Add new EM/PM"
    End With

    'Add headings
    'DestSh.Cells("A1:A200").Value = ""
    DestSh.Cells(2, 1).Value = "No."
    DestSh.Cells(2, 2).Value = "EM/PM"
    DestSh.Cells(2, 3).Value = "Customer"
    DestSh.Cells(2, 4).Value = "Project"
    DestSh.Cells(2, 5).Value = "Percent"
    DestSh.Cells(2, 6).Value = "Project ID (PID):"
    DestSh.Cells(2, 7).Value = "Start Date"
    DestSh.Cells(2, 8).Value = "End Date"
    DestSh.Cells(2, 9).Value = "Comments"
    DestSh.Cells(2, 10).Value = "=(TODAY()+(7-WEEKDAY(TODAY(),2)+1))-7"
    DestSh.Cells(2, 11).Value = "=J2+7"
    DestSh.Range("K2:BQ2").FillRight
    
    'Remove Hyperlinks
    With DestSh.Range("B2:C2")
        .Hyperlinks.Delete
        .font.Size = 9
    End With
    
    
    'Loop Through Rows to Remove Blanks and Format
    last = Lastrow(DestSh)
    Firstrow = ActiveSheet.UsedRange.Cells(1).Row
    Lrow = last + Firstrow - 1
    
    With DestSh
        .DisplayPageBreaks = False
            
            For Lrow = last To Firstrow Step -1
                If IsError(.Cells(Lrow, "D").Value) Then
                    'MsgBox (Lrow)
                
                'This Method Leaves the total row for each individual -
                'Row will be deleted if no information is provided in column B- E
                ElseIf .Cells(Lrow, "B").Value = "" And .Cells(Lrow, "C").Value = "" And .Cells(Lrow, "D").Value = "" _
                And .Cells(Lrow, "E").Value = "" Then
                    .Rows(Lrow).EntireRow.Delete
                End If
            Next
            Counter = 1
            last = Lastrow(DestSh)
            For Firstrow = 3 To last
                If .Cells(Firstrow, "C").Value <> "Total" Then
                DestSh.Cells(Firstrow, 1).Value = Counter
                Counter = Counter + 1
                'Clear Formatting on C - I
                With Range(Cells(Firstrow, "C"), Cells(Firstrow, "I"))
                    With .font
                        .Name = "Calibri"
                        .Size = 9
                    End With
                    With .Interior
                        .Pattern = xlNone
                    End With
                    With .Borders
                        .LineStyle = xlNone
                    End With
                End With
                              
                Else
                DestSh.Cells(Firstrow, 1).Value = ""
                End If
            Next
            
            
    End With
    
    'Add Formats
    With Range("J:BQ")
        .ColumnWidth = 4.4
        '.NumberFormat = "#%"
    End With
    With Range("J2:BQ2")
        .NumberFormat = "m/d;@"
    End With
    
    'define Table
    last = Lastrow(DestSh)
    Set mastertable = Range(Cells(2, 1), Cells(last, "BQ"))
    Debug.Print (mastertable.Address)
    'format table
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(mastertable.Address), , xlYes).Name = "Master Table"
    ActiveSheet.ListObjects("Master Table").TableStyle = "TableStyleLight2"
    
    'remove autofilter
    Range("K2:BQ2").AutoFilter
    
    'Add Format for HOT
    With Range("J3", Cells(last, "BQ"))
        With .font
            .Color = red
        End With
    End With
    
    Range("J3").Select
    ActiveWindow.FreezePanes = True
    
    With Range(Cells(3, 2), Cells(last, 2))
        With .font
            .Bold = True
            .Color = Black
        End With
    End With
    
    
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With
End Sub
Function Lastrow(Sh As Worksheet)
'This Fucntion takes a worksheet as an input and returns the last used row in the sheet

    On Error Resume Next
    Lastrow = Sh.Cells.Find(What:="*", _
                            after:=Sh.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    On Error GoTo 0
End Function


Sub ClearSheet()
    ActiveWorkbook.Worksheets("Master").Range("A:BT").Clear
End Sub


Sub AddNewTeamMemberTab()
'This Sub Creates a copy of a sheet template with the name of the new team member and adds all of the macros and formulas etc
Dim NewSheet As Worksheet
Dim TeamMemberName As String
Application.ScreenUpdating = False


    'TeamMemberName = InputBox(Prompt:="What is the name of the new team member?", Title:="Add new team member tab?", Default:="")
    TeamMemberName = InputBox("Please enter the name of the person you wish to add.", "Add new team member tab?", "")
    
    'MsgBox (TeamMemberName)
            
        If TeamMemberName = "" Then

            GoTo Cleanup
                       
        Else
            'Add a new copy of the team member template
            ActiveWorkbook.Sheets("TEAM-MEMBER-TEMPLATE").Visible = xlSheetVisible
            ActiveWorkbook.Sheets("TEAM-MEMBER-TEMPLATE").Copy _
               after:=ActiveWorkbook.Sheets(Sheets.Count)
               
            Set NewSheet = Sheets("TEAM-MEMBER-TEMPLATE (2)")
            NewSheet.Name = TeamMemberName
            NewSheet.Visible = xlSheetVisible
            
            NewSheet.Activate
            NewSheet.Cells(3, 2).Select
            
        End If

Cleanup:
    'MsgBox ("this is Cleanup")
    With Worksheets("Master")
     .Hyperlinks.Add Anchor:=.Range("B1"), _
     Address:="", _
     ScreenTip:="Add a new Tab for a new EM/PM", _
     TextToDisplay:="Add new EM/PM"
    End With
    ActiveWorkbook.Sheets("TEAM-MEMBER-TEMPLATE").Visible = xlSheetVeryHidden
    Application.ScreenUpdating = True
End Sub

Sub unhideTemplate()
 Sheet16.Visible = xlSheetVisible
End Sub

Sub SetConditionalFormatsComplex()
Dim Sh As Worksheet
Dim cs As ColorScale
Set Sh = ActiveSheet
Dim Rng As Range
Set Rng = Sh.Range("I3:BP16")

'clear existing conditional formats
Rng.FormatConditions.Delete


    'Set Formating of individual project ranges
    With Sh.Range("I3:BP14").FormatConditions _
        .Add(Type:=xlExpression, Formula1:="=I3<>""""")
        With .font
        .Bold = True
        .Italic = False
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        End With
        With .Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        End With
    End With
    
    '''''''''''''''''''''''''''''
    'Set Formatting for total row
    
    'Set the Range For the Color Scale
    Set cs = Sh.Range("I16:BP16").FormatConditions.AddColorScale(ColorScaleType:=3) '

    ' Format the first color as Green
    With cs.ColorScaleCriteria(1)
        .Type = xlConditionValuePercent
        .Value = 0
        With .FormatColor
            .Color = &H7BBE63
            .TintAndShade = 0
        End With
    End With
    
    ' Format the Acceptable Range as yellow
    With cs.ColorScaleCriteria(2)
        .Type = xlConditionValuePercent
        .Value = 60
        With .FormatColor
            .Color = &H84EBFF
            .TintAndShade = 0
        End With
    End With
    
    ' Format the Over Utilized as red
    With cs.ColorScaleCriteria(3)
        .Type = xlConditionValuePercent
        .Value = 100
        With .FormatColor
            .Color = &H6B69F8
            .TintAndShade = 0
        End With
    End With
        
End Sub

Sub SetConditionalFormatsSimple()
Dim Sh As Worksheet
Dim cs As ColorScale
Set Sh = ActiveSheet
Dim Rng As Range
Set Rng = Sh.Range("I3:BP30")

'clear existing conditional formats
Rng.FormatConditions.Delete


    'Set Formating of Total ranges
    'With Sh.Range("I29:BP29").FormatConditions _
    '    .Add(Type:=xlExpression, Formula1:="=I29<>""""")
    '    With .font
    '   .Bold = True
    '   .Italic = False
    '   .ThemeColor = xlThemeColorDark1
    '    .TintAndShade = 0
    '    End With
    '    With .Interior
     '   .PatternColorIndex = xlAutomatic
    '    .ThemeColor = xlThemeColorDark1
     '   .TintAndShade = -0.499984740745262
     '   End With
    'End With
    With Sh.Range("I29:BP29").FormatConditions _
        .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=.75")
        With .font
        .Bold = True
        .Italic = False
        .Color = -16776961
        .TintAndShade = 0
        End With
        'With .Interior
        '.PatternColorIndex = xlAutomatic
        '.ThemeColor = xlThemeColorDark1
        '.TintAndShade = -0.499984740745262
        'End With
    End With
    
    'Set Formating of individual project ranges
    With Sh.Range("I3:BP27").FormatConditions _
        .Add(Type:=xlExpression, Formula1:="=I3<>""""")
        With .font
        .Bold = False
        .Italic = False
        .Color = Black
        .TintAndShade = 0
        End With
    End With
        
End Sub

Sub AddProjectRow()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim Sh As Worksheet
Dim bformula As String
Dim cformula As String
Dim dformula As String
Dim bottom As Long

Set Sh = ActiveSheet

'Find the current last row not including total row
last = Lastrow(Sh) - 2

'Insert row below current last cell
Cells(last, 2).Offset(1).EntireRow.Insert

'reset last
last = Lastrow(Sh) - 2
bottom = Lastrow(Sh)

'Redo numbering
With Sh.Range(Cells(3, 1), Cells(last, 1))
    .Value = "=ROW(A1)"
    .Value = .Value
End With

'build Dateformula


'Add Formulas
DateFormula = "=IF(AND(I$2>=$F" & last & ",I$2<=$G" & last & "),$D" & last & "," & """""" & ")"
bformula = "=IF(COUNTA(B3:B" & last & ")<>0," & """Total""" & "," & """""" & ")"
'original Cformula = =IF(COUNTA(C3:C27)<>0,COUNTA(C3:C27)&" Project(s)","")
cformula = "=IF(COUNTA(C3:C" & last & ")<>0,COUNTA(C3:C" & last & ")" & "&"" Project(s)""" & "," & """""" & ")"
'Original Dformula = =IF(COUNTA(D3:D27)<>0,SUM(D3:D27),"")
dformula = "=IF(COUNTA(D3:D" & last & ")<>0,SUM(D3:D" & last & ")," & """""" & ")"
Sh.Cells(last, 9).Value = DateFormula
Sh.Range(Cells(last, 9), Cells(last, "BP")).FillRight
Sh.Cells(bottom, 2).Value = bformula
Sh.Cells(bottom, 3).Value = cformula
Sh.Cells(bottom, 4).Value = dformula
'MsgBox (dformula)

'rewrite hyperline
With Sh.Cells(bottom - 1, 2)
    .Hyperlinks.Add Anchor:=Sh.Cells(bottom - 1, 2), _
     Address:="", _
     ScreenTip:="Add another row with formulas to the entry table.", _
     TextToDisplay:="Add Row"
     .font.Size = 8
End With

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub


Sub DeleteProjectRow()

'Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim Sh As Worksheet
Dim SelectedRange As Range
Dim DeleteRow As Range
Set Sh = ActiveSheet

last = Lastrow(Sh)
'Prompt for Row
'' need to test if multiple rows will wotk
Set SelectedRange = Application.InputBox( _
    Prompt:="Select row(s) you would like to remove.", _
    Title:="Remove Project Rows", _
    Type:=8)


'select entire row
If SelectedRange Is Nothing Then
    'no Range Selected
Else
    Set DeleteRow = SelectedRange.EntireRow
    DeleteRow.EntireRow.Select
End If

'confirm delete and read back row Number
confirm = MsgBox( _
    Prompt:="Are you sure you want to delete the selected row(s)?", _
    Buttons:=vbYesNo, _
    Title:="Confirm delete rows")

Application.ScreenUpdating = False
   
If confirm = vbNo Then
ElseIf confirm = vbYes Then
    DeleteRow.EntireRow.Delete
End If

'call AddProjectRow
Call AddProjectRow

'rewrite hyperlink
With Sh.Cells(1, 2)
    .Hyperlinks.Add Anchor:=Sh.Cells(1, 2), _
     Address:="", _
     ScreenTip:="Remove row(s) from the data entry table.", _
     TextToDisplay:="Remove Row"
     .font.Size = 8
End With

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub

