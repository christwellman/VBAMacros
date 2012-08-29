Sub CleanData()
    Dim sh As Worksheet
    Dim Prompt As String
    Dim DestSh As Worksheet
    Dim Firstrow As Integer
    Dim Last As Long
    Dim shLast As Long
    Dim checkrange As Range
    Dim checkarea As Range
    Dim checkrow As Range
    Dim StartRow As Long

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    

    ' Delete unneccssary sheets.
    Application.DisplayAlerts = False
    On Error Resume Next
    'ActiveWorkbook.Worksheets("LOVs").Delete
    'ActiveWorkbook.Worksheets("Complete - Presales - Scoping").Delete
    'ActiveWorkbook.Worksheets("Cold Projects").Delete
    'ActiveWorkbook.Worksheets("Closed Projects").Delete
    'ActiveWorkbook.Worksheets("Project Tracking - GPAGE").Delete
    'ActiveWorkbook.Worksheets("Archive Desk Complete").Delete
    'ActiveWorkbook.Worksheets("Archive Cold Projects").Delete
    'ActiveWorkbook.Worksheets("Archive Closed Projects").Delete
    'ActiveWorkbook.Worksheets("Project Pipeline").Delete
    'ActiveWorkbook.Worksheets("New WM Mapping").Delete
    
    On Error GoTo 0
    

    ' Choose Sheet to Clean
    Prompt = ("Which worksheet do you want to clean?")
    DynamicForm.PromptLabel.Caption = Prompt
    DynamicForm.DynamicComboBox.RowSource = "Sheets"

    DynamicForm.Show
    
    
    'destshname = InputBox(Prompt:="Which worksheet do you want to clean?", Title:="Clean which sheet?", Default:="")
    Set DestSh = ActiveWorkbook.Worksheets(CheckValue)
    DynamicForm.DynamicComboBox = ""
    
    ' Fill in the start row.--- Not Used See below check to get seccond row of data
    'If DestSh.Name = "Project Pipeline" Then
        StartRow = InputBox(Prompt:="Which row would you like to start checking at?", Title:="Start Row?", Default:=2)
    'Else
        'StartRow = 2
    'End If
    Last = Lastrow(DestSh)
    Firstrow = StartRow
    
    Set checkrange = Range(Cells(Firstrow, "A"), Cells(Last, "AP"))
    'MsgBox ("Range " & checkrange.Address)

    
    With DestSh
        .DisplayPageBreaks = False
 
                For Each checkrow In checkrange.Rows
                    
                    ' Check if request preceeds Desk
                    If IsDate(.Cells(checkrow.Row, "AK")) = True Then
                        If CDATE(.Cells(checkrow.Row, "AK")) < CDATE(.Cells(checkrow.Row, "D")) Then
							.cells(checkrow.row,"AK").value = .cells(checkrow.row,"D").value
						End If
                    End If
                Next
    End With


    ' AutoFit the column width in the summary sheet.
    'DestSh.Columns.AutoFit
    'DestSh.Rows.Height = 35
    
    With Application
        .Calculation = CalcMode
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    Unload DynamicForm
End Sub
