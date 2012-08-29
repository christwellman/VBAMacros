Sub CheckDates()
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
		.DisplayAlerts = False
    End With
    
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
                    If .Cells(checkrow.Row, "D").Value = "Before Desk" Or .Cells(checkrow.Row, "D").Value < CDate("8/8/2011") Then
                        
                    Else
                    'Proceed with data checks if request came in after desk was established
                        ''''''''''''''''''''
                        'Check request Date
                        '''''''''''''''''''''''''''''
                        If IsDate(.Cells(checkrow.Row, "D")) = False Then
                            .Cells(checkrow.Row, "D").Value = InputBox(Prompt:="Please provide a valid request date for customer: " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "F"), Title:="Request Date?", Default:=.Cells(checkrow.Row, "D").Value)
                        End If
                        
                        
                        ''''''''''''''''''''''''''''''''''''
                        ' Check Follow Up Date
                        '''''''''''''''''''''''''''''''''''''''
                        If IsEpmty(.Cells(checkrow.Row, "AK")) = True Or .Cells(checkrow.Row, "AK") < .Cells(checkrow.Row, "D")) Then
                            .Cells(checkrow.Row, "AK").Value = InputBox(Prompt:="Please provide a valid follow-up date for customer: " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "F"), Title:="Initial Follow up Date?", Default:=.Cells(checkrow.Row, "AK").Value)
                        End If
                        
                        ''''''''''''''''''''''''''''''''''''
                        ' Check DCPM Assigned Date
                        '''''''''''''''''''''''''''''''''''''''
                        If IsEmpty(.Cells(checkrow.Row, "AL").Value) = True And IsEmpty(.Cells(checkrow.Row, "T").Value) = False Then
                            .Cells(checkrow.Row, "AL").Value = InputBox(Prompt:="Please provide a valid DCPM Assignment date for customer: " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "F"), Title:="PM Assigned Date?", Default:=.Cells(checkrow.Row, "AL").Value)
						ElseIF IsEpmty(.Cells(checkrow.Row, "AL")) = False And .Cells(checkrow.Row, "AL") < .Cells(checkrow.Row, "D")) Then
                            .Cells(checkrow.Row, "AL").Value = InputBox(Prompt:="Please provide a valid DCPM Assignment date for customer: " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "F"), Title:="PM Assigned Date?", Default:=.Cells(checkrow.Row, "AL").Value)	
                        End If
                        
                        ''''''''''''''''''''''''''''''''''''
                        ' Check Completion Date
                        '''''''''''''''''''''''''''''''''''''''
                        If IsEmpty(.Cells(checkrow.Row, "AP").Value) = True And .Cells(checkrow.Row, "G").Value = "Closed" Then
                            .Cells(checkrow.Row, "AP").Value = InputBox(Prompt:="Please provide a valid project close date for customer: " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "F"), Title:="Project Close Date?", Default:=.Cells(checkrow.Row, "AP").Value)
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