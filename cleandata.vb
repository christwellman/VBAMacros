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
    DynamicForm.DynamicComboBox.AddItem "Project Pipeline"
    DynamicForm.DynamicComboBox.AddItem "Project Tracking - BBODEN"
    DynamicForm.DynamicComboBox.AddItem "Project Tracking - SMILESNI"
    DynamicForm.DynamicComboBox.AddItem "Project Tracking - EVOGEL"
    DynamicForm.DynamicComboBox.AddItem "Project Tracking - MILASKIN"
    DynamicForm.DynamicComboBox.AddItem "Complete - Presales - Scoping"
    DynamicForm.DynamicComboBox.AddItem "Cold Projects"
    DynamicForm.DynamicComboBox.AddItem "Closed Projects"
    
    DynamicForm.Show
    
    
    'destshname = InputBox(Prompt:="Which worksheet do you want to clean?", Title:="Clean which sheet?", Default:="")
    Set DestSh = ActiveWorkbook.Worksheets(CheckSheet)
    
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
                        
                        ''''''''''''''''''''''''''''''''''''''
                        'Check PID string length for valid PID
                        ''''''''''''''''''''''''''''''''''''''
                        If Len(.Cells(checkrow.Row, "F").Value) = 6 Then
                        
                        Else
                            .Cells(checkrow.Row, "F").Value = InputBox(Prompt:="Please provide a PID for: " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "F"), Title:="PID?", Default:=.Cells(checkrow.Row, "F").Value)
                        End If
                        
                        ''''''''''''''''''''''''''''''''''''''
                        'Check PID Status
                        ''''''''''''''''''''''''''''''''''''''
                        If .Cells(checkrow.Row, "G").Value = "Pipeline" Or .Cells(checkrow.Row, "G").Value = "Presales" Or .Cells(checkrow.Row, "G").Value = "Active" Or .Cells(checkrow.Row, "G").Value = "Cancelled" Or .Cells(checkrow.Row, "G").Value = "Closed" Or .Cells(checkrow.Row, "G").Value = "Delivery Close" Or .Cells(checkrow.Row, "G").Value = "On Hold" Or .Cells(checkrow.Row, "G").Value = "Not Available" Then
                        
                        Else
                            .Cells(checkrow.Row, "G").Value = InputBox(Prompt:="Please provide a PID Status for: " & .Cells(checkrow.Row, "F"), Title:="PID Status?", Default:=.Cells(checkrow.Row, "G").Value)
                        End If
                        
                        ''''''''''''''''''''''''''''''''''''''
                        'Check Technology Status
                        ''''''''''''''''''''''''''''''''''''''
                        If .Cells(checkrow.Row, "K").Value = "CIAC" Or .Cells(checkrow.Row, "K").Value = "UCS" Or .Cells(checkrow.Row, "K").Value = "DCN" Or .Cells(checkrow.Row, "K").Value = "SAN" Or .Cells(checkrow.Row, "K").Value = "VDI" Then
                        
                        Else
                            .Cells(checkrow.Row, "K").Value = InputBox(Prompt:="Please provide a Valid Technology for: " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "F"), Title:="DCV Technology?", Default:=.Cells(checkrow.Row, "L").Value)
                        End If
                        
                        ''''''''''''''''''''''''''''''''''''''
                        'Check Service Type
                        ''''''''''''''''''''''''''''''''''''''
                        If IsEmpty(.Cells(checkrow.Row, "L").Value) = False Then
                        
                        Else
                            .Cells(checkrow.Row, "L").Value = InputBox(Prompt:="Please provide a Service Type for: " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "F"), Title:="Service Type?", Default:=.Cells(checkrow.Row, "L").Value)
                        End If
                        
                        ''''''''''''''''''''''''''''''''''''''
                        'Check Request Type
                        ''''''''''''''''''''''''''''''''''''''
                        If .Cells(checkrow.Row, "Q").Value = "Scoping" Or .Cells(checkrow.Row, "Q").Value = "Presales" Or .Cells(checkrow.Row, "Q").Value = "Delivery" Or .Cells(checkrow.Row, "Q").Value = "SME" Then
                        
                        Else
                            .Cells(checkrow.Row, "Q").Value = InputBox(Prompt:="Please provide a request Type for: " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "F"), Title:="Request Type?", Default:=.Cells(checkrow.Row, "Q").Value)
                        End If
                        
                        ''''''''''''''''''''''''''''''''''''''''
                        'Check Project Type
                        ''''''''''''''''''''''''''''''''''''''''
                        If .Cells(checkrow.Row, "R").Value = "Subscription" Or .Cells(checkrow.Row, "R").Value = "Transaction" Or .Cells(checkrow.Row, "R").Value = "AS Fixed" Or .Cells(checkrow.Row, "R").Value = "CAP" Or .Cells(checkrow.Row, "R").Value = "Unknown" Then
                        
                        Else
                            .Cells(checkrow.Row, "R").Value = InputBox(Prompt:="Please provide a Valid Project Type: " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "F"), Title:="Project Type?", Default:=.Cells(checkrow.Row, "R").Value)
                        End If
                        
                        '''''''''''''''''''''''''''''
                        'Check for Work Manager
                        '''''''''''''''''''''''''''''
                        If IsEmpty(.Cells(checkrow.Row, "S").Value) = False Then
                        
                        Else
                            .Cells(checkrow.Row, "S").Value = InputBox(Prompt:="Please provide a Work Manager for: " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "F"), Title:="WM?", Default:=.Cells(checkrow.Row, "S").Value)
                        End If
                        
                        '''''''''''''''''''''''''''''
                        'Check for Project Manager/DCPM
                        '''''''''''''''''''''''''''''''
                        If IsEmpty(.Cells(checkrow.Row, "T").Value) = False Then
                        
                        Else
                            .Cells(checkrow.Row, "T").Value = InputBox(Prompt:="Please provide a Project Manager/DCPM for: " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "F"), Title:="DCPM?", Default:=.Cells(checkrow.Row, "T").Value)
                        End If
                        
                        
                        ''''''''''''''''''''''''''''''''
                        'Services Revenue should be updated via PID health Report and is ommitted here intentionally for now
                        ''''''''''''''''''''''''''''''''
                        
                        '''''''''''''''''''''''''''''''''''''''''''''''''''
                        'Dynamically Update DCPM Status
                        '''''''''''''''''''''''''''''''''''''''''''''''''''
                        If (.Cells(checkrow.Row, "G").Value = "Active" And IsEmpty(.Cells(checkrow.Row, "T").Value) = True) Or (.Cells(checkrow.Row, "G").Value = "Pipeline") Then
                            .Cells(checkrow.Row, "AI").Value = "Pending Assignment"
                        ElseIf .Cells(checkrow.Row, "G").Value = "Presales" And IsEmpty(.Cells(checkrow.Row, "S").Value) = False Then
                            .Cells(checkrow.Row, "AI").Value = "In Progress"
                        ElseIf .Cells(checkrow.Row, "G").Value = "On Hold" Then
                            .Cells(checkrow.Row, "AI").Value = "On Hold"
                        ElseIf .Cells(checkrow.Row, "G").Value = "Delivery Close" Then
                            .Cells(checkrow.Row, "AI").Value = "Delivery Close"
                        ElseIf .Cells(checkrow.Row, "G").Value = "Cancelled" Then
                            .Cells(checkrow.Row, "AI").Value = "Complete"
                        ElseIf .Cells(checkrow.Row, "G").Value = "Closed" Then
                            .Cells(checkrow.Row, "AI").Value = "Closed"
                        Else
                            .Cells(checkrow.Row, "AI").Value = InputBox(Prompt:="Please provide a valid DCPM Status for: " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "F"), Title:="DCPM Status?", Default:=.Cells(checkrow.Row, "I").Value)
                        End If
                        
                        
                        ''''''''''''''''''''''''''''''''''''
                        ' Check Follow Up Date
                        '''''''''''''''''''''''''''''''''''''''
                        If .Cells(checkrow.Row, "AK").Value = "Fixed" Then
                        
                        ElseIf IsDate(.Cells(checkrow.Row, "AK")) = False Then
                            .Cells(checkrow.Row, "AK").Value = InputBox(Prompt:="Please provide a valid follow-up date for customer: " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "F"), Title:="Initial Follow up Date?", Default:=.Cells(checkrow.Row, "AK").Value)
                        End If
                        
                        ''''''''''''''''''''''''''''''''''''
                        ' Check DCPM Assigned Date
                        '''''''''''''''''''''''''''''''''''''''
                        If IsEmpty(.Cells(checkrow.Row, "AL").Value) = True And IsEmpty(.Cells(checkrow.Row, "T").Value) = False Then
                            .Cells(checkrow.Row, "AL").Value = InputBox(Prompt:="Please provide a valid DCPM Assignment date for customer: " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "F"), Title:="PM Assigned Date?", Default:=.Cells(checkrow.Row, "AL").Value)
                        End If
                        
                        ''''''''''''''''''''''''''''''''''''
                        ' Check Closure Date
                        '''''''''''''''''''''''''''''''''''''''
                        If IsEmpty(.Cells(checkrow.Row, "AP").Value) = True And .Cells(checkrow.Row, "G").Value = "Closed" Then
                            .Cells(checkrow.Row, "AP").Value = InputBox(Prompt:="Please provide a valid project close date for customer: " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "H") & " " & .Cells(checkrow.Row, "F"), Title:="Project Close Date?", Default:=.Cells(checkrow.Row, "AP").Value)
                        End If
                            
                    End If
                    .Cells(checkrow.Row, "AY").Value = Now
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