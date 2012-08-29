Sub DeleteRowsSecondFastest()
''''''''''''''''''''''''''
'Written by www.ozgrid.com
''''''''''''''''''''''''''
Dim rTable As Range
Dim rCol As Range, rCell As Range
Dim lCol As Long
Dim xlCalc As XlCalculation
Dim vCriteria

On Error Resume Next


   'Determine the table range
     With Selection
         If .Cells.Count > 1 Then
             Set rTable = Selection
             MsgBox rTable, vbOKOnly
             
         Else
             Set rTable = .CurrentRegion
             On Error GoTo 0
         End If
    End With
   
    'Determine if table range is valid

    If rTable Is Nothing Or rTable.Cells.Count = 1 Or WorksheetFunction.CountA(rTable) < 2 Then
        MsgBox "Could not determine you table range.", vbCritical, "Ozgrid.com"
        Exit Sub
    End If

    'Get the criteria in the form of text or number.

    vCriteria1 = "Presales"
    vCriteria2 = "Active"
    vCriteria3 = "Delivery Close"
    vCriteria4 = "PID Status:"

    'vCriteria = Application.InputBox(Prompt:="Type in the criteria that matching rows should be deleted. " _
    '& "If the criteria is in a cell, point to the cell with your mouse pointer", _
    'Title:="CONDITIONAL ROW DELETION CRITERIA", Type:=1 + 2)

    'Go no further if they Cancel.

    'If vCriteria1 = "Active" Then Exit Sub

    'Get the relative column number where the criteria should be found

    'CT INPUT
    lCol = 7
    
    'lCol = Application.InputBox(Prompt:="Type in the relative number of the column where " _
    '& "the criteria can be found.", Title:="CONDITIONAL ROW DELETION COLUMN NUMBER", Type:=1)

    'Cancelled
    If lCol = 0 Then Exit Sub
        'Set rCol to the column where criteria should be found
        Set rCol = rTable.Columns(lCol)
        'Set rCell to the first data cell in rCol
        Set rCell = rCol.Cells(2, 1)

    'Store current Calculation then switch to manual.
    'Turn off events and screen updating
    With Application
        xlCalc = .Calculation
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .ScreenUpdating = False
    End With

   

  'Loop and delete as many times as vCriteria exists in rCol
    For Each Row In rTable.Rows
  
  
        'For lCol = 1 To WorksheetFunction.CountIf(rCol, vCriteria)
        If rTable.Cells(Row.Row, lCol) <> vCriteria1 And rTable.Cells(Row.Row, lCol) <> vCriteria2 And rTable.Cells(Row.Row, lCol) <> vCriteria3 And rTable.Cells(Row.Row, lCol) <> vCriteria4 Then
            MsgBox "delete this row" & rTable.Cells(Row.Row, lCol).Value
            
        Else
            'rTable.Cells.Offset(0, 0).EntireRow.Delete
            MsgBox "this row stays" & rTable.Cells(Row.Row, lCol).Value
        End If
         
    Next Row

    With Application
        .Calculation = xlCalc
        .EnableEvents = True
        .ScreenUpdating = True
    End With
   On Error GoTo 0

End Sub
