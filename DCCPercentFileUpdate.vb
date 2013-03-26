Sub LastColumnInOneRow()
'Original File
'Find the last used column in a Row: row 1 in this example
    Dim LastCol As Integer
    With ActiveSheet
        LastCol = .Cells(6, .Columns.Count).End(xlToLeft).Column
    End With
    MsgBox LastCol
End Sub
Sub LastRow_Example()
    Dim LastRow As Long
    Dim rng As Range

    ' Use all cells on the sheet
    Set rng = Sheets("Sheet1").Cells

    'Use a range on the sheet
    'Set rng = Sheets("Sheet1").Range("A1:D30")

    ' Find the last row
    LastRow = Last(1, rng)

    ' After the last row with data change the value of the cell in Column A
    'rng.Parent.Cells(LastRow + 1, 1).Value = "Hi there"
    MsgBox (rng.Parent.Cells(LastRow + 1, 1).Address)

End Sub


Sub LastColumn_Example()
    Dim LastCol As Long
    Dim rng As Range

    ' Use all cells on the sheet
    Set rng = Sheets("Sheet3").Cells

    'Or use a range on the sheet
    'Set rng = Sheets("Project Time Entry Summary").Range("N7:AS7")

    ' Find the last column
    LastCol = Last(2, rng)

    ' After the last column with data change the value of the cell in row 1
    'rng.Parent.Cells(1, LastCol + 1).Value = "Hi there"
    MsgBox (rng.Parent.Cells(1, LastCol).Address)
    MsgBox (rng.Parent.Cells(1, LastCol).Column)
    MsgBox (Cells(4, rng.Parent.Cells(1, LastCol).Column) & " " & Cells(6, rng.Parent.Cells(1, LastCol).Column))
End Sub


Sub LastCell_Example()
    Dim LastCell As String
    Dim rng As Range

    ' Use all cells on the sheet
    Set rng = Sheets("Sheet1").Cells

    'Or use a range on the sheet
    'Set rng = Sheets("Sheet1").Range("A1:D30")

    ' Find the last cell
    LastCell = Last(3, rng)

    ' Select from A1 till the last cell in Rng
    With rng.Parent
        .Select
        .Range("A1", LastCell).Select
    End With
End Sub

'This is the function we use in the macro's above
Function Last(choice As Long, rng As Range)
'Ron de Bruin, 5 May 2008
' 1 = last row
' 2 = last column
' 3 = last cell
    Dim lrw As Long
    Dim lcol As Long

    Select Case choice

    Case 1:
        On Error Resume Next
        Last = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Row
        On Error GoTo 0

    Case 2:
        On Error Resume Next
        Last = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        On Error GoTo 0

    Case 3:
        On Error Resume Next
        lrw = rng.Find(What:="*", _
                       After:=rng.Cells(1), _
                       lookat:=xlPart, _
                       LookIn:=xlFormulas, _
                       SearchOrder:=xlByRows, _
                       SearchDirection:=xlPrevious, _
                       MatchCase:=False).Row
        On Error GoTo 0

        On Error Resume Next
        lcol = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        On Error GoTo 0

        On Error Resume Next
        Last = rng.Parent.Cells(lrw, lcol).Address(False, False)
        If Err.Number > 0 Then
            Last = rng.Cells(1).Address(False, False)
            Err.Clear
        End If
        On Error GoTo 0

    End Select
End Function

Sub GetTimeEntryData()
With Application
    .ScreenUpdating = False
    .Cursor = xlWait
    .DisplayStatusBar = True
    .StatusBar = "Updating Sheet..."
End With
On Error GoTo Cleanup


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Declare Variables                                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Val As String
Dim Srcbk As Workbook
Dim SrcSh As Worksheet
Dim DestSh As Worksheet
Dim rng As Range
Dim rngX As Range
Dim PID As Integer
Dim week As Integer
Dim LastCol As Long
Dim colnum As Integer 'this is a counter used to get the row to read from the src sheet
Dim SrcWkb As Workbook
Dim StrMsg As String
Dim CopyRng As Range


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Variables                                                      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Set Destination Sheet
Set DestSh = ActiveSheet
'Source File name
strSheet = "Fly BY Wire - Data Center" & ".xls"
'Source File Path
strpath = "C:\Users\ctwellma\Documents\AS\Reports\FlyByWire\" & strSheet
'Open source workbook
Application.Workbooks.Open (strpath)
Set SrcWkb = ActiveWorkbook
'MsgBox (SrcWkb.name)
sFileName = ActiveWorkbook.name

'Set Source Sheet
Set SrcSh = SrcWkb.Worksheets("Project Time Entry Summary")


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Code Execution                                                     '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''Kick Off Timer/Progress Bar'''''''''''''''''''''''''


'find last cell in Source sheet
    Set rng = SrcSh.Cells
    EndOfSource = Last(1, rng)
    Debug.Print (EndOfSource & " " & SrcSh.name)
    'Get PIDS
    Set CopyRng = SrcSh.Range(Cells(7, 2), Cells(EndOfSource, 2))
    CopyRng.Copy Destination:=DestSh.Range("A2")
    'Get PID Statuses
    Set CopyRng = SrcSh.Range(Cells(7, 9), Cells(EndOfSource, 9))
    CopyRng.Copy Destination:=DestSh.Range("B2")
    'Get Total SC Hours
    Set CopyRng = SrcSh.Range(Cells(7, 46), Cells(EndOfSource, 46))
    CopyRng.Copy Destination:=DestSh.Range("C2")
    'Get Customer name
    Set CopyRng = SrcSh.Range(Cells(7, 11), Cells(EndOfSource, 11))
    CopyRng.Copy Destination:=DestSh.Range("F2")
    'Get Project name
    Set CopyRng = SrcSh.Range(Cells(7, 10), Cells(EndOfSource, 10))
    CopyRng.Copy Destination:=DestSh.Range("G2")
    
    With DestSh.UsedRange
        .NumberFormat = "General"
        .Value = .Value
        .Borders.LineStyle = xlNone
        .Font.Bold = False
    End With
    
    
    'Loop Through Each PID in the lookup sheet
    '   starts at row 2 for now
        For PID = 2 To EndOfSource + 1
        
        'Define Value you want to look up
        Val = DestSh.Cells(PID, 1).Value
        
                'Define where to Search For PID
                Set rngX = SrcSh.Range("B:B").Find(Val, lookat:=xlPart)

                'Set Search Range for last Time Entry
                Set rng = SrcSh.Range(Cells(rngX.Row, 14), Cells(rngX.Row, 45))

                ' Find the last column in the above defined range
                LastCol = Last(2, rng)
                'Write the last time entry date to the corresponding column
                DestSh.Cells(PID, 8).Value = Cells(4, rng.Parent.Cells(1, LastCol).Column) & " " & Cells(6, rng.Parent.Cells(1, LastCol).Column)
        Next
MsgBox ("Exiting Last time")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Change Source Sheet to get Total Hour information                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set SrcSh = SrcWkb.Worksheets("Active, Delivery Close, On Hold")
MsgBox ("New source Set")
'Loop Through Each PID in the lookup sheet
    '   starts at row 2 for now
        For PID = 2 To 3 ' EndOfSource + 1
        
        'Define Value you want to look up
        Val = DestSh.Cells(PID, 1).Value
        
                'Define where to Search For PID
                Set rngX = SrcSh.Range("K:K").Find(Val, lookat:=xlPart)
                MsgBox (rngX.Offset(0, 50).Value)

                'Write the total hours to the corresponding column
                DestSh.Cells(PID, 4).Value = rngX.Offset(0, 50).Value
        Next
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Cleanup                                                            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Cleanup:

Workbooks(SrcWkb.name).Close SaveChanges:=False
'ActiveWorkbook.Worksheets("FY13Q2").Activate

With Application
    .ScreenUpdating = True
    .Cursor = xlDefault
    .DisplayStatusBar = False
End With
    
End Sub

Sub Contiuned()

'find last cell in destination sheet - this should return the total number of weeks data is avaiable for
    endofpage = LastCol(SrcSh) - 4
    proceed = MsgBox(endofpage & " weeks of data available in " & SrcSh.name & ". Continue with update?", vbYesNo, _
        "Proceed with update?")
If proceed = vbYes Then
    'Loop Through Each name in the lookup sheet starts at row 3 for now
        For name = 3 To Last
        
        'Define Value you want to look up
        Val = DestSh.Cells(name, 1).Value
        
        'Define how many columns of data to pull from Source Sheet
        
        If Not Val = "" Then
        'MsgBox (Val)
                'Search Range
                Set rngX = SrcSh.Range("B:B").Find(Val, lookat:=xlPart)
                'Set rngX = SrcSh.Range("B:B").Find(Val, lookat:=xlWhole)
                'MsgBox (rngX.Address)
                For destcol = 1 To ((endofpage * 7) - 6) Step 7
                Application.StatusBar = "Updating Sheet..." & name / Last & "%"
                    'MsgBox (destCol)
                    If Not rngX Is Nothing Then
                        'Compliance
                        DestSh.Cells(name, destcol + 1).Value = rngX.Offset(0, (destcol / 7) + 2).Value
                        'Billable
                        DestSh.Cells(name, destcol + 2).Value = rngX.Offset(1, (destcol / 7) + 2).Value
                        'Non-Billable
                        DestSh.Cells(name, destcol + 3).Value = rngX.Offset(2, (destcol / 7) + 2).Value
                        'CFU
                        DestSh.Cells(name, destcol + 4).Value = rngX.Offset(3, (destcol / 7) + 2).Value
                        'Internal
                        DestSh.Cells(name, destcol + 5).Value = rngX.Offset(4, (destcol / 7) + 2).Value
                        'Admin
                        DestSh.Cells(name, destcol + 6).Value = rngX.Offset(5, (destcol / 7) + 2).Value
                        'Overall
                        DestSh.Cells(name, destcol + 7).Value = rngX.Offset(6, (destcol / 7) + 2).Value
                    Else
                        'MsgBox ("Cannot retrieve data for " & Val & ". Name was not found in " & SrcSh.name)
                        StrMsg = StrMsg & "• " & Val & vbNewLine
                        GoTo NextName
                    End If
                Next
        End If
NextName:
    Next
Else
    GoTo Cleanup
End If

If Not StrMsg = "" Then
    MsgBox ("The follwing names couldn't be found in " & sFileName & ". No data has been entered for them:" & vbNewLine _
          & StrMsg)
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Cleanup                                                            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Cleanup:

Workbooks(SrcWkb.name).Close SaveChanges:=False
'ActiveWorkbook.Worksheets("FY13Q2").Activate

With Application
    .ScreenUpdating = True
    .Cursor = xlDefault
    .DisplayStatusBar = False
End With
     
End Sub
