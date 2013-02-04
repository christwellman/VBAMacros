Function ActivateWB(wbname As String)
  'This function takes a workbook name as string as input and opens
  'Open wbname.

  Workbooks(wbname).Activate

End Function
Function Lastrow(Sh As Worksheet)
'This Fucntion takes a worksheet as an input and returns the last used row in the sheet

    On Error Resume Next
    Lastrow = Sh.Cells.Find(What:="*", _
                            after:=Sh.Range("A1"), _
                            lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    On Error GoTo 0
End Function

Function LastCol(Sh As Worksheet)
'This Fucntion takes a worksheet as an input and returns the last used row in the sheet
'SearchOrder:=xlByRows, _


    On Error Resume Next
    LastCol = Sh.Cells.Find(What:="*", _
                            after:=Sh.Range("A1"), _
                            lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
    On Error GoTo 0
End Function

Function findVal(Val As String)
'This Fucntion takes a worksheet as an input and returns the last used row in the sheet

Dim CheckCol As Range
Set CheckCol = ActiveSheet.Range("A:A")
MsgBox (CheckCol)

    On Error Resume Next
    findVal = CheckCol.Find(What:=Val, _
                            after:=Sh.Range("A1"), _
                            lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    On Error GoTo 0
End Function


Sub GetData()
With Application
    .ScreenUpdating = False
End With


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Declare Variables                                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Val As String
Dim Srcbk As Workbook
Dim SrcSh As Worksheet
Dim DestSh As Worksheet
Dim rngX As Range
Dim name As Integer
Dim week As Integer
Dim colnum As Integer 'this is a counter used to get the row to read from the src sheet
Dim SrcWkb As Workbook

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Variables                                                      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Set Destination Sheet
Set DestSh = ActiveWorkbook.Worksheets("Q2FY13")

'Source File name
StrSheet = "lookUpBook.xlsm"
'Source File Path
strPath = "C:\Users\ctwellma\SkyDrive\" & StrSheet
'MsgBox (strPath)
'set Source Workbook

'Set SrcWkb = appExcel.ActiveWorkbook
'Set Source Sheet
Set SrcSh = ActiveWorkbook.Worksheets("Resource Detailed Report")
'Set SrcSh = lookUpBook.Worksheets("Resource Detailed Report")
'Search Range
'Set rngX = SrcSh.Range("B:B").Find(Val, lookat:=xlPart)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Code Execution                                                     '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'find last cell in destination sheet
    last = Lastrow(DestSh)
'find last cell in destination sheet - this should return the total number of weeks data is avaiable for
    endofpage = LastCol(SrcSh) - 4
    
'Loop Through Each name in the lookup sheet starts at row 3 for now
    For name = 3 To last
    
    'Define Value you want to look up
    Val = DestSh.Cells(name, 1).Value
    If Not Val = "" Then
    'MsgBox (Val)
            'Search Range
            Set rngX = SrcSh.Range("B:B").Find(Val, lookat:=xlPart)
            'MsgBox (rngX.Address)
            For destCol = 1 To 85 Step 7
                'MsgBox (destCol)
                If Not rngX Is Nothing Then
                    'Compliance
                    DestSh.Cells(name, destCol + 1).Value = rngX.Offset(0, (destCol / 7) + 2).Value
                    'Billable
                    DestSh.Cells(name, destCol + 2).Value = rngX.Offset(1, (destCol / 7) + 2).Value
                    'Non-Billable
                    DestSh.Cells(name, destCol + 3).Value = rngX.Offset(2, (destCol / 7) + 2).Value
                    'CFU
                    DestSh.Cells(name, destCol + 4).Value = rngX.Offset(3, (destCol / 7) + 2).Value
                    'Internal
                    DestSh.Cells(name, destCol + 5).Value = rngX.Offset(4, (destCol / 7) + 2).Value
                    'Admin
                    DestSh.Cells(name, destCol + 6).Value = rngX.Offset(5, (destCol / 7) + 2).Value
                    'Overall
                    DestSh.Cells(name, destCol + 7).Value = rngX.Offset(6, (destCol / 7) + 2).Value
                Else
                    MsgBox ("Cannot retrieve data for " & Val & ". Name was not found in " & SrcSh.name)
                    GoTo NextName
                End If
            Next
    End If
    
NextName:
    Next
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Cleanup                                                            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Cleanup:

'Workbooks("C:\Users\ctwellma\SkyDrive\datepicker.xls").Worksheets("Sheet1").Activate

With Application
    .ScreenUpdating = True
End With
     
End Sub
