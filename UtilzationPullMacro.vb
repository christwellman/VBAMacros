Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    If Target.name = "Update Sheet" Then
        Call GetData
    ElseIf Target.name = "Add" Then
    Else
        Exit Sub
    End If

End Sub

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
On Error GoTo Cleanup


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
Dim StrMsg As String

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Variables                                                      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Set Destination Sheet
Set DestSh = ActiveSheet
'Source File name
strSheet = ActiveSheet.name & "Latest.xlsx"
'Source File Path
strpath = "C:\Users\ctwellma\Documents\AS\Reports\Global Utilization Report\" & strSheet
'Open source workbook
Application.Workbooks.Open (strpath)
Set SrcWkb = ActiveWorkbook
sfileName = ActiveWorkbook.name

'set Source Workbook
'Workbooks.Open Filename:=strPath
'Set SrcWkb = ActiveWorkbook
'src
'Set SrcWkb = appExcel.ActiveWorkbook
'Set Source Sheet
'MsgBox (SrcWkb.name)
Set SrcSh = SrcWkb.Worksheets("Resource Detailed Report")
'Set SrcSh = lookUpBook.Worksheets("Resource Detailed Report")
'Search Range
'Set rngX = SrcSh.Range("B:B").Find(Val, lookat:=xlPart)
StrMsg = ""

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Code Execution                                                     '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'find last cell in destination sheet
    last = Lastrow(DestSh)
'find last cell in destination sheet - this should return the total number of weeks data is avaiable for
    endofpage = LastCol(SrcSh) - 4
    proceed = MsgBox(endofpage & " weeks of data available in " & SrcSh.name & ". Continue with update?", vbYesNo, _
        "Proceed with update?")
If proceed = vbYes Then
    'Loop Through Each name in the lookup sheet starts at row 3 for now
        For name = 3 To last
        
        'Define Value you want to look up
        Val = DestSh.Cells(name, 1).Value
        
        'Define how many columns of data to pull from Source Sheet
        
        If Not Val = "" Then
        'MsgBox (Val)
                'Search Range
                Set rngX = SrcSh.Range("B:B").Find(Val, lookat:=xlPart)
                'Set rngX = SrcSh.Range("B:B").Find(Val, lookat:=xlWhole)
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

MsgBox ("The follwing names couldn't be found in " & sfileName & ". No data has been entered for them:" & vbNewLine _
          & StrMsg)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Cleanup                                                            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Cleanup:

Workbooks(SrcWkb.name).Close SaveChanges:=False
'ActiveWorkbook.Worksheets("FY13Q2").Activate

With Application
    .ScreenUpdating = True
End With
     
End Sub