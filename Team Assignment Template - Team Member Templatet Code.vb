VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
'This Updates the last Update Field if user updates value in the given range
    Dim Sh As Worksheet
    Dim ChangeRange As Range
    Set Sh = ActiveSheet
    Dim Isect As Range
    
    
    Set ChangeRange = Sh.Range("B2", "BP27")
    
    Set Isect = Intersect(Target, ChangeRange)
        'MsgBox (Isect.Address)
        
        If Isect Is Nothing Then
            'MsgBox ("Out of range")
        Else
            'MsgBox ("Out of range")
            Sh.Cells(1, 6).Value = Now()
        End If
   
End Sub

Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)

    If Target.Name = "Reset" Then
        response = MsgBox("If yes, data in columns I:BP will be reset to their initial formulas.", vbYesNo, "Are you Sure you want to reset sheet formulas?")
        
        If response = vbYes Then
         Call ResetSheet
         
        Else
            Exit Sub
        End If
        
        Exit Sub
    End If

End Sub

Sub ResetSheet() 'ByVal Target As Excel.Range)
'This macro resets the formulas for a sheet which may have been changed by accident
Dim currentcell As Range
Dim Sh As Worksheet
Dim DateFormula As String

Set currentcell = activecell

Set Sh = ActiveSheet
DateFormula = "=IF(AND(I$2>=$F3,I$2<=$G3),$D3," & """""" & ")"
'MsgBox (Sh.Name)
'MsgBox (DateFormula)

'rewrite hyperline
    With Worksheets(Sh.Name)
     .Hyperlinks.Add Anchor:=.Range("C1"), _
     Address:="", _
     ScreenTip:="Refresh worksheet Data", _
     TextToDisplay:="Reset"
    End With

'Write Formulas to cells and fill across
Sh.Cells(1, 6).Value = Now()
Sh.Cells(3, 9).Value = DateFormula
Sh.Range("I3:I27").FillDown
Sh.Range("I3:BP27").FillRight

'freeze panes
Sh.Range("I3").Select
ActiveWindow.FreezePanes = True

Call SetConditionalFormatsSimple

Application.Calculation = xlCalculationAutomatic

End Sub


