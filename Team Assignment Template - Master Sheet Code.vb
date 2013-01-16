VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    If Target.Name = "Update Sheet" Then '$C$4 will vary depending on where the hyperlink is
        Call UpdateMaster
        'MsgBox ("you clicked it")
        
        Exit Sub
    ElseIf Target.Name = "Add new EM/PM" Then
        'call me maybe
        Call AddNewTeamMemberTab
    End If

End Sub

Sub Test()
MsgBox ("procedure Called")

End Sub
