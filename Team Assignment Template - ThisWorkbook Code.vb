VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_NewSheet(ByVal Sh As Object)
'This Sub Creates a copy of a sheet template with the name of the new team member and adds all of the macros and formulas etc
Dim NewSheet As Worksheet
Set NewSheet = Sh
Dim TeamMemberName As String
With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
End With

response = MsgBox("If you would like to add a new team member tab click Yes, to create a new standard tab click No?", vbYesNo, "Add Team Member?")
        
        If response = vbYes Then

        Sh.Delete
        
            TeamMemberName = InputBox(Prompt:="What is the name of the new team member?", Title:="Add new team member tab?", Default:="")
            If TeamMemberName <> "" Then
            
                'Add a new copy of the team member template
                ActiveWorkbook.Sheets("TEAM-MEMBER-TEMPLATE").Visible = xlSheetVisible
                ActiveWorkbook.Sheets("TEAM-MEMBER-TEMPLATE").Copy _
                   after:=ActiveWorkbook.Sheets(Sheets.Count)
                
                Set NewSheet = Sheets("TEAM-MEMBER-TEMPLATE (2)")
                NewSheet.Name = TeamMemberName
                NewSheet.Visible = xlSheetVisible
                
                'NewSheet.Activate
                NewSheet.Cells(3, 2).Select
                
                
            Else
                Exit Sub
            End If
            
        Else
            Exit Sub
        End If
        
ActiveWorkbook.Sheets("TEAM-MEMBER-TEMPLATE").Visible = xlSheetVeryHidden
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
End With
End Sub

Private Sub Workbook_Open()
Sheets("TEAM-MEMBER-TEMPLATE").Visible = xlSheetVeryHidden
Sheets("Master").Activate

End Sub
