VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TaskInfoForm 
   Caption         =   "Task Information"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9495
   OleObjectBlob   =   "TaskInfoForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TaskInfoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exitButton_Click()
Unload Me
End Sub

Private Sub SaveButton_Click()
'Read from the Form
On Error Resume Next
If TaskInfoForm.taskNamebox.Value <> "" Then
    TaskName = TaskInfoForm.taskNamebox.Value
Else: TaskName = ""
End If

If TaskInfoForm.effortbox.Value <> "" Then
    ApprovedEffort = TaskInfoForm.effortbox.Value
Else: ApprovedEffort = 0
End If

If TaskInfoForm.costbox.Value <> "" Then
    ApprovedCost = TaskInfoForm.costbox.Value
Else: ApprovedCost = 0
End If

Me.Hide
End Sub

Private Sub UserForm_Initialize()
TaskInfoForm.taskNamebox.SetFocus
End Sub
