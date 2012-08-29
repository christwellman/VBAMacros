VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DynamicForm 
   Caption         =   "Dynamic Form"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5145
   OleObjectBlob   =   "DynamicForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DynamicForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub ExitButton_Click()
Unload Me
End Sub

Private Sub PromptLabel_Click()

End Sub

Private Sub SaveButton_Click()
CheckValue = DynamicForm.DynamicComboBox.Value
Me.Hide

End Sub

Private Sub UserForm_Click()

End Sub
