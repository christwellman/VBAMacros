VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DateValidationForm 
   Caption         =   "Date Validation Form"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5775
   OleObjectBlob   =   "DateValidationForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DateValidationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelButton_Click()
Unload Me
End Sub



Private Sub UpdateButton_Click()
    Dim ws0 As Worksheet
    Dim ws1 As Worksheet
    Dim RowNumber As Integer

    Set ws0 = Sheets("LOVs")
    Set ws1 = ActiveSheet

    RowNumber = ActiveCell.Row

    '''''''''''''''''''''''''''
    'Perform Date Validations '
    '''''''''''''''''''''''''''
    If IsDate(ValidRequestDate) = False Then
        DateValidationForm.RequestBox.ForeColor = &HFF&

    
        If IsDate(ValidFollowUpDate) = True Then
            If CDate(ValidFollowUpDate) >= CDate(ValidRequestDate) Then
            End If
        Else
            DateValidationForm.FollowUpDateLabel.ForeColor = &HFF&
        End If

    Else
        DateValidationForm.RequestDateLabel.ForeColor = &H80000012
        DateValidationForm.FollowUpDateLabel.ForeColor = &H80000012
        DateValidationForm.DCPMAssignedLabel.ForeColor = &H80000012
        DateValidationForm.SAAssignedLabel.ForeColor = &H80000012
        DateValidationForm.WalkerSurveyLabel.ForeColor = &H80000012
        DateValidationForm.DeliveryCloseLabel.ForeColor = &H80000012
        DateValidationForm.ProjectCloseLabel.ForeColor = &H80000012
        Me.Hide
        Call DateValidationForm_Initialize
        
    End If
    
    ws1.Cells(RowNumber, "D").Value = DateValidationForm.RequestBox.Value
    ws1.Cells(RowNumber, "AJ").Value = DateValidationForm.FollowUpBox.Value
    ws1.Cells(RowNumber, "AK").Value = DateValidationForm.DCPMAssignedBox.Value
    'ValidSAAssignedDate = DateValidationForm.SAAssignedBox.Value
    'ValidWalkerSentDate = DateValidationForm.WalkerSentBox.Value
    'ValidDeliveryCloseDate = DateValidationForm.DeliveryClose.Value
    'ValidProjectCloseDate = DateValidationForm.ProjectCloseBox.Value
    
End Sub


Sub DateValidationForm_Initialize()
    Dim ws0 As Worksheet
    Dim ws1 As Worksheet
    Dim RowNumber As Integer

    Set ws0 = Sheets("LOVs")
    Set ws1 = ActiveSheet
    MsgBox (ActiveSheet.Name)

    RowNumber = ActiveCell.Row
    
    MsgBox ("Active Row " & RowNumber)
    
    RequestBox.Text = ws1.Cells(RowNumber, "D").Value
    FollowUpBox.Text = ws1.Cells(RowNumber, "AJ").Value
    DCPMAssignedBox.Text = ws1.Cells(RowNumber, "AK").Value
    
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub WalkerSentBox_Change()

End Sub

Private Sub WalkerSurveyLabel_Click()

End Sub
