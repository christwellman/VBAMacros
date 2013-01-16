VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgressBar 
   Caption         =   "Processing..."
   ClientHeight    =   435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4425
   OleObjectBlob   =   "frmProgressBar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub UserForm_Activate()
    Call ShowProgressBarWithoutPercentage
End Sub
  
Sub ShowProgressBarWithoutPercentage()
    Dim Percent As Integer
    Dim PercentComplete As Single
    Dim MaxRow, MaxCol As Integer
    Dim iRow, iCol As Integer
    MaxRow = 500
    MaxCol = 500
    Percent = 0
'Initially Set the width of the Label as Zero
    frmProgressBar.LabelProgress.Width = 0
    For iRow = 1 To MaxRow
        For iCol = 1 To MaxCol
            Worksheets("Master").Cells(iRow, iCol).Value = iRow * iCol
  
        Next
        PercentComplete = iRow / MaxRow
        frmProgressBar.LabelProgress.Width = PercentComplete * frmProgressBar.Width
                frmProgressBar.LabelProgress.Caption = Format(PercentComplete, "0%")
        DoEvents
    Next
    Unload frmProgressBar
End Sub
