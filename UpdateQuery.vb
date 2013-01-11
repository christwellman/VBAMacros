Sub LookUp()
Dim Sh As Worksheet
Dim DestSh As Worksheet
Dim last As Long
Dim CopyRange As Range
Dim CopyValue As String

'File name
'strSheet = "DataCenterPracticeNewMetricsDatasheet.xlsm"
'File Path
'strPath = "C:\Users\ctwellma\Desktop\"


Set Sh = ActiveWorkbook.Worksheets("Sheet2")
Set DestSh = ActiveWorkbook.Worksheets("Sheet1")

'Loop Through Rows to Remove Blanks and Format
    last = Lastrow(Sh)
    endofpage = LastCol(Sh)
    Debug.Print endofpage
    firstrow = Sh.UsedRange.Cells(1).Row
    lrow = last + firstrow - 1
    
    With Sh 'Sheet 2 Loop through Each Name on Sheet 2
        MsgBox ("Outerloop" & Sh.Name)
        .DisplayPageBreaks = False
            For lrow = last To firstrow Step -1
            'Get Value to Check Against/Look Up
                last2 = Lastrow(DestSh)
                firstrow2 = DestSh.UsedRange.Cells(1).Row
                Lrow2 = last2 + firstrow2 - 1
                MsgBox (.Cells(lrow, "A").Value & " " & Sh.Name)
                
                With DestSh ' Try to match name from Sheet 2 with list in Sheet 1
                    MsgBox ("Innerloop" & DestSh.Name)
                    For Lrow2 = last2 To firstrow2 Step -1
                    '
                        'MsgBox (.Cells(Lrow2, "A").Value & " " & DestSh.Name)
                        
                        If .Cells(lrow, "A").Value = .Cells(Lrow2, "A").Value Then
                        MsgBox ("found one" & .Cells(lrow, "A").Value)
                        
                       '     MsgBox ("Match Found" & "$A$" & Lrow & .Cells(Lrow, "A").Value & " " & Sh.Name & " = " & "$A$" & Lrow2 & .Cells(Lrow2, "A").Value & " " & DestSh.Name)
                       '     '.Cells(Lrow2, "B").Value = "$A$" & lrow
                       ' Else
                       '     'MsgBox ("no Match")
                        End If
                        
                    Next
                End With
            Next
            
            
    End With

End Sub