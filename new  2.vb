Sub StaffAndMove ()
'use this to move a project to another tab on action

Dim ws0 As Worksheet
Dim ws1 As Worksheet
Dim ws2 As Worksheet
Dim ws3 As Worksheet
Dim ws4 As Worksheet
Dim ws5 As Worksheet
Dim ws6 As Worksheet
Dim ws7 As Worksheet
Dim ws8 As Worksheet

Dim LR As Long
Dim C As Range
Dim Brendon As Range
Dim David As Range
Dim Erik As Range
Dim copyrange As Range

Set ws0 = Sheets("LOVs")
Set ws1 = Sheets("Project Pipeline")
Set ws2 = Sheets("Project Tracking - BBODEN")
Set ws3 = Sheets("Project Tracking - SMILESNI")
Set ws4 = Sheets("Project Tracking - EVOGEL")
Set ws5 = Sheets("Closed Projects")
Set ws6 = Sheets("Project Tracking - MILASKIN")
Set ws7 = Sheets("Project Tracking - AMALAN")
Set ws8 = Sheets("Complete - Presales - Scoping")
Set Brendon = ws0.Range("BrendonsTeam")
Set Scott = ws0.Range("ScottsTeam")
Set Erik = ws0.Range("EriksTeam")
    
    
'Automaticly move for staffing hinges on Column "R"
    If Target.Column = 17 Then
        Thisrow = Target.Row
        
        assignedPM = Target.Value
        If assignedPM <> vbNullString Then
        
            'Update DCPM Assigned Date
            Range("AH" & Thisrow).Value = Now()
            
            'Test For Brendon
            If (WorksheetFunction.CountIf(Brendon, assignedPM) > 0) Then
                m = MsgBox("Confirm Assignment and move to Brendon's Tab", vbYesNo + vbMsgBoxSetForeground, "Assign and Move")
                If m = vbYes Then
                    'Update DCPM Assigned Date
                    Range("AH" & Thisrow).Value = Now()
                    'Call Copy to
                    ws1.Cells(Thisrow, 1).Resize(1, 52).Copy
                    ws2.Cells(200, 1).PasteSpecial xlPasteAll
                End If
                
            'Test for Scott
            ElseIf (WorksheetFunction.CountIf(Scott, assignedPM) > 0) Then
                m = MsgBox("Confirm Assignment and move to Scott's Tab", vbYesNo + vbMsgBoxSetForeground, "Assign and Move")
                If m = vbYes Then
                    'Update DCPM Assigned Date
                    Range("AH" & Thisrow).Value = Now()
                    'Call Copy to
                    ws1.Cells(Thisrow, 1).Resize(1, 52).Copy
                    ws3.Cells(200, 1).PasteSpecial xlPasteAll
                End If
                
            'Test for Erik
            ElseIf (WorksheetFunction.CountIf(Erik, assignedPM) > 0) Then
                m = MsgBox("Confirm Assignment and move to Erik's Tab", vbYesNo + vbMsgBoxSetForeground, "Assign and Move")
                If m = vbYes Then
                    'Update DCPM Assigned Date
                    Range("AH" & Thisrow).Value = Now()
                    'Call Copy to
                    'What to Copy
                    ws1.Cells(Thisrow, 1).Resize(1, 52).Cut
                    'Where to paste
                    ws4.Cells(200, 1).PasteSpecial xlPasteAll
                End If

               
            'Else
            '   Range("B" & ThisRow).Interior.ColorIndex = xlColorIndexNone
            End If
            
        End If
    End If
    
 'Move to Closed tab if PID is closed
    If Target.Column = 7 Then
        Thisrow = Target.Row
        
        'Check to see if PID is Delivery Close
        If ActiveSheet.Cells(Thisrow, 7).Value = "Delivery Close" Then
        
            'Update Delivery Close Date
            ActiveSheet.Cells(Thisrow, 41).Value = Now()
            
        ElseIf ActiveSheet.Cells(Thisrow, 7).Value = "Closed" Then
        m = MsgBox("Are you sure you want to move this to the 'Closed' tab?", vbYesNo + vbMsgBoxSetForeground, "Move Project?")
            If m = vbYes Then
        
            'Update Closed Date
            ActiveSheet.Cells(Thisrow, 42).Value = Now()
            
            'Update DCPM Status
            ActiveSheet.Cells(Thisrow, 35).Value = "Closed"
            
            'Find the last row on teh Closed Tab
            moverow = Lastrow(ws5) + 1
        
            'copy used range in active row
            ActiveSheet.Cells(Thisrow, 1).Resize(1, 52).Copy
            ws5.Cells(moverow, 1).PasteSpecial xlPasteAllExceptBorders
            
            'Remove source row
            ActiveSheet.Cells(Thisrow, 1).EntireRow.Delete
            Else
            
            End If
        ElseIf ActiveSheet.Cells(Thisrow, 7).Value = "Cancelled" Then
        m = MsgBox("Are you sure you want to move this to the 'Complete' tab?", vbYesNo + vbMsgBoxSetForeground, "Move Project?")
            If m = vbYes Then
        
            'Update Closed Date
            ActiveSheet.Cells(Thisrow, 42).Value = Now()
            
            'Update DCPM Status
            ActiveSheet.Cells(Thisrow, 35).Value = "Cancelled"
            
            'Find the last row on teh Closed Tab
            moverow = Lastrow(ws5) + 1
         
            'copy used range in active row
            ActiveSheet.Cells(Thisrow, 1).Resize(1, 52).Copy
            ws8.Cells(moverow, 1).PasteSpecial xlPasteAllExceptBorders
            
            'Remove source row
            ActiveSheet.Cells(Thisrow, 1).EntireRow.Delete
            End If
            
        End If

    End If
                
End Sub