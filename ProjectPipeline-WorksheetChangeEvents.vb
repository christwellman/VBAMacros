Sub Worksheet_Change(ByVal Target As Excel.Range)
'use this to automatically update cells and move rows when values in teh sheet are changed

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
'Set ws2 = Sheets("Project Tracking - BBODEN")
Set ws3 = Sheets("Project Tracking - SMILESNI")
Set ws4 = Sheets("Project Tracking - EVOGEL")
Set ws5 = Sheets("Closed Projects")
Set ws6 = Sheets("Project Tracking - MILASKIN")
Set ws7 = Sheets("Project Tracking - AMALAN")
Set ws8 = Sheets("Complete - Presales - Scoping")
'Set Brendon = ws0.Range("BrendonsTeam")
Set Scott = ws0.Range("ScottsTeam")
Set Erik = ws0.Range("EriksTeam")

    
    
'Automaticly move for staffing hinges on Column "R"
    If Target.Column = 20 Then
        ThisRow = Target.Row
        
        assignedPM = Target.Value
        If assignedPM <> vbNullString Then
        

ConfirmAssigned:
            'Test For Ready For Staffing
            If ActiveSheet.Cells(ThisRow, 20).Value = "Needs DCPM" Then
                
                        'Update Date Ready for Staffing
                        If ActiveSheet.Cells(ThisRow, 50).Value = vbNullString Then
                            ActiveSheet.Cells(ThisRow, 50).Value = Now()
                        End If
                        ActiveSheet.Cells(ThisRow, 35).Value = "Pipeline - Pending Assignment"
                    
                        'Call Copy to
                        ws1.Cells(ThisRow, 1).Resize(1, 52).Interior.ColorIndex = 43
                        
      
                
            'Test for Scott
            ElseIf (WorksheetFunction.CountIf(Scott, assignedPM) > 0) Then
                M = InputBox(Prompt:="Confirm Assignment time and move to Scott's Tab", Title:="Assign and Move", Default:=Now())
                If StrPtr(M) = 0 Then
                    Target.Value = vbNullString
                Else
                    If M = vbNullString Then
                        GoTo ConfirmAssigned
                    Else
                        'Update DCPM Assigned Date
                        If ActiveSheet.Cells(ThisRow, 38).Value = vbNullString Then
                            ActiveSheet.Cells(ThisRow, 38).Value = M
                        End If
                        ActiveSheet.Cells(ThisRow, 35).Value = "Assigned"
                    
                        'Call Copy to
                        ws1.Cells(ThisRow, 1).Resize(1, 52).Copy
                    
                        'find the last row on teh destination sheet
                        moverow = Lastrow(ws3) + 1
                        ws3.Cells(moverow, 1).PasteSpecial xlPasteAll
                    
                        'Remove source row
                        ws1.Cells(ThisRow, 1).EntireRow.Delete
                    End If
                End If

            'Test for Erik
            ElseIf (WorksheetFunction.CountIf(Erik, assignedPM) > 0) Then
                M = InputBox(Prompt:="Confirm Assignment time and move to Erik's Tab", Title:="Assign and Move", Default:=Now())
                If StrPtr(M) = 0 Then
                    Target.Value = vbNullString
                Else
                    If M = vbNullString Then
                        GoTo ConfirmAssigned
                    Else
                        'Update DCPM Assigned Date
                        If ActiveSheet.Cells(ThisRow, 38).Value = vbNullString Then
                            ActiveSheet.Cells(ThisRow, 38).Value = M
                        End If
                        ActiveSheet.Cells(ThisRow, 35).Value = "Assigned"
                    
                        'Call Copy to
                        ws1.Cells(ThisRow, 1).Resize(1, 52).Copy
                    
                        'find the last row on teh destination sheet
                        moverow = Lastrow(ws4) + 1
                        ws4.Cells(moverow, 1).PasteSpecial xlPasteAll
                    
                        'Remove source row
                        ws1.Cells(ThisRow, 1).EntireRow.Delete
                    End If
                End If

            End If
        Else
            M = MsgBox("Confirm clear PM assiged data?", vbYesNo + vbMsgBoxSetForeground, "Clear Assignement Data")
                If M = vbYes Then
                    'Update DCPM Assigned Date
                    ActiveSheet.Cells(ThisRow, 38).Value = ""
                    ws1.Cells(ThisRow, 1).Resize(1, 52).Interior.ColorIndex = xlNone
                End If
        End If
    End If
'Exit Sub
    
 'Move to Closed tab if PID is closed
    If Target.Column = 7 Then
        ThisRow = Target.Row
        
        'Check to see if PID is Delivery Close
        If ActiveSheet.Cells(ThisRow, 7).Value = "Delivery Close" Then
        
            'Update Delivery Close Date
            ActiveSheet.Cells(ThisRow, 41).Value = Now()
            
        ElseIf ActiveSheet.Cells(ThisRow, 7).Value = "Closed" Then
        M = MsgBox("Are you sure you want to move this to the 'Closed' tab?", vbYesNo + vbMsgBoxSetForeground, "Move Project?")
            If M = vbYes Then
        
            'Update Closed Date
            ActiveSheet.Cells(ThisRow, 42).Value = Now()
            
            'Update DCPM Status
            ActiveSheet.Cells(ThisRow, 35).Value = "Closed"
            
            'Find the last row on teh Closed Tab
            moverow = Lastrow(ws5) + 1
        
            'copy used range in active row
            ActiveSheet.Cells(ThisRow, 1).Resize(1, 52).Copy
            ws5.Cells(moverow, 1).PasteSpecial xlPasteAllExceptBorders
            
            'Remove source row
            ActiveSheet.Cells(ThisRow, 1).EntireRow.Delete
            Else
            
            End If
        ElseIf ActiveSheet.Cells(ThisRow, 7).Value = "Cancelled" Then
        M = MsgBox("Are you sure you want to move this to the 'Complete' tab?", vbYesNo + vbMsgBoxSetForeground, "Move Project?")
            If M = vbYes Then
        
            'Update Closed Date
            ActiveSheet.Cells(ThisRow, 42).Value = Now()
            
            'Update DCPM Status
            ActiveSheet.Cells(ThisRow, 35).Value = "Cancelled"
            
            'Find the last row on teh Closed Tab
            moverow = Lastrow(ws5) + 1
         
            'copy used range in active row
            ActiveSheet.Cells(ThisRow, 1).Resize(1, 52).Copy
            ws8.Cells(moverow, 1).PasteSpecial xlPasteAllExceptBorders
            
            'Remove source row
            ActiveSheet.Cells(ThisRow, 1).EntireRow.Delete
            
            ElseIf M = vbNo Then Exit Sub
            
            End If
            
        End If

    End If
    
    'PM Status Changes
    If Target.Column = 35 Then
        ThisRow = Target.Row
        
        If ActiveSheet.Cells(ThisRow, 35).Value = "Complete - Presales Only" Or _
            ActiveSheet.Cells(ThisRow, 35).Value = "Complete - Handled out of Practice" Or _
            ActiveSheet.Cells(ThisRow, 35).Value = "Complete - DCV work finished" Then
                M = MsgBox("Are you sure you want to move this to the 'Complete' tab?", vbYesNo + vbMsgBoxSetForeground, "Move Project?")
                If M = vbYes Then
        
                'Update Closed Date
                ActiveSheet.Cells(ThisRow, 42).Value = Now()
            
                'Update DCPM Status
                ActiveSheet.Cells(ThisRow, 35).Value = "Complete"
            
                'Find the last row on teh Closed Tab
                moverow = Lastrow(ws8) + 1
        
                'copy used range in active row
                ActiveSheet.Cells(ThisRow, 1).Resize(1, 52).Copy
                ws8.Cells(moverow, 1).PasteSpecial xlPasteAllExceptBorders
            
                'Remove source row
                ActiveSheet.Cells(ThisRow, 1).EntireRow.Delete
                End If
        
        ElseIf ActiveSheet.Cells(ThisRow, 35).Value = "Cold" Or _
            ActiveSheet.Cells(ThisRow, 35).Value = "Duplicate Request" Or _
            ActiveSheet.Cells(ThisRow, 35).Value = "Cancelled" Then
                M = MsgBox("Are you sure you want to move this to the 'Complete' tab?", vbYesNo + vbMsgBoxSetForeground, "Move Project?")
                If M = vbYes Then
        
                'Update Closed Date
                ActiveSheet.Cells(ThisRow, 42).Value = Now()
            
                'Update DCPM Status
                ActiveSheet.Cells(ThisRow, 35).Value = "Complete"
            
                'Find the last row on teh Closed Tab
                moverow = Lastrow(ws8) + 1
        
                'copy used range in active row
                ActiveSheet.Cells(ThisRow, 1).Resize(1, 52).Copy
                ws8.Cells(moverow, 1).PasteSpecial xlPasteAllExceptBorders
            
                'Remove source row
                ActiveSheet.Cells(ThisRow, 1).EntireRow.Delete
                End If
        End If
    End If

End Sub