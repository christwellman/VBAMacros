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
dim UpdatedCell as Integer
dim PracticeDM as String
dim PIDStatus as string

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

'**************************************************************
ThisRow = Target.Row
UpdatedCell = Target.Column
Select Case UpdatedCell
    'PM assignement or Ready to staff determined by column "T" or 20
	Case 20
        assignedPM = Target.Value
		Select Case assingedPM
			Case VBNullString
				'If PM name is changed to blank to you want to clear the input data?
				m = MsgBox("Confirm clear PM assiged data?", vbYesNo + vbMsgBoxSetForeground, "Clear Assignement Data")
                If m = vbYes Then
                    'Update DCPM Assigned Date
                    ActiveSheet.Cells(ThisRow, 38).Value = ""
                    ws1.Cells(ThisRow, 1).Resize(1, 52).Interior.ColorIndex = xlNone
                End If
			Case "Needs DCPM"
				'Project is ready for staffing
				' Update Ready for staffing date if its blank
				If ActiveSheet.Cells(ThisRow, 50).Value = vbNullString Then
					ActiveSheet.Cells(ThisRow, 50).Value = Now()
				End If
				'Update PM status
				ActiveSheet.Cells(ThisRow, 35).Value = "Pending Assignment"
				'Add Pretty color to indicate this ones ready
				ws1.Cells(ThisRow, 1).Resize(1, 52).Interior.ColorIndex = 43
			Case Else
				If (WorksheetFunction.CountIf(Scott, assignedPM) > 0) Then
                m = InputBox(Prompt:="Confirm Assignment time and move to Scott's Tab", Title:="Assign and Move", Default:=Now())
					If StrPtr(m) = 0 Then
						Target.Value = vbNullString
					Else
						If m = vbNullString Then
							GoTo ConfirmAssigned
						Else
							'Update DCPM Assigned Date
							If ActiveSheet.Cells(ThisRow, 38).Value = vbNullString Then
								ActiveSheet.Cells(ThisRow, 38).Value = m
							End If
							ActiveSheet.Cells(ThisRow, 21).Value = "Scott Milesnick"
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
                m = InputBox(Prompt:="Confirm Assignment time and move to Erik's Tab", Title:="Assign and Move", Default:=Now())
					If StrPtr(m) = 0 Then
						Target.Value = vbNullString
					Else
						If m = vbNullString Then
							GoTo ConfirmAssigned
						Else
							'Update DCPM Assigned Date
							If ActiveSheet.Cells(ThisRow, 38).Value = vbNullString Then
								ActiveSheet.Cells(ThisRow, 38).Value = m
							End If
							ActiveSheet.Cells(ThisRow, 35).Value = "Assigned"
							ActiveSheet.Cells(ThisRow, 21).Value = "Erik Vogel"
						
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
		End Select
	'WOrksheet CHange events for PIDStatus Changes	
    Case 7
		PIDStatus = Activesheet.Cells(thisrow,7).Value
		Select Case PIDStatus
			Case "Delivery Close"
				'Update Delivery Close Date
				ActiveSheet.Cells(ThisRow, 41).Value = Now()
			Case "Closed"
				m = MsgBox("Are you sure you want to move this to the 'Closed' tab?", vbYesNo + vbMsgBoxSetForeground, "Move Project?")
				If m = vbYes Then
			
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
			Case "Cancelled"
				m = MsgBox("Are you sure you want to move this to the 'Complete' tab?", vbYesNo + vbMsgBoxSetForeground, "Move Project?")
				If m = vbYes Then
			
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
				End If
        End Select
    Case 35
		If ActiveSheet.Cells(ThisRow, 35).Value = "Complete - Presales Only" Or _
            ActiveSheet.Cells(ThisRow, 35).Value = "Complete - Handled out of Practice" Or _
            ActiveSheet.Cells(ThisRow, 35).Value = "Complete - DCV work finished" Then
				m = MsgBox("Are you sure you want to move this to the 'Complete' tab?", vbYesNo + vbMsgBoxSetForeground, "Move Project?")
				If m = vbYes Then

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
	
	Case Else
        
End Select
    
    
	
'***************************************************************
End Sub

Function Lastrow(Sh As Worksheet)
    On Error Resume Next
    Lastrow = Sh.Cells.Find(What:="*", _
                            After:=Sh.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    On Error GoTo 0
End Function
