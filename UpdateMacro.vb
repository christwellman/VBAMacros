Sub UpdateQuery()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Updates Data across worksheets from OP export (gppr, etc) '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Sh As Worksheet
    Dim DestSh As Worksheet
    Dim Sourcefile As Workbook
    Dim Lookupfile As Workbook
    Dim Last As Long
    Dim shLast As Long
    Dim CopyRng As Range
    Dim StartRow As Long
    Dim CheckRow As Range
    Dim CheckRange As Range
    Dim offsetaddress As Range
    'Dim SearchParam As Long
    Dim PIDStatus As String

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    'Open the Data Source
    'Set Sourcefile = "C:\Users\ctwellma\Documents\AS\PMO Pipeline\Tester.xlsm"
    'Workbooks.Open Filename:= _
    '   "C:\Users\ctwellma\Documents\AS\Reports\GPPR\Global Profitability by Project Report - by PID 8-21-12.xls"
    

     ' Hide unneccssary sheets.
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets("LOVs").Visible = xlSheetHidden
    ActiveWorkbook.Worksheets("Complete - Presales - Scoping").Visible = xlSheetHidden
    ActiveWorkbook.Worksheets("Cold Projects").Visible = xlSheetHidden
    ActiveWorkbook.Worksheets("Closed Projects").Visible = xlSheetHidden
    ActiveWorkbook.Worksheets("Project Tracking - BBODEN").Visible = xlSheetHidden
    ActiveWorkbook.Worksheets("Project Tracking - GPAGE").Visible = xlSheetHidden
    ActiveWorkbook.Worksheets("Archive Desk Complete").Visible = xlSheetHidden
    ActiveWorkbook.Worksheets("Archive Cold Projects").Visible = xlSheetHidden
    ActiveWorkbook.Worksheets("Archive Closed Projects").Visible = xlSheetHidden
    ActiveWorkbook.Worksheets("New WM Mapping").Visible = xlSheetHidden
    
    On Error GoTo 0
    

    ' Fill in the start row.
    StartRow = 2
    

    'Loop Through Sheets
    For Each Sh In Workbooks("Tester.xlsm").Worksheets
        With Sh
            If Sh.Visible <> xlSheetHidden Then
            Last = Lastrow(Sh)
            Firstrow = StartRow
            Set CheckRange = Range(Cells(Firstrow, "A"), Cells(Last, "AP"))
                'For Lrow = Last To Firstrow Step -1
                For Each CheckRow In CheckRange.Rows
                            If Len(Cells(CheckRow.Row, "F").Value) = 6 Then
                                'Set Search Variable
                                SearchParam = Cells(CheckRow.Row, "F").Value
                                'MsgBox (Sh.Name & " " & SearchParam)
                                
                                'Search Through GPPR
								'
								'Need to reference the GPPR Report instead of exsitng worksheet
								'
								
                                With Worksheets("Project Pipeline").Range("f1:f500")
                                    Set C = .Find(SearchParam, LookIn:=xlValues)
                                    If Not C Is Nothing Then
                                        Set offsetaddress = Range(C.Address).Offset(0, 1)
                                        MsgBox (SearchParam & " " & offsetaddress.Value)

                                    End If
                                End With
                            Else
                                'Skip this incase it's not avaiable
                            End If
                Next 'Next Row
            End If
        End With
    Next 'WorkSheet

ExitTheSub:

End Sub

