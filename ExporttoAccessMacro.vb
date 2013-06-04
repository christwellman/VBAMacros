Sub ExportMailByFolder()
  'Export specified fields from each mail
  'item in selected folder.
  
  
  '******************* Pre Requisites ***************************'
  ' Must have reference to Microsoft ActiveX Data Objects X.X library
  ' Must have reference to Mircosoft ADO Ext.6.0 for DDL and Security
  ' Must setup OBDC Source for DB to connect with outlook 2010 need to find 32 bit file:
  ' Compatibility issue with 32 bit vs 64 bit Odbcad32.exe):
  '     The 32-bit version of the Odbcad32.exe file is located in the "%systemdrive%\Windows\SysWoW64" folder.
  '     The 64-bit version of the Odbcad32.exe file is located in the "%systemdrive%\Windows\System32" folder.
  '     Also need to ensure that Microsoft Access Database Engine 2010 Redistributable has been installed -- i needed 32 Bit version:
  '     http://www.microsoft.com/en-us/download/confirmation.aspx?id=13255
  ' Tutorial for setting up the connection see: http://www.interfaceware.com/manual/setting_up_odbc_datasource.html
  ' Database and Table must exist before macro can run
  ' System DSN Data Source: PEM-Database
  ' DB Name: P:\DB_FrontEnd_Current.accdb
  ' Table Name: "Requests"
  
  
  Dim ns As Outlook.NameSpace
  Dim objFolder As Outlook.MAPIFolder
  Set ns = GetNamespace("MAPI")
  Set objFolder = ns.PickFolder
  Dim adoConn As ADODB.Connection
  Dim adoRS As ADODB.Recordset
  Dim intCounter As Integer
  Set adoConn = CreateObject("ADODB.Connection")
  Set adoRS = CreateObject("ADODB.Recordset")
  'DSN and target file must exist.
  adoConn.Open "DSN=PEM-Database;"
  adoRS.Open "SELECT * FROM Requests", adoConn, _
       adOpenDynamic, adLockOptimistic
  'Cycle through selected folder.
  For intCounter = objFolder.Items.Count To 1 Step -1
   With objFolder.Items(intCounter)
   'Copy property value to corresponding fields
   'in target file.
    If .Class = olMail Then
      adoRS.AddNew
            'adoRS("User ID") = ParseTextLinePair(.Body, "UserID:")
            adoRS("Requestor") = .SenderName
            adoRS("Request Date") = Mid(ParseTextLinePair(.Body, "Date and Time of Submission:"), 1, 19)
            adoRS("Customer Name") = ParseTextLinePair(.Body, "Customer Name:")
            adoRS("Location") = ParseTextLinePair(.Body, "Customer Site Location:")
            adoRS("Customer Primary Contact") = ParseTextLinePair(.Body, "Customer Primary Contact:")
            'adoRS("Project Type") = ParseTextLinePair(.Body, "Project Type:")
            adoRS("Request Type") = ParseTextLinePair(.Body, "Request Type:")
            adoRS("Requestor Description") = ParseTextLinePair(.Body, "Service Description:") & Chr(13) & _
                    "Specific PM req:" & ParseTextLinePair(.Body, "specific PM requirements:") & Chr(13) & _
                    "Assigned NCE" & ParseTextLinePair(.Body, "Assigned NCE's:")
            'adoRS("Services Revenue") = ParseTextLinePair(.Body, "Services Revenue:")
            adoRS("Project ID (PID)") = ParseTextLinePair(.Body, "PID:")
            adoRS("SO Number") = ParseTextLinePair(.Body, "Sales Order Nbr:")
            adoRS("Deal ID") = ParseTextLinePair(.Body, "DID:")
            adoRS("Start Date") = ParseTextLinePair(.Body, "Project Start Date:")
            adoRS("End Date") = ParseTextLinePair(.Body, "End Date:")
            'adoRS("Kick Off Date") = ParseTextLinePair(.Body, "Project Kick-Off Meeting:")
            'adoRS("On Site PM req") = ParseTextLinePair(.Body, "On site PM requirements(# of Days):")
            'adoRS("Customer Providing PM") = ParseTextLinePair(.Body, "customer providing PM:")
            '1-adoRS("Specific PM req") = ParseTextLinePair(.Body, "specific PM requirements:")
            adoRS("Sales Representative") = ParseTextLinePair(.Body, "DCN Delivery Manager:")
            '2-adoRS("Assigned NCE") = ParseTextLinePair(.Body, "Assigned NCE's:")
            'adoRS("Segment") = ParseTextLinePair(.Body, "Theather/Market: Mkt Seg - ")
            'adoRS("Geography") = ParseTextLinePair(.Body, "US Enterprise Geography:")
            adoRS("Funding") = ParseTextLinePair(.Body, "Funding:")
      adoRS.Update
     End If
    End With
   Next
  adoRS.Close
  Set adoRS = Nothing
  Set adoConn = Nothing
  Set ns = Nothing
  Set objFolder = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Cool little procedure to parse data where it is on a single line  '
' followed by carriage return                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function ParseTextLinePair(strSource As String, strLabel As String)
    Dim intLocLabel As Integer
    Dim intLocCRLF As Integer
    Dim intLenLabel As Integer
    Dim strText As String
     
    ' locate the label in the source text
    intLocLabel = InStr(strSource, strLabel)
    intLenLabel = Len(strLabel)
        If intLocLabel > 0 Then
        intLocCRLF = InStr(intLocLabel, strSource, vbCrLf)
        If intLocCRLF > 0 Then
            intLocLabel = intLocLabel + intLenLabel
            strText = Mid(strSource, _
                            intLocLabel, _
                            intLocCRLF - intLocLabel)
                            
        Else
            intLocLabel = Mid(strSource, intLocLabel + intLenLabel)
        End If
    End If
    ParseTextLinePair = Trim(strText)
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Parse Text between to given strings                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetBetween(sSearch As String, sStart As String, sStop As String, Optional ByRef lSearch As Long = 1) As String
    
    lSearch = InStr(lSearch, sSearch, sStart)
    If lSearch > 0 Then
        lSearch = lSearch + Len(sStart)
        Dim lTemp As Long
        lTemp = InStr(lSearch, sSearch, sStop)
        If lTemp > lSearch Then
            GetBetween = Mid$(sSearch, lSearch, lTemp - lSearch)
        End If
    End If
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This procedure uses the ParseTextLinePair function to pull        '
' information from the Engagment Request Form                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function ParseFormInfo()
    Dim objItem As Object
    Dim intLocAddress As Integer
    Dim intLocCRLF As Integer
    Dim Requestor As String
    Dim RequestInfo() As String

    
    
       
            
    Set objItem = GetCurrentItem()
    If objItem.Class = olMail Then
        ' Pull Requestor info from Message Header
            Requestor = objItem.SenderName
        ' Parse Information from the Email Body
            UserID = ParseTextLinePair(objItem.Body, "UserID:")
            submitTime = ParseTextLinePair(objItem.Body, "Date and Time of Submission:")
            CustName = ParseTextLinePair(objItem.Body, "Customer Name:")
            CustSiteLoc = ParseTextLinePair(objItem.Body, "Customer Site Location:")
            CustPriContact = ParseTextLinePair(objItem.Body, "Customer Primary Contact:")
            projType = ParseTextLinePair(objItem.Body, "Project Type:")
            RequestType = ParseTextLinePair(objItem.Body, "Request Type")
            ServiceDesc = ParseTextLinePair(objItem.Body, "Service Description:")
            Scoped = ParseTextLinePair(objItem.Body, "Has engagement been scoped:")
            ServiceRev = ParseTextLinePair(objItem.Body, "Services Revenue:")
            PID = ParseTextLinePair(objItem.Body, "PID:")
            SONum = ParseTextLinePair(objItem.Body, "Sales Order Nbr:")
            DealID = ParseTextLinePair(objItem.Body, "DID:")
            StartDate = ParseTextLinePair(objItem.Body, "Project Start Date:")
            EndDate = ParseTextLinePair(objItem.Body, "End Date:")
            KickOffDate = ParseTextLinePair(objItem.Body, "Project Kick-Off Meeting:")
            OnSitePM = ParseTextLinePair(objItem.Body, "On site PM requirements(# of Days):")
            CustPM = ParseTextLinePair(objItem.Body, "customer providing PM:")
            PMRequirements = ParseTextLinePair(objItem.Body, "specific PM requirements:")
            AcctManager = ParseTextLinePair(objItem.Body, "DCN Delivery Manager:")
            AssignedNCEs = ParseTextLinePair(objItem.Body, "Assigned NCE's:")
            Segment = ParseTextLinePair(objItem.Body, "Theather/Market: Mkt Seg - ")
            Geography = ParseTextLinePair(objItem.Body, "US Enterprise Geography:")
            Funding = ParseTextLinePair(objItem.Body, "Funding:")
            
        'For De-Bug read back what you found
        MsgBox Requestor & Chr(10) & UserID & Chr(10) & submitTime & Chr(10) & CustName & Chr(10) & CustSiteLoc & Chr(10) & CustPriContact & Chr(10) & projType & Chr(10) & ServiceDesc & Chr(10) & Scoped & Chr(10) & ServiceRev & Chr(10) & PID & Chr(10) & SONum & Chr(10) & DealID & Chr(10) & StartDate & Chr(10) & EndDate & Chr(10) & KickOffDate & Chr(10) & OnSitePM & Chr(10) & CustPM & Chr(10) & PMRequirements & Chr(10) & DeliveryMgr & Chr(10) & AssignedNCEs & Chr(10) & Segment & Chr(10) & Geography, vbAbortRetryIgnore, "Check"

            
    End If
    
    Set objItem = Nothing
End Function

