Sub ExportMailByFolder()
  'Export specified fields from each mail
  'item in selected folder.
  
  
  '******************* Pre Requisites ***************************'
  ' Must have reference to Microsoft ActiveX Data Objects X.X library
  ' Must have reference to Mircosoft ADO Ext.6.0 for DDL and Security
  ' Must setup OBDC Source for DB to connect with outlook 2010 need to find 32 bit file:
  ' Compatibility issue with 32 bit vs 64 bit Odbcad32.exe):
  '     The 32-bit version of the Odbcad32.exe file is located in the %systemdrive%\Windows\SysWoW64 folder.
  '     The 64-bit version of the Odbcad32.exe file is located in the %systemdrive%\Windows\System32 folder.
  ' Tutorial for setting up teh connection see: http://www.interfaceware.com/manual/setting_up_odbc_datasource.html
  ' Database and Table must exist before macro can run
  ' DB Name: C:\MyDatabase.mdb
  '	Table Name: "email"
  
  
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
  adoConn.Open "DSN=OutlookData;"
  adoRS.Open "SELECT * FROM email", adoConn, _
       adOpenDynamic, adLockOptimistic
  'Cycle through selected folder.
  For intCounter = objFolder.Items.Count To 1 Step -1
   With objFolder.Items(intCounter)
   'Copy property value to corresponding fields
   'in target file.
    If .Class = olMail Then
      adoRS.AddNew
      adoRS("Subject") = .subject
      adoRS("Body") = .Body
      adoRS("FromName") = .SenderName
      adoRS("ToName") = .To
      adoRS("FromAddress") = .SenderEmailAddress
      adoRS("FromType") = .SenderEmailType
      adoRS("CCName") = .CC
      adoRS("BCCName") = .BCC
      adoRS("Importance") = .Importance
      adoRS("Sensitivity") = .Sensitivity
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