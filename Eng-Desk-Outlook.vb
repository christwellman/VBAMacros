''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    List of Macros for Engagement Desk Outlook                                                                          '
'	1) Find [EDRepresenative(0) = "Chris Twellman"] and replace with [EDRepresenative(0) = "<First Name> <Last Name>"]   '
'	2) Find [EDRepresenative(1) = "ctwellma@cisco.com"] and replace with [EDRepresenative(1) = "<username>@cisco.com"]   '
'	3) Find [EDRepresenative(2) = "+1 919 392 6154"] and replace with [EDRepresenative(2) = "<Phone Number>"]            '
'	4) Find [TemplatePath = "C:\Users\ctwellma\AppData\Roaming\Microsoft\Templates\"] and replace with _                 '
'				[TemplatePath = "<path to Cloud response Template>"]                                                     '                                                   '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This macro creates a reply to an email to specific parties to     '
' follow up on Workflow Notifications for ASF projects              '
'                                                                   '
'        !!!Relies on the GetCurrentItem() Function!!!              '
'        !!!Relies on the ParseTextLinePair() Function!!!           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ASFWorkflowReply()

'On Error GoTo ErrHandler


'          Define Engagement Desk Sender details                   '
'------------------------------------------------------------------'
    Dim EDRepresenative(0 To 2) As String
        EDRepresenative(0) = "Chris Twellman" 'Name
        EDRepresenative(1) = "ctwellma@cisco.com" 'Email
        EDRepresenative(2) = "+1 919 392 6154" 'Phone


    
    Dim objItem As Object
    Dim Name As Variant
    Dim MailAttach As MailItem
    Dim SalesContact As String
    Dim objReply As MailItem
    Dim objFWD As MailItem
    Dim TestSubject As String
    
    
    
    
    'Test the subject of the message against the string below to see if it qualifies for this macro
    TestSubject = "New AS-Fixed project has been created and is ready for delivery"
        
    Set objItem = GetCurrentItem()
    If objItem.Class = olMail Then
        
        'Test to make sure this is a workflow mailer as fixed notification
        If (objItem.SenderName = "Workflow Mailer") And InStr(1, objItem.subject, TestSubject, vbTextCompare) Then
        
            ' find the requestor address
            subject = objItem.subject

                   
            ' create the reply, add the address and display
            Set objReply = objItem.Reply
            'Reply to:
            objReply.To = ParseTextLinePair(objItem.Body, "Delivery Manager:")
            'Get Recipient Name
            Name = Split(ParseTextLinePair(objItem.Body, "Delivery Manager:"))

            
            'CC addresses
            objReply.CC = "as-dcn-pmo-request@cisco.com; "
            'Reply Subject
            objReply.subject = "RE: " & subject
            'Reply Importance
            'objReply.Importance = olImportanceHigh
            'Reply Attachement from File
            'objReply.Attachments.Add ("C:\Users\")
            
            'Reply Msg Body Composition:
            objReply.HTMLBody = "<p>Hi " & Name(1) & "," & "<br></p>" & _
            "<p>Just received notice of this AS Fixed project being created -- I'm assigned to all of the UCS SKU's as a Work Mananger " & _
            "to facilitate resource assignment from the practices. I wanted to confirm with you that you don't have resources you" & _
            " are planning to assign before putting a practice DCV PM and SA on this.</p>" & _
            "<p>Also please let me know if you have the customer and/or sales contact information we should work with to set " & _
            "a delivery timeline.</p>" & _
            "<p>Thank you,</p>" & _
            "<p>" & EDRepresenative(0) & "<br>" & _
            "PROJECT SPECIALIST DCV ENGAGEMENT DESK <br>" & _
            ".:|:.:|:.  Cisco | Data Center & Virtualization Practice | Solutions Delivery Management Team <br>" & _
            EDRepresenative(1) & " | " & EDRepresenative(2) & "<br>" & _
            "<p>-----Original Message-----" & _
            objItem.HTMLBody & "</p>"
            
            With objReply
                
                ' Resolve Emails (same as hitting "Check Names" button
                Call .Recipients.ResolveAll
                
                ' Show the message
                Call .Display
            End With
    
        Else
            MsgBox ("This macro is intended to create a response to AS-Fixed workflow notifications.  The selected message " & _
            "does not meet that criteria.")
            Exit Sub
            
        End If
    End If
    
    Set objReply = Nothing
    Set objItem = Nothing
    
    
'ErrHandler:

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This macro creates a reply to an email to specific parties to                                                                  '
' follow up on cloud workshop inquiries                                                                                          '
'                                                                                                                                '
'        !!!Relies on the GetCurrentItem() Function!!!                                                                           '
'        !!!Relies on the ParseTextLinePair() Function!!!                                                                        '
'        !!!Relies on File attachment as defined in the File attachement Parameters                                              '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CloudWorkshopFollowUp()

'On Error GoTo ErrHandler
    
    Dim objItem As Object
    Dim Name As Variant
    Dim MailAttach As MailItem
    Dim SalesContact As String
    Dim objReply As MailItem
    Dim TestSubject As String
	
'          Define File Attachement Parameters                      '
'------------------------------------------------------------------'
	Dim TemplateName as String
	Dim TemplatePath As string
	Dim FileLocSting as string
	'Template file Name
	TemplateName = "CloudWorkshopResponse.oft"
	'Template Path
	TemplatePath = "C:\Users\ctwellma\AppData\Roaming\Microsoft\Templates\"
	'File:
	FileLocSting = TemplatePath & TemplateName
	
        
    Set objItem = GetCurrentItem()
    If objItem.Class = olMail Then
        
        'Test to make sure this is a workflow mailer as fixed notification
        'If ParseTextLinePair(objItem.Body, "Request Type:") = "Cloud Workshop" Then
        
            ' find the requestor address
            subject = objItem.subject
                   
            ' create the reply, add the address and display
            Set objReply = CreateItemFromTemplate(FileLocSting)
            'et objReply = objItem.createfromtemplate
            'Reply to:
            objReply.To = objItem.SenderName
            'Get Recipient Name
            Name = Split(objItem.SenderName)
            'CC addresses
            objReply.CC = "as-dcn-pmo-request@cisco.com; "
            'Reply Subject
            objReply.subject = "Cloud Strategy Workshop Inquiry: " & ParseTextLinePair(objItem.Body, "Customer Name:")
            'Reply Importance
            'objReply.Importance = olImportanceHigh
            'Reply Attachement from File
            'objReply.Attachments.Add ("")

            'Customize Msg Body Composition:
            
            
            With objReply
                
                ' Resolve Emails (same as hitting "Check Names" button
                Call .Recipients.ResolveAll
                
                ' Show the message
                Call .Display
            End With
        'End if
        
            With objReply
                'This line adds a line to the top of the message including the Requestor and Customer Names
                .HTMLBody = "<p>" & Name(0) & ", thank you for your inquiring about a Cloud Strategy Workshop for " & ParseTextLinePair(objItem.Body, "Customer Name:") & ".<br></p>" & objReply.HTMLBody
            End With
            'objReply.HTMLBody = "<p>" & Name(0) & ", thank you for your inquiring about Cloud Strategy Workshop for " & ParseTextLinePair(objItem.Body, "Customer Name:") & ".<br></p>"
            
    End If
    
    Set objReply = Nothing
    Set objItem = Nothing
    
    
'ErrHandler:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This macro creates a reply to an email to specific parties with                                                                             '
' specific attachments, and asks for the staffing owner information                                                                           '
' It will then create another message to teh staffing owner with                                                                              '
' the original request information so that they can submit the                                                                                '
' requst                                                                                                                                      '
'                                                                                                                                             '
'        !!!Relies on the GetCurrentItem() Function!!!                                                                                        '
'        !!!Relies on File attachment as defined in the File attachement Parameters                                                          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub OPRMRejectionReply()

On Error GoTo ErrHandler


'          Define Engagement Desk Sender details                   '
'------------------------------------------------------------------'
    Dim EDRepresenative(0 To 2) As String
        EDRepresenative(0) = "Chris Twellman"
        EDRepresenative(1) = "ctwellma@cisco.com"
        EDRepresenative(2) = "+1 919 392 6154"
		
'          Define File Attachement Parameters                      '
'------------------------------------------------------------------'
	Dim TemplateName as String
	Dim TemplatePath As string
	Dim FileLocSting as string
	'Template file Name
	TemplateName = "Resource Management Initiative - Engaging with AS Practices.msg"
	'Template Path
	TemplatePath = "C:\Users\ctwellma\AppData\Roaming\Microsoft\Templates\"
	'File:
	FileLocSting = TemplatePath & TemplateName

    
    Dim objItem As Object
    Dim Name As Variant
    Dim MailAttach As MailItem
    Dim StaffingOwner As String
    Dim objReply As MailItem
    Dim objFWD As MailItem
    Dim name2 As Variant
    
        
    Set objItem = GetCurrentItem()
    If objItem.Class = olMail Then
        ' find the requestor address
        subject = objItem.subject
        'Split Sender Name into array
        Name = Split(objItem.SenderName)
        
        'Define Staffing Owner
        StaffingOwner = InputBox(Prompt:="Who should be the staffing owner for PID: " & ParseTextLinePair(objItem.Body, "PID:"), Title:="Staffing Owner?")
        
        If StaffingOwner = "" Then
            FollowUp = ""
            Set objFWD = Nothing
        Else
            FollowUp = "<p>We believe that person is: " & StaffingOwner & "</p>"
            name2 = Split(StaffingOwner)
            'If Staffing Owner is identified:
                'Create forward email
                Set objFWD = objItem.Forward
                'Forward to:
                objFWD.To = StaffingOwner
                'CC addresses
                objFWD.CC = objItem.SenderName & "; as-dcn-pmo-request@cisco.com"
                'Reply Subject
                'objFWD.subject = "DCV Resource Request Rejection: " & subject & ", Need to use OPRM"
                'Reply Importance
                objFWD.Importance = olImportanceHigh
                
                objFWD.HTMLBody = "<p>Hi " & name2(0) & "," & "<br></p>" & _
                "<p>We've received the staffing request below from " & Name(0) & " " & Name(1) & ", however; as you " & _
                "know AS is now using OPRM for project resourcing, so we'll take no further action until a staffing requirement from OPRM reaches us.</p>" & _
                "<p>Please use the information below as needed for completing the request in the system." & _
                "<p>Thank you,</p>" & _
				"<p>" & EDRepresenative(0) & "<br>" & _
				"PROJECT SPECIALIST DCV ENGAGEMENT DESK <br>" & _
				".:|:.:|:.  Cisco | Data Center & Virtualization Practice | Solutions Delivery Management Team <br>" & _
				EDRepresenative(1) & " | " & EDRepresenative(2) & "<br>" & _
                "<p>-----Original Message-----" & _
                objItem.HTMLBody & "</p>"
        End If
               
        ' create the reply, add the address and display
        Set objReply = objItem.Reply
        'Reply to:
        objReply.To = objItem.SenderName
        'CC addresses
        objReply.CC = objItem.To & "; pegore@cisco.com;"
        'Reply Subject
        objReply.subject = "DCV Resource Request Rejection: " & subject & ", Need to use OPRM"
        'Reply Importance
        objReply.Importance = olImportanceHigh
        'Reply Attachement from File
        objReply.Attachments.Add (FileLocSting)
        
        
        'Reply Msg Body Composition:
        objReply.HTMLBody = "<p>Hi " & Name(0) & "," & "<br></p>" & _
        "<p>Thank you for submitting your request below to the DCV Engagement Desk.  However, as of 9/17/12 the Americas Enterprise, " & _
        "GET, Public Sector, Architectures, Advisory and GSP teams are using Oracle Projects Resource Management (OPRM) to centrally manage" & _
        " resources, and staffing requests should originate from the segment delivery team for this account.</p>" & _
        "<p>We'll forward this on to the responsible person (ex. Enterprise, Theater DM) and CC you, <u>but you are responsible for following " & _
        "up with that person to ensure the OPRM staffing process is utilized.  We will not assign resources until we receive the OPRM request.</u></p>" & _
        FollowUp & _
        "<p>Thank you,</p>" & _
		"<p>" & EDRepresenative(0) & "<br>" & _
		"PROJECT SPECIALIST DCV ENGAGEMENT DESK <br>" & _
		".:|:.:|:.  Cisco | Data Center & Virtualization Practice | Solutions Delivery Management Team <br>" & _
		EDRepresenative(1) & " | " & EDRepresenative(2) & "<br>" & _
        "<p>-----Original Message-----" & _
        objItem.HTMLBody & "</p>"
        
        With objReply
            
            ' Resolve Emails (same as hitting "Check Names" button
            Call .Recipients.ResolveAll
            
            ' Show the message
            Call .Display
        End With
        
        With objFWD
            
            ' Resolve Emails (same as hitting "Check Names" button
            Call .Recipients.ResolveAll
            
            ' Show the message
            Call .Display
        End With
            
    End If
    
    Set objReply = Nothing
    Set objFWD = Nothing
    Set objItem = Nothing
    
    
ErrHandler:     If Err.Number = 91 Then
                    MsgBox ("A staffing owner name was not provided, so no forward email will be automatically generated")
                End If
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Simple function to get the current Mail items based on the Explorer   '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application
         
    Set objApp = CreateObject("Outlook.Application")
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
        Case Else
            ' anything else will result in an error, which is
            ' why we have the error handler above
    End Select
     
    Set objApp = Nothing
End Function