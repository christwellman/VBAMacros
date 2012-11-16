'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This macro creates a reply to an email to specific parties to     '
' follow up on cloud workshop inquiries                             '
'                                                                   '
'        !!!Relies on the GetCurrentItem() Function!!!              '
'        !!!Relies on the ParseTextLinePair() Function!!!           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CloudWorkshopFollowUp()

'On Error GoTo ErrHandler
    
    Dim objItem As Object
    Dim Name As Variant
    Dim MailAttach As MailItem
    Dim SalesContact As String
    Dim objReply As MailItem
    Dim TestSubject As String
    
    'Test the subject of the message against the string below to see if it qualifies for this macro
    'TestSubject = "New AS-Fixed project has been created and is ready for delivery"
        
    Set objItem = GetCurrentItem()
    If objItem.Class = olMail Then
        
        'Test to make sure this is a workflow mailer as fixed notification
        'If ParseTextLinePair(objItem.Body, "Request Type:") = "Cloud Workshop" Then
        
            ' find the requestor address
            subject = objItem.subject
                   
            ' create the reply, add the address and display
            Set objReply = CreateItemFromTemplate("C:\Users\ctwellma\AppData\Roaming\Microsoft\Templates\CloudWorkshopResponse.oft")
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