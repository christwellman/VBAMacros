'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This macro creates a reply to an email to specific parties with   '
' specific attachments, and asks for the staffing owner information '
' It will then create another message to teh staffing owner with    '
' the original request information so that they can submit the      '
' requst                                                            '
'                                                                   '
'        !!!Relies on the GetCurrentItem() Function!!!              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub OPRMRejectionReply()

On Error GoTo ErrHandler
    
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
                "know AS is now using OPRM for project resourcing, so we take no further action until a staffing requirement from OPRM reaches us.</p>" & _
                "<p>Please use the information below as needed for completing the request in the system." & _
                "<p>Thank you,</p>" & _
                "<p>CHRIS TWELLMAN <br>" & _
                "PROJECT SPECIALIST DCV ENGAGEMENT DESK <br>" & _
                ".:|:.:|:.  Cisco | Data Center & Virtualization Practice | Solutions Delivery Management Team <br>" & _
                "ctwellma@cisco.com | +1 919 392 6154" & "<br>" & _
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
        objReply.Attachments.Add ("C:\Users\ctwellma\AppData\Roaming\Microsoft\Templates\Resource Management Initiative - Engaging with AS Practices.msg")
        
        
        'Reply Msg Body Composition:
        objReply.HTMLBody = "<p>Hi " & Name(0) & "," & "<br></p>" & _
        "<p>Thank you for submitting your request below to the DCV Engagement Desk.  However, as of 9/17/12 the Americas Enterprise, " & _
        "GET, Public Sector, Architectures, Advisory and GSP teams are using Oracle Projects Resource Management (OPRM) to centrally manage" & _
        " resources, and staffing requests should originate from the segment delivery team for this account.</p>" & _
        "<p>We'll forward this on to the responsible person (ex. Enterprise, Theater DM) and CC you, <u>but you are responsible for following " & _
        "up with that person to ensure the OPRM staffing process is utilized.  We will not assign resources until we receive the OPRM request.</u></p>" & _
        FollowUp & _
        "<p>Thank you,</p>" & _
        "<p>CHRIS TWELLMAN <br>" & _
        "PROJECT SPECIALIST DCV ENGAGEMENT DESK <br>" & _
        ".:|:.:|:.  Cisco | Data Center & Virtualization Practice | Solutions Delivery Management Team <br>" & _
        "ctwellma@cisco.com | +1 919 392 6154" & "<br>" & _
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