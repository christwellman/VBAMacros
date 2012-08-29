VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} projectInfo 
   Caption         =   "Project Information"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10560
   OleObjectBlob   =   "projectInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "projectInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub AddNotesButton_Click()
    Dim NewNote As String
    Dim OldNote As String
    Dim Update As String
    Dim Complete As String
    
    'OldDeskNotes = Me.NotesBox.Value
    
    OldNote = Me.NotesBox.Value
    'Me.NotesBox.AutoWordSelect
    
    
    'Create Timestamp
     vNow = Now()
     vMthStr = CStr(Month(vNow))
     vDayStr = CStr(Day(vNow))
     vHourStr = CStr(Hour(vNow))
     vMinuteStr = CStr(Minute(vNow))
    'Add leading zeroes to month, day, hour, minutes
     If Len(vMthStr) = 1 Then
        vMthStr = "0" & vMthStr
     End If
     If Len(vDayStr) = 1 Then
        vDayStr = "0" & vDayStr
     End If
     If Len(vHourStr) = 1 Then
        vHourStr = "0" & vHourStr
     End If
     If Len(vMinuteStr) = 1 Then
        vMinuteStr = "0" & vMinuteStr
     End If
    'Get date string in yyyymmddhhnn format.
     vDateStr = vHourStr & ":" & vMinuteStr & " " & vMthStr & "/" & vDayStr & "/" & Year(vNow)
            
    Update = InputBox(Prompt:="Please provide you update below:", Default:="")
        'Timestamp the note
        NewNote = (vDateStr & " - " & Update)
        Me.NotesBox.Value = (NewNote & vbCrLf & OldNote)
        Complete = (NewNote & vbCrLf & OldNote)
        Me.NotesBox.Value = Complete
        
     
    
End Sub

Private Sub AddUpdateButton_Click()
Dim wks As Excel.Worksheet
'appExcel.Application.Visible = True
ActiveWorkbook.Sheets("Project Pipeline").Activate
Range("A2").Select

'check for a PID number
'If Trim(Me.PIDBox.Value) = "" Then
'  Me.PIDBox.SetFocus
'  MsgBox "Please enter a PID"
'  Exit Sub
'End If


    Do
        
    If IsEmpty(ActiveCell) = False Then
        ActiveCell.Offset(1, 4).Select
    End If
    Loop Until IsEmpty(ActiveCell) = True
    
    
    'Update the datasheet
    ActiveCell.Offset(0, 4).Value = Me.requestBox.Value
    ActiveCell.Offset(0, 3).Value = Me.submittedBox.Value
    ActiveCell.Offset(0, 17).Value = Me.RequestTypeCombo.Value
    ActiveCell.Offset(0, 33).Value = Me.SegmentCombo.Value
    ActiveCell.Offset(0, 8).Value = Me.NotesBox.Value
    ActiveCell.Offset(0, 7).Value = Me.CustomerNameBox.Value
    ActiveCell.Offset(0, 33).Value = Me.CustomerContactBox.Value
    ActiveCell.Offset(0, 12).Value = Me.CityBox.Value
    'ActiveCell.Offset(0, 13).Value = Me.StateBox.Value
    ActiveCell.Offset(0, 43).Value = Me.SalesContactBox.Value
    ActiveCell.Offset(0, 26).Value = Me.ProjectNameBox.Value
    ActiveCell.Offset(0, 10).Value = Me.TechnologyBox.Value
    ActiveCell.Offset(0, 5).Value = Me.PIDBox.Value
    ActiveCell.Offset(0, 6).Value = Me.StatusBox.Value
    ActiveCell.Offset(0, 13).Value = Me.StartDateBox.Value
    ActiveCell.Offset(0, 17).Value = Me.ProjectTypeBox.Value
    ActiveCell.Offset(0, 14).Value = Me.KickOffDateBox.Value
    ActiveCell.Offset(0, 23).Value = Me.ProjectDetailsBox.Value
    ActiveCell.Offset(0, 15).Value = Me.EndDateBox.Value
    ActiveCell.Offset(0, 18).Value = Me.WMBox.Value
    ActiveCell.Offset(0, 19).Value = Me.DCPMBOX.Value
    ActiveCell.Offset(0, 34).Value = Me.DCPMStatusBox.Value
    ActiveCell.Offset(0, 30).Value = Me.MarginAnalysisCheck.Value
    ActiveCell.Offset(0, 31).Value = Me.QuoteCheck.Value
    ActiveCell.Offset(0, 31).Value = Me.SOWCheck.Value
        
        
    If StatusBox.Value = "Cancelled" Then
        PIDCloseDate = InputBox(Prompt:="When was the PID Cancelled?", Default:=DateValue(Now))
        DeliveryCloseDate = PIDCloseDate
        
    ElseIf StatusBox.Value = "Closed" Then
        PIDCloseDate = InputBox(Prompt:="When was the PID Closed?", Default:=DateValue(Now))
        
    ElseIf StatusBox.Value = "Delivery Close" Then
        DeliveryCloseDate = InputBox(Prompt:="When was the PID moved to Delivery Close?", Default:=DateValue(Now))
        
    End If
    
    ActiveCell.Offset(0, 38).Value = PIDCloseDate
    ActiveCell.Offset(0, 39).Value = DeliveryCloseDate
    ActiveCell.Offset(0, 35).Value = Me.PMAssignedBox.Value
    ActiveCell.Offset(0, 34).Value = Me.FollowUpBox.Value
    ActiveCell.Offset(0, 8).Value = Me.ServiceBox.Value

    'StatusBox.Value = ""
    Range("A2").Select
    
    Unload Me
    
End Sub

Private Sub CancelButton_Click()
  Unload Me
End Sub



Private Sub ProjectTypeBox_Change()
    'initialize set all checkboxes available
    MarginAnalysisCheck.Enabled = True
    SOWCheck.Enabled = True
    QuoteCheck.Enabled = True

    If ProjectTypeBox.Value = "Subscription" Then
        MarginAnalysisCheck.Enabled = False
        SOWCheck.Enabled = False
    ElseIf ProjectTypeBox.Value = "Transaction" Then
        QuoteCheck.Enabled = False
    ElseIf ProjectTypeBox.Value = "Fixed" Then
        MarginAnalysisCheck.Enabled = False
        SOWCheck.Enabled = False
        QuoteCheck.Enabled = False
    End If
    End Sub

Private Sub QuoteCheck_Click()

End Sub

Private Sub RequestTypeCombo_Change()

End Sub

Private Sub ServiceBox_Change()

End Sub

Private Sub SOWCheck_Click()

End Sub

Sub StatusBox_Change()

End Sub

Private Sub submittedBox_Change()

End Sub

Private Sub UserForm_Initialize()

PrevButton.Visible = False
NextButton.Visible = False
SearchButton.Visible = False


    
    requestBox.Text = ""
    submittedBox.Text = ""
    NotesBox.Text = ""
    CustomerNameBox.Text = ""
    CustomerContactBox.Text = ""
    CityBox.Text = ""
    SalesContactBox.Text = ""
    ProjectNameBox.Text = ""
    PIDBox.Text = ""
    StatusBox.Text = ""
    StartDateBox.Text = ""
    ProjectTypeBox.Text = ""
    KickOffDateBox.Text = ""
    ProjectDetailsBox.Text = ""
    EndDateBox.Text = ""
    MarginAnalysisCheck.Value = False
    QuoteCheck.Value = False
    SOWCheck.Value = False


'Request Cobmobox
RequestTypeCombo.AddItem "Scoping"
RequestTypeCombo.AddItem "Presales"
RequestTypeCombo.AddItem "Delivery"
RequestTypeCombo.AddItem "SME"
RequestTypeCombo.Value = ""

'''Segment Info
SegmentCombo.AddItem "Global Enterprise"
SegmentCombo.AddItem "US Ent, Comm, Canada"
SegmentCombo.AddItem "US Public Sector"
SegmentCombo.AddItem "US Service Provider"
SegmentCombo.Value = ""

'''Theater


'''Technology
TechnologyBox.AddItem "CIAC"
TechnologyBox.AddItem "UCS"
TechnologyBox.AddItem "DCN"
TechnologyBox.AddItem "SAN"
TechnologyBox.Value = ""

'''Services
ServiceBox.AddItem "ACE Planning and Design Service"
ServiceBox.AddItem "ACE Planning and Design Service"
ServiceBox.AddItem "ASF-ULT2-UCS-AA - Architecture Assessment"
ServiceBox.AddItem "ASF-ULT2-UCS-ADS - UCS Accelerated Deployment"
ServiceBox.AddItem "ASF-ULT2-UCS-OES - UCS Startup Accelerator"
ServiceBox.AddItem "ASF-ULT2-UCS-PP - UCS Pre-Production Pilot"
ServiceBox.AddItem "ASF-ULT2-UCS-VA - UCS Virtualization Assessment"
ServiceBox.AddItem "ASF-ULT2-UCS-VAS - UCS Virtualization Accelerator"
ServiceBox.AddItem "CIAC Deployment"
ServiceBox.AddItem "CIAC PDI"
ServiceBox.AddItem "Cisco Intelligent Automation Consulting Support Service"
ServiceBox.AddItem "Cisco Network Operations Automation Service"
ServiceBox.AddItem "Cisco Security Architecture Assessment Service"
ServiceBox.AddItem "Cisco Security Posture Assessment Service"
ServiceBox.AddItem "Cisco Tidal Deployment Service for Tidal Enterprise Scheduler"
ServiceBox.AddItem "Cisco Tidal Deployment Service for Tidal Enterprise Scheduler"
ServiceBox.AddItem "Cisco Tidal Deployment Service for Tidal Enterprise Transporter"
ServiceBox.AddItem "Cisco Tidal Solution Connector Development and Integration Service"
ServiceBox.AddItem "CON-AS-DCN-PD - Subscription Data Center Planning and Design Services"
ServiceBox.AddItem "CON-AS-IPC-PD  -UCS Plan and Design Subscription Services"
ServiceBox.AddItem "CON-NSST-1 - Network Optimization Renewal"
ServiceBox.AddItem "Data Center Health Check"
ServiceBox.AddItem "Data Center Migration"
ServiceBox.AddItem "Data Center Networking Planning and Design Service"
ServiceBox.AddItem "Data Center Value Analysis and Strategy and Architecture"
ServiceBox.AddItem "DCN-OPT UCS - Unified Computing Systems Optimization  "
ServiceBox.AddItem "Director Class SAN Planning and Design"
ServiceBox.AddItem "DMM Enablement"
ServiceBox.AddItem "eCDS PDI Services"
ServiceBox.AddItem "End-to-End Data Center Assessments and Design Service"
ServiceBox.AddItem "High Level Operation Assessment"
ServiceBox.AddItem "Knowledge Management"
ServiceBox.AddItem "MDS Install"
ServiceBox.AddItem "Network Optimization Service"
ServiceBox.AddItem "Nexus HLD / LLD"
ServiceBox.AddItem "Nexus Optimization"
ServiceBox.AddItem "Nexus PDI"
ServiceBox.AddItem "Nexus Planning and Design Service"
ServiceBox.AddItem "Nexus SME"
ServiceBox.AddItem "Presales"
ServiceBox.AddItem "SAN"
ServiceBox.AddItem "Tidal / NewScale"
ServiceBox.AddItem "UC on UCS"
ServiceBox.AddItem "UCS Consulting"
ServiceBox.AddItem "UCS POC"
ServiceBox.AddItem "Unified Computing Systems Planning and Design Service"
ServiceBox.AddItem "V Block"
ServiceBox.AddItem "VDI Workshop"
ServiceBox.AddItem "WAAS Assessment Services"
ServiceBox.AddItem "WAAS LLD"
ServiceBox.AddItem "WAAS Wide Area Application Services Planning and Design"
ServiceBox.AddItem "AS_DCN_CNSLT"
ServiceBox.Value = ""

'''States
'StateBox.AddItem ("AL")
'StateBox.AddItem ("AK")
'StateBox.AddItem ("AZ")
'StateBox.AddItem ("AR")
'StateBox.AddItem ("CA")
'StateBox.AddItem ("CO")
'StateBox.AddItem ("CT")
'StateBox.AddItem ("DE")
'StateBox.AddItem ("FL")
'StateBox.AddItem ("GA")
'StateBox.AddItem ("HI")
'StateBox.AddItem ("ID")
'StateBox.AddItem ("IL")
'StateBox.AddItem ("IN")
'StateBox.AddItem ("IA")
'StateBox.AddItem ("KS")
'StateBox.AddItem ("KY")
'StateBox.AddItem ("LA")
'StateBox.AddItem ("ME")
'StateBox.AddItem ("MD")
'StateBox.AddItem ("MA")
'StateBox.AddItem ("MI")
'StateBox.AddItem ("MN")
'StateBox.AddItem ("MS")
'StateBox.AddItem ("MO")
'StateBox.AddItem ("MT")
'StateBox.AddItem ("NE")
'StateBox.AddItem ("NV")
'StateBox.AddItem ("NH")
'StateBox.AddItem ("NJ")
'StateBox.AddItem ("NM")
'StateBox.AddItem ("NY")
'StateBox.AddItem ("NC")
'StateBox.AddItem ("ND")
'StateBox.AddItem ("OH")
'StateBox.AddItem ("OK")
'StateBox.AddItem ("OR")
'StateBox.AddItem ("PA")
'StateBox.AddItem ("RI")
'StateBox.AddItem ("SC")
'StateBox.AddItem ("SD")
'StateBox.AddItem ("TN")
'StateBox.AddItem ("TX")
'StateBox.AddItem ("UT")
'StateBox.AddItem ("VT")
'StateBox.AddItem ("VA")
'StateBox.AddItem ("WA")
'StateBox.AddItem ("WV")
'StateBox.AddItem ("WI")
'StateBox.AddItem ("WY")
'StateBox.AddItem ("??")
'StateBox.Value = ""

'''PID Status
StatusBox.AddItem "Pipeline"
StatusBox.AddItem "Presales"
StatusBox.AddItem "Active"
StatusBox.AddItem "Cancelled"
StatusBox.AddItem "Closed"
StatusBox.AddItem "Delivery Close"
StatusBox.AddItem "On Hold"
StatusBox.AddItem "Not Available"
StatusBox.Value = ""


'''Project Type
ProjectTypeBox.AddItem "Subscription"
ProjectTypeBox.AddItem "Transaction"
ProjectTypeBox.AddItem "AS Fixed"
ProjectTypeBox.AddItem "Internal"
ProjectTypeBox.AddItem "CAP"
ProjectTypeBox.AddItem "Unknown"
ProjectTypeBox.Value = ""


'''DCPM Status Box
DCPMStatusBox.AddItem "Technical Close"
DCPMStatusBox.AddItem "Reroute"
DCPMStatusBox.AddItem "Radar"
DCPMStatusBox.AddItem "New"
DCPMStatusBox.AddItem "In Progress"
DCPMStatusBox.AddItem "Hold"
DCPMStatusBox.AddItem "Delivery Complete"
DCPMStatusBox.AddItem "Deliverable Upload Audit"
DCPMStatusBox.AddItem "Complete - Presales Only"
DCPMStatusBox.AddItem "Complete - Handled out of Practice"
DCPMStatusBox.AddItem "Cold"
DCPMStatusBox.AddItem "Closed"
DCPMStatusBox.AddItem "Cancelled"
DCPMStatusBox.AddItem "Backlog"
DCPMStatusBox.AddItem "Duplicate Request"
DCPMStatusBox.Value = ""


'''DCPM Assigned
DCPMBOX.AddItem "Bill Falk"
DCPMBOX.AddItem "Bill McGillick"
DCPMBOX.AddItem "Birhanie Robinson"
DCPMBOX.AddItem "Brendon Boden"
DCPMBOX.AddItem "Caren Woodruff"
DCPMBOX.AddItem "Carmen Medeiros"
DCPMBOX.AddItem "Dale Singh"
DCPMBOX.AddItem "Dan Allison"
DCPMBOX.AddItem "Dan Mathisen"
DCPMBOX.AddItem "Daniel Brienen"
DCPMBOX.AddItem "Darryl Wortham"
DCPMBOX.AddItem "David Bregman"
DCPMBOX.AddItem "Farooq Raza"
DCPMBOX.AddItem "Felisha Spivey"
DCPMBOX.AddItem "Gabriel Castillo"
DCPMBOX.AddItem "Jerry Smessaert"
DCPMBOX.AddItem "Katie Waddell"
DCPMBOX.AddItem "Mark Cooper"
DCPMBOX.AddItem "Michael Dyer"
DCPMBOX.AddItem "Randy Bentele"
DCPMBOX.AddItem "Rick Scouler"
DCPMBOX.AddItem "Rob Batten"
DCPMBOX.AddItem "Robert Batten"
DCPMBOX.AddItem "Roger Ellsworth"
DCPMBOX.AddItem "Russ Brockman"
DCPMBOX.AddItem "Scott Milesnick"
DCPMBOX.AddItem "Shahir Ahang"
DCPMBOX.AddItem "Troy Tanaka"
DCPMBOX.AddItem "Warren Beck"
DCPMBOX.AddItem "Wasif Kazmi"
DCPMBOX.AddItem "Pending Assignment - Brendon"
DCPMBOX.AddItem "Pending Assignment - David"
DCPMBOX.AddItem "Pending Assignment - Erik"
DCPMBOX.AddItem "Outside of Practice"
DCPMBOX.AddItem "Needs DCPM"
DCPMBOX.Value = ""


''' Work Manager
WMBox.AddItem "Will Hill"
WMBox.AddItem "Chris Twellman"
WMBox.AddItem "Karla Puga"
WMBox.AddItem "Carlos Bolanos"
WMBox.AddItem "Ernesto Escobar"
WMBox.AddItem "Outside Practice"
WMBox.AddItem "Theater PM"
WMBox.Value = ""

End Sub

Private Sub GetData()
Dim Selection As Range
Set Selection = A3


MsgBox Selection


    'Me.requestBox.Value = ActiveCell.Offset(0, 1).Value
    'ActiveCell.Offset(0, 0).Value = Me.submittedBox.Value
    'ActiveCell.Offset(0, 2).Value = Me.RequestTypeCombo.Value
    'ActiveCell.Offset(0, 23).Value = Me.SegmentCombo.Value
    'ActiveCell.Offset(0, 27).Value = Me.NotesBox.Value
    'ActiveCell.Offset(0, 3).Value = Me.CustomerNameBox.Value
    'ActiveCell.Offset(0, 8).Value = Me.CustomerContactBox.Value
    'ActiveCell.Offset(0, 6).Value = Me.CityBox.Value
    'ActiveCell.Offset(0, 7).Value = Me.StateBox.Value
    'ActiveCell.Offset(0, 28).Value = Me.SalesContactBox.Value
    'ActiveCell.Offset(0, 12).Value = Me.ProjectNameBox.Value
    'ActiveCell.Offset(0, 5).Value = Me.TechnologyBox.Value
    'ActiveCell.Offset(0, 13).Value = Me.PIDBox.Value
    'ActiveCell.Offset(0, 16).Value = Me.StatusBox.Value
    'ActiveCell.Offset(0, 17).Value = Me.StartDateBox.Value
    'ActiveCell.Offset(0, 20).Value = Me.ProjectTypeBox.Value
    'ActiveCell.Offset(0, 18).Value = Me.KickOffDateBox.Value
    'ActiveCell.Offset(0, 9).Value = Me.ProjectDetailsBox.Value
    'ActiveCell.Offset(0, 19).Value = Me.EndDateBox.Value
    'ActiveCell.Offset(0, 31).Value = Me.WMBox.Value
    'ActiveCell.Offset(0, 29).Value = Me.DCPMBOX.Value
    'ActiveCell.Offset(0, 24).Value = Me.DCPMStatusBox.Value
    'ActiveCell.Offset(0, 21).Value = Me.MarginAnalysisCheck.Value
    'ActiveCell.Offset(0, 22).Value = Me.QuoteCheck.Value
    'ActiveCell.Offset(0, 22).Value = Me.SOWCheck.Value

End Sub

Private Sub WMBox_Change()

End Sub
