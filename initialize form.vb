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
StateBox.AddItem ("AL")
StateBox.AddItem ("AK")
StateBox.AddItem ("AZ")
StateBox.AddItem ("AR")
StateBox.AddItem ("CA")
StateBox.AddItem ("CO")
StateBox.AddItem ("CT")
StateBox.AddItem ("DE")
StateBox.AddItem ("FL")
StateBox.AddItem ("GA")
StateBox.AddItem ("HI")
StateBox.AddItem ("ID")
StateBox.AddItem ("IL")
StateBox.AddItem ("IN")
StateBox.AddItem ("IA")
StateBox.AddItem ("KS")
StateBox.AddItem ("KY")
StateBox.AddItem ("LA")
StateBox.AddItem ("ME")
StateBox.AddItem ("MD")
StateBox.AddItem ("MA")
StateBox.AddItem ("MI")
StateBox.AddItem ("MN")
StateBox.AddItem ("MS")
StateBox.AddItem ("MO")
StateBox.AddItem ("MT")
StateBox.AddItem ("NE")
StateBox.AddItem ("NV")
StateBox.AddItem ("NH")
StateBox.AddItem ("NJ")
StateBox.AddItem ("NM")
StateBox.AddItem ("NY")
StateBox.AddItem ("NC")
StateBox.AddItem ("ND")
StateBox.AddItem ("OH")
StateBox.AddItem ("OK")
StateBox.AddItem ("OR")
StateBox.AddItem ("PA")
StateBox.AddItem ("RI")
StateBox.AddItem ("SC")
StateBox.AddItem ("SD")
StateBox.AddItem ("TN")
StateBox.AddItem ("TX")
StateBox.AddItem ("UT")
StateBox.AddItem ("VT")
StateBox.AddItem ("VA")
StateBox.AddItem ("WA")
StateBox.AddItem ("WV")
StateBox.AddItem ("WI")
StateBox.AddItem ("WY")
StateBox.AddItem ("??")
StateBox.Value = ""

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
