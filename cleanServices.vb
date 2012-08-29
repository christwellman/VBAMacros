Sub CleanServices2()
    Dim sh As Worksheet
    Dim Prompt As String
    Dim DestSh As Worksheet
    Dim Firstrow As Integer
    Dim Last As Long
    Dim shLast As Long
    Dim checkrange As Range
    Dim checkarea As Range
    Dim checkrow As Range
    Dim StartRow As Long

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    

    ' Delete unneccssary sheets.
    Application.DisplayAlerts = False
    On Error Resume Next
    'ActiveWorkbook.Worksheets("LOVs").Delete
    'ActiveWorkbook.Worksheets("Complete - Presales - Scoping").Delete
    'ActiveWorkbook.Worksheets("Cold Projects").Delete
    'ActiveWorkbook.Worksheets("Closed Projects").Delete
    'ActiveWorkbook.Worksheets("Project Tracking - GPAGE").Delete
    'ActiveWorkbook.Worksheets("Archive Desk Complete").Delete
    'ActiveWorkbook.Worksheets("Archive Cold Projects").Delete
    'ActiveWorkbook.Worksheets("Archive Closed Projects").Delete
    'ActiveWorkbook.Worksheets("Project Pipeline").Delete
    'ActiveWorkbook.Worksheets("New WM Mapping").Delete
    
    On Error GoTo 0
    

    ' Choose Sheet to Clean
    Prompt = ("Which worksheet do you want to clean?")
    DynamicForm.PromptLabel.Caption = Prompt
    DynamicForm.DynamicComboBox.RowSource = "Sheets"

    DynamicForm.Show
    
    
    'destshname = InputBox(Prompt:="Which worksheet do you want to clean?", Title:="Clean which sheet?", Default:="")
    Set DestSh = ActiveWorkbook.Worksheets(CheckValue)
    DynamicForm.DynamicComboBox = ""
    
    ' Fill in the start row.--- Not Used See below check to get seccond row of data
    'If DestSh.Name = "Project Pipeline" Then
        StartRow = InputBox(Prompt:="Which row would you like to start checking at?", Title:="Start Row?", Default:=2)
    'Else
        'StartRow = 2
    'End If
    Last = Lastrow(DestSh)
    Firstrow = StartRow
    
    Set checkrange = Range(Cells(Firstrow, "A"), Cells(Last, "AP"))
    'MsgBox ("Range " & checkrange.Address)

    
    With DestSh
        .DisplayPageBreaks = False
 
                For Each checkrow In checkrange.Rows
                    
                    ' Service Types Rename
                    If .Cells(checkrow.Row, "L") = "3 Month UCS Optimization" Then
						.cells(checkrow.row,"L").value = "Unified Computing Optimization Service"
					End IF

					If .Cells(checkrow.Row, "L") = "Accelerated Deployment" Then
						.cells(checkrow.row,"L").value = "Fixed-price Cisco Unified Computing Accelerated Deployment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Accelerated Design/Deployment services for SAN Fabric" Then
						.cells(checkrow.row, "L").value = "Unified Computing Accelerated Deployment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "ACE PD" Then
						.cells(checkrow.row, "L").value = "Application Control Engine Planning and Design Service (ACE)"
					End IF
					If .Cells(checkrow.Row, "L") = "ACE PDI" Then
						.cells(checkrow.row, "L").value = "Application Control Engine Planning and Design Service (ACE)"
					End IF
					If .Cells(checkrow.Row, "L") = "ADS" Then
						.cells(checkrow.row, "L").value = "Application Dependency Mapping Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Architectural Value Analysis" Then
						.cells(checkrow.row, "L").value = "Architecture Value Analysis Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Architecture Assessment" Then
						.cells(checkrow.row, "L").value = "Architecture Assessment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "ASF-ULT2-UCS-AA - Architecture Assessment" Then
						.cells(checkrow.row, "L").value = "Architecture Assessment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "ASF-ULT2-UCS-ADS - UCS Accelerated Deployment" Then
						.cells(checkrow.row, "L").value = "Fixed-price Cisco Unified Computing Accelerated Deployment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "ASF-ULT2-UCS-PP" Then
						.cells(checkrow.row, "L").value = "Fixed-price Cisco Unified Computing Preproduction Pilot Service"
					End IF
					If .Cells(checkrow.Row, "L") = "ASF-ULT2-UCS-PP - UCS Pre-Production Pilot" Then
						.cells(checkrow.row, "L").value = "Fixed-price Cisco Unified Computing Preproduction Pilot Service"
					End IF
					If .Cells(checkrow.Row, "L") = "CCP" Then
						.cells(checkrow.row, "L").value = "IA: Cloud Portal"
					End IF
					If .Cells(checkrow.Row, "L") = "Cisco Tidal Deployment Service for Tidal Enterprise Scheduler" Then
						.cells(checkrow.row, "L").value = "IA: Enterprise Scheduler"
					End IF
					If .Cells(checkrow.Row, "L") = "Cisco Workplace Portal" Then
						.cells(checkrow.row, "L").value = "IA: Workplace"
					End IF
					If .Cells(checkrow.Row, "L") = "Cisco Workplace Portal enhancements" Then
						.cells(checkrow.row, "L").value = "IA: Workplace"
					End IF
					If .Cells(checkrow.Row, "L") = "CON-AS-DCN-PD - Subscription Data Center Planning and Design Services" Then
						.cells(checkrow.row, "L").value = "Nexus Planning and Design Service"
					End IF
					If .Cells(checkrow.Row, "L") = "CON-AS-UCS-PD: UCS Quote - PDI" Then
						.cells(checkrow.row, "L").value = "Unified Computing Planning, Design and Implementation Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Data Center Architecture Services " Then
						.cells(checkrow.row, "L").value = "Architecture Assessment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Data Center Assesment" Then
						.cells(checkrow.row, "L").value = "Architecture Assessment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Data Center Assessment " Then
						.cells(checkrow.row, "L").value = "Architecture Assessment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Data Center Assment" Then
						.cells(checkrow.row, "L").value = "Architecture Assessment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Data Center Migration" Then
						.cells(checkrow.row, "L").value = "Data Center Migration Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Data Center Network Assessment" Then	
					.cells(checkrow.row, "L").value = "Data Center Network Assessment"
					End IF
					If .Cells(checkrow.Row, "L") = "Data Center Network Assessment High-Level and Low-Level Design Review Knowledge Transfer (remote)" Then
						.cells(checkrow.row, "L").value = "Data Center Network Assessment"
					End If
					If .Cells(checkrow.Row, "L") = "Data Center Networking Assessment" Then
						.cells(checkrow.row, "L").value = "Data Center Network Assessment"
					End IF
					If .Cells(checkrow.Row, "L") = "Data Center Optimization" Then
						.cells(checkrow.row, "L").value = "Data Center Optimization Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Data Center Value Analysis and Strategy and Architecture" Then
						.cells(checkrow.row, "L").value = "Architecture Value Analysis Service"
					End IF
					If .Cells(checkrow.Row, "L") = "DCN Optimization" Then
						.cells(checkrow.row, "L").value = "Data Center Optimization Service"
					End IF
					If .Cells(checkrow.Row, "L") = "DCN PD" Then
						.cells(checkrow.row, "L").value = "Data Center Networking Planning and Design Service"
					End IF
					If .Cells(checkrow.Row, "L") = "DCN Scoping" Then
						.cells(checkrow.row, "L").value = "DCN Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus implementation" Then
						.cells(checkrow.row, "L").value = "Nexus implementation"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus Implementation Service for up to 8 Cisco Nexus 7000 " Then
						.cells(checkrow.row, "L").value = "Nexus implementation"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus PD" Then
						.cells(checkrow.row, "L").value = "Nexus Planning and Design Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus Planning and Design" Then
						.cells(checkrow.row, "L").value = "Nexus Planning and Design Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Planning and Design services for Nexus 1000v" Then
						.cells(checkrow.row, "L").value = "Nexus Planning and Design Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Preproduction pilot" Then
						.cells(checkrow.row, "L").value = "Fixed-price Cisco Unified Computing Preproduction Pilot Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Presales" Then
						.cells(checkrow.row, "L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "Presales " Then
						.cells(checkrow.row, "L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "Presales/Scoping" Then
						.cells(checkrow.row, "L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "PresalesRFP" Then
						.cells(checkrow.row, "L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "PresalesSME" Then
						.cells(checkrow.row, "L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "PresalesUCS POC" Then
						.cells(checkrow.row, "L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "Remote Presalesengament management for large Nexus,UCS,VDI & ASA Project -- remote meetings" Then
						.cells(checkrow.row, "L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "RFI" Then
						.cells(checkrow.row, "L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "RFP" Then
						.cells(checkrow.row, "L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "SAN Assessment" Then
						.cells(checkrow.row, "L").value = "SAN Health Check Service"
					End IF
					If .Cells(checkrow.Row, "L") = "SAN Health Check" Then
						.cells(checkrow.row, "L").value = "SAN Health Check Service"
					End IF
					If .Cells(checkrow.Row, "L") = "SAN Healthcheck" Then
						.cells(checkrow.row, "L").value = "SAN Health Check Service"
					End IF
					If .Cells(checkrow.Row, "L") = "SAN Optimization" Then
						.cells(checkrow.row, "L").value = "SAN Optimization"
					End IF
					If .Cells(checkrow.Row, "L") = "SAN Scoping" Then
						.cells(checkrow.row, "L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "Scoping" Then
						.cells(checkrow.row, "L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "Scoping ADM" Then
						.cells(checkrow.row, "L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "Scoping Data Center Migration - Phase 2" Then
						.cells(checkrow.row, "L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "SME for RFP" Then
						.cells(checkrow.row, "L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "Solaris Migration and Private IaaS Cloud Acceleration Services" Then
						.cells(checkrow.row, "L").value = "Cloud Enablement Services for Building IaaS Clouds"
					End IF
					If .Cells(checkrow.Row, "L") = "SOW Based Accelerated Deployment" Then
						.cells(checkrow.row, "L").value = "Unified Computing Accelerated Deployment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Start-Up Accelerator" Then
						.cells(checkrow.row, "L").value = "Fixed-price Cisco Unified Computing Start-Up Accelerator Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Migration" Then
						.cells(checkrow.row, "L").value = "Unified Computing Migration and Transition Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Opportunity" Then
						.cells(checkrow.row, "L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Optimization" Then
						.cells(checkrow.row, "L").value = "Unified Computing Optimization Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Optimziation (3 Year)" Then
						.cells(checkrow.row, "L").value = "Unified Computing Optimization Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS PDI" Then
						.cells(checkrow.row, "L").value = "Unified Computing Planning, Design and Implementation Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS PPP" Then
						.cells(checkrow.row, "L").value = "Unified Computing Preproduction Pilot Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Pre-Production Pilot" Then
						.cells(checkrow.row, "L").value = "Unified Computing Preproduction Pilot Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Scoping" Then
						.cells(checkrow.row, "L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Start Up Accelerator" Then
						.cells(checkrow.row, "L").value = "Fixed-price Cisco Unified Computing Start-Up Accelerator Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Startup Accelerator" Then
						.cells(checkrow.row, "L").value = "Fixed-price Cisco Unified Computing Start-Up Accelerator Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Start-Up Accelerator" Then
						.cells(checkrow.row, "L").value = "Fixed-price Cisco Unified Computing Start-Up Accelerator Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Unified Computing Optimization Service" Then
						.cells(checkrow.row, "L").value = "Unified Computing Optimization Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Unified Fabric Optimization Service" Then
						.cells(checkrow.row, "L").value = "Unified Fabric Optimization"
					End IF
					If .Cells(checkrow.Row, "L") = "Unified Fabric Optimization Service  " Then
						.cells(checkrow.row, "L").value = "Unified Fabric Optimization"
					End IF
					If .Cells(checkrow.Row, "L") = "WAAS" Then
						.cells(checkrow.row, "L").value = "Wide Area Application Services Planning and Design Service (WAAS)"
					End IF
					'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					If .Cells(checkrow.Row, "L") = "ACE PD" Then
						.cells(checkrow.row,"L").value = "Application Control Engine Planning and Design Service (ACE)"
					End IF
					If .Cells(checkrow.Row, "L") = "ACE PDI" Then 
						.cells(checkrow.row,"L").value = "Application Control Engine Planning, Design and Implmentation Service (ACE)"
					End IF
					If .Cells(checkrow.Row, "L") = "ADS" Then 
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "ANS" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Architectural Value Analysis" Then
						.cells(checkrow.row,"L").value = "Architecture Value Analysis Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Architecture Evaluation" Then
						.cells(checkrow.row,"L").value = "Architecture Assessment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "AS DCN CNSLT" Then
						.cells(checkrow.row,"L").value = "AS DCN CNSLT"
					End IF
					If .Cells(checkrow.Row, "L") = "AS_DCN_CNSLT" Then
						.cells(checkrow.row,"L").value = "AS DCN CNSLT"
					End IF
					If .Cells(checkrow.Row, "L") = "AS-DCN-SME" Then
						.cells(checkrow.row,"L").value = "SME"
					End IF
					If .Cells(checkrow.Row, "L") = "ASF-ULT2-UCS-AA - Architecture Assessment" Then
						.cells(checkrow.row,"L").value = "Architecture Assessment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "AS-UCS-CNSLT - RISC Migration Acceleration Service" Then
						.cells(checkrow.row,"L").value = "Unified Computing Database Accelerator Service"
					End IF
					If .Cells(checkrow.Row, "L") = "BMC Deployment " Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "CAP" Then
						.cells(checkrow.row,"L").value = "CAP"
					End IF
					If .Cells(checkrow.Row, "L") = "CAP B440 Replacement " Then
						.cells(checkrow.row,"L").value = "CAP"
					End IF
					If .Cells(checkrow.Row, "L") = "CIAC" Then
						.cells(checkrow.row,"L").value = "CIAC"
					End IF
					If .Cells(checkrow.Row, "L") = "CIA-C" Then
						.cells(checkrow.row,"L").value = "CIAC"
					End IF
					If .Cells(checkrow.Row, "L") = "CIAC - Plan, Build, Design" Then
						.cells(checkrow.row,"L").value = "IA: Full Cloud"
					End IF
					If .Cells(checkrow.Row, "L") = "CIAC - SME" Then
						.cells(checkrow.row,"L").value = "SME"
					End IF
					If .Cells(checkrow.Row, "L") = "CIAC Cloud Builder Accelerator" Then
						.cells(checkrow.row,"L").value = "CIAC"
					End IF
					If .Cells(checkrow.Row, "L") = "CIAC Consultant" Then
						.cells(checkrow.row,"L").value = "SME"
					End IF
					If .Cells(checkrow.Row, "L") = "CIAC Deployment" Then
						.cells(checkrow.row,"L").value = "CIAC"
					End IF
					If .Cells(checkrow.Row, "L") = "CIAC Implementation" Then
						.cells(checkrow.row,"L").value = "CIAC"
					End IF
					If .Cells(checkrow.Row, "L") = "CIAC Implmentation" Then
						.cells(checkrow.row,"L").value = "CIAC"
					End IF
					If .Cells(checkrow.Row, "L") = "CIAC Integration" Then
						.cells(checkrow.row,"L").value = "CIAC"
					End IF
					If .Cells(checkrow.Row, "L") = "CIA-C PaaS" Then
						.cells(checkrow.row,"L").value = "CIAC"
					End IF
					If .Cells(checkrow.Row, "L") = "CIAC PDI" Then
						.cells(checkrow.row,"L").value = "IA: Full Cloud"
					End IF
					If .Cells(checkrow.Row, "L") = "CIAC PDI" Then
						.cells(checkrow.row,"L").value = "IA: Full Cloud"
					End IF
					If .Cells(checkrow.Row, "L") = "CIAC Plan and Build Services" Then
						.cells(checkrow.row,"L").value = "IA: Full Cloud"
					End IF
					If .Cells(checkrow.Row, "L") = "CIA-C POC" Then
						.cells(checkrow.row,"L").value = "CIAC POC"
					End IF
					If .Cells(checkrow.Row, "L") = "CIAC Presales" Then
						.cells(checkrow.row,"L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "CIAC Public" Then
						.cells(checkrow.row,"L").value = "CIAC"
					End IF
					If .Cells(checkrow.Row, "L") = "CIAC SME" Then
						.cells(checkrow.row,"L").value = "SME"
					End IF
					If .Cells(checkrow.Row, "L") = "CIA-C Start" Then
						.cells(checkrow.row,"L").value = "CIAC"
					End IF
					If .Cells(checkrow.Row, "L") = "CISCO IT OESE: PAAS TIDAL" Then
						.cells(checkrow.row,"L").value = "CIAC"
					End IF
					If .Cells(checkrow.Row, "L") = "Cisco Tidal Deployment Service for Tidal Enterprise Scheduler" Then
						.cells(checkrow.row,"L").value = "IA: Enterprise Scheduler"
					End IF
					If .Cells(checkrow.Row, "L") = "CITIES Express" Then
						.cells(checkrow.row,"L").value = "CIAC"
					End IF
					If .Cells(checkrow.Row, "L") = "Cloud" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Cloud Advisory Services" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Cloud build" Then
						.cells(checkrow.row,"L").value = "IA: Full Cloud"
					End IF
					If .Cells(checkrow.Row, "L") = "Cloud Builder Accelerator Program" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Cloud Computing Strategy Consulting Services " Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Cloud growth" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Cloud Portal Demo" Then
						.cells(checkrow.row,"L").value = "IA: Cloud Portal"
					End IF
					If .Cells(checkrow.Row, "L") = "Cloud Portal Plan & Build" Then
						.cells(checkrow.row,"L").value = "IA: Cloud Portal"
					End IF
					If .Cells(checkrow.Row, "L") = "CNOAS" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "CON-AS-DCN-PD - Subscription Data Center Planning and Design Services" Then
						.cells(checkrow.row,"L").value = "Data Center Networking Planning and Design Service"
					End IF
					If .Cells(checkrow.Row, "L") = "CON-AS-UCS-PD: UCS Quote - PDI" Then
						.cells(checkrow.row,"L").value = "Unified Computing Planning, Design and Implementation Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Consulting Assistance" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Data Center Assesment" Then
						.cells(checkrow.row,"L").value = "Data Center Facilities Assessment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Data Center Assessment " Then
						.cells(checkrow.row,"L").value = "Data Center Facilities Assessment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Data Center Assment" Then
						.cells(checkrow.row,"L").value = "Data Center Facilities Assessment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Data Center Migration" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Data Center Networking Assessment " Then
						.cells(checkrow.row,"L").value = "Data Center Network Assessment"
					End IF
					If .Cells(checkrow.Row, "L") = "Data Center Strategy and Architecture Blueprint" Then
						.cells(checkrow.row,"L").value = "Data Center Architecture Workshop"
					End IF
					If .Cells(checkrow.Row, "L") = "DCN Bundle PDI" Then
						.cells(checkrow.row,"L").value = "Data Center Networking Planning, Design and Implemntation Service"
					End IF
					If .Cells(checkrow.Row, "L") = "DCN Consulting" Then
						.cells(checkrow.row,"L").value = "AS DCN CNSLT"
					End IF
					If .Cells(checkrow.Row, "L") = "DCN DI" Then
						.cells(checkrow.row,"L").value = "Data Center Networking Planning, Design and Implemntation Service"
					End IF
					If .Cells(checkrow.Row, "L") = "DCN Optimization" Then
						.cells(checkrow.row,"L").value = "Data Center Optimization Service"
					End IF
					If .Cells(checkrow.Row, "L") = "DCN OTV" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "DCN PD" Then
						.cells(checkrow.row,"L").value = "Data Center Networking Planning and Design Service"
					End IF
					If .Cells(checkrow.Row, "L") = "DCN PDI" Then
						.cells(checkrow.row,"L").value = "Data Center Networking Planning, Design and Implemntation Service"
					End IF
					If .Cells(checkrow.Row, "L") = "DCN PDI" Then
						.cells(checkrow.row,"L").value = "Data Center Networking Planning, Design and Implemntation Service"
					End IF
					If .Cells(checkrow.Row, "L") = "DCN Scoping" Then
						.cells(checkrow.row,"L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "DCN SME" Then
						.cells(checkrow.row,"L").value = "SME"
					End IF
					If .Cells(checkrow.Row, "L") = "DCN SME" Then
						.cells(checkrow.row,"L").value = "SME"
					End IF
					If .Cells(checkrow.Row, "L") = "Demonstrator" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Disaster Recovery Gap Analysis" Then
						.cells(checkrow.row,"L").value = "Disaster Recovery Planning and Implementation"
					End IF
					If .Cells(checkrow.Row, "L") = "Disaster Recovery Planning and Implementation" Then
						.cells(checkrow.row,"L").value = "Disaster Recovery Planning and Implementation"
					End IF
					If .Cells(checkrow.Row, "L") = "eCDS PDI Services" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "eCDS PDI Services" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "GLaaS workshops - Phase1" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "HLD" Then
						.cells(checkrow.row,"L").value = "AS DCN CNSLT"
					End IF
					If .Cells(checkrow.Row, "L") = "LDAP/SSO customi" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "MDS Data Migration" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Multiple Servcies - UCS, DCN, TS" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "N1K SME" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "NewScale" Then
						.cells(checkrow.row,"L").value = "IA: Enterprise Scheduler"
					End IF
					If .Cells(checkrow.Row, "L") = "NewScale Enablement Services" Then
						.cells(checkrow.row,"L").value = "IA: Enterprise Scheduler"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus 1K install & config" Then
						.cells(checkrow.row,"L").value = "Nexus implementation"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus Implementation Service for up to 8 Cisco Nexus 7000 " Then
						.cells(checkrow.row,"L").value = "Nexus implementation"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus IPD" Then
						.cells(checkrow.row,"L").value = "Data Center Networking Planning, Design and Implemntation Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus LLD review" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus Migration" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus Optimization" Then
						.cells(checkrow.row,"L").value = "Data Center Optimization Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus Optimization" Then
						.cells(checkrow.row,"L").value = "Data Center Optimization Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus Optimization " Then
						.cells(checkrow.row,"L").value = "Data Center Optimization Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus PD" Then
						.cells(checkrow.row,"L").value = "Nexus Planning and Design Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus PDI" Then
						.cells(checkrow.row,"L").value = "Data Center Networking Planning, Design and Implemntation Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus PDI" Then
						.cells(checkrow.row,"L").value = "Data Center Networking Planning, Design and Implemntation Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus PDI" Then
						.cells(checkrow.row,"L").value = "Data Center Networking Planning, Design and Implemntation Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus PDI" Then
						.cells(checkrow.row,"L").value = "Data Center Networking Planning, Design and Implemntation Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus PDI" Then
						.cells(checkrow.row,"L").value = "Data Center Networking Planning, Design and Implemntation Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus PDI" Then
						.cells(checkrow.row,"L").value = "Data Center Networking Planning, Design and Implemntation Service"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus SME" Then
						.cells(checkrow.row,"L").value = "SME"
					End IF
					If .Cells(checkrow.Row, "L") = "Nexus SME" Then
						.cells(checkrow.row,"L").value = "SME"
					End IF
					If .Cells(checkrow.Row, "L") = "Not Booked 5/3 Nexus LLD" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "OTV SME" Then
						.cells(checkrow.row,"L").value = "SME"
					End IF
					If .Cells(checkrow.Row, "L") = "Phase 5/6" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Plan and Design Subscription Services - Wireless LAN" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Planning and Design services for Nexus 1000v" Then
						.cells(checkrow.row,"L").value = "Nexus Planning and Design Service"
					End IF
					If .Cells(checkrow.Row, "L") = "POC" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "POL - UCS POC" Then
						.cells(checkrow.row,"L").value = "UCS POC"
					End IF
					If .Cells(checkrow.Row, "L") = "Presales" Then
						.cells(checkrow.row,"L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "RC Implementation" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "RC Implementation" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "RISC" Then
						.cells(checkrow.row,"L").value = "Unified Computing Database Accelerator Service"
					End IF
					If .Cells(checkrow.Row, "L") = "RISC to UCS Migration" Then
						.cells(checkrow.row,"L").value = "Unified Computing Database Accelerator Service"
					End IF
					If .Cells(checkrow.Row, "L") = "RISC to UCS Migration Acceleration Service " Then
						.cells(checkrow.row,"L").value = "Unified Computing Database Accelerator Service"
					End IF
					If .Cells(checkrow.Row, "L") = "SAN" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "SAN Health Check" Then
						.cells(checkrow.row,"L").value = "SAN Health Check Service"
					End IF
					If .Cells(checkrow.Row, "L") = "SAN PD" Then
						.cells(checkrow.row,"L").value = "SAN Planning and Design Service"
					End IF
					If .Cells(checkrow.Row, "L") = "SAN PD" Then
						.cells(checkrow.row,"L").value = "SAN Planning and Design Service"
					End IF
					If .Cells(checkrow.Row, "L") = "SAN PDI" Then
						.cells(checkrow.row,"L").value = "SAN Planning, Design and Implementation Service"
					End IF
					If .Cells(checkrow.Row, "L") = "SAP HANA" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "SAP HANA on UCS" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Security" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "SME" Then
						.cells(checkrow.row,"L").value = "SME"
					End IF
					If .Cells(checkrow.Row, "L") = "SME for RFP" Then
						.cells(checkrow.row,"L").value = "Presales/Scoping"
					End IF
					If .Cells(checkrow.Row, "L") = "TES / SAP Deals" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Tidal" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Tidal / NewScale" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Tidal PDI" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Tidal Phase 2" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "UC on UCS" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS & DCN SME" Then
						.cells(checkrow.row,"L").value = "SME"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS / Oracle POC" Then
						.cells(checkrow.row,"L").value = "UCS POC"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Accelerated Deployment" Then
						.cells(checkrow.row,"L").value = "Fixed-price Cisco Unified Computing Accelerated Deployment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Accelerated Deployment" Then
						.cells(checkrow.row,"L").value = "Fixed-price Cisco Unified Computing Accelerated Deployment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Accelerated Deployment" Then
						.cells(checkrow.row,"L").value = "Fixed-price Cisco Unified Computing Accelerated Deployment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Accelerated Deployment" Then
						.cells(checkrow.row,"L").value = "Fixed-price Cisco Unified Computing Accelerated Deployment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Assessment and Planning" Then
						.cells(checkrow.row,"L").value = "Unified Computing Architecture Assessment Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS CAP" Then
						.cells(checkrow.row,"L").value = "CAP"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Escalation" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Lab Support" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Optimization" Then
						.cells(checkrow.row,"L").value = "Unified Computing Optimization Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS PD" Then
						.cells(checkrow.row,"L").value = "Unified Computing Planning and Design Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS POC" Then
						.cells(checkrow.row,"L").value = "UCS POC"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS POC" Then
						.cells(checkrow.row,"L").value = "UCS POC"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Pre-Production Pilot" Then
						.cells(checkrow.row,"L").value = "Fixed-price Cisco Unified Computing Preproduction Pilot Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS SAP HANA SME" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS SME" Then
						.cells(checkrow.row,"L").value = "SME"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS SME" Then
						.cells(checkrow.row,"L").value = "SME"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS SME" Then
						.cells(checkrow.row,"L").value = "SME"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Start Up Accelerator" Then
						.cells(checkrow.row,"L").value = "Fixed-price Cisco Unified Computing Start-Up Accelerator Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Startup Accelerator" Then
						.cells(checkrow.row,"L").value = "Fixed-price Cisco Unified Computing Start-Up Accelerator Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Start-Up Accelerator" Then
						.cells(checkrow.row,"L").value = "Fixed-price Cisco Unified Computing Start-Up Accelerator Service"
					End IF
					If .Cells(checkrow.Row, "L") = "UCS Workshop" Then
						.cells(checkrow.row,"L").value = "UCS POC"
					End IF
					If .Cells(checkrow.Row, "L") = "UI Project" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "Unified Fabric Optimization Service" Then
						.cells(checkrow.row,"L").value = "Unified Fabric Optimization"
					End IF
					If .Cells(checkrow.Row, "L") = "Unified Fabric Optimization Service  " Then
						.cells(checkrow.row,"L").value = "Unified Fabric Optimization"
					End IF
					If .Cells(checkrow.Row, "L") = "V Block" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "VDI - Custom SOW" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "VDI pilot testing" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "VDI POC" Then
						.cells(checkrow.row,"L").value = "VDI POC"
					End IF
					If .Cells(checkrow.Row, "L") = "VDI Program Management" Then
						.cells(checkrow.row,"L").value = "Other"
					End IF
					If .Cells(checkrow.Row, "L") = "VDI SME" Then
						.cells(checkrow.row,"L").value = "SME"
					End IF
					If .Cells(checkrow.Row, "L") = "VDI Workshop" Then
						.cells(checkrow.row,"L").value = "VDI POC"
					End IF
					If .Cells(checkrow.Row, "L") = "VDI Workshop" Then
						.cells(checkrow.row,"L").value = "VDI POC"
					End IF
					If .Cells(checkrow.Row, "L") = "Virtualization Accelerator" Then
						.cells(checkrow.row,"L").value = "Unified Computing Virtualization Accelerator Service"
					End IF
					If .Cells(checkrow.Row, "L") = "WAAS" Then
						.cells(checkrow.row,"L").value = "Wide Area Application Services Planning and Design Service (WAAS)"
					End IF
					If .Cells(checkrow.Row, "L") = "WAAS Assessment" Then
						.cells(checkrow.row,"L").value = "Wide Area Application Services Planning and Design Service (WAAS)"
					End IF
					If .Cells(checkrow.Row, "L") = "WAAS Assessment" Then
						.cells(checkrow.row,"L").value = "Wide Area Application Services Planning and Design Service (WAAS)"
					End IF
					If .Cells(checkrow.Row, "L") = "WAAS LLD" Then
						.cells(checkrow.row,"L").value = "Wide Area Application Services Planning and Design Service (WAAS)"
					End IF
					If .Cells(checkrow.Row, "L") = "WAAS Optimization" Then
						.cells(checkrow.row,"L").value = "Wide Area Application Services Optimization"
					End IF
					If .Cells(checkrow.Row, "L") = "WAAS PDI" Then
						.cells(checkrow.row,"L").value = "Wide Area Application Services Planning Design and Implementation Service (WAAS)"
					End IF
					If .Cells(checkrow.Row, "L") = "Wide Area Application Services Planning, Design  And Implementation Service (WAAS)" Then
						.cells(checkrow.row,"L").value = "Wide Area Application Services Planning Design and Implementation Service (WAAS)"
					End IF

					'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Next
    End With


    ' AutoFit the column width in the summary sheet.
    'DestSh.Columns.AutoFit
    'DestSh.Rows.Height = 35
    
    With Application
        .Calculation = CalcMode
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    Unload DynamicForm
End Sub