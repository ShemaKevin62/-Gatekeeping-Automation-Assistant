# -Gatekeeping-Automation-Assistant
Streamlining Operational Task Assignment &amp; Audit Workflow with Excel, Power Automate, and Outlook
🧠 Executive Summary
The Email Automation Assistant is a low-code digital solution designed to automate the triaging, logging, and tracking of incoming work requests received via email. Built using Microsoft Power Automate, Excel, and Outlook, this system replaces an error-prone, manual task allocation method with a streamlined, auditable, and dashboard-driven workflow.

This project addresses operational inefficiencies that previously led to lost work, inconsistent accountability, and lack of visibility. It introduces automation to ensure every task is logged, assigned fairly, and tracked throughout its lifecycle — empowering both team leads and executive management with accurate, real-time insights.

💼 Business Problem
The team received client instructions via a shared Outlook inbox. Gatekeeping was manual:

A team member manually downloaded emails

Saved large email attachments to OneDrive

Created folders in employee names

Assigned emails by dragging files into folders

🔴 Problems Faced:

No automated log of received emails

Some emails were missed or never assigned

Employees deleted or moved files, erasing traceability

No clear ownership or audit trail for who was assigned what

Daily reports were compiled manually (or skipped)

🎯 Project Goals
Eliminate manual gatekeeping and assignment process

Ensure every email received is logged and processed

Automate fair and trackable employee task assignment

Create a file lifecycle trail (received → assigned → completed)

Provide a real-time dashboard and daily reports for management

Build audit readiness by logging every action and user involved

⚙️ Tools & Technologies
Microsoft Outlook — Email intake

Power Automate — Trigger and data logging

Microsoft Excel — Assignment log, audit trail, dashboard

VBA Macros — Random assignment logic, file lifecycle update

🛠 Solution Process Overview
✅ Step-by-Step Automation Flow:
Email Received in shared Outlook inbox

Power Automate Trigger captures metadata (subject, sender, timestamp)

Logs to Excel Table → creates a new record

VBA Macro Checks attendance status ("present"/"absent")

Task Assignment is randomly done among present employees

Task Status (e.g., "Assigned", "In Progress", "Completed") is updated by employees

Excel Dashboard updates live

Automated Daily Summary is sent to management

Audit Trail includes:

Who was assigned what

Timestamp logs

Status changes with usernames

📊 Visual Flow Diagram
mermaid
Copy
Edit
flowchart LR
A[Shared Outlook Inbox] --> B[Power Automate: Email Trigger]
B --> C[Log to Excel: Subject, Sender, Time]
C --> D[Check Employee Availability (Excel)]
D --> E[Randomly Assign to Present Employee]
E --> F[Excel Log Updates]
F --> G[Employee Updates Status]
G --> H[Dashboard and Audit Trail Update]
H --> I[Daily Summary Sent to Management]
🔁 Before vs After
Step	Before (Manual)	After (Automated)
Email receipt	Manually viewed in Outlook	Auto-captured by Power Automate
Task assignment	Done by dragging files to folders	VBA assigns based on availability
Tracking of email lifecycle	None; files could be deleted or moved	Full log with assignment + status updates
Accountability	Unknown who was assigned	Every task has an assignee + timestamp
Reporting	Compiled manually at day-end	Live dashboard + auto-generated report
Audit-readiness	Not auditable	Centralized log with status and changes tracked

📈 Outcome & Measurable Impact
🕒 Saved 2–3 hours per day previously spent on manual routing and follow-up

🧾 Introduced full audit trail of task flow, from email to completion

📉 Reduced unassigned/missed emails to zero

🔍 Increased task ownership and accountability across employees

📊 Empowered management with up-to-date performance dashboards

⚙️ Reduced dependencies on a single gatekeeper

🔍 Features That Matter to PMs
✅ No emails missed — complete digital paper trail

📤 Automatic reporting to management every day

🧠 Fair distribution based on presence (can later be expanded to skill-based matching)

🔒 Audit logs for compliance and traceability

💬 Employee statuses (Assigned, In Progress, Completed) track lifecycle

📁 Centralized record of tasks with who/what/when

📷 Visual Suggestions for GitHub
Place these inside a /images folder and reference in your README:

Visual	Description
email_flow.png	Diagram showing email-to-dashboard journey
excel_log.png	Screenshot of Excel table logging Subject, Time, Assignee
attendance_sheet.png	Table showing employee presence/absence
status_dashboard.png	Visual dashboard with task counts, employee stats
email_assign_demo.gif	10–20 sec screen recording: email → Excel → dashboard
audit_trail_log.png	Snapshot of how assignments and status changes are tracked

📊 Dashboard Metrics
Your dashboard could include:

Total emails received today

Assignment distribution by employee

% Completed vs In Progress

Time taken from assignment to completion (for SLAs)

of overdue or stalled tasks
Filter by day, week, month

📁 Folder Structure
pgsql
Copy
Edit
email-automation-assistant/
│
├── README.md
├── flow_diagram.mmd
├── images/
│   ├── email_flow.png
│   ├── excel_log.png
│   ├── dashboard_view.png
│   ├── attendance_sheet.png
├── excel/
│   └── email_task_log.xlsx
├── power_automate/
│   └── outlook_trigger_flow.json
├── vba/
│   └── assign_and_track.bas
📌 How to Use
Download the Excel file and customize your employee list.

Import Power Automate flow and connect it to your shared inbox.

Open Excel daily and use the VBA macro to assign tasks.

Share the dashboard via email or Teams using macro or Power Automate.

Let employees update statuses as they progress.

Use dashboard or exported reports to track, audit, and manage performance.

🧠 Lessons Learned
Built automation across Outlook, Excel, and Power Automate without complex code

Identified and solved real-world audit and visibility issues

Strengthened reporting workflows with auto-dashboards

Realized the need for future skill-based assignment logic

🚧 Limitations
📆 Accuracy of employee attendance is essential

⚖️ Assignment is random — not yet skill-aware

📨 Email body content not currently parsed (future version)

⚙️ Not yet cloud-native — relies on local Excel file

🚀 Future Enhancements
Skill-based assignment logic (mapping tasks to employee specialties)

Outlook email parsing (for task complexity or urgency tagging)

Integration with Teams/Slack for real-time alerts

SLA monitoring and escalation alerts

Web dashboard (Power BI or Excel Online)





-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

'=========================
'  PUBLIC BUTTON MACROS
'=========================
Public Sub GenerateEmail()
    Dim ui As UIInfo, productBase As String, isJoint As Boolean
    If Not LoadUI(ui, productBase, isJoint) Then Exit Sub

    Dim rules As ProductRule: rules = GetProductRule(productBase)
    Dim info As HostInfo
    If Not ReadHostFields(info) Then
        MsgBox "Couldn't read required fields from Host screen.", vbCritical: Exit Sub
    End If

    Dim rec As Recommendation
    rec = CalculateRecommendation(productBase, info, rules)
    If Not rec.IsReductionPossible Then
        MsgBox "No reduction possible under current rules (new limit ≥ current limit).", vbInformation: Exit Sub
    End If

    BuildOutlookEmail BuildProductLabel(productBase, isJoint), info, rules, rec, ui
End Sub

Public Sub GenerateDocument()
    Dim ui As UIInfo, productBase As String, isJoint As Boolean
    If Not LoadUI(ui, productBase, isJoint) Then Exit Sub

    Dim rules As ProductRule: rules = GetProductRule(productBase)
    Dim info As HostInfo
    If Not ReadHostFields(info) Then
        MsgBox "Couldn't read required fields from Host screen.", vbCritical: Exit Sub
    End If

    Dim rec As Recommendation
    rec = CalculateRecommendation(productBase, info, rules)
    If Not rec.IsReductionPossible Then
        MsgBox "No reduction possible under current rules (new limit ≥ current limit).", vbInformation: Exit Sub
    End If

    BuildWordDocument BuildProductLabel(productBase, isJoint), isJoint, info, rules, rec, ui
End Sub

Public Sub GenerateBoth()
    GenerateDocument
    GenerateEmail
End Sub

'=========================
'   UI / CONFIG HELPERS
'=========================
Private Type UIInfo
    OPC As String
    ClientName As String
    Province As String
    ProductRaw As String
End Type

Private Function LoadUI(ByRef ui As UIInfo, ByRef productBase As String, ByRef isJoint As Boolean) As Boolean
    ' Tries named ranges first; falls back to Control!E4:E7 (as in your screenshot)
    ui.OPC = NzS(GetCellValuePreferNames(Array("OPC_Number"), "Control", "E4"))
    ui.ClientName = NzS(GetCellValuePreferNames(Array("Client_Name"), "Control", "E5"))
    ui.Province = NzS(GetCellValuePreferNames(Array("Province"), "Control", "E6"))
    ui.ProductRaw = NzS(GetCellValuePreferNames(Array("ProductChoice"), "Control", "E7"))

    If Len(ui.ProductRaw) = 0 Then
        MsgBox "Choose a product in the dropdown (e.g., Joint LOC / Joint Heloc / Joint Flexline).", vbExclamation
        Exit Function
    End If

    ParseProduct ui.ProductRaw, productBase, isJoint
    LoadUI = True
End Function

Private Sub ParseProduct(ByVal productText As String, ByRef baseName As String, ByRef isJoint As Boolean)
    Dim s As String: s = Trim$(UCase$(productText))
    isJoint = (InStr(1, s, "JOINT", vbTextCompare) > 0)

    If InStr(1, s, "HELOC", vbTextCompare) > 0 Then
        baseName = "HELOC FlexLine"
    ElseIf InStr(1, s, "FLEXLINE", vbTextCompare) > 0 Then
        baseName = "FlexLine"
    Else
        baseName = "LOC"
    End If
End Sub

Private Function BuildProductLabel(ByVal baseName As String, ByVal isJoint As Boolean) As String
    If isJoint Then
        If UCase$(baseName) = "LOC" Then
            BuildProductLabel = "Joint LOC"
        ElseIf UCase$(baseName) Like "*HELOC*" Then
            BuildProductLabel = "Joint Heloc"
        Else
            BuildProductLabel = "Joint Flexline"
        End If
    Else
        BuildProductLabel = baseName
    End If
End Function

' One-time config creator (unchanged defaults; tweak in sheet)
Public Sub SetupConfig()
    Dim ws As Worksheet
    If Not SheetExists("LimitConfig") Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = "LimitConfig"
    Else
        Set ws = Worksheets("LimitConfig"): ws.Cells.Clear
    End If

    With ws
        .Range("A1:E1").Value = Array("Product", "CushionAmount", "CushionPercent", "RoundStep", "MinRevolving")
        .Range("A2:E2").Value = Array("LOC", 500, 0, 100, 0)
        .Range("A3:E3").Value = Array("FlexLine", 1000, 0, 100, 1000)
        .Range("A4:E4").Value = Array("HELOC FlexLine", 1000, 0, 500, 5000)
        .Columns("A:E").AutoFit
    End With
    MsgBox "LimitConfig prepared.", vbInformation
End Sub

Private Type ProductRule
    CushionAmount As Double
    CushionPercent As Double
    RoundStep As Double
    MinRevolving As Double
End Type

Private Function GetProductRule(ByVal productName As String) As ProductRule
    Dim ws As Worksheet, r As Long, lastRow As Long
    If Not SheetExists("LimitConfig") Then Err.Raise vbObjectError + 3000, , "Run SetupConfig first."
    Set ws = Worksheets("LimitConfig")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow
        If StrComp(Trim(ws.Cells(r, 1).Value), Trim(productName), vbTextCompare) = 0 Then
            With GetProductRule
                .CushionAmount = NzD(ws.Cells(r, 2).Value)
                .CushionPercent = NzD(ws.Cells(r, 3).Value)
                .RoundStep = IIf(NzD(ws.Cells(r, 4).Value) <= 0, 100, NzD(ws.Cells(r, 4).Value))
                .MinRevolving = Application.Max(0, NzD(ws.Cells(r, 5).Value))
            End With
            Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 3001, , "No matching Product in LimitConfig for '" & productName & "'."
End Function

'=========================
'  IBM HOST SCREEN READ
'=========================
' Coordinates you mapped:
' Total credit O/S:   r5 c20–34
' Available credit:   r6 c20–34
' Credit limit:       r5 c59–75
' Current plan limit: r6 c59–75
' Revolving Cr:       r9 c20–35
' Term portions:      r11–20 c20–35 (one per row)

Private Const RBAL_R As Integer = 5, RBAL_C1 As Integer = 20, RBAL_C2 As Integer = 34
Private Const RAVC_R As Integer = 6, RAVC_C1 As Integer = 20, RAVC_C2 As Integer = 34
Private Const RLIM_R As Integer = 5, RLIM_C1 As Integer = 59, RLIM_C2 As Integer = 75
Private Const RPLAN_R As Integer = 6, RPLAN_C1 As Integer = 59, RPLAN_C2 As Integer = 75
Private Const RREV_R As Integer = 9, RREV_C1 As Integer = 20, RREV_C2 As Integer = 35
Private Const TPORT_R1 As Integer = 11, TPORT_RN As Integer = 20, TPORT_C1 As Integer = 20, TPORT_C2 As Integer = 35

Private Type HostInfo
    TotalOS As Double
    AvailableCredit As Double
    CreditLimit As Double
    PlanLimit As Double
    RevolvingBalance As Double
    TermPortions() As Double
    TermSum As Double
End Type

Private Function ReadHostFields(ByRef info As HostInfo) As Boolean
    On Error GoTo fail
    Dim sess As Object, oia As Object, ps As Object

    Set sess = CreateObject("PCOMM.autECLSession")
    sess.SetConnectionByName "A"
    sess.StartCommunication

    Set oia = CreateObject("PCOMM.autECLOIA")
    oia.SetConnectionByHandle sess.Handle
    oia.WaitForInputReady 10000, 1

    Set ps = CreateObject("PCOMM.autECLPS")
    ps.SetConnectionByHandle sess.Handle

    info.TotalOS = ToAmount(ps.GetText(RBAL_R, RBAL_C1, RBAL_C2 - RBAL_C1 + 1))
    info.AvailableCredit = ToAmount(ps.GetText(RAVC_R, RAVC_C1, RAVC_C2 - RAVC_C1 + 1))
    info.CreditLimit = ToAmount(ps.GetText(RLIM_R, RLIM_C1, RLIM_C2 - RLIM_C1 + 1))
    info.PlanLimit = ToAmount(ps.GetText(RPLAN_R, RPLAN_C1, RPLAN_C2 - RPLAN_C1 + 1))
    info.RevolvingBalance = ToAmount(ps.GetText(RREV_R, RREV_C1, RREV_C2 - RREV_C1 + 1))

    Dim r As Integer, s As String, amt As Double, idx As Integer
    ReDim info.TermPortions(1 To (TPORT_RN - TPORT_R1 + 1))
    For r = TPORT_R1 To TPORT_RN
        s = ps.GetText(r, TPORT_C1, TPORT_C2 - TPORT_C1 + 1)
        amt = ToAmount(s)
        idx = idx + 1
        info.TermPortions(idx) = amt
        info.TermSum = info.TermSum + amt
    Next r

    ReadHostFields = True
    Exit Function
fail:
    ReadHostFields = False
End Function

'=========================
'   CALC RECOMMENDER
'=========================
Private Type Recommendation
    RevolvingTarget As Double
    NewLimit As Double
    Reduction As Double
    CapApplied As Double
    IsReductionPossible As Boolean
End Type

Private Function CalculateRecommendation(ByVal productName As String, ByRef info As HostInfo, ByRef rules As ProductRule) As Recommendation
    Dim effCap As Double: effCap = IIf(info.PlanLimit > 0, info.PlanLimit, info.CreditLimit)
    Dim revTargetRaw As Double, revTarget As Double, newLimitRaw As Double, newLimitRounded As Double

    Select Case UCase$(productName)
        Case "LOC"
            newLimitRaw = info.TotalOS + rules.CushionAmount + (rules.CushionPercent * info.TotalOS)
            newLimitRounded = RoundUpTo(newLimitRaw, rules.RoundStep)
            revTarget = newLimitRounded

        Case "FLEXLINE", "HELOC FLEXLINE"
            revTargetRaw = info.RevolvingBalance + rules.CushionAmount + (rules.CushionPercent * info.RevolvingBalance)
            revTarget = RoundUpTo(Application.Max(revTargetRaw, rules.MinRevolving), rules.RoundStep)
            newLimitRounded = revTarget + info.TermSum

        Case Else
            Err.Raise vbObjectError + 3100, , "Unknown product: " & productName
    End Select

    Dim capped As Double: capped = Application.Min(newLimitRounded, effCap)
    Dim red As Double: red = info.CreditLimit - capped

    With CalculateRecommendation
        .RevolvingTarget = IIf(revTarget = 0, newLimitRounded, revTarget)
        .NewLimit = capped
        .Reduction = IIf(red > 0, red, 0)
        .CapApplied = effCap
        .IsReductionPossible = (red > 0.0001)
    End With
End Function

'=========================
'   OUTLOOK EMAIL
'=========================
Private Sub BuildOutlookEmail(ByVal productLabel As String, ByRef info As HostInfo, ByRef rules As ProductRule, ByRef rec As Recommendation, ByRef ui As UIInfo)
    On Error Resume Next
    Dim ol As Object: Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0

    Dim body As String
    body = ""
    body = body & "Team," & vbCrLf & vbCrLf
    body = body & "Recommendation: Reduce " & productLabel & " credit limit." & vbCrLf & vbCrLf
    body = body & "Client: " & ui.ClientName & " | OPC: " & ui.OPC & " | Province: " & ui.Province & vbCrLf & vbCrLf

    body = body & "— Current snapshot —" & vbCrLf
    body = body & "   • Total O/S:        " & Fmt(info.TotalOS) & vbCrLf
    body = body & "   • Available Credit: " & Fmt(info.AvailableCredit) & vbCrLf
    body = body & "   • Credit Limit:     " & Fmt(info.CreditLimit) & vbCrLf
    If info.PlanLimit > 0 Then body = body & "   • Plan Limit:       " & Fmt(info.PlanLimit) & vbCrLf
    If UCase$(productLabel) <> "LOC" And InStr(1, productLabel, "LOC", vbTextCompare) = 0 Then
        body = body & "   • Revolving O/S:    " & Fmt(info.RevolvingBalance) & vbCrLf
        body = body & "   • Term Portions Σ:  " & Fmt(info.TermSum) & vbCrLf
    End If
    body = body & vbCrLf

    body = body & "— Rule parameters —" & vbCrLf
    body = body & "   • Cushion Amount:   " & Fmt(rules.CushionAmount) & vbCrLf
    body = body & "   • Cushion Percent:  " & FormatPercent(rules.CushionPercent, 2) & vbCrLf
    body = body & "   • Round Step:       $" & FormatNumber(rules.RoundStep, 0) & vbCrLf
    If UCase$(productLabel) <> "LOC" And InStr(1, productLabel, "LOC", vbTextCompare) = 0 Then _
        body = body & "   • Min Revolving:    " & Fmt(rules.MinRevolving) & vbCrLf
    body = body & vbCrLf

    body = body & "— Computation —" & vbCrLf
    If InStr(1, UCase$(productLabel), "LOC") > 0 And InStr(1, UCase$(productLabel), "FLEX") = 0 Then
        body = body & "   • Proposed Limit:   " & Fmt(rec.NewLimit) & vbCrLf
    Else
        body = body & "   • Revolving Target: " & Fmt(rec.RevolvingTarget) & vbCrLf
        body = body & "   • Proposed Total:   " & Fmt(rec.NewLimit) & vbCrLf
    End If
    body = body & "   • Reduction:        " & Fmt(rec.Reduction) & vbCrLf
    If rec.NewLimit < rec.CapApplied Then
        body = body & "   • (Capped by plan/credit limit at " & Fmt(rec.CapApplied) & ")." & vbCrLf
    End If
    body = body & vbCrLf & "Please review before sending." & vbCrLf

    With ol.CreateItem(0)
        .Subject = "Limit Reduction Recommendation – " & productLabel
        .Body = body
        .Display
    End With
End Sub

'=========================
'   WORD DOC GENERATION
'=========================
Private Sub BuildWordDocument(ByVal productLabel As String, ByVal isJoint As Boolean, _
                              ByRef info As HostInfo, ByRef rules As ProductRule, _
                              ByRef rec As Recommendation, ByRef ui As UIInfo)

    Dim outDir As String, outPath As String, templatePath As String
    outDir = EnsureFolder(ThisWorkbook.Path & "\Output")
    templatePath = ThisWorkbook.Path & "\Templates\LimitReductionTemplate.docx"

    outPath = outDir & "\" & Format(Now, "yyyymmdd") & "_" & ToFilenameSafe(ui.OPC & "_" & productLabel) & "_LimitReduction.docx"

    Dim wd As Object, doc As Object
    Set wd = CreateObject("Word.Application")

    If Dir(templatePath, vbNormal) <> "" Then
        Set doc = wd.Documents.Open(templatePath, ReadOnly:=True)
        doc.SaveAs2 outPath
        FillBM doc, "OPC", ui.OPC
        FillBM doc, "ClientName", ui.ClientName
        FillBM doc, "Province", ui.Province
        FillBM doc, "Product", productLabel
        FillBM doc, "IsJoint", IIf(isJoint, "Yes", "No")
        FillBM doc, "TotalOS", Fmt(info.TotalOS)
        FillBM doc, "AvailCredit", Fmt(info.AvailableCredit)
        FillBM doc, "CreditLimit", Fmt(info.CreditLimit)
        FillBM doc, "PlanLimit", IIf(info.PlanLimit > 0, Fmt(info.PlanLimit), "—")
        FillBM doc, "RevolvingOS", Fmt(info.RevolvingBalance)
        FillBM doc, "TermSum", Fmt(info.TermSum)
        FillBM doc, "ProposedLimit", Fmt(rec.NewLimit)
        FillBM doc, "Reduction", Fmt(rec.Reduction)
        FillBM doc, "RunDate", Format(Now, "yyyy-mm-dd")
    Else
        ' No template: build a clean doc with a table
        Set doc = wd.Documents.Add
        With doc.Content
            .InsertAfter "Credit Limit Reduction Recommendation" & vbCrLf
            .InsertAfter productLabel & IIf(isJoint, " (Joint)", "") & vbCrLf & vbCrLf
        End With

        Dim tbl As Object
        Set tbl = doc.Tables.Add(Range:=doc.Content, NumRows:=14, NumColumns:=2)
        tbl.Range.ParagraphFormat.SpaceAfter = 6

        PutKV tbl, 1, "Client Name", ui.ClientName
        PutKV tbl, 2, "OPC", ui.OPC
        PutKV tbl, 3, "Province", ui.Province
        PutKV tbl, 4, "Product", productLabel
        PutKV tbl, 5, "Total O/S", Fmt(info.TotalOS)
        PutKV tbl, 6, "Available Credit", Fmt(info.AvailableCredit)
        PutKV tbl, 7, "Credit Limit", Fmt(info.CreditLimit)
        PutKV tbl, 8, "Plan Limit", IIf(info.PlanLimit > 0, Fmt(info.PlanLimit), "—")
        PutKV tbl, 9, "Revolving O/S", Fmt(info.RevolvingBalance)
        PutKV tbl, 10, "Term Portions Σ", Fmt(info.TermSum)
        PutKV tbl, 11, "Cushion Amount", Fmt(rules.CushionAmount)
        PutKV tbl, 12, "Cushion Percent", FormatPercent(rules.CushionPercent, 2)
        PutKV tbl, 13, "Rounding Step", "$" & FormatNumber(rules.RoundStep, 0)
        PutKV tbl, 14, "Proposed Limit / Reduction", Fmt(rec.NewLimit) & "  (↓ " & Fmt(rec.Reduction) & ")"

        doc.Content.InsertAfter vbCrLf & "Generated: " & Format(Now, "yyyy-mm-dd")
        doc.SaveAs2 outPath
    End If

    wd.Visible = True
    wd.Activate
    MsgBox "Document ready: " & outPath, vbInformation
End Sub

Private Sub FillBM(ByVal doc As Object, ByVal name As String, ByVal textVal As String)
    On Error Resume Next
    If doc.Bookmarks.Exists(name) Then
        doc.Bookmarks(name).Range.Text = textVal
        doc.Bookmarks.Add name, doc.Bookmarks(name).Range
    End If
    On Error GoTo 0
End Sub

Private Sub PutKV(ByRef tbl As Object, ByVal row As Long, ByVal key As String, ByVal val As String)
    tbl.Cell(row, 1).Range.Text = key
    tbl.Cell(row, 2).Range.Text = val
End Sub

'=========================
'        UTILITIES
'=========================
Private Function SheetExists(ByVal name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function NzD(v) As Double
    If IsError(v) Or IsEmpty(v) Or v = "" Then NzD = 0 Else NzD = CDbl(v)
End Function

Private Function NzS(v) As String
    If IsError(v) Or IsEmpty(v) Then NzS = "" Else NzS = CStr(v)
End Function

Private Function ToAmount(ByVal s As String) As Double
    Dim t As String: t = Trim$(s)
    Dim neg As Boolean
    If InStr(1, t, "(") > 0 And InStr(1, t, ")") > 0 Then neg = True
    If InStr(1, t, "-") > 0 Then neg = True
    t = Replace$(t, "CR", "", , , vbTextCompare)
    t = Replace$(t, "DR", "", , , vbTextCompare)
    t = Replace$(t, "$", "")
    t = Replace$(t, ",", "")
    t = Replace$(t, "(", "")
    t = Replace$(t, ")", "")
    t = Replace$(t, " ", "")
    If Len(t) = 0 Then
        ToAmount = 0
    Else
        ToAmount = Val(t)
        If neg Then ToAmount = -ToAmount
    End If
End Function

Private Function Fmt(ByVal amt As Double) As String
    Fmt = FormatCurrency(amt, 2)
End Function

Private Function RoundUpTo(ByVal value As Double, ByVal stepSize As Double) As Double
    If stepSize <= 0 Then stepSize = 100
    RoundUpTo = stepSize * WorksheetFunction.Ceiling_Precise(value / stepSize, 1)
End Function

Private Function EnsureFolder(ByVal path As String) As String
    If Dir(path, vbDirectory) = "" Then MkDir path
    EnsureFolder = path
End Function

Private Function ToFilenameSafe(ByVal s As String) As String
    Dim bad As Variant: bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, bad(i), "_")
    Next i
    ToFilenameSafe = s
End Function

Private Function GetCellValuePreferNames(possibleNames As Variant, ByVal fallbackSheet As String, ByVal fallbackAddr As String) As Variant
    Dim nm As Variant
    For Each nm In possibleNames
        On Error Resume Next
        GetCellValuePreferNames = Evaluate(nm)
        If Err.Number = 0 And Not IsEmpty(GetCellValuePreferNames) Then
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    Next nm
    GetCellValuePreferNames = Worksheets(fallbackSheet).Range(fallbackAddr).Value
End Function


