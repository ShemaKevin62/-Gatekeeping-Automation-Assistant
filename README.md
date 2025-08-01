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

