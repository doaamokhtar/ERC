"""Builds the WB Business Requirements & SLA template (Excel) that ERC fills in.

Yellow cells are for the customer; the grey italic row in each table is an example of the
expected format. Dropdowns come from the hidden 'Lists' sheet, so a value added there shows up
in every dropdown that uses it.

Usage: python requirements-template.py <out.xlsx>
"""
import sys
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from pathlib import Path

OUT = sys.argv[1]
HERE = Path(__file__).parent

NAVY, MINT, MINT_DARK, GREY, INK = "024638", "36B78E", "228848", "5E716B", "1D2B27"  # Expo 2030 Riyadh palette
FILL_ME = PatternFill("solid", fgColor="FFF8D6")      # customer input
HEAD = PatternFill("solid", fgColor=NAVY)
EXAMPLE = PatternFill("solid", fgColor="F1F6F3")
BAND = PatternFill("solid", fgColor="EEF7F3")
thin = Side(style="thin", color="D6E6DF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
F = lambda **k: Font(name="Arial", size=k.pop("size", 10), **k)
WRAP = Alignment(wrap_text=True, vertical="top")
ROWS = 60                                              # empty input rows per table

wb = Workbook()

# ---------------------------------------------------------------- Lists (hidden)
lists = {
    "Area": ["Intake / Portal", "Classification", "Routing", "Investigation", "Notifications",
             "SLA & Escalation", "Reporting", "Security & Access", "Data & Retention", "Other"],
    "Priority": ["Must have", "Should have", "Could have", "Won't have (now)"],
    "ReqStatus": ["New", "In review", "Approved", "Rejected", "Implemented"],
    "ReportCategory": ["All categories",
                       "Fraud, Theft, Embezzlement or Asset Misappropriation",
                       "Bribery, Corruption, Kickbacks or Trading in Influence",
                       "Money Laundering, Terrorist Financing or Sanctions Concerns",
                       "Procurement Manipulation, Collusion, Bid Rigging or Conflict of Interest",
                       "Harassment, Discrimination, Bullying or Abuse of Authority",
                       "Health, Safety, Environmental or Heritage-Protection Breach",
                       "Information Security, Data Protection or Systems Misuse",
                       "Other Regulatory or License Compliance Breach",
                       "Retaliation Against a Reporter or Witness", "Other"],
    "Severity": ["All severities", "Critical", "High", "Medium", "Low"],
    "Stage": ["Compliance Decision", "Investigation Decision", "Employee Relations Approval (Retaliation)",
              "First contact with reporter", "Case closure", "Other"],
    "Unit": ["Working hours", "Working days", "Calendar hours", "Calendar days"],
    "Calendar": ["ERC working week (Sun-Thu 08:00-16:00)", "24/7"],
    "Channel": ["Email", "Email + Microsoft Teams", "Microsoft Teams", "SMS"],
    "YesNo": ["Yes", "No"],
    "Requesters": ["All ERC employees", "Specific departments (list in Notes)", "Specific group (name in Notes)",
                   "Anyone - no sign-in (anonymous)", "External customers / partners"],
    "FilledBy": ["Requester (portal)", "Internal team only"],
    # field types in business words, and what each becomes in Jira (same order)
    "FieldType": ["Single-line text", "Multi-line text", "Number", "Date", "Date and time", "Dropdown - one choice",
                  "Dropdown - many choices", "Checkboxes", "Yes / No", "Person (pick a user)", "Email address",
                  "Phone number", "Attachment", "Automatic - requester name", "Automatic - date submitted"],
    "JiraField": ["Short text (Text Field - single line)", "Paragraph (Text Field - multi-line)", "Number Field",
                  "Date Picker", "Date Time Picker", "Select List (single choice)", "Select List (multiple choices)",
                  "Checkboxes", "Radio Buttons (Yes, No)", "User Picker (single user)", "Short text + email validation (Form)",
                  "Short text (Form)", "System: Attachment", "System: Reporter - no field needed",
                  "System: Created - no field needed"],
}
ls = wb.active
ls.title = "Lists"
for c, (name, values) in enumerate(lists.items(), start=1):
    ls.cell(row=1, column=c, value=name).font = F(bold=True)
    for r, v in enumerate(values, start=2):
        ls.cell(row=r, column=c, value=v)
    letter = get_column_letter(c)
    wb.defined_names[name] = __import__("openpyxl").workbook.defined_name.DefinedName(
        name, attr_text=f"Lists!${letter}$2:${letter}${len(values) + 1}")
ls.sheet_state = "hidden"


def sheet(title, tab_color):
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = tab_color
    ws.sheet_view.showGridLines = False
    return ws


def banner(ws, title, subtitle, last_col):
    ws.row_dimensions[1].height = 44
    ws.row_dimensions[2].height = 22
    end = get_column_letter(last_col)
    ws.merge_cells(f"B1:{end}1")
    ws.merge_cells(f"B2:{end}2")
    for r in (1, 2):
        for c in range(2, last_col + 1):
            ws.cell(row=r, column=c).fill = HEAD
    ws["B1"] = title
    ws["B1"].font = F(size=16, bold=True, color="FFFFFF")
    ws["B1"].alignment = Alignment(vertical="center")
    ws["B2"] = subtitle
    ws["B2"].font = F(size=10, color=MINT)
    ws["B2"].alignment = Alignment(vertical="center")
    try:
        logo = Image(str(HERE / "erc-icon.png"))
        logo.width, logo.height = 46, 41
        ws.add_image(logo, "A1")
    except Exception:
        pass


def table(ws, start_row, headers, widths, example, validations, notes=None, rows=ROWS, header_height=34):
    """headers row, one grey example row, then `rows` yellow input rows."""
    ws.row_dimensions[start_row].height = header_height
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=start_row, column=i, value=h)
        c.font = F(bold=True, color="FFFFFF")
        c.fill = HEAD
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = BOX
        ws.column_dimensions[get_column_letter(i)].width = w
        if notes and h.split("\n")[0] in notes:
            from openpyxl.comments import Comment
            c.comment = Comment(notes[h.split("\n")[0]], "Ejada")
    ex_row = start_row + 1
    for i, v in enumerate(example, start=1):
        c = ws.cell(row=ex_row, column=i, value=v)
        c.font = F(italic=True, color=GREY)
        c.fill = EXAMPLE
        c.alignment = WRAP
        c.border = BOX
    ws.cell(row=ex_row, column=1).value = example[0]
    first, last = ex_row + 1, ex_row + rows
    for r in range(first, last + 1):
        for i in range(1, len(headers) + 1):
            c = ws.cell(row=r, column=i)
            c.fill = FILL_ME
            c.border = BOX
            c.alignment = WRAP
            c.font = F(color=INK)
    for col_index, rule in validations.items():
        letter = get_column_letter(col_index)
        dv = DataValidation(**rule)
        dv.add(f"{letter}{first}:{letter}{last}")
        ws.add_data_validation(dv)
    ws.freeze_panes = ws.cell(row=start_row + 1, column=3)
    return first, last


lst = lambda name: {"type": "list", "formula1": f"={name}", "allow_blank": True,
                    "showErrorMessage": True, "errorTitle": "Pick from the list",
                    "error": "Choose a value from the dropdown, or add it to the list first."}
email_rule = {"type": "custom", "allow_blank": True, "showErrorMessage": True,
              "errorTitle": "Email address", "error": "Enter a valid email address, e.g. name@erc.sa."}
pct = {"type": "decimal", "operator": "between", "formula1": "0", "formula2": "1", "allow_blank": True,
       "showErrorMessage": True, "error": "Enter a percentage between 0% and 100%."}
num = {"type": "decimal", "operator": "greaterThan", "formula1": "0", "allow_blank": True,
       "showErrorMessage": True, "error": "Enter a number greater than 0."}

# ---------------------------------------------------------------- Instructions
ins = sheet("Instructions", NAVY)
ins.column_dimensions["A"].width = 16
ins.column_dimensions["B"].width = 34
ins.column_dimensions["C"].width = 70
banner(ins, "Whistleblowing Reporting Portal - Requirements Template",
       "Service catalog, request forms, SLA and groups - to be completed by ERC", 3)
r = 4
ins.cell(row=r, column=2, value="How to fill in this workbook").font = F(size=12, bold=True, color=NAVY); r += 1
legend = [
    (FILL_ME, "Yellow cell", "For you to fill in."),
    (EXAMPLE, "Grey italic row", "An example of the expected format. Overwrite it or leave it."),
    (None, "Dropdown", "Click the cell and pick from the list. Lists can be extended on request."),
    (None, "Red note marker", "Hover over a column header with a red corner for guidance."),
]
for fill, k, v in legend:
    a = ins.cell(row=r, column=2, value=k); a.font = F(bold=True); a.border = BOX
    if fill: a.fill = fill
    b = ins.cell(row=r, column=3, value=v); b.font = F(); b.border = BOX; b.alignment = WRAP
    r += 1
r += 1
ins.cell(row=r, column=2, value="Sheets").font = F(size=12, bold=True, color=NAVY); r += 1
for k, v in [("Service Catalog", "One row per service on the portal: who can request it, who approves (up to 3 levels) and which team delivers it."),
             ("Request Form Fields", "One row per field on a service's request form: label (English and Arabic), type, required or not, dropdown options."),
             ("Other Requirements", "Anything that is not a service or a field: notifications, reports, access rules, data retention."),
             ("SLA", "One row per service level: what is measured, the target, and who is warned and escalated to (group + email) at each level."),
             ("Groups & Contacts", "Who belongs to each group, their email, and whether they need a Jira licence.")]:
    a = ins.cell(row=r, column=2, value=k); a.font = F(bold=True); a.border = BOX
    b = ins.cell(row=r, column=3, value=v); b.font = F(); b.border = BOX; b.alignment = WRAP
    r += 1
r += 1
ins.cell(row=r, column=2, value="Document details").font = F(size=12, bold=True, color=NAVY); r += 1
details_start = r
for k, v in [("Customer", "Expo 2030 Riyadh Company (ERC)"), ("Service", "Whistleblowing Reporting Portal (Jira Service Management)"),
             ("Template version", "1.2"), ("Prepared by (Ejada)", "Doaa Mokhtar"),
             ("Completed by (name, role)", None), ("Date completed", None), ("Approved by (name, role)", None)]:
    a = ins.cell(row=r, column=2, value=k); a.font = F(bold=True); a.border = BOX
    b = ins.cell(row=r, column=3, value=v); b.font = F(); b.border = BOX
    if v is None: b.fill = FILL_ME
    r += 1
r += 1
ins.cell(row=r, column=2, value="Progress").font = F(size=12, bold=True, color=NAVY); r += 1
for k, formula in [("Services entered", "=COUNTA('Service Catalog'!C7:C66)"),
                   ("Form fields entered", "=COUNTA('Request Form Fields'!D7:D66)"),
                   ("Other requirements entered", "=COUNTA('Other Requirements'!C7:C66)"),
                   ("Other requirements approved", "=COUNTIF('Other Requirements'!K7:K66,\"Approved\")"),
                   ("SLA rows entered", "=COUNTA(SLA!B7:B66)"),
                   ("Group members entered", "=COUNTA('Groups & Contacts'!C7:C66)")]:
    a = ins.cell(row=r, column=2, value=k); a.font = F(bold=True); a.border = BOX
    b = ins.cell(row=r, column=3, value=formula); b.font = F(); b.border = BOX; b.alignment = Alignment(horizontal="left")
    r += 1
r += 1
note = ins.cell(row=r, column=2, value="Confidentiality: do not enter details of real reports in this workbook - only requirements and contacts.")
ins.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
note.font = F(bold=True, color=MINT_DARK); note.fill = BAND

# ---------------------------------------------------------------- Service Catalog
# Same columns, order and bilingual headers as the business's own service-catalog sheet; the grey
# 'Jira' columns are worked out from them, so each row maps straight onto one request type.
free = lambda name: dict(lst(name), showErrorMessage=False)  # dropdown, but typing a group name is allowed
sc = sheet("Service Catalog", MINT)
sc_headers = ["Service ID", "Service Group / Category\nفئة الخدمة", "Service Name\nاسم الخدمة",
              "Service Description\nوصف الخدمة", "Who can request?\nمن لديه صلاحية رفع الطلب؟",
              "Level 1 Approval\nالمستوى الأول للموافقة", "Level 2 Approval\nالمستوى الثاني للموافقة",
              "Level 3 Approval\nالمستوى الثالث للموافقة", "L1 Implementation Responsible Team\nفريق التنفيذ - المستوى الأول",
              "L2 Implementation Responsible Team\nفريق التنفيذ - المستوى الثاني", "Notes\nملاحظات",
              "Jira: approval steps", "Jira: form fields", "Ejada notes"]
sc_widths = [12, 26, 30, 42, 26, 24, 24, 24, 28, 28, 30, 13, 12, 30]
banner(sc, "Service Catalog", "One row per service - each row becomes one request type on the portal", len(sc_headers))
sc["A4"] = ("Row 6 is a worked example - copy its style. Repeat the category on every row (do not merge cells). "
            "Approvals and teams are free text: write the role or the team name, or '-' if none. Grey columns fill in automatically.")
sc["A4"].font = F(italic=True, color=GREY)
first, last = table(sc, 5, sc_headers, sc_widths,
    ["SVC-01", "Digital & Technology Demand Intake", "New System or Capability",
     "ERC sectors can ask for a new system or capability.", "All ERC employees",
     "Requester's Department Head", "-", "-", "PTO Team", "EA Team",
     "One approval only. PTO Team reviews the demand first, then hands it to EA Team for the architecture review.",
     None, None, "Request type: New System or Capability. Approver taken from the requester's department head."],
    {5: free("Requesters")},
    notes={"Who can request?": "Anonymous = a public portal with no sign-in. Specific departments or groups: name them in Notes.",
           "Level 1 Approval": "Free text: the role (e.g. Requester's Department Head) or the team/person who approves first. Write '-' if no approval is needed.",
           "Level 2 Approval": "Free text: who approves after Level 1. Write '-' if none.",
           "Level 3 Approval": "Free text: who approves after Level 2. Write '-' if none.",
           "L1 Implementation Responsible Team": "Free text: the team that receives and works the request first.",
           "L2 Implementation Responsible Team": "Free text: the team it is handed to next. Write '-' if none.",
           "Jira: approval steps": "Automatic: how many approval levels this service needs.",
           "Jira: form fields": "Automatic: how many fields are listed for this service on the 'Request Form Fields' sheet."},
    header_height=48)
for r_ in range(first - 1, last + 1):
    if r_ >= first:
        sc.cell(row=r_, column=1, value=f'=IF(C{r_}="","","SVC-"&TEXT(ROW()-{first - 2},"00"))')
    sc.cell(row=r_, column=12, value=f'=IF(C{r_}="","",SUMPRODUCT((TRIM(F{r_}:H{r_})<>"")*(TRIM(F{r_}:H{r_})<>"None")*(TRIM(F{r_}:H{r_})<>"-")))')
    sc.cell(row=r_, column=13, value=f"=IF(C{r_}=\"\",\"\",COUNTIF('Request Form Fields'!$B$6:$B$66,C{r_}))")
    for col in (1, 12, 13):
        c = sc.cell(row=r_, column=col); c.fill = EXAMPLE; c.font = F(color=GREY, italic=(r_ == first - 1))
    for col in (12, 13):
        sc.cell(row=r_, column=col).alignment = Alignment(horizontal="center", vertical="top")
wb.defined_names["ServiceNames"] = __import__("openpyxl").workbook.defined_name.DefinedName(
    "ServiceNames", attr_text=f"'Service Catalog'!$C${first - 1}:$C${last}")

# ---------------------------------------------------------------- Request Form Fields
ff = sheet("Request Form Fields", NAVY)
ff_headers = ["Field ID", "Service Name\nاسم الخدمة", "Order", "Field label (English)", "Field label (Arabic)\nاسم الحقل",
              "Field type\nنوع الحقل", "Required?\nإلزامي؟", "Dropdown options (separate with ;)\nالخيارات",
              "Help text shown under the field", "Who fills it?", "Jira: field type", "Ejada notes"]
ff_widths = [11, 30, 8, 30, 26, 24, 11, 36, 34, 18, 34, 28]
banner(ff, "Request Form Fields", "One row per field on a service's request form, in the order they appear", len(ff_headers))
ff["A4"] = ("Row 6 is a worked example. Do not add 'Requester name' or 'Date submitted' as fields - Jira records both automatically "
            "(choose the 'Automatic' field types if you want them listed).")
ff["A4"].font = F(italic=True, color=GREY)
first, last = table(ff, 5, ff_headers, ff_widths,
    ["FLD-001", "New System or Capability", 1, "Business Priority", "أولوية الطلب", "Dropdown - one choice", "Yes",
     "High; Medium; Low", "How important this demand is to your sector.", "Requester (portal)", None,
     "New field: Business Priority (select list)."],
    {2: {"type": "list", "formula1": "=ServiceNames", "allow_blank": True, "showErrorMessage": True,
         "errorTitle": "Unknown service", "error": "Pick a service from the 'Service Catalog' sheet."},
     3: dict(num, type="whole", formula1="0"), 6: lst("FieldType"), 7: lst("YesNo"), 10: lst("FilledBy")},
    notes={"Field type": "Pick the closest type; the grey 'Jira: field type' column shows what it becomes in Jira.",
           "Dropdown options (separate with ;)": "Only for Dropdown and Checkboxes, e.g. High; Medium; Low.",
           "Who fills it?": "Internal team only = hidden from the requester, filled in by the team that works the request.",
           "Jira: field type": "Automatic, from 'Field type'."},
    header_height=48)
for r_ in range(first - 1, last + 1):
    if r_ >= first:
        ff.cell(row=r_, column=1, value=f'=IF(D{r_}="","","FLD-"&TEXT(ROW()-{first - 2},"000"))')
    ff.cell(row=r_, column=11, value=f'=IF(F{r_}="","",IFERROR(INDEX(JiraField,MATCH(F{r_},FieldType,0)),"Ejada to review"))')
    for col in (1, 11):
        c = ff.cell(row=r_, column=col); c.fill = EXAMPLE; c.font = F(color=GREY, italic=(r_ == first - 1))

# ---------------------------------------------------------------- Other Requirements
br = sheet("Other Requirements", MINT)
br_headers = ["Req ID", "Area", "Requirement title", "Description", "Business justification", "Priority",
              "Requested by (name)", "Department", "Acceptance criteria", "Target date", "Status", "Ejada notes"]
br_widths = [14, 20, 30, 45, 35, 16, 22, 20, 40, 14, 14, 30]
banner(br, "Other Requirements", "Anything that is not a service or a form field - one row per requirement", len(br_headers))
br["A4"] = "Tip: keep one requirement per row. Use 'Acceptance criteria' to say how we will know it is done."
br["A4"].font = F(italic=True, color=GREY)
first, last = table(br, 5, br_headers, br_widths,
    ["BR-001", "Notifications", "Notify Compliance on every new report",
     "When a report is submitted, every member of GRC-Compliance receives an email with the reference number and a link.",
     "Reports must be seen within the same working day.", "Must have", "Head of Compliance", "Compliance",
     "A test report generates one email per GRC-Compliance member within 5 minutes; the email has no report details.",
     "2026-11-15", "Implemented", "Live in sandbox."],
    {2: lst("Area"), 6: lst("Priority"), 11: lst("ReqStatus"),
     10: {"type": "date", "operator": "greaterThan", "formula1": "DATE(2026,1,1)", "allow_blank": True,
          "showErrorMessage": True, "error": "Enter a date (dd/mm/yyyy)."}},
    notes={"Priority": "Must have = go-live blocker. Should have = important, can follow shortly after. Could have = nice to have.",
           "Acceptance criteria": "Describe the test that proves the requirement is met.",
           "Status": "Filled in jointly with Ejada during review."})
for r_ in range(first, last + 1):
    br.cell(row=r_, column=1, value=f'=IF(C{r_}="","","BR-"&TEXT(ROW()-{first - 2},"000"))').font = F(color=GREY)
    br.cell(row=r_, column=1).fill = EXAMPLE
    br.cell(row=r_, column=10).number_format = "dd/mm/yyyy"
br.cell(row=first - 1, column=10).number_format = "dd/mm/yyyy"
# requirements ERC has already raised; justification, priority, requester and date are left for ERC
ANONYMITY_NOTICE = ("Please note that choosing to remain anonymous may limit our ability to clarify facts "
                    "or follow up with you during the investigation.")
raised = [
    {2: "Intake / Portal", 3: "Rename the portal to \"Speak Up\"",
     4: "Rename the portal from \"Whistleblowing Reporting Portal\" to \"Speak Up\" wherever reporters or staff see it: "
        "the help center portal name, the portal page title, the notification emails and the user guide.",
     9: "The portal, the emails and the user guide show \"Speak Up\"; \"Whistleblowing Reporting Portal\" no longer "
        "appears anywhere a reporter or staff member can see it.",
     11: "New",
     12: "Portal name and emails are configuration only. The portal link does not change, so links already "
         "shared keep working."},
    {2: "Intake / Portal", 3: "Anonymity notice next to the anonymous option",
     4: "Next to the anonymity selection option on the reporting form, show this statement: \"" + ANONYMITY_NOTICE + "\"",
     9: "Opening the reporting form shows the statement, word for word, directly next to the anonymity option, "
        "before the reporter submits.",
     11: "New",
     12: "Shown as the help text of the anonymity option. Arabic text to confirm: \"يرجى العلم بأن اختيار عدم الإفصاح "
         "عن هويتك قد يحدّ من قدرتنا على استيضاح الوقائع أو المتابعة معك أثناء التحقيق.\""},
]
for i, req in enumerate(raised):
    for col, value in req.items():
        br.cell(row=first + i, column=col, value=value)

# ---------------------------------------------------------------- SLA
sla = sheet("SLA", NAVY)
sla_headers = ["SLA ID", "SLA name", "Stage measured", "Applies to report category", "Severity",
               "Clock starts when", "Clock stops when", "Clock pauses when", "Target", "Unit", "Calendar",
               "Warning at (% of target)", "Warning - group", "Warning - email",
               "Escalation L1 - after breach", "L1 - group", "L1 - email", "L1 - channel",
               "Escalation L2 - after breach", "L2 - group", "L2 - email", "L2 - channel",
               "Escalation L3 - after breach", "L3 - group", "L3 - email", "L3 - channel", "Notes"]
sla_widths = [14, 28, 22, 30, 14, 30, 30, 22, 9, 15, 26, 13, 22, 26, 16, 22, 26, 16, 16, 22, 26, 16, 16, 22, 26, 16, 30]
banner(sla, "Service Levels (SLA) and Escalation",
       "One row per SLA level - add a row for each category or severity that needs its own target", len(sla_headers))
sla["A4"] = ("Escalation timing is after the breach, in the same unit as the target (e.g. 1 = one working day after the target is missed). "
             "Groups come from the 'Groups & Contacts' sheet; add a personal email only if someone outside a group must be told.")
sla["A4"].font = F(italic=True, color=GREY)
first, last = table(sla, 5, sla_headers, sla_widths,
    ["SLA-01", "Time to Compliance Decision", "Compliance Decision", "All categories", "All severities",
     "Report created, or returned to Back to Compliance", "Under Investigation or Rejected by Compliance", "Never",
     2, "Working days", "ERC working week (Sun-Thu 08:00-16:00)", 0.75, "GRC-Compliance", "",
     0, "HO-Compliance", "doaamokhtar@gmail.com", "Email", 1, "", "", "", "", "", "", "",
     "Today: L1 only, on breach. L1 email is a placeholder until the real Head of Compliance is named."],
    {3: lst("Stage"), 4: lst("ReportCategory"), 5: lst("Severity"), 9: num, 10: lst("Unit"), 11: lst("Calendar"),
     12: pct, 13: {"type": "list", "formula1": "=GroupNames", "allow_blank": True},
     16: {"type": "list", "formula1": "=GroupNames", "allow_blank": True},
     20: {"type": "list", "formula1": "=GroupNames", "allow_blank": True},
     24: {"type": "list", "formula1": "=GroupNames", "allow_blank": True},
     18: lst("Channel"), 22: lst("Channel"), 26: lst("Channel")},
    notes={"Warning at (% of target)": "e.g. 75% sends a warning when three quarters of the time is used, before the breach.",
           "Escalation L1 - after breach": "0 = at the moment of breach. 1 = one unit (e.g. one working day) after.",
           "Severity": "Leave 'All severities' unless targets differ by severity - then add one row per severity.",
           "Applies to report category": "Leave 'All categories' unless a category needs its own target - then add one row per category."})
for col in (14, 17, 21, 25):
    letter = get_column_letter(col)
    rule = dict(email_rule, formula1=f'=ISNUMBER(SEARCH("@",{letter}{first}))')
    dv = DataValidation(**rule); dv.add(f"{letter}{first}:{letter}{last}"); sla.add_data_validation(dv)
for r_ in range(first - 1, last + 1):
    sla.cell(row=r_, column=12).number_format = "0%"
    if r_ >= first:
        sla.cell(row=r_, column=1, value=f'=IF(B{r_}="","","SLA-"&TEXT(ROW()-{first - 2},"00"))').font = F(color=GREY)
        sla.cell(row=r_, column=1).fill = EXAMPLE
# second example: the investigation SLA as configured today
ex2 = ["", "Time to Investigation Decision", "Investigation Decision", "All categories", "All severities",
       "Report enters Under Investigation", "Investigation Approved or Back to Compliance", "Never",
       2, "Working days", "ERC working week (Sun-Thu 08:00-16:00)", None, "", "", 0, "HO-Compliance",
       "doaamokhtar@gmail.com", "Email", None, "", "", "", None, "", "", "", "Configured in the sandbox today."]
for i, v in enumerate(ex2, start=1):
    if i > 1:
        c = sla.cell(row=first, column=i)
        c.value = v
        c.font = F(italic=True, color=GREY)

# ---------------------------------------------------------------- Groups & Contacts
gc = sheet("Groups & Contacts", MINT)
gc_headers = ["Group", "Role in the process", "Member name", "Member email", "Job title / department",
              "Needs Jira agent licence?", "Also receives escalations?", "Notes"]
gc_widths = [32, 40, 26, 30, 28, 16, 16, 30]
banner(gc, "Groups and Contacts", "One row per member - repeat the group name for each member", len(gc_headers))
gc["A4"] = "A person should belong to one investigation group only, so they cannot review the same report twice."
gc["A4"].font = F(italic=True, color=GREY)
groups = [("GRC-Compliance", "Compliance Manager: classifies new reports, approves or rejects them"),
          ("Internal Audit", "Investigates fraud, bribery and procurement reports"),
          ("Human Resources", "Investigates harassment; first approval on retaliation reports"),
          ("Employee Relations", "Second approval on retaliation reports"),
          ("Health and Safety Representative", "Investigates health, safety, environmental and heritage reports"),
          ("Information Security", "Investigates information security and data protection reports"),
          ("HO-Compliance", "Head of Compliance: receives SLA escalations"),
          ("FollowUP TKT", "Oversight: sees all reports")]
gl = len(lists) + 1
ls.cell(row=1, column=gl, value="GroupNames").font = F(bold=True)
for i, (g, _) in enumerate(groups, start=2):
    ls.cell(row=i, column=gl, value=g)
gletter = get_column_letter(gl)
wb.defined_names["GroupNames"] = __import__("openpyxl").workbook.defined_name.DefinedName(
    "GroupNames", attr_text=f"Lists!${gletter}$2:${gletter}${len(groups) + 1}")
first, last = table(gc, 5, gc_headers, gc_widths,
    ["GRC-Compliance", "Compliance Manager", "Fotoun Kurdi", "f.kurdi@erc.sa", "GRC", "Yes", "No", "Example only."],
    {1: {"type": "list", "formula1": "=GroupNames", "allow_blank": True}, 6: lst("YesNo"), 7: lst("YesNo")})
dv = DataValidation(**dict(email_rule, formula1=f'=ISNUMBER(SEARCH("@",D{first}))')); dv.add(f"D{first}:D{last}"); gc.add_data_validation(dv)
# pre-fill one row per group so the customer only adds names
for i, (g, role) in enumerate(groups):
    gc.cell(row=first + i, column=1, value=g)
    gc.cell(row=first + i, column=2, value=role)

wb.move_sheet("Lists", offset=len(wb.sheetnames))
wb.active = 0
for ws in wb.worksheets:
    ws.sheet_view.zoomScale = 90
wb.save(OUT)
print("written", OUT)
