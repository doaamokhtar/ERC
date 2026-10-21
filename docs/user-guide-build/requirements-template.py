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
             ("BR-006 Impact", "What the escalation branch (BR-006) changes in today's setup, and the decisions needed - answer in the yellow 'ERC answer' cells."),
             ("BR-007 Impact", "The same for the five new investigation functions (BR-007)."),
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
REPORTABLE_MATTERS = "\n".join([
    "A report may concern wrongdoing that has occurred, is occurring, or is likely to occur. "
    "Reportable matters include, without limitation:",
    "Fraud, theft, embezzlement, misappropriation or misuse of assets, funds or information.",
    "Bribery, kickbacks, facilitation payments, trading in influence, extortion or any other form of corruption.",
    "Money laundering, terrorist financing, sanctions evasion or suspicious financial activity.",
    "Collusion, bid rigging, conflict of interest, favoritism or abuse of position in procurement, contracting, land, "
    "permitting, licensing or recruitment decisions.",
    "Falsification, concealment, destruction or manipulation of records, accounts, reports or evidence.",
    "Breach of law, regulation, license conditions, contractual obligation or internal policy.",
    "Harassment, bullying, discrimination, abuse of authority or other serious workplace misconduct.",
    "Health, safety or environmental danger, and breaches of construction, heritage or site-protection requirements.",
    "Information security breaches, unauthorized disclosure of confidential or personal data, and misuse of systems.",
    "Deliberate circumvention or override of internal controls.",
    "Retaliation against a whistleblower, witness or any person assisting a case, or any attempt to obstruct or "
    "interfere with a case.",
])
NEW_FUNCTIONS = ["Legal", "Procurement", "Finance", "Cybersecurity", "ARC Chairman"]
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
    {2: "Intake / Portal", 3: "\"What can I report?\" guidance behind an info button",
     4: "The portal should present a summary of reportable matters to guide reporters. It must not be static text on "
        "the form: an info icon or button (e.g. \"What can I report?\" or an \"i\" next to the title) opens the guidance "
        "in BR-005 as an expandable panel, popup or tooltip, and it can be collapsed again so it does not clutter the form.",
     9: "The form shows only the button; clicking or tapping it shows the BR-005 text in full; it can be closed again; "
        "it works on desktop and mobile.",
     11: "New",
     12: "Jira's portal has no native info-icon popup. Closest standard option: a \"What can I report?\" toggle on the "
         "form that shows and hides a section holding the text. A true icon or popup needs a portal customization - "
         "to agree with ERC."},
    {2: "Intake / Portal", 3: "Reportable matters - guidance text",
     4: "Text to show in the \"What can I report?\" guidance (BR-004), word for word:\n" + REPORTABLE_MATTERS,
     9: "The guidance shows this text word for word, including all eleven items.",
     11: "New",
     12: "Arabic version to be provided or confirmed by ERC, as the form is bilingual."},
    {2: "Routing", 3: "Escalation questions on the reporting page (new Internal Audit branch)",
     4: "Add two questions to the reporting page. This changes the existing workflow and adds a new branch.\n"
        "Q1: \"Is the reported case related to the Board of Directors?\"\n"
        "- Yes: the report goes to the Internal Audit team. Flow: Reject, or Under Investigation and then Closed.\n"
        "- No: ask Q2.\n"
        "Q2: \"Is it related to the Governance, Risk & Compliance or Cybersecurity department?\"\n"
        "- Yes: the report goes to Internal Audit.\n"
        "- No: continue with the normal workflow.\n"
        "Impact on current areas: see the 'BR-006 Impact' sheet.",
     9: "A Yes to either question sends the report straight to Internal Audit, with no Compliance review, and "
        "GRC-Compliance can neither see it nor receive an email about it. A No to both follows today's workflow unchanged.",
     11: "New",
     12: "11 impacted areas and the decisions ERC needs to make are listed on the 'BR-006 Impact' sheet."},
    {2: "Routing", 3: "Five more investigation functions under Report Category \"Other\"",
     4: "When the Report Category is \"Other\", the Compliance Manager chooses the Investigating Function. "
        "Add these functions, each with its own investigation group:\n" + "\n".join(
            f"{n}. {name}" for n, name in enumerate(NEW_FUNCTIONS, start=1)) + "\n"
        "This affects the investigation group, the assignee and the user groups. "
        "Impact on current areas: see the 'BR-007 Impact' sheet.",
     9: "Choosing each new function on a test report routes it to that function's group only: its members get the "
        "assignment email, can be assigned and can see the report; no other investigation group can.",
     11: "New",
     12: "Groups added to the 'Groups & Contacts' sheet for ERC to fill in the members. Decisions needed are on "
         "the 'BR-007 Impact' sheet."},
]
for i, req in enumerate(raised):
    for col, value in req.items():
        br.cell(row=first + i, column=col, value=value)

# ---------------------------------------------------------------- BR-006 Impact
# What the escalation branch touches in today's setup. Escalated reports concern GRC or the Board,
# so most of the rows come down to one rule: nothing about them may reach GRC-Compliance.
def impact_sheet(name, title, subtitle, note, rows):
    """One row per area of today's setup: what it is, what changes, what ERC must decide, and ERC's answer.
    A row may carry a fifth item - an answer already agreed - which is shown instead of a yellow cell."""
    im = sheet(name, NAVY)
    headers = ["#", "Area", "Today", "Change needed", "Decision needed from ERC", "ERC answer"]
    widths = [5, 24, 42, 52, 46, 40]
    banner(im, title, subtitle, len(headers))
    im["A4"] = note
    im["A4"].font = F(italic=True, color=GREY)
    im.row_dimensions[5].height = 24
    for i, h in enumerate(headers, start=1):
        c = im.cell(row=5, column=i, value=h)
        c.font = F(bold=True, color="FFFFFF"); c.fill = HEAD; c.border = BOX
        c.alignment = Alignment(vertical="center")
        im.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    for n, row in enumerate(rows, start=1):
        r_ = 5 + n
        for i, v in enumerate((n,) + row[:4], start=1):
            c = im.cell(row=r_, column=i, value=v)
            c.font = F(color=INK, bold=(i == 2)); c.border = BOX; c.alignment = WRAP
            if n % 2 == 0: c.fill = BAND
        agreed = row[4] if len(row) > 4 else None
        answer = im.cell(row=r_, column=6, value=agreed)
        answer.border = BOX; answer.alignment = WRAP
        answer.font = F(color=MINT_DARK, bold=True) if agreed else F(color=INK)
        answer.fill = BAND if agreed else (FILL_ME if row[3] != "-" else EXAMPLE)
    im.freeze_panes = "C6"


impact_006 = [
    ("Reporting form",
     "One request type: summary, phone number, details, attachments and the pre-filled anonymous email.",
     "Add the two Yes/No questions (English and Arabic). Q2 shows only when Q1 is No. Both mandatory. The conditional "
     "question needs a form on the portal.",
     "Confirm the Arabic wording and that both questions are mandatory."),
    ("Workflow",
     "Open -> Compliance review -> Under Investigation -> Investigation Approved. A team rejection goes Back to Compliance.",
     "New branch for escalated reports: Open -> Under Investigation (Internal Audit) -> Closed, or Rejected with a "
     "reason. No path back to Compliance on this branch.",
     "Should Internal Audit record an Investigation Outcome when closing, as on the normal branch?"),
    ("Routing (automation)",
     "The investigation group is set from the Report Category when the Compliance Manager approves.",
     "New rule on submission: if either answer is Yes, assign Internal Audit at once and skip Compliance review. "
     "Today's routing rule is unchanged for the normal branch.",
     "-"),
    ("Confidentiality (issue security)",
     "GRC-Compliance can see every report. FollowUP TKT sees all reports.",
     "Escalated reports get a new security level that GRC-Compliance cannot see (Internal Audit only).",
     "Who else may see escalated reports (e.g. the Audit Committee)? Should FollowUP TKT still see them?"),
    ("Notifications",
     "Every new report emails GRC-Compliance. The reporter gets an acknowledgement.",
     "Escalated reports email Internal Audit only; the GRC new-report rule must skip them. The reporter's "
     "acknowledgement does not change.",
     "-"),
    ("SLA and escalation",
     "Compliance decision: 2 working days. Investigation decision: 2 working days. Breaches email HO-Compliance.",
     "The Compliance SLA does not apply. Breaches on escalated reports must not go to HO-Compliance.",
     "Who receives SLA escalations for escalated reports? Is 2 working days right for Internal Audit's decision?"),
    ("Classification",
     "The Compliance Manager records the Report Category in the Compliance Approval form.",
     "Escalated reports skip that form, so no Report Category is recorded unless Internal Audit sets it.",
     "Should Internal Audit classify escalated reports, so they appear in the statistics by category?"),
    ("Queues and reporting (incl. Power BI)",
     "Queues and reports cover all reports, by category.",
     "New queue: Escalated - Internal Audit. Reports and Power BI see escalated cases only if their account is "
     "allowed on the new security level.",
     "Should escalated cases appear in management reporting, and in how much detail?"),
    ("Reports about Internal Audit",
     "-",
     "The two questions do not cover a report about Internal Audit itself; it would follow the normal flow to GRC.",
     "Is that acceptable, or is a third question needed?",
     "Agreed (21/10/2026): acceptable. A report about Internal Audit follows the normal workflow as built today; "
     "no third question."),
    ("Documentation and training",
     "User guide v1.1, workflow diagram and SLA template describe one branch.",
     "Update the guide, the diagram and the SLA template; brief Internal Audit on the new branch.",
     "-"),
    ("Reports already submitted",
     "-",
     "No change: the questions apply to new reports only.",
     "-"),
]
impact_sheet("BR-006 Impact", "BR-006 - Impact of the escalation branch",
             "What changes in the current setup when a report concerns the Board, GRC or Cybersecurity",
             "Escalated reports concern GRC itself, so the rule behind most rows is: nothing about them may reach "
             "GRC-Compliance - not the report, not an email, not an SLA escalation.", impact_006)

# ---------------------------------------------------------------- BR-007 Impact
# Access to a report follows its Investigation Group field (security level and assignable-user permission),
# so a new function mostly means a new group, a routing entry and licensed members.
impact_007 = [
    ("Investigating Function (form question)",
     "Shown when the Report Category is Other. Options: Internal Audit, Human Resources, Health and Safety "
     "representative, Information Security, GRC-Compliance Manager.",
     "Add the five options to the Compliance Approval form and to the Investigating Function field: "
     + ", ".join(NEW_FUNCTIONS) + ".",
     "Confirm the Arabic labels for the five options."),
    ("Investigation Group",
     "Set automatically from the Investigating Function when the Compliance Manager submits the form; "
     "one group per function.",
     "Five new Jira groups, and the five function-to-group pairs added to the routing rule.",
     "Group names. Jira groups are shared across ERC's whole Atlassian organisation: a generic name such as "
     "\"Finance\" may already exist with all finance staff in it, and every one of them would see these reports. "
     "Dedicated names (e.g. \"WB-Finance\") are safer."),
    ("Group members",
     "-",
     "One row per new group added to the 'Groups & Contacts' sheet.",
     "Members and emails for each group."),
    ("Assignee",
     "Members of GRC-Compliance and of the report's Investigation Group can be assigned.",
     "No permission change: members of a new group can be assigned as soon as a report is routed to it.",
     "Who in each group should be assigned first, if anyone (default assignee)?"),
    ("Confidentiality (issue security)",
     "A report is visible to GRC-Compliance and to its Investigation Group only.",
     "No change to the security level: each new group sees only the reports routed to it.",
     "-"),
    ("Licences and sign-in",
     "Every investigator needs a Jira Service Management agent licence and must be assigned to the "
     "\"Jira - SSO\" enterprise app in Entra ID.",
     "Members of the five groups need agent licences and the SSO assignment before go-live.",
     "Are enough agent licences available for the new members?"),
    ("ARC Chairman",
     "-",
     "A single person, possibly outside ERC's staff directory. As an investigation function they would need a "
     "Jira account, a licence, and would see the full report in Jira.",
     "Should the ARC Chairman work in Jira, or receive these cases another way (e.g. through Internal Audit)?"),
    ("Cybersecurity and Information Security",
     "Information Security investigates information security and data protection reports.",
     "Cybersecurity is added as a separate function. Under BR-006, a report about the Cybersecurity department "
     "goes to Internal Audit, so Cybersecurity never investigates a report about itself.",
     "Is Cybersecurity a different team from Information Security, or the same group?"),
    ("Notifications",
     "The assignment email goes to the report's Investigation Group.",
     "Works for the new groups once they are in the routing rule.",
     "-"),
    ("SLA and escalation",
     "Investigation decision: 2 working days, for every investigation group.",
     "Applies to the new groups automatically.",
     "-"),
    ("Documentation and reporting",
     "The user guide lists the groups and the category-to-function table. Power BI breaks reports down by group.",
     "Update the guide's tables; the new functions appear in reporting once used.",
     "-"),
]
impact_sheet("BR-007 Impact", "BR-007 - Impact of the five new investigation functions",
             "What changes when Legal, Procurement, Finance, Cybersecurity and ARC Chairman can investigate",
             "Access to a report follows its Investigation Group, so each new function needs a group, a routing entry "
             "and licensed members - nothing else in the workflow changes.", impact_007)

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
          ("FollowUP TKT", "Oversight: sees all reports")] + [
          (f, f"Investigates reports routed to {f} (Report Category Other) - new, BR-007") for f in NEW_FUNCTIONS]
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
