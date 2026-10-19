// Builds "WB User Guide v1 (DRAFT)" as a Word file in the Expo 2030 Riyadh theme.
// Footer: page number only. Header: ERC icon + document title + DRAFT.
// Usage: node build.js <out.docx> [report-form.png]
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType, ShadingType,
  BorderStyle, AlignmentType, HeadingLevel, Header, Footer, PageNumber, LevelFormat, PageBreak,
  TableOfContents, VerticalAlign, TableLayoutType,
} = require('docx');

const dir = __dirname;
const img = f => fs.readFileSync(path.isAbsolute(f) ? f : path.join(dir, f));
const out = process.argv[2];
const reportForm = process.argv[3] && fs.existsSync(process.argv[3]) ? process.argv[3] : null;

// Expo 2030 Riyadh palette (docs/Color Code.txt + logo)
const DARK = '024638', GREEN = '228848', MINT = '36B78E', INK = '1D2B27', GREY = '5E716B', LINE = 'D6E6DF', TINT = 'EEF7F3', ROW = 'F6FAF8';
const FONT = 'Segoe UI';
const PAGE_W = 11906, MARGIN = 1134, CONTENT_W = PAGE_W - 2 * MARGIN; // A4, 2 cm margins

const t = (text, o = {}) => new TextRun({ text, font: FONT, ...o });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [t(runs)], spacing: { after: 120, line: 288 }, ...o });
const h1 = text => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [t(text)], spacing: { before: 360, after: 160 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: MINT, space: 4 } } });
const h2 = text => new Paragraph({ children: [t(text, { bold: true, size: 24, color: GREEN })], spacing: { before: 200, after: 100 } });
const bullet = (runs, level = 0) => new Paragraph({ numbering: { reference: 'bullets', level }, children: Array.isArray(runs) ? runs : [t(runs)], spacing: { after: 80, line: 276 } });
const step = (runs, level = 0, ref = 'steps') => new Paragraph({ numbering: { reference: ref, level }, children: Array.isArray(runs) ? runs : [t(runs)], spacing: { after: 80, line: 276 } });
const b = text => t(text, { bold: true });

const picture = (file, w, h, maxW = 620) => {
  const scale = Math.min(1, maxW / w);
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 },
    children: [new ImageRun({ type: 'png', data: img(file), transformation: { width: Math.round(w * scale), height: Math.round(h * scale) },
      altText: { title: path.basename(file), description: path.basename(file), name: path.basename(file) } })] });
};
const caption = text => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [t(text, { italics: true, size: 18, color: GREY })] });

const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const borders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };
const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const table = (headers, rows, widths) => {
  const total = widths.reduce((a, c) => a + c, 0);
  const mk = (text, w, fill, run) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.CENTER,
    shading: fill ? { type: ShadingType.CLEAR, color: 'auto', fill } : undefined, margins: { top: 80, bottom: 80, left: 110, right: 110 },
    children: [new Paragraph({ children: [t(text, run)] })] });
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, layout: TableLayoutType.FIXED,
    rows: [new TableRow({ tableHeader: true, cantSplit: true, children: headers.map((x, i) => mk(x, widths[i], DARK, { bold: true, color: 'FFFFFF', size: 19 })) }),
           ...rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((x, i) => mk(x, widths[i], ri % 2 ? ROW : null, { size: 19, color: INK })) }))] });
};
const note = (label, text, color = GREEN, fill = TINT) => new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
  rows: [new TableRow({ cantSplit: true, children: [new TableCell({ width: { size: CONTENT_W, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill }, borders: { top: none, bottom: none, right: none, left: { style: BorderStyle.SINGLE, size: 24, color } },
    margins: { top: 120, bottom: 120, left: 200, right: 160 },
    children: [new Paragraph({ children: [t(label + ' ', { bold: true, color }), t(text, { color: INK })] })] })] })] });
const spacer = (after = 120) => new Paragraph({ spacing: { after }, children: [] });

// ---------- cover ----------
const coverBand = new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
  rows: [new TableRow({ cantSplit: true, children: [new TableCell({ width: { size: CONTENT_W, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: DARK }, borders: { top: none, bottom: none, left: none, right: none },
    margins: { top: 700, bottom: 700, left: 520, right: 520 },
    children: [
      new Paragraph({ spacing: { after: 160 }, children: [t('●  WHISTLEBLOWING · USER GUIDE', { bold: true, color: MINT, size: 22 })] }),
      new Paragraph({ spacing: { after: 80 }, children: [t('Whistleblowing', { bold: true, color: 'FFFFFF', size: 64 })] }),
      new Paragraph({ spacing: { after: 240 }, children: [t('Reporting Portal', { bold: true, color: 'FFFFFF', size: 64 })] }),
      new Paragraph({ spacing: { after: 320 }, children: [t('How to submit, review and investigate reports in Jira Service Management', { color: 'CFE3DB', size: 26 })] }),
      new Paragraph({ children: [t('  DRAFT · SANDBOX  ', { bold: true, color: DARK, size: 20, shading: { type: ShadingType.CLEAR, color: 'auto', fill: MINT } })] }),
    ] })] })] });

const meta = table(['Item', 'Detail'], [
  ['Document', 'Whistleblowing Reporting Portal – User Guide'], ['Version', '1.1 (draft)'], ['Date', '19 October 2026'],
  ['Prepared by', 'Doaa Mokhtar'], ['Client', 'Expo 2030 Riyadh Company (ERC)'],
  ['Environment', 'Sandbox (erc-sd-sandbox.atlassian.net)'], ['Status', 'Draft – to be updated with live ERC data before release'],
], [2600, CONTENT_W - 2600]);

const cover = [
  new Paragraph({ spacing: { after: 500 }, children: [new ImageRun({ type: 'png', data: img('expo-logo.png'), transformation: { width: 280, height: 110 },
    altText: { title: 'Expo 2030 Riyadh', description: 'Expo 2030 Riyadh logo', name: 'logo' } })] }),
  coverBand, spacer(480), meta, new Paragraph({ children: [new PageBreak()] }),
];

// ---------- body ----------
const LINK = 'https://erc-sd-sandbox.atlassian.net/helpcenter/WhistleblowingReporting/portal/169/group/274/create/474?email=Anonymous%40Anonymous.com';
const body = [
  new Paragraph({ children: [t('Contents', { bold: true, size: 32, color: DARK })], spacing: { after: 160 } }),
  new TableOfContents('Contents', { hyperlink: true, headingStyleRange: '1-1' }),
  p([t('Right-click the list and choose Update Field if page numbers are missing.', { italics: true, size: 18, color: GREY })]),
  note('DRAFT.', 'This guide is based on the sandbox environment (erc-sd-sandbox). Links, groups and screenshots will be updated with the live ERC data before release.'),
  new Paragraph({ children: [new PageBreak()] }),

  h1('1. About the portal'),
  p('The Whistleblowing Reporting Portal lets anyone report misconduct at ERC without signing in and without giving their name. Every report is classified first by the Compliance Manager, then investigated by the team responsible for that type of concern.'),
  bullet([b('Anonymous by default. '), t('The email field is pre-filled with Anonymous@Anonymous.com. Leave it as it is to stay anonymous.')]),
  bullet([b('Confidential. '), t('Only the Compliance team and the team assigned to investigate can see a report. Other staff, and other investigation teams, cannot.')]),
  bullet([b('Optional contact. '), t('A reporter who wants to be contacted can leave a phone number; nothing else identifies them.')]),
  spacer(60),
  table(['Who', 'Section to read'], [
    ['Anyone reporting a concern', '2. How to submit a report'], ['Compliance Manager (GRC-Compliance)', '4. For the Compliance Manager'],
    ['Investigation teams', '5. For investigation teams'], ['Head of Compliance (HO-Compliance)', '7. Service levels and escalation'],
  ], [4400, CONTENT_W - 4400]),

  h1('2. How to submit a report'),
  p('Anyone can submit a report in about two minutes, with no account and no sign-in.'),
  step([t('Open the reporting form: '), t(LINK, { color: GREEN, underline: {} })]),
  step('Use this exact link. Opening the portal from its home page leaves the email field empty.', 1),
  step([b('Email confirmation to '), t('is already filled with Anonymous@Anonymous.com. Leave it to stay anonymous, or replace it with your own email.')]),
  step([b('Summary | الموضوع: '), t('a short title, for example "Supplier invoice inflated".')]),
  step([b('Phone number | رقم الجوال '), t('(optional): only if you want the Compliance team to be able to contact you.')]),
  step([b('What are the details of your request? '), t('What happened, when, where and who was involved.')]),
  step([b('Attachments '), t('(optional): documents or photos that support the report.')]),
  step([t('Click '), b('Send'), t('. A reference number such as WB-36 is shown. Keep it if you want to refer to your report later.')]),
  ...(reportForm
    ? [picture(reportForm, 846, 739, 520), caption('Figure 1 – The reporting form with the email pre-filled')]
    : [spacer(60), note('Screenshot to add:', 'the reporting form with the email pre-filled, and the confirmation screen after Send.')]),

  h1('3. What happens after you submit'),
  p('Every report gets two decisions, each within 2 working days: the Compliance Manager decides whether it is investigated, then the assigned team approves or rejects the investigation.'),
  picture('workflow.png', 2000, 952), caption((reportForm ? 'Figure 2' : 'Figure 1') + ' – Report workflow'),
  bullet([b('Approve '), t('and '), b('Reject '), t('are buttons next to the report status. Each opens its own screen: Approve asks for the Investigation Outcome, Reject asks for a Rejection Reason.')]),
  bullet([b('Retaliation reports '), t('need two approvals: Human Resources first, then Employee Relations.')]),
  bullet([t('A team that rejects sends the report '), b('Back to Compliance'), t(' with its reason; the Compliance Manager can correct the classification and approve again, or close it.')]),

  h1('4. For the Compliance Manager'),
  p('The Compliance Manager classifies every new report and decides whether it goes to investigation. You have 2 working days. You receive an email for every new report.'),
  step([t('Open the queue and open a report with status '), b('Open'), t(' (or '), b('Back to Compliance'), t(').')], 0, 'steps2'),
  step([t('Under '), b('Attached forms'), t(', open the '), b('Compliance Approval'), t(' form. It is internal: the reporter never sees it.')], 0, 'steps2'),
  picture('form-attached.png', 960, 520, 560), caption((reportForm ? 'Figure 3' : 'Figure 2') + ' – The Compliance Approval form on a new report, not yet submitted'),
  step([t('Choose '), b('Report Category'), t('.')], 0, 'steps2'),
  step([t('If you choose '), b('Other'), t(', a second question appears: '), b('Investigating Function'), t('. Choose which team investigates. The form will not submit without it.')], 1, 'steps2'),
  step([t('Click '), b('Submit'), t(' on the form. Within a few seconds Investigating Function is filled in from the agreed table.')], 0, 'steps2'),
  picture('form-submitted.png', 960, 590, 560), caption((reportForm ? 'Figure 4' : 'Figure 3') + ' – The form after submission: Report Category recorded'),
  step([t('Click '), b('Approve'), t(' next to the status: the report moves to Under Investigation, is routed to the right team, and that team gets an email.')], 0, 'steps2'),
  step([t('Or click '), b('Reject'), t(': choose a '), b('Rejection Reason'), t(' (required) and add a comment; the report closes as Rejected by Compliance.')], 0, 'steps2'),
  spacer(60),
  note('Good to know:', 'Approve is blocked until the form is submitted. A submitted form is read-only: to change the classification of a report that came Back to Compliance, open the form menu (⋯) and choose Reopen, then submit it again. Investigating Function and Investigation Group cannot be edited directly.'),

  h1('5. For investigation teams'),
  p('An investigation team sees a report only after the Compliance Manager approves it and routes it to them. You have 2 working days to decide.'),
  step([t('You receive an email: '), b('New report assigned to your team'), t('. Click Open the report, or find it in the queue.')], 0, 'steps3'),
  step('Read the report and its attachments.', 0, 'steps3'),
  step([t('Click '), b('Approve'), t(' next to the status. Choose the '), b('Investigation Outcome'), t(' (Substantiated, Partially substantiated or Not substantiated), add a comment and confirm. The report moves to Investigation Approved.')], 0, 'steps3'),
  step([t('Or click '), b('Reject'), t(', choose a '), b('Rejection Reason'), t(' (insufficient information, duplicate, out of scope or Wrong category), add a comment and confirm. The report returns to the Compliance Manager as Back to Compliance.')], 0, 'steps3'),
  h2('Retaliation reports'),
  bullet([b('Human Resources '), t('approves first, with the Investigation Outcome. The report moves to '), b('Pending Employee Relations Approval'), t('.')]),
  bullet([b('Employee Relations '), t('then approves (the report moves to Investigation Approved) or rejects with a reason (Back to Compliance).')]),
  spacer(60),
  note('Good to know:', 'Neither Approve nor Reject can be saved without its required field. Teams cannot change the category, the investigating function, the group or the form; if the category is wrong, reject with "Wrong category". Teams cannot see reports assigned to other teams.'),
  spacer(60),
  note('Screenshot to add:', 'the Approve screen with Investigation Outcome, and the Reject screen with Rejection Reason.'),

  h1('6. Groups: who sees what and who can do what'),
  p('Access comes only from group membership; each investigation team sees only the reports routed to it.'),
  table(['Group', 'Role', 'Sees', 'Can do'], [
    ['GRC-Compliance', 'Compliance Manager', 'All reports', 'Fill and submit the Compliance Approval form, Approve or Reject new reports, receives the new-report email'],
    ['Internal Audit', 'Investigation team', 'Reports routed to Internal Audit', 'Approve with an Investigation Outcome, Reject with a reason, receives the assignment email'],
    ['Human Resources', 'Investigation team', 'Reports routed to HR, including Retaliation', 'Same as Internal Audit; first approval on Retaliation'],
    ['Employee Relations', 'Investigation team', 'Retaliation reports', 'Second approval or rejection on Retaliation'],
    ['Health and Safety Representative', 'Investigation team', 'Reports routed to Health and Safety', 'Same as Internal Audit'],
    ['Information Security', 'Investigation team', 'Reports routed to Information Security', 'Same as Internal Audit'],
    ['HO-Compliance', 'Head of Compliance', 'Escalation emails', 'Receives an escalation email when a stage passes its 2-working-day target'],
    ['FollowUP TKT', 'Oversight', 'All reports', 'View and edit; keep this group small'],
  ], [2000, 1650, 2400, CONTENT_W - 6050]),
  spacer(160),
  p('Which team investigates is set by the report category:'),
  table(['Report category', 'Investigating function', 'Group'], [
    ['Fraud, Theft, Embezzlement or Asset Misappropriation', 'Internal Audit', 'Internal Audit'],
    ['Bribery, Corruption, Kickbacks or Trading in Influence', 'Internal Audit', 'Internal Audit'],
    ['Money Laundering, Terrorist Financing or Sanctions Concerns', 'GRC-Compliance Manager', 'GRC-Compliance'],
    ['Procurement Manipulation, Collusion, Bid Rigging or Conflict of Interest', 'Internal Audit', 'Internal Audit'],
    ['Harassment, Discrimination, Bullying or Abuse of Authority', 'Human Resources', 'Human Resources'],
    ['Health, Safety, Environmental or Heritage-Protection Breach', 'Health and Safety representative', 'Health and Safety Representative'],
    ['Information Security, Data Protection or Systems Misuse', 'Information Security', 'Information Security'],
    ['Other Regulatory or License Compliance Breach', 'GRC-Compliance Manager', 'GRC-Compliance'],
    ['Retaliation Against a Reporter or Witness', 'Human Resources', 'Human Resources, then Employee Relations'],
    ['Other', 'Chosen by the Compliance Manager in the form', 'Follows the chosen function'],
  ], [4300, 2600, CONTENT_W - 6900]),
  spacer(120),
  note('For administrators:', 'every member needs a Jira Service Management agent licence. GRC-Compliance must also be in the Service Desk Team role of the project, or its members cannot edit the form. Do not add investigation groups to that role, and do not put one person in more than one investigation group unless intended.'),

  h1('7. Service levels and escalation'),
  p('Each decision stage has 2 working days (16 working hours, Sunday to Thursday, 08:00–16:00 Riyadh time). When a stage passes its target, the Head of Compliance is emailed automatically.'),
  table(['Service level', 'Clock runs while the report is', 'Target', 'When the target is missed'], [
    ['Time to Compliance Decision', 'Open, or Back to Compliance (restarts on each return)', '2 working days', 'Escalation email to HO-Compliance'],
    ['Time to Investigation Decision', 'Under Investigation and Pending Employee Relations Approval', '2 working days', 'Escalation email to HO-Compliance'],
  ], [2400, 3200, 1500, CONTENT_W - 7100]),
  spacer(120),
  bullet('Fridays, Saturdays and the hours outside 08:00–16:00 do not count.'),
  bullet('Every automatic email (new report, assignment, escalation) shows only the reference number, the status and who it is waiting on, with a link. It never includes the report’s content.'),
  bullet('The time remaining for each service level is shown on the report in Jira.'),
];

// ---------- header / footer ----------
const header = new Header({ children: [
  new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: MINT, space: 6 } }, tabStops: [{ type: 'right', position: CONTENT_W }],
    children: [new ImageRun({ type: 'png', data: img('erc-icon.png'), transformation: { width: 36, height: 32 }, altText: { title: 'ERC', description: 'ERC icon', name: 'icon' } }),
               t('\tWhistleblowing Reporting Portal – User Guide  ', { size: 16, color: GREY }),
               t('|  DRAFT', { size: 16, bold: true, color: GREEN })] }),
] });
const footer = new Footer({ children: [ new Paragraph({ alignment: AlignmentType.CENTER,
  children: [t('Page ', { size: 16, color: GREY }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GREY }),
             t(' of ', { size: 16, color: GREY }), new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: GREY })] }) ] });

const numbering = { config: [
  { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { run: { color: GREEN }, paragraph: { indent: { left: 540, hanging: 270 } } } },
                                   { level: 1, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 270 } } } }] },
  ...['steps', 'steps2', 'steps3'].map(ref => ({ reference: ref, levels: [
    { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { run: { bold: true, color: DARK }, paragraph: { indent: { left: 540, hanging: 360 } } } },
    { level: 1, format: LevelFormat.LOWER_LETTER, text: '%2.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 360 } } } }] })),
] };

const doc = new Document({
  creator: 'Doaa Mokhtar', title: 'Whistleblowing Reporting Portal – User Guide v1.1 (Draft)', description: 'Draft user guide based on the sandbox environment',
  styles: { default: { document: { run: { font: FONT, size: 21, color: INK } } },
    paragraphStyles: [{ id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 32, bold: true, color: DARK }, paragraph: { outlineLevel: 0 } }] },
  numbering, features: { updateFields: true },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: 1300, bottom: 1100, left: MARGIN, right: MARGIN, header: 500, footer: 500 } }, titlePage: true },
    headers: { default: header, first: new Header({ children: [new Paragraph({ children: [] })] }) },
    footers: { default: footer, first: new Footer({ children: [new Paragraph({ children: [] })] }) },
    children: [...cover, ...body],
  }],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(out, buf); console.log('written', out, buf.length, 'bytes', reportForm ? '(with report form screenshot)' : '(report form screenshot pending)'); });
