// Builds "WB SLA Requirements Template" in the Expo 2030 Riyadh theme: a form the customer (ERC) fills in.
// Grey italic cells = today's sandbox values, given as examples; empty cells are for the customer.
// Usage: node sla-template.js <out.docx>
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType, ShadingType,
  BorderStyle, AlignmentType, HeadingLevel, Header, Footer, PageNumber, LevelFormat, PageBreak,
  VerticalAlign, TableLayoutType, PageOrientation,
} = require('docx');

const dir = __dirname;
const img = f => fs.readFileSync(path.join(dir, f));
const out = process.argv[2];

const NAVY = '024638', MINT = '36B78E', MINT_DARK = '228848', INK = '1D2B27', GREY = '5E716B', LINE = 'D6E6DF', FILL_ME = 'FFF8D6';  // Expo 2030 Riyadh palette
const FONT = 'Segoe UI';
const PAGE_W = 11906, MARGIN = 1134, CONTENT_W = PAGE_W - 2 * MARGIN;

const t = (text, o = {}) => new TextRun({ text, font: FONT, ...o });
const ex = text => t(text, { italics: true, color: GREY });          // example value
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [t(runs)], spacing: { after: 120, line: 288 }, ...o });
const h1 = text => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [t(text)], spacing: { before: 360, after: 160 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: MINT, space: 4 } } });
const h2 = text => new Paragraph({ children: [t(text, { bold: true, size: 24, color: NAVY })], spacing: { before: 240, after: 100 } });
const bullet = runs => new Paragraph({ numbering: { reference: 'bullets', level: 0 }, children: Array.isArray(runs) ? runs : [t(runs)], spacing: { after: 80, line: 276 } });
const b = text => t(text, { bold: true });
const spacer = (after = 120) => new Paragraph({ spacing: { after }, children: [] });
const box = '☐ ';

const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const borders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };
// cell value: string = plain, {ex:'..'} = grey example, {fill:true} = empty box to fill
const cellRuns = v => v == null ? [] : typeof v === 'string' ? [t(v, { size: 18, color: INK })] : v.ex ? [ex(v.ex)].map(r => r) : [];
const table = (headers, rows, widths) => {
  const total = widths.reduce((a, c) => a + c, 0);
  const mk = (v, w, head, odd) => new TableCell({
    width: { size: w, type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.CENTER,
    shading: head ? { type: ShadingType.CLEAR, color: 'auto', fill: NAVY } : (v && v.fill) ? { type: ShadingType.CLEAR, color: 'auto', fill: FILL_ME } : odd ? { type: ShadingType.CLEAR, color: 'auto', fill: 'F6FAF8' } : undefined,
    margins: { top: 70, bottom: 70, left: 100, right: 100 },
    children: [new Paragraph({ children: head ? [t(v, { bold: true, color: 'FFFFFF', size: 18 })] : (v && v.ex) ? [t(v.ex, { italics: true, color: GREY, size: 18 })] : cellRuns(v) })] });
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, layout: TableLayoutType.FIXED,
    rows: [new TableRow({ tableHeader: true, cantSplit: true, children: headers.map((x, i) => mk(x, widths[i], true)) }),
           ...rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((x, i) => mk(x, widths[i], false, ri % 2 === 1)) }))] });
};
const F = { fill: true };            // empty cell for the customer
const E = s => ({ ex: s });          // example value
const blank = (n, cols) => Array.from({ length: n }, () => Array.from({ length: cols }, () => F));

const note = (label, text) => new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
  rows: [new TableRow({ cantSplit: true, children: [new TableCell({ width: { size: CONTENT_W, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: 'EEF7F3' },
    borders: { top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
               left: { style: BorderStyle.SINGLE, size: 24, color: MINT_DARK } },
    margins: { top: 120, bottom: 120, left: 200, right: 160 },
    children: [new Paragraph({ children: [t(label + ' ', { bold: true, color: MINT_DARK }), t(text, { color: INK })] })] })] })] });
const choice = (q, options) => [p([b(q)], { spacing: { before: 120, after: 60 } }), ...options.map(o => p([t(box + o)], { indent: { left: 360 }, spacing: { after: 40 } }))];

// ---------- cover ----------
const coverBand = new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
  rows: [new TableRow({ cantSplit: true, children: [new TableCell({ width: { size: CONTENT_W, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: NAVY },
    borders: { top: { style: BorderStyle.NONE, size: 0, color: NAVY }, bottom: { style: BorderStyle.NONE, size: 0, color: NAVY }, left: { style: BorderStyle.NONE, size: 0, color: NAVY }, right: { style: BorderStyle.NONE, size: 0, color: NAVY } },
    margins: { top: 700, bottom: 700, left: 520, right: 520 },
    children: [
      new Paragraph({ spacing: { after: 160 }, children: [t('●  SERVICE DESK · REQUIREMENTS TEMPLATE', { bold: true, color: MINT, size: 22 })] }),
      new Paragraph({ spacing: { after: 80 }, children: [t('Service Level (SLA)', { bold: true, color: 'FFFFFF', size: 60 })] }),
      new Paragraph({ spacing: { after: 240 }, children: [t('Requirements', { bold: true, color: 'FFFFFF', size: 60 })] }),
      new Paragraph({ spacing: { after: 320 }, children: [t('Speak Up – targets, levels and escalation, to be completed by ERC', { color: 'CFE3DB', size: 26 })] }),
      new Paragraph({ children: [t('  TO BE COMPLETED BY THE CUSTOMER  ', { bold: true, color: NAVY, size: 20, shading: { type: ShadingType.CLEAR, color: 'auto', fill: MINT } })] }),
    ] })] })] });

const cover = [
  new Paragraph({ spacing: { after: 600 }, children: [new ImageRun({ type: 'png', data: img('expo-logo.png'), transformation: { width: 280, height: 110 }, altText: { title: 'Expo 2030 Riyadh', description: 'Expo 2030 Riyadh logo', name: 'logo' } })] }),
  coverBand, spacer(480),
  table(['Item', 'Detail'], [
    ['Customer', 'Expo 2030 Riyadh Company (ERC)'], ['Service', 'Speak Up (Jira Service Management)'],
    ['Template version', '1.0'], ['Prepared by', 'Doaa Mokhtar'],
    ['Completed by (name, role)', F], ['Date completed', F], ['Approved by (name, role)', F],
  ], [3000, CONTENT_W - 3000]),
  new Paragraph({ children: [new PageBreak()] }),
];

// ---------- body ----------
const W4 = [2300, 2600, 2500, CONTENT_W - 7400];
const body = [
  h1('1. How to fill in this template'),
  bullet([b('Pale yellow cells '), t('are for you to complete.')]),
  bullet([b('Grey italic values '), t('show what is configured today in the sandbox. Keep them, change them, or delete them.')]),
  bullet([t('Tick '), b(box.trim()), t(' the options that apply. Add rows wherever you need more levels.')]),
  bullet([t('Targets are measured in '), b('working time'), t(' on the calendar in section 2, unless you state otherwise.')]),
  spacer(60),
  note('Why levels matter:', 'one target for every report is simple, but most organisations want shorter targets for serious concerns (for example retaliation or fraud) and a staged escalation: a warning before the target, then escalation to a manager, then to a higher level if still unresolved.'),

  h1('2. Working calendar'),
  table(['Item', 'Your answer', 'Today (example)'], [
    ['Working days', F, E('Sunday – Thursday')],
    ['Working hours', F, E('08:00 – 16:00')],
    ['Time zone', F, E('Riyadh (GMT+3)')],
    ['Do weekends and holidays count?', F, E('No')],
    ['Is a 24/7 calendar needed for any level?', F, E('No')],
  ], [3200, 3300, CONTENT_W - 6500]),
  h2('Official holidays to exclude'),
  table(['Holiday', 'From', 'To'], [[E('Eid al-Fitr'), F, F], [E('Eid al-Adha'), F, F], [E('Founding Day'), F, F], [E('National Day'), F, F], ...blank(2, 3)],
    [4200, 2400, CONTENT_W - 6600]),

  h1('3. What is measured (SLA metrics)'),
  p('Each row is one clock. Add a row for every additional stage you want measured.'),
  table(['SLA name', 'Clock starts when', 'Clock stops when', 'Clock pauses when'], [
    [E('Time to Compliance Decision'), E('Report created, or returned to Back to Compliance'), E('Under Investigation or Rejected by Compliance'), E('Never')],
    [E('Time to Investigation Decision'), E('Report enters Under Investigation'), E('Investigation Approved or Back to Compliance'), E('Never')],
    [E('e.g. Time to first contact with reporter'), F, F, F],
    [E('e.g. Time to close the case'), F, F, F],
    ...blank(2, 4),
  ], W4),
  spacer(60),
  ...choice('Should any clock pause while waiting for the reporter or a third party?', ['No pause', 'Pause while waiting for the reporter (requires a "Waiting for reporter" status)', 'Pause while waiting for a third party', 'Other: ____________________']),

  h1('4. Targets by level'),
  ...choice('How should targets differ? (tick one)', ['One target for all reports (today)', 'By severity or priority (fill in table 4A)', 'By report category (fill in table 4B)', 'By both severity and category']),
  h2('4A. Targets by severity'),
  table(['Severity', 'Definition', 'Compliance Decision target', 'Investigation Decision target'], [
    [E('Critical'), E('e.g. immediate risk to people, large fraud'), F, F],
    [E('High'), F, F, F],
    [E('Medium'), F, E('2 working days'), E('2 working days')],
    [E('Low'), F, F, F],
  ], [1600, 3400, 2300, CONTENT_W - 7300]),
  h2('4B. Targets by report category'),
  table(['Report category', 'Compliance Decision target', 'Investigation Decision target'], [
    ['Fraud, Theft, Embezzlement or Asset Misappropriation', F, F],
    ['Bribery, Corruption, Kickbacks or Trading in Influence', F, F],
    ['Money Laundering, Terrorist Financing or Sanctions Concerns', F, F],
    ['Procurement Manipulation, Collusion, Bid Rigging or Conflict of Interest', F, F],
    ['Harassment, Discrimination, Bullying or Abuse of Authority', F, F],
    ['Health, Safety, Environmental or Heritage-Protection Breach', F, F],
    ['Information Security, Data Protection or Systems Misuse', F, F],
    ['Other Regulatory or License Compliance Breach', F, F],
    ['Retaliation Against a Reporter or Witness', F, F],
    ['Other', F, F],
  ], [5200, 2200, CONTENT_W - 7400]),
  spacer(60),
  ...choice('Who sets the severity of a report?', ['Compliance Manager, in the Compliance Approval form', 'Set automatically from the category', 'Other: ____________________']),

  h1('5. Warnings and escalation levels'),
  p('Fill in one table per SLA. Each row is one level. Today a single level exists: an email to HO-Compliance when the target is missed.'),
  ...['Time to Compliance Decision', 'Time to Investigation Decision'].flatMap((sla, i) => [
    h2(`5${String.fromCharCode(65 + i)}. ${sla}`),
    table(['Level', 'When', 'Notify (role / group / email)', 'Channel', 'Action expected'], [
      [E('Warning'), E('e.g. 75% of target used'), E('e.g. the assigned group'), E('Email'), E('Act before breach')],
      [E('Level 1'), E('On breach'), E('HO-Compliance (today: Doaa Mokhtar)'), E('Email'), E('Follow up with the owner')],
      [E('Level 2'), E('e.g. 1 working day after breach'), F, F, F],
      [E('Level 3'), E('e.g. 3 working days after breach'), F, F, F],
      ...blank(1, 5),
    ], [1200, 2100, 2700, 1200, CONTENT_W - 7200]),
  ]),
  spacer(60),
  note('Confidentiality:', 'escalation emails carry only the reference number, status, owner and a link. Please confirm below if any recipient outside Jira (for example a board member) must receive them, as they would need Jira access to open the report.'),
  spacer(60),
  table(['Recipient outside the Compliance team', 'Role', 'Email', 'Needs Jira access? (Y/N)'], blank(3, 4), [2800, 2000, 2600, CONTENT_W - 7400]),

  h1('6. Reporting'),
  table(['Item', 'Your answer', 'Today (example)'], [
    ['SLA report needed?', F, E('Not configured')],
    ['Frequency', F, E('e.g. monthly')],
    ['Recipients', F, F],
    ['Content', F, E('e.g. % within target per SLA, breaches by category and group')],
    ['Format', F, E('e.g. Jira dashboard / emailed PDF')],
  ], [3200, 3300, CONTENT_W - 6500]),

  h1('7. Other requirements'),
  table(['#', 'Requirement'], blank(4, 2), [700, CONTENT_W - 700]),

  h1('8. Sign-off'),
  table(['Role', 'Name', 'Signature', 'Date'], [
    ['Customer (ERC) – Head of Compliance', F, F, F],
    ['Customer (ERC) – IT / Service owner', F, F, F],
    ['Ejada – Implementation', E('Doaa Mokhtar'), F, F],
  ], [3600, 2400, 2000, CONTENT_W - 8000]),
];

// ---------- header / footer ----------
const header = new Header({ children: [ new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: MINT, space: 6 } }, tabStops: [{ type: 'right', position: CONTENT_W }],
  children: [new ImageRun({ type: 'png', data: img('erc-icon.png'), transformation: { width: 36, height: 32 }, altText: { title: 'ERC', description: 'ERC icon', name: 'icon' } }),
             t('\tSpeak Up – SLA Requirements Template', { size: 16, color: GREY })] }) ] });
const footer = new Footer({ children: [ new Paragraph({ alignment: AlignmentType.CENTER,
  children: [t('Page ', { size: 16, color: GREY }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GREY }),
             t(' of ', { size: 16, color: GREY }), new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: GREY })] }) ] });

const doc = new Document({
  creator: 'Doaa Mokhtar', title: 'Speak Up – SLA Requirements Template',
  styles: { default: { document: { run: { font: FONT, size: 20, color: INK } } },
    paragraphStyles: [{ id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 30, bold: true, color: NAVY }, paragraph: { outlineLevel: 0 } }] },
  numbering: { config: [{ reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { run: { color: MINT_DARK }, paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: 1300, bottom: 1100, left: MARGIN, right: MARGIN, header: 500, footer: 500 } }, titlePage: true },
    headers: { default: header, first: new Header({ children: [new Paragraph({ children: [] })] }) },
    footers: { default: footer, first: new Footer({ children: [new Paragraph({ children: [] })] }) },
    children: [...cover, ...body],
  }],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(out, buf); console.log('written', out, buf.length, 'bytes'); });
