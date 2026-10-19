// HTML bodies for the WB automation emails, in the Expo 2030 Riyadh palette.
// Table layout + inline styles so Outlook and Gmail render them the same.
// Confidentiality: only the key, status, owner and a link - never the report's content.
const DARK = '#024638', GREEN = '#228848', MINT = '#36B78E', INK = '#1D2B27', MUTED = '#5E716B', LINE = '#D6E6DF', TINT = '#EEF7F3';

const shell = ({ eyebrow, title, intro, rows, button, en, ar, alert }) => `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#F3F7F5;padding:28px 12px;font-family:Segoe UI,Tahoma,Arial,sans-serif;">
<tr><td align="center">
<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;width:100%;background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid ${LINE};">
<tr><td style="background:${DARK};padding:22px 30px;color:#ffffff;">
<div style="font-size:12px;letter-spacing:1.5px;color:${MINT};font-weight:bold;">${eyebrow}</div>
<div style="font-size:21px;font-weight:bold;margin-top:6px;">${title}</div>
</td></tr>
<tr><td style="height:4px;background:${alert ? '#C0392B' : GREEN};line-height:4px;font-size:0;">&nbsp;</td></tr>
<tr><td style="padding:24px 30px 6px 30px;color:${INK};font-size:15px;line-height:1.6;" dir="ltr">${intro}</td></tr>
<tr><td style="padding:4px 30px 8px 30px;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid ${LINE};border-radius:10px;background:${TINT};">
${rows.map(([k, v, strong], i) => `<tr><td style="padding:12px 16px;color:${MUTED};font-size:13px;width:42%;${i < rows.length - 1 ? `border-bottom:1px solid ${LINE};` : ''}">${k}</td>
<td style="padding:12px 16px;color:${strong ? DARK : INK};font-size:${strong ? 18 : 14}px;font-weight:bold;${i < rows.length - 1 ? `border-bottom:1px solid ${LINE};` : ''}">${v}</td></tr>`).join('\n')}
</table></td></tr>
<tr><td align="center" style="padding:18px 30px 8px 30px;">
<a href="{{issue.url}}" style="display:inline-block;background:${GREEN};color:#ffffff;text-decoration:none;font-weight:bold;font-size:15px;padding:12px 28px;border-radius:8px;">${button}</a>
</td></tr>
<tr><td style="padding:6px 30px 4px 30px;color:${MUTED};font-size:12px;line-height:1.6;" dir="ltr">${en}</td></tr>
<tr><td style="padding:0 30px;"><div style="border-top:1px solid ${LINE};margin:14px 0;"></div></td></tr>
<tr><td style="padding:0 30px 20px 30px;color:${INK};font-size:15px;line-height:1.8;text-align:right;" dir="rtl">${ar}</td></tr>
<tr><td style="background:${DARK};padding:14px 30px;color:#CFE3DB;font-size:12px;line-height:1.6;text-align:center;">
<span style="color:${MINT};font-weight:bold;">EXPO 2030 RIYADH</span> &middot; Whistleblowing Reporting Portal<br>
For confidentiality, report details are available in Jira only. This is an automated message.<br>
<span dir="rtl">حفاظًا على السرية، تفاصيل البلاغ متاحة داخل Jira فقط. هذه رسالة آلية.</span>
</td></tr>
</table></td></tr></table>`;

exports.newReport = () => shell({
  eyebrow: 'WHISTLEBLOWING REPORTING PORTAL', title: 'New report waiting for review',
  intro: '<p style="margin:0 0 12px 0;">Dear Compliance team,</p><p style="margin:0;">A new report has been submitted. Please complete the <b>Compliance Approval</b> form and approve or reject it within 2 working days.</p>',
  rows: [['Reference', '{{issue.key}}', true], ['Status', '{{issue.status.name}}']],
  button: 'Open the report &nbsp;|&nbsp; فتح البلاغ',
  en: 'You are receiving this because you are a member of GRC-Compliance.',
  ar: '<p style="margin:0 0 10px 0;">فريق الالتزام الكريم،</p><p style="margin:0;">تم تقديم بلاغ جديد برقم <b>{{issue.key}}</b>. يرجى استكمال نموذج <b>Compliance Approval</b> واعتماده أو رفضه خلال يومي عمل.</p>',
});

exports.assigned = () => shell({
  eyebrow: 'WHISTLEBLOWING REPORTING PORTAL', title: 'New report assigned to your team',
  intro: '<p style="margin:0 0 12px 0;">Dear team,</p><p style="margin:0;">The Compliance Manager has approved a report and assigned it to your team for investigation.</p>',
  rows: [['Reference', '{{issue.key}}', true], ['Investigating function', '{{issue.customfield_12000.value}}']],
  button: 'Open the report &nbsp;|&nbsp; فتح البلاغ',
  en: 'Use <b>Approve</b> (record the Investigation Outcome) or <b>Reject</b> (choose a Rejection Reason) from the status button on the report. Target: 2 working days.',
  ar: '<p style="margin:0 0 10px 0;">فريقنا الكريم،</p><p style="margin:0 0 10px 0;">اعتمد مدير الالتزام البلاغ رقم <b>{{issue.key}}</b> وتم إسناده إلى فريقكم للتحقيق.</p><p style="margin:0;color:' + MUTED + ';font-size:13px;">استخدم زر <b>Approve</b> مع تسجيل نتيجة التحقيق، أو <b>Reject</b> مع اختيار سبب الرفض، خلال يومي عمل.</p>',
});

exports.escalation = (stage, owner) => shell({
  eyebrow: 'WHISTLEBLOWING REPORTING PORTAL &middot; ESCALATION', title: `SLA breached: ${stage}`, alert: true,
  intro: `<p style="margin:0 0 12px 0;">Dear Head of Compliance,</p><p style="margin:0;">A report has passed its 2-working-day target at the <b>${stage}</b> stage and needs your attention.</p>`,
  rows: [['Reference', '{{issue.key}}', true], ['Current status', '{{issue.status.name}}'], ['Waiting on', owner]],
  button: 'Open the report &nbsp;|&nbsp; فتح البلاغ',
  en: 'This escalation is sent to the HO-Compliance group.',
  ar: `<p style="margin:0 0 10px 0;">رئيس الالتزام الكريم،</p><p style="margin:0;">تجاوز البلاغ رقم <b>{{issue.key}}</b> المهلة المحددة بيومي عمل في مرحلة <b>${stage}</b> ويحتاج إلى متابعتكم.</p>`,
});
