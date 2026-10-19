/**
 * Raise a Ticket - Google Apps Script web app (REST version)
 *
 * The form is not defined here. It is read from Jira: request type "Submit a request or
 * incident" in Anonymous TKT (AT). Rename a field or change its help text in Jira and the
 * page follows within FORM_CACHE_SECONDS, so the portal and this page cannot drift apart.
 *
 * Each submission is created through the JSM REST API and raised on behalf of one fixed
 * customer (FIXED_EMAIL). The requester never types an email, and the email the page shows
 * is not trusted from the browser - the server always uses REPORTER_ID.
 *
 * Script Properties (Project Settings -> Script Properties), never in this file:
 *   JIRA_USER     email of the Atlassian account that owns the API token
 *   JIRA_TOKEN    API token (id.atlassian.com -> Security -> API tokens)
 *   REPORTER_ID   accountId of FIXED_EMAIL - tools/ensure-anonymous-customer.ps1 prints it
 */

const JIRA_BASE = 'https://erc-sd.atlassian.net';
const SERVICE_DESK_ID = '169';   // Anonymous TKT (AT)
const REQUEST_TYPE_ID = '474';   // Submit a request or incident | تقديم طلب أو بلاغ
const FIXED_EMAIL = 'Anonymous@Anonymous.com';

const MAX_FILES = 3;
const MAX_FILE_BYTES = 5 * 1024 * 1024;
const FORM_CACHE_SECONDS = 600;
const RATE_LIMIT = 20;            // submissions per RATE_WINDOW_SECONDS, whole form
const RATE_WINDOW_SECONDS = 600;

function doGet() {
  const page = HtmlService.createTemplateFromFile('Index');
  // Escaped so a "</script>" inside a Jira help text cannot close the page's script tag.
  page.formJson = JSON.stringify(getForm_()).replace(/</g, '\\u003c');
  page.fixedEmail = FIXED_EMAIL;
  return page.evaluate()
    .setTitle('Raise a Ticket')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Called from the page with google.script.run.submitTicket(payload)
 * payload = { values: { fieldId: text }, website, files: [{ name, type, data(base64) }] }
 */
function submitTicket(payload) {
  if (!payload) throw new Error('Empty request.');

  // Honeypot: real users never fill this hidden field. Pretend it worked.
  if (payload.website) return { ok: true, key: null };

  // Only fields Jira says belong to the request type are sent - never what the browser chose.
  const form = getForm_();
  const input = payload.values || {};
  const values = {};
  form.fields.forEach(function (f) {
    if (f.kind === 'files') return;
    const v = clean_(input[f.id], f.id === 'summary' ? 255 : 5000);
    if (f.required && !v) throw new Error(english_(f.name) + ' is required.');
    if (v) values[f.id] = v;
  });

  // Decoded and checked before the ticket exists, so a bad file never leaves half a request.
  const files = (payload.files || []).slice(0, MAX_FILES).map(function (f) {
    const bytes = Utilities.base64Decode(f.data);
    if (bytes.length > MAX_FILE_BYTES) throw new Error('File "' + f.name + '" is larger than 5 MB.');
    return Utilities.newBlob(bytes, f.type || 'application/octet-stream', clean_(f.name, 120));
  });

  rateLimit_();

  const created = jira_('POST', '/rest/servicedeskapi/request', {
    serviceDeskId: SERVICE_DESK_ID,
    requestTypeId: REQUEST_TYPE_ID,
    requestFieldValues: values,
    raiseOnBehalfOf: prop_('REPORTER_ID')
  });

  // The ticket exists from here on. A failed upload is reported, not thrown: throwing would
  // tell the user to submit again and leave the team with two copies of the same request.
  const failed = [];
  files.forEach(function (blob) {
    try {
      jira_('POST', '/rest/api/2/issue/' + created.issueKey + '/attachments', null, { file: blob });
    } catch (e) {
      failed.push(blob.getName());
    }
  });

  return { ok: true, key: created.issueKey, failedFiles: failed };
}

function getForm_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('form');
  if (hit) return JSON.parse(hit);

  const base = '/rest/servicedeskapi/servicedesk/' + SERVICE_DESK_ID + '/requesttype/' + REQUEST_TYPE_ID;
  const rt = jira_('GET', base);
  const def = jira_('GET', base + '/field');

  const form = {
    name: rt.name,
    description: rt.description || '',
    fields: def.requestTypeFields
      .filter(function (f) { return f.visible !== false && kind_(f); })
      .map(function (f) {
        return { id: f.fieldId, name: f.name, description: f.description || '', required: !!f.required, kind: kind_(f) };
      })
  };
  cache.put('form', JSON.stringify(form), FORM_CACHE_SECONDS);
  return form;
}

/** How the page draws a Jira field. null = a type this page does not support, so it is left out. */
function kind_(f) {
  const s = f.jiraSchema || {};
  if (s.system === 'attachment') return 'files';
  if (s.system === 'summary') return 'text';
  if (s.system === 'description' || /:textarea$/.test(s.custom || '')) return 'textarea';
  if (s.type === 'string' && !s.custom) return 'text';
  if (/:textfield$/.test(s.custom || '')) return 'text';
  return null;
}

function rateLimit_() {
  const cache = CacheService.getScriptCache();
  const count = Number(cache.get('rate') || 0);
  if (count >= RATE_LIMIT) throw new Error('Too many requests right now. Please try again in a few minutes.');
  cache.put('rate', String(count + 1), RATE_WINDOW_SECONDS);
}

function jira_(method, path, body, multipart) {
  const auth = Utilities.base64Encode(prop_('JIRA_USER') + ':' + prop_('JIRA_TOKEN'), Utilities.Charset.UTF_8);
  const opts = {
    method: method,
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Basic ' + auth,
      Accept: 'application/json',
      'X-Atlassian-Token': 'no-check',   // required by the attachments endpoint
      'X-ExperimentalApi': 'opt-in'
    }
  };
  if (multipart) {
    opts.payload = multipart;            // an object holding a Blob is sent as multipart/form-data
  } else if (body) {
    opts.contentType = 'application/json; charset=utf-8';
    opts.payload = JSON.stringify(body);
  }

  const res = UrlFetchApp.fetch(JIRA_BASE + path, opts);
  const code = res.getResponseCode();
  if (code >= 300) {
    // The detail goes to Executions, not to the page: it can name fields, accounts and ids.
    console.error(method + ' ' + path + ' -> ' + code + ' ' + res.getContentText());
    throw new Error('The service desk could not accept the request right now. Please try again later.');
  }
  const text = res.getContentText();
  return text ? JSON.parse(text) : null;
}

function prop_(name) {
  const v = PropertiesService.getScriptProperties().getProperty(name);
  if (!v) throw new Error('Script property ' + name + ' is not set.');
  return v;
}

/** "Summary | الموضوع" -> "Summary" for error messages. */
function english_(label) {
  return String(label).split('|')[0].trim();
}

function clean_(value, max) {
  return String(value || '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max);
}
