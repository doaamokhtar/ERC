# Raise a Ticket (REST version)

A public page that creates requests in **Anonymous TKT (AT)** without sign-in.

- The page's fields are read from Jira, from request type `474` *Submit a request or incident*.
  If you change a field's name or help text in Jira, the page shows the change within 10 minutes.
- The email field is filled in with `FIXED_EMAIL` and locked. The server ignores whatever the
  browser sends and always raises the request as `REPORTER_ID`.
- The ticket is created through the JSM REST API. The requester sees the ticket key (AT-xx) on
  the success screen.
- Visibility does not change: the permission scheme and the *Internal only* security level
  still decide who sees the ticket.

## Deploy

1. Create the fixed customer once and copy the `REPORTER_ID` it prints:
   ```powershell
   .\tools\ensure-anonymous-customer.ps1 -DryRun
   .\tools\ensure-anonymous-customer.ps1
   ```
2. In script.google.com, open the project (or create *Raise a Ticket*). Paste `Code.gs`, add an HTML
   file named exactly `Index` and paste `Index.html`. Under *Project Settings*, enable
   "Show appsscript.json" and paste `appsscript.json`.
3. In *Project Settings → Script Properties*, add:

   | Property | Value |
   |---|---|
   | `JIRA_USER` | email of the account that owns the token |
   | `JIRA_TOKEN` | API token from id.atlassian.com |
   | `REPORTER_ID` | the accountId from step 1 |

   The token goes here and never in `Code.gs`. Anyone who can edit the script can read it, so
   keep the project's editors to the owner only.
4. Go to *Deploy → Manage deployments → Edit → New version*. Editing the existing deployment
   keeps the same `/exec` link.

## Check

- Open the link in an incognito window. The email is already filled in and greyed out.
- Submit. The success screen shows `AT-xx`, and the ticket in Jira has the fixed customer as
  Reporter and *Internal only* as its security level.
- If it fails: Apps Script → *Executions* shows the exact response from Jira. The page itself
  only shows a general message.
