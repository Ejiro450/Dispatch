# EAM Dispatch Board (installable web app)

A standalone version of the dispatch board. It signs in with Microsoft, reads the admin
Outlook calendar, and saves truck, crew and subcontractor choices back to Outlook as hidden
tags (the visible calendar entries are never changed). No database and no server code.

## Files
- `index.html` the app
- `config.js` the only file to edit (Microsoft IDs)
- `msal-browser.min.js` Microsoft's sign-in library (v4.30.0)
- `manifest.webmanifest`, `sw.js`, `icons/` make it installable and let it open offline

## Setup, once
1. **Host the folder** on any HTTPS static host (Cloudflare Pages, GitHub Pages, Netlify).
   Note the address, including the trailing slash, e.g. `https://dispatch.eamglobalmovers.com.au/`.
2. **Register the app with Microsoft** at entra.microsoft.com, signed in as a Microsoft 365 admin:
   - App registrations, New registration. Name: EAM Dispatch Board. Accounts in this organisation only.
   - Redirect URI: platform **Single-page application (SPA)**, value = the address from step 1.
   - API permissions, Add, Microsoft Graph, Delegated: `Calendars.ReadWrite`. Then Grant admin consent.
   - Copy the **Application (client) ID** and **Directory (tenant) ID** from the Overview page.
3. **Edit `config.js`**: paste the client ID into `clientId` and the tenant ID into `tenant`. Upload it.
4. Open the address, sign in as admin@eamglobalmovers.com.au, and check today's jobs appear.

If the app is opened before step 3 it shows the exact redirect address to register.

## Installing on a phone
- iPhone: open the address in Safari, Share, Add to Home Screen.
- Android: open it in Chrome, menu, Install app (or Add to Home screen).

## Notes
- Everyone signs in with the admin Outlook account. If staff should use their own accounts,
  set `mailbox` in `config.js` to the admin address, share that calendar with them with edit
  rights, and add the `Calendars.ReadWrite.Shared` permission instead.
- The trucks and crew list is stored on the calendar itself; job assignments are stored on each event.
- To ship an update: replace the files and bump `VERSION` in `sw.js`.
