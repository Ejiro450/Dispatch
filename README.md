# EAM Global Movers app (installable web app)

One app with two sections that share a Microsoft sign-in:

- **Dispatch**: trucks, crews and subcontractors against the jobs in the admin Outlook calendar.
- **Quotes**: the Quote Builder (new quote, saved quotes, reports, settings, PDFs, job sheets).

Everything is stored in the company's Outlook account. There is no database and no server code.

## Files
- `index.html` the app shell and the dispatch board
- `quotes.html`, `quotes-bridge.js` the Quote Builder and its Outlook storage
- `config.js` the only file to edit (Microsoft IDs). Do not overwrite it when updating.
- `msal-browser.min.js` Microsoft's sign-in library (v4.30.0)
- `manifest.webmanifest`, `sw.js`, `icons/` make it installable and let it open offline

## What is deliberately NOT in these files
Rates, off-peak deals, deposit rules, wording and bank details. They are loaded from Outlook
after sign-in. The first time the Quotes section is opened it asks for the settings file
(`eam-quote-settings.json`). Keep that file private. Never upload it to GitHub.

## Where things are kept in Outlook
- Truck, crew and subcontractor choices: hidden tags on each job in the main calendar.
- Trucks and crew list: a hidden tag on the main calendar.
- Saved quotes and Quote Builder settings: entries in a separate calendar called **EAM Quotes**
  (created automatically). Do not edit or delete those entries by hand.
- "Add to calendar" in Quotes creates the booking in the main calendar, where the dispatch
  board picks it up.

## Setup, once
1. **Host the folder** on any HTTPS static host. Note the address, including the trailing slash.
2. **Register the app with Microsoft** at entra.microsoft.com, signed in as a Microsoft 365 admin:
   - App registrations, New registration. Name: EAM Dispatch Board. Accounts in this organisation only.
   - Redirect URI: platform **Single-page application (SPA)**, value = the address from step 1.
   - API permissions, Add, Microsoft Graph, Delegated: `Calendars.ReadWrite`. Then Grant admin consent.
   - Copy the **Application (client) ID** and **Directory (tenant) ID** from the Overview page.
3. **Edit `config.js`**: paste the client ID into `clientId` and the tenant ID into `tenant`.
4. Open the address and sign in as the admin Outlook account.
5. Open **Quotes** and choose `eam-quote-settings.json` when asked.

## Installing on a phone
- iPhone: open the address in Safari, Share, Add to Home Screen.
- Android: open it in Chrome, menu, Install app (or Add to Home screen).

## Updating
Replace the changed files (never `config.js`) and bump `VERSION` in `sw.js`.
