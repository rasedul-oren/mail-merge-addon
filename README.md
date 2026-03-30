# Mail Merge Add-on for Google Sheets

A Google Apps Script add-on that lets you send personalised bulk emails directly from Google Sheets — with open & click tracking, reusable templates, bounce detection, and a live dashboard.

## Features

- **Personalised Bulk Email** — Merge any spreadsheet column into your email body and subject line.
- **Open & Click Tracking** — Know who opened your email and which links they clicked, logged back into your sheet.
- **Reusable Templates** — Save, manage, and switch between email templates with merge fields.
- **Live Dashboard** — View delivery stats, open rates, and bounce counts from the sidebar.
- **Bounce Detection** — Automatically flags bounced emails to keep your list clean.
- **Simple Setup** — Install with `clasp push`, refresh your sheet, and go. No external services required.

## Quick Start

1. Install [clasp](https://github.com/google/clasp): `npm install -g @google/clasp`
2. Run `clasp login` and sign in with your Google account.
3. Create a Google Sheet and open **Extensions > Apps Script**.
4. Copy the Script ID from the Apps Script editor URL.
5. Run `clasp clone YOUR_SCRIPT_ID` then `clasp push`.
6. Deploy as a **Web App** for tracking support (see [DEPLOY.md](DEPLOY.md) for full instructions).
7. Refresh your sheet — the **Mail Merge** menu appears automatically.

## Sheet Format

Your main sheet needs an `Email` column header plus any columns you want as merge fields:

| Email | First Name | Last Name | Company | Designation |
|-------|-----------|-----------|---------|-------------|
| john@example.com | John | Smith | Acme Corp | CEO |

## Built By

[**Oren**](https://www.orennow.com/) — ESG & Sustainability Consulting
