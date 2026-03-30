# Mail Merge Add-on — Deployment Guide

## Step 1: Install clasp (one-time)

```bash
npm install -g @google/clasp
```

## Step 2: Login to Google (one-time)

```bash
clasp login
```

This opens your browser. Sign in with the Google account where you want the add-on.

## Step 3: Enable Apps Script API (one-time)

Go to: https://script.google.com/home/usersettings
Toggle "Google Apps Script API" to ON.

## Step 4: Create your Google Sheet

1. Go to https://sheets.google.com and create a new spreadsheet
2. Name it something like "Mail Merge"
3. Open **Extensions > Apps Script**
4. In the Apps Script editor URL, copy the Script ID:
   `https://script.google.com/macros/s/d/SCRIPT_ID_HERE/edit`

## Step 5: Connect and push

```bash
cd mail-merge-addon
clasp clone YOUR_SCRIPT_ID
clasp push
```

If `clasp clone` creates duplicate files, just run `clasp push` — it will overwrite.

## Step 6: Deploy as Web App (needed for tracking)

In the Apps Script editor:
1. Click **Deploy > New deployment**
2. Select type: **Web app**
3. Execute as: **Me**
4. Who has access: **Anyone**
5. Click **Deploy**
6. Copy the Web App URL (you'll need this — the code reads it automatically via `ScriptApp.getService().getUrl()`)

## Step 7: Use it

1. Go back to your Google Sheet
2. Refresh the page
3. You'll see a **Mail Merge** menu in the menu bar
4. Click **Mail Merge > Open Sidebar**
5. First time: click **Mail Merge > Setup Sheets** to create the log and template sheets

## Sheet Setup

Your main sheet should have an "Email" column header (required) plus any other columns you want as merge fields:

| Email | First Name | Last Name | Company | Designation |
|-------|-----------|-----------|---------|-------------|
| john@example.com | John | Smith | Acme Corp | CEO |
