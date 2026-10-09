# Unreplied Email Triage - Apps Script

Google Apps Script project for triaging Dominic Ford's unreplied emails.

## Files

- `EmailTriage.gs` - Main script (paste as Code.gs in Apps Script editor)
- `appsscript.json` - Project manifest (Project Settings > show appsscript.json)

## Setup

1. Create a standalone Apps Script project at https://script.google.com
2. Paste EmailTriage.gs as Code.gs
3. Paste appsscript.json as the manifest
4. Enable the Gmail API advanced service (Services > Gmail API)
5. Set Script Properties:
   - GEMINI_API_KEY
   - AVA_WEBHOOK_URL
   - AVA_WEBHOOK_SECRET (optional)
6. Run setup() once and approve scopes
7. setup() creates Email_System_DB spreadsheet, labels, and 15-minute trigger

## Architecture

- Runs every 15 minutes via time trigger
- Searches Gmail for `label:inbox -label:Ava/Processed`
- Deterministic filters run first (skip patterns, already replied, calendar invites)
- Unknown senders classified with Gemini Flash
- Threads needing Dominic's reply get Ava/Needs-Reply label + webhook POST
- All state tracked in Email_System_DB spreadsheet

See the code header in EmailTriage.gs for full documentation.
