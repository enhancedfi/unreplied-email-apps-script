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
   - AVA_WEBHOOK_SECRET (optional HMAC key)

   `GEMINI_API_KEY` and `AVA_WEBHOOK_SECRET` are read only from Script Properties. Values pasted into the Config sheet for those keys are ignored. When the secret is set, the webhook signs `timestamp + "." + raw JSON body` with HMAC-SHA256 and sends `X-Webhook-Timestamp` plus `X-Webhook-Signature: sha256=<hex>`. The raw secret is not sent. Receivers should reject timestamps older than five minutes.
6. Run setup() once and approve scopes
7. setup() creates Email_System_DB spreadsheet, labels, and 15-minute trigger

## Architecture

- Runs every 15 minutes via time trigger
- Searches Gmail for `label:inbox -label:Ava/Processed`
- Before that search, strips `Ava/Processed` from inbox threads that received a newer inbound message (so a later client reply is triaged again)
- Deterministic filters run first (skip patterns, already replied, calendar invites)
- A thread counts as already replied only when Dominic's latest message is newer than the latest inbound message
- Unknown senders classified with Gemini Flash
- Threads needing Dominic's reply get Ava/Needs-Reply label + webhook POST
- All state tracked in Email_System_DB spreadsheet

See the code header in EmailTriage.gs for full documentation.
