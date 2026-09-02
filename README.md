# WhatsApp Web Automation Bot

Sends personalized WhatsApp messages to a list of contacts and extracts their last few replies — built with Playwright.

## What it does
1. Logs into WhatsApp Web (scan the QR code once; session is cached).
2. Reads contacts from `contacts.xlsx`.
3. For each contact: opens their chat, sends a personalized message, confirms it sent, screenshots it, and grabs their last 3 messages.
4. Saves a full report (JSON) and a summary (Excel), both dated.

## Setup
```bash
pip install playwright openpyxl
playwright install chromium
```

## contacts.xlsx format
| Name | Phone | Message |
|------|-------|---------|
| Manoj | 9941419751 | Hi {name}, good morning! |

- **Phone**: with or without country code. If missing, `+91` is added automatically (change `DEFAULT_COUNTRY_CODE` in the script for a different default).
- **Message**: optional. `{name}` is replaced with the contact's name. Leave blank to use a default greeting.
- Format the Phone column as **Text** in Excel/Sheets to avoid numbers getting mangled.

## Run
```bash
python playwright_assign.py
```
First run: scan the QR code shown in the browser. Later runs reuse the saved session (`wa_session/` folder).

## Output
- `whatsapp_report_YYYY-MM-DD.json` — full details per contact, including extracted messages.
- `whatsapp_report_YYYY-MM-DD.xlsx` — one-row-per-contact summary (status, message, screenshot path, error).
- `screenshots/` — a screenshot of each successfully sent message.

## Notes
- Random 2–5s delays between actions to avoid looking automated.
- Failed/not-found contacts are logged with an error, not skipped silently — one bad contact won't crash the run.
- WhatsApp Web's HTML changes periodically. If something stops working, right-click the relevant element in the browser → **Inspect**, and update the matching selector in the script.
