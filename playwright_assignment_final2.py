from playwright.sync_api import sync_playwright
import openpyxl
import json
import random
import time
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# SETTINGS
# Change things here instead of hunting through the code below.
# ---------------------------------------------------------------------------
CONTACTS_FILE = "/Users/deepika/Documents/GenAI Program/Playwright/Contacts.xlsx"
SESSION_FOLDER = "wa_session"          # keeps you logged in between runs
SCREENSHOTS_FOLDER = Path("/Users/deepika/Documents/GenAI Program/Playwright/screenshots")
MIN_DELAY_SECONDS = 2                  # random pause range between actions,
MAX_DELAY_SECONDS = 5                  # so the bot doesn't act "robotic"

TODAY = datetime.now().strftime("%Y-%m-%d")
JSON_REPORT_FILE = f"whatsapp_report_{TODAY}.json"
EXCEL_REPORT_FILE = f"whatsapp_report_{TODAY}.xlsx"
DEFAULT_COUNTRY_CODE = "+91"


# ---------------------------------------------------------------------------
# STEP 1: LOGIN
# ---------------------------------------------------------------------------
def login(page):
    """
    Opens WhatsApp Web in the browser.

    First time ever: a QR code appears -> you scan it with your phone.
    Every time after: WhatsApp remembers you (because we reuse the same
    browser profile folder), so it skips straight to your chats.
    """
    page.goto("https://web.whatsapp.com/", wait_until="domcontentloaded", timeout=60000)

    qr_code = "canvas[aria-label='Scan me!']"
    chat_list = "div[role='grid']"

    # Wait for whichever one shows up first: QR code, or the chat list.
    page.wait_for_selector(f"{qr_code}, {chat_list}", timeout=60000)

    if page.locator(qr_code).is_visible():
        print("Please scan the QR code shown in the browser window with your phone.")
        page.wait_for_selector(chat_list, timeout=60000)

    print("Logged in to WhatsApp Web.")


# ---------------------------------------------------------------------------
# STEP 2: READ CONTACTS FROM EXCEL
# ---------------------------------------------------------------------------
def clean_phone_number(raw_value):
    """
    Turns whatever Excel/Sheets gave us for a phone number into a clean,
    WhatsApp-ready number with a country code.
 
    Handles two common issues:
      1. If the Phone column isn't formatted as text, Excel stores the
         number as a float (e.g. 9941419751.0). Converting that straight
         to text with str() would leave a stray ".0" on the end, which
         corrupts the number once WhatsApp strips out the "." (it keeps
         the trailing 0). We convert floats to int first to avoid that.
      2. If the number doesn't already start with "+" (i.e. no country
         code was entered), we add DEFAULT_COUNTRY_CODE automatically.
 
    Examples (with DEFAULT_COUNTRY_CODE = "+91"):
        9941419751.0      (float)  -> "+919941419751"
        "9941419751"                -> "+919941419751"
        "+91 99414 19751"           -> "+919941419751"   (already had a code)
    """
    if isinstance(raw_value, float) and raw_value.is_integer():
        raw_value = int(raw_value)  # 9941419751.0 -> 9941419751, no ".0"
 
    text = str(raw_value).strip()
    digits_and_plus = "".join(ch for ch in text if ch.isdigit() or ch == "+")
 
    if not digits_and_plus.startswith("+"):
        digits_and_plus = DEFAULT_COUNTRY_CODE + digits_and_plus
 
    return digits_and_plus
 
 
def read_contacts(filepath=CONTACTS_FILE):
    """
    Reads the contacts.xlsx file and turns each row into a simple dictionary:
        {"name": "Ravi", "phone": "+919876543210", "message": "Hi {name}!"}
 
    Expects a header row with columns named exactly: Name, Phone, Message.
    """
    workbook = openpyxl.load_workbook(filepath)
    sheet = workbook.active
 
    header_row = [cell.value for cell in sheet[1]]
    name_column = header_row.index("Name")
    phone_column = header_row.index("Phone")
    message_column = header_row.index("Message")
 
    contacts = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        # Skip blank rows (e.g. trailing empty rows at the end of the sheet),
        # which would otherwise show up as a contact named "None".
        if row[name_column] is None or row[phone_column] is None:
            continue
 
        contacts.append({
            "name": str(row[name_column]).strip(),
            "phone": clean_phone_number(row[phone_column]),
            "message": str(row[message_column]).strip() if row[message_column] else "",
        })
 
    return contacts
 
# ---------------------------------------------------------------------------
# STEP 3: SEND MESSAGES (helper functions used inside the main loop)
# ---------------------------------------------------------------------------
def pause_like_a_human():
    """Sleeps for a random amount of time so actions don't fire back-to-back."""
    time.sleep(random.uniform(MIN_DELAY_SECONDS, MAX_DELAY_SECONDS))


def open_chat(page, phone_number):
    """
    Opens a chat directly using WhatsApp's "click to chat" link, which looks like:
        https://web.whatsapp.com/send?phone=919876543210

    Returns:
        True  -> chat opened, message box is ready
        False -> WhatsApp says this phone number is invalid / not on WhatsApp
    """
    digits_only = phone_number.replace("+", "").replace(" ", "")
    page.goto(f"https://web.whatsapp.com/send?phone={digits_only}")

    message_box = "div[data-testid='conversation-compose-box-input']"
    invalid_number_text = "text=Phone number shared via url is invalid."

    try:
        page.wait_for_selector(message_box, timeout=15000)
        return True
    except Exception:
        if page.locator(invalid_number_text).count() > 0:
            return False
        raise 

def send_message(page, message_text):

    """
    Types `message_text` into the open chat and presses Enter to send it.
 
    After sending, it waits for a "sent" tick icon to appear on the message
    bubble. That's how we confirm WhatsApp actually sent it, instead of just
    assuming it worked.
 
    Returns:
        True  -> tick icon appeared, message confirmed sent
        False -> tick never appeared within 10 seconds (send may have failed)
    """""
    # Target the actual editable field, not just anything with this title —
    # WhatsApp sometimes wraps the editable field inside a non-editable div
    # that also carries the same title, and clicking the wrapper doesn't
    # always focus the text field inside it.
    message_box = page.locator("div[data-testid='conversation-compose-box-input']")
    message_box.wait_for(state="visible", timeout=10000)
 
    message_box.click()
    message_box.focus()  # force focus explicitly, don't just rely on the click landing right
    page.wait_for_timeout(300)  # let WhatsApp finish any load animation before we type
 
    # Type the message line by line, so multi-line templates work correctly.
    for line in message_text.split("\n"):
        page.keyboard.type(line, delay=30)
        page.keyboard.press("Shift+Enter")
    page.keyboard.press("Backspace")  # remove the one extra blank line we added above
 
    page.keyboard.press("Enter")  # actually send it
 
    page.wait_for_timeout(500)
    box_is_now_empty = message_box.inner_text().strip() == ""
    return box_is_now_empty
 


def take_screenshot(page, contact_name):
    """
    Saves a screenshot of the most recently sent message bubble.
    Returns the file path so we can record it in the report.
    """
    SCREENSHOTS_FOLDER.mkdir(exist_ok=True)

    safe_filename = "".join(ch if ch.isalnum() else "_" for ch in contact_name)
    timestamp = datetime.now().strftime("%H%M%S")
    save_path = SCREENSHOTS_FOLDER / f"{safe_filename}_{timestamp}.png"

    sent_message_bubbles = page.locator("div.message-out")
    try:
        sent_message_bubbles.last.screenshot(path=str(save_path))
    except Exception:
        # If we can't find the exact bubble for some reason, screenshot
        # the whole page instead of failing the whole contact.
        page.screenshot(path=str(save_path))

    return str(save_path)


def extract_last_messages(page, how_many=3):
    """
    Reads the last few messages visible in the currently open chat
    (a mix of messages we sent and messages they sent).

    Returns a list like:
        [{"direction": "sent", "text": "Hi Ravi!"},
         {"direction": "received", "text": "Thanks!"}]
    """
    all_bubbles = page.locator("div.message-in, div.message-out")
    total_messages = all_bubbles.count()
    first_index_to_read = max(0, total_messages - how_many)

    messages = []
    for i in range(first_index_to_read, total_messages):
        bubble = all_bubbles.nth(i)
        text_parts = bubble.locator("span.selectable-text").all_inner_texts()

        if not text_parts:
            continue  # skip non-text messages (images, stickers, etc.)

        bubble_class = bubble.get_attribute("class") or ""
        direction = "sent" if "message-out" in bubble_class else "received"

        messages.append({"direction": direction, "text": " ".join(text_parts)})

    return messages


def process_one_contact(page, contact):
    """
    Runs the full flow for a single contact: open chat -> send message ->
    screenshot -> extract their last messages. Any error here is caught so
    it doesn't stop the whole batch.

    Returns a dictionary describing exactly what happened, which later gets
    written into the report files.
    """
    name = contact["name"]
    phone = contact["phone"]
    personalized_message = (
        contact["message"].replace("{name}", name) if contact["message"] else f"Hi {name}!"
    )

    result = {
        "name": name,
        "phone": phone,
        "message": personalized_message,
        "status": "failed",       # will be updated below as we go
        "error": None,
        "screenshot": None,
        "last_messages": [],
        "timestamp": datetime.now().isoformat(timespec="seconds"),
    }

    try:
        chat_found = open_chat(page, phone)
        if not chat_found:
            result["status"] = "not_found"
            result["error"] = "Phone number shared via url is invalid."
            return result

        pause_like_a_human()

        was_sent = send_message(page, personalized_message)
        result["status"] = "sent" if was_sent else "failed"
        if not was_sent:
            result["error"] = "Sent-confirmation tick was not detected."

        pause_like_a_human()

        result["screenshot"] = take_screenshot(page, name)
        result["last_messages"] = extract_last_messages(page, how_many=3)

    except Exception as error:
        # Catch-all so one bad contact (missing element, timeout, etc.)
        # never crashes the whole run.
        result["status"] = "failed"
        result["error"] = str(error)

    return result


# ---------------------------------------------------------------------------
# STEP 4: SAVE REPORTS
# ---------------------------------------------------------------------------
def save_reports(all_results):
    """
    Writes two report files summarizing every contact processed:
      - a JSON file with full details (including extracted messages)
      - an Excel file with a simple one-row-per-contact summary
    """
    with open(JSON_REPORT_FILE, "w", encoding="utf-8") as f:
        json.dump({"run_date": TODAY, "contacts": all_results}, f, indent=2, ensure_ascii=False)

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Summary"
    sheet.append(["Name", "Phone", "Status", "Message", "Screenshot", "Timestamp", "Error"])

    for result in all_results:
        sheet.append([
            result["name"],
            result["phone"],
            result["status"],
            result["message"],
            result["screenshot"] or "",
            result["timestamp"],
            result["error"] or "",
        ])

    workbook.save(EXCEL_REPORT_FILE)

    print(f"\nReports saved:\n  {JSON_REPORT_FILE}\n  {EXCEL_REPORT_FILE}")


# ---------------------------------------------------------------------------
# MAIN: ties all 4 steps together
# ---------------------------------------------------------------------------
def main():
    with sync_playwright() as playwright:
        # This reuses the SESSION_FOLDER so you don't have to scan the QR
        # code every single time you run the script.
        context = playwright.chromium.launch_persistent_context(
            SESSION_FOLDER,
            headless=False,
            viewport={"width": 1280, "height": 800},
        )
        page = context.pages[0] if context.pages else context.new_page()

        login(page)
        contacts = read_contacts()

        all_results = []
        for contact in contacts:
            print(f"\nProcessing {contact['name']} ({contact['phone']})...")
            result = process_one_contact(page, contact)
            print(f"  -> {result['status']}")

            all_results.append(result)
            pause_like_a_human()  # wait a bit before moving to the next contact

        context.close()

    save_reports(all_results)


if __name__ == "__main__":
    main()