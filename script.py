import os
import sys
import win32com.client
from openpyxl import Workbook
from datetime import datetime

# ── helpers ──────────────────────────────────────────────────────────────────

def get_test_folder(inbox):
    """Return the 'test' subfolder of *inbox*, or None if not found."""
    for folder in inbox.Folders:
        if folder.Name.strip().lower() == "test":
            return folder
    return None

def iter_mail_items(folder):
    """Yield MailItem objects from *folder* (skips non-mail items)."""
    items = folder.Items
    items.Sort("[ReceivedTime]", True)          # newest first
    item = items.GetFirst()
    while item:
        try:
            # olMailItem = 43
            if item.Class == 43:
                yield item
        except Exception:
            pass
        item = items.GetNext()

# ── main ─────────────────────────────────────────────────────────────────────

def main():
    # Connect to a running Outlook instance (or launch one)
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mapi    = outlook.GetNamespace("MAPI")
    except Exception as exc:
        sys.exit(f"[ERROR] Could not connect to Outlook: {exc}")

    # olFolderInbox = 6  →  always the *current user's* inbox, never shared
    try:
        inbox = mapi.GetDefaultFolder(6)
    except Exception as exc:
        sys.exit(f"[ERROR] Could not open default Inbox: {exc}")

    # Locate the 'test' subfolder
    test_folder = get_test_folder(inbox)
    if test_folder is None:
        sys.exit("[ERROR] No subfolder named 'test' was found inside your Inbox.")

    print(f"[INFO] Reading emails from: Inbox \\ {test_folder.Name}")

    # Collect email data
    rows = []
    for mail in iter_mail_items(test_folder):
        try:
            subject      = mail.Subject or "(no subject)"
            received     = mail.ReceivedTime          # pywintypes.datetime
            date_str     = received.strftime("%Y-%m-%d")
            time_str     = received.strftime("%H:%M:%S")
            rows.append((subject, date_str, time_str))
        except Exception as e:
            print(f"[WARNING] Skipped one item: {e}")

    print(f"[INFO] Found {len(rows)} email(s).")

    # Build the output path
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    out_dir   = os.path.join(downloads, "test folder")
    os.makedirs(out_dir, exist_ok=True)
    out_path  = os.path.join(out_dir, "outlook_emails.xlsx")

    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Emails"

    ws.append(["Subject", "Date", "Time"])

    for row in rows:
        ws.append(list(row))

    # Save
    wb.save(out_path)
    print(f"[OK] Excel file saved to: {out_path}")


if __name__ == "__main__":
    main()