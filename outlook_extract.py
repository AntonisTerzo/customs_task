import os
import tkinter as tk
from tkinter import messagebox
import win32com.client
from openpyxl import Workbook


FORMULA_CHARS = {"=", "+", "-", "@", "\t", "\r"}

# BE task configuration: (folder name in Inbox, sheet name in Excel)
BE_FOLDERS = [
    ("POA Registration",     "POA"),
    ("BEBRU Export",         "BEBRU Export"),
    ("BEBRU 3rdparty ECS",   "ECS"),
]
# ── Shared helpers ────────────────────────────────────────────────────────────

def sanitize_subject(subject):
    """Replace a leading formula-injection character with underscore."""
    if subject and subject[0] in FORMULA_CHARS:
        subject = "_" + subject[1:]
    return subject


def resolve_output_path(out_dir, base_name):
    """
    Returns a unique file path inside out_dir.
    If <base_name>.xlsx already exists, tries <base_name>(1).xlsx,
    (2).xlsx, and so on until a free slot is found.
    """
    candidate = os.path.join(out_dir, f"{base_name}.xlsx")
    if not os.path.exists(candidate):
        return candidate
    counter = 1
    while True:
        candidate = os.path.join(out_dir, f"{base_name}({counter}).xlsx")
        if not os.path.exists(candidate):
            return candidate
        counter += 1

def read_folder_emails(folder):
    """
    Returns (rows, skipped) for a given Outlook folder.
    rows is a list of (subject, datetime_str) tuples.
    """
    rows = []
    skipped = 0
    items = folder.Items
    items.Sort("[ReceivedTime]", True)
    for i in range(items.Count):
        try:
            item = items.Item(i + 1)
            if item.Class == 43:              # olMailItem = 43
                subject      = sanitize_subject(item.Subject or "(no subject)")
                received     = item.ReceivedTime
                datetime_str = received.strftime("%Y/%m/%d %H:%M:%S")
                rows.append((subject, datetime_str))
        except Exception:
            skipped += 1
            continue
    return rows, skipped

def connect_outlook():
    """Returns (mapi, error_message)."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mapi = outlook.GetNamespace("MAPI")
        mapi.Logon()
        return mapi, None
    except Exception as exc:
        return None, f"Could not connect to Outlook:\n{exc}"

# ── PGTS task ─────────────────────────────────────────────────────────────────

def get_pgts_folder(mapi):
    """
    Navigates to Archive/PGTS in the personal mailbox only.
    Returns (folder, error_message). If found, error_message is None.
    """
    try:
        inbox = mapi.GetDefaultFolder(6)
        store_root = inbox.Parent          # root of the personal mailbox
    except Exception as exc:
        return None, f"Could not access personal mailbox:\n{exc}"

    archive_folder = None
    for i in range(store_root.Folders.Count):
        f = store_root.Folders.Item(i + 1)
        if f.Name.strip().lower() == "archive":
            archive_folder = f
            break

    if archive_folder is None:
        return None, "The 'Archive' folder was not found in your personal mailbox."

    for i in range(archive_folder.Folders.Count):
        f = archive_folder.Folders.Item(i + 1)
        if f.Name.strip().upper() == "PGTS":
            return f, None

    return None, "The 'PGTS' subfolder was not found inside Archive."

def run_pgts_task():
    pgts_btn.config(state=tk.DISABLED, text="Running...")
    be_btn.config(state=tk.DISABLED)
    root.update()

    mapi, err = connect_outlook()
    if err:
        messagebox.showerror("Error", err)
        reset_buttons()
        return

    pgts_folder, err = get_pgts_folder(mapi)
    if pgts_folder is None:
        messagebox.showerror("Folder Not Found", err)
        reset_buttons()
        return

    try:
        rows, skipped = read_folder_emails(pgts_folder)
    except Exception as exc:
        messagebox.showerror("Error", f"Failed to read emails:\n{exc}")
        reset_buttons()
        return

    if not rows:
        messagebox.showwarning(
            "Empty Folder",
            "The Archive/PGTS folder was found but contains no emails."
        )
        reset_buttons()
        return

    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    out_dir   = os.path.join(downloads, "PGTS")
    os.makedirs(out_dir, exist_ok=True)
    out_path  = resolve_output_path(out_dir, "outlook_emails")

    wb = Workbook()
    ws = wb.active
    ws.title = "Emails"
    ws.append(["Subject", "Received"])
    for row in rows:
        ws.append(list(row))
    wb.save(out_path)

    summary = f"{len(rows)} email(s) exported to:\n{out_path}"
    if skipped:
        summary += f"\n\n{skipped} item(s) were skipped (non-email items)."
    messagebox.showinfo("Done", summary)
    reset_buttons()

# ── BE task ───────────────────────────────────────────────────────────────────

def find_inbox_subfolder(inbox, target_name):
    """
    Finds a direct subfolder of Inbox by name (case-insensitive, whitespace-tolerant).
    Returns the folder object or None.
    """
    target = target_name.strip().lower()
    for i in range(inbox.Folders.Count):
        f = inbox.Folders.Item(i + 1)
        if f.Name.strip().lower() == target:
            return f
    return None


def run_be_task():
    be_btn.config(state=tk.DISABLED, text="Running...")
    pgts_btn.config(state=tk.DISABLED)
    root.update()

    mapi, err = connect_outlook()
    if err:
        messagebox.showerror("Error", err)
        reset_buttons()
        return

    try:
        inbox = mapi.GetDefaultFolder(6)
    except Exception as exc:
        messagebox.showerror("Error", f"Could not access Inbox:\n{exc}")
        reset_buttons()
        return

    # Build workbook with one sheet per configured folder
    wb = Workbook()
    wb.remove(wb.active)      # remove default empty sheet

    report_lines = []
    total_rows = 0
    total_skipped = 0

    for folder_name, sheet_name in BE_FOLDERS:
        ws = wb.create_sheet(title=sheet_name)
        ws.append(["Subject", "Received"])

        folder = find_inbox_subfolder(inbox, folder_name)
        if folder is None:
            report_lines.append(f"- {folder_name}: NOT FOUND (empty sheet created)")
            continue

        try:
            rows, skipped = read_folder_emails(folder)
        except Exception as exc:
            report_lines.append(f"- {folder_name}: ERROR reading ({exc})")
            continue

        for row in rows:
            ws.append(list(row))

        total_rows += len(rows)
        total_skipped += skipped

        if not rows:
            report_lines.append(f"- {folder_name}: EMPTY (headers only)")
        else:
            note = f"- {folder_name}: {len(rows)} email(s)"
            if skipped:
                note += f", {skipped} skipped"
            report_lines.append(note)

    # Save workbook
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    out_dir   = os.path.join(downloads, "BE")
    os.makedirs(out_dir, exist_ok=True)
    out_path  = resolve_output_path(out_dir, "be_emails")
    wb.save(out_path)

    summary  = "BE export complete.\n\n"
    summary += "\n".join(report_lines)
    summary += f"\n\nTotal: {total_rows} email(s) exported"
    if total_skipped:
        summary += f", {total_skipped} skipped"
    summary += f"\nSaved to:\n{out_path}"
    messagebox.showinfo("Done", summary)
    reset_buttons()

# ── GUI ───────────────────────────────────────────────────────────────────────

def reset_buttons():
    pgts_btn.config(state=tk.NORMAL, text="PGTS")
    be_btn.config(state=tk.NORMAL, text="BE")

root = tk.Tk()
root.title("Outlook -> Excel")
root.resizable(False, False)
root.geometry("320x180")

lbl = tk.Label(root, text="Export Outlook emails to Excel", pady=10,
               font=("Arial", 10))
lbl.pack()

pgts_btn = tk.Button(root, text="PGTS", command=run_pgts_task,
                     width=20, height=2, bg="#0078D4", fg="white",
                     font=("Arial", 10, "bold"), relief=tk.FLAT, cursor="hand2")
pgts_btn.pack(pady=4)

be_btn = tk.Button(root, text="BE", command=run_be_task,
                   width=20, height=2, bg="#107C10", fg="white",
                   font=("Arial", 10, "bold"), relief=tk.FLAT, cursor="hand2")
be_btn.pack(pady=4)

root.mainloop()
