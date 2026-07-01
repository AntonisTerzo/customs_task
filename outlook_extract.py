import os
import ctypes
import tkinter as tk
from tkinter import scrolledtext, messagebox
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

def get_display_name():
    """Return the user's Windows display name, trimmed and safe."""
    try:
        GetUserNameEx = ctypes.windll.secur32.GetUserNameExW
        NameDisplay = 3
        size = ctypes.pointer(ctypes.c_ulong(0))
        GetUserNameEx(NameDisplay, None, size)
        name_buffer = ctypes.create_unicode_buffer(size.contents.value)
        GetUserNameEx(NameDisplay, name_buffer, size)
        username = name_buffer.value or ""
        # Trim company suffix, e.g. "Last, First / Company" -> "Last, First"
        if "/" in username:
            username = username.split("/")[0].strip()
        if not username:
            username = os.environ.get("USERNAME", "User")
        return username
    except Exception:
        return os.environ.get("USERNAME", "User")

def sanitize_subject(subject):
    """Replace a leading formula-injection character with underscore."""
    if subject and subject[0] in FORMULA_CHARS:
        subject = "_" + subject[1:]
    return subject

def resolve_output_path(out_dir, base_name):
    """
    Return a unique file path inside out_dir.
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
    Return (rows, skipped) for a given Outlook folder.
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
    """Return (mapi, error_message)."""
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
    Navigate to Archive/PGTS in the personal mailbox only.
    Return (folder, error_message). If found, error_message is None.
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

def run_pgts_task(log):
    """Return (success, out_folder_path)."""
    log("Starting PGTS task...")
    log("Connecting to Outlook...")

    mapi, err = connect_outlook()
    if err:
        log(f"ERROR: {err}")
        messagebox.showerror("Error", err)
        return False, None

    log("Locating Archive/PGTS...")
    pgts_folder, err = get_pgts_folder(mapi)
    if pgts_folder is None:
        log(f"ERROR: {err}")
        messagebox.showerror("Folder Not Found", err)
        return False, None

    log("Reading emails...")
    try:
        rows, skipped = read_folder_emails(pgts_folder)
    except Exception as exc:
        log(f"ERROR: Failed to read emails: {exc}")
        messagebox.showerror("Error", f"Failed to read emails:\n{exc}")
        return False, None

    if not rows:
        log("Folder is empty.")
        messagebox.showwarning(
            "Empty Folder",
            "The Archive/PGTS folder was found but contains no emails."
        )
        return False, None

    log(f"Found {len(rows)} email(s). Writing Excel file...")
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

    log(f"Saved to: {out_path}")
    if skipped:
        log(f"{skipped} item(s) skipped (non-email items).")
    log("PGTS task complete.")

    summary = f"{len(rows)} email(s) exported to:\n{out_path}"
    if skipped:
        summary += f"\n\n{skipped} item(s) were skipped (non-email items)."
    messagebox.showinfo("PGTS Complete", summary)
    return True, out_dir

# ── BE task ───────────────────────────────────────────────────────────────────

def find_inbox_subfolder(inbox, target_name):
    """Return the direct subfolder of Inbox matching name, or None."""
    target = target_name.strip().lower()
    for i in range(inbox.Folders.Count):
        f = inbox.Folders.Item(i + 1)
        if f.Name.strip().lower() == target:
            return f
    return None

def run_be_task(log):
    """Return (success, out_folder_path)."""
    log("Starting BE task...")
    log("Connecting to Outlook...")

    mapi, err = connect_outlook()
    if err:
        log(f"ERROR: {err}")
        messagebox.showerror("Error", err)
        return False, None

    try:
        inbox = mapi.GetDefaultFolder(6)
    except Exception as exc:
        log(f"ERROR: Could not access Inbox: {exc}")
        messagebox.showerror("Error", f"Could not access Inbox:\n{exc}")
        return False, None

    wb = Workbook()
    wb.remove(wb.active)          # remove default empty sheet

    report_lines = []
    total_rows = 0
    total_skipped = 0

    for folder_name, sheet_name in BE_FOLDERS:
        log(f"Processing '{folder_name}'...")
        ws = wb.create_sheet(title=sheet_name)
        ws.append(["Subject", "Received"])

        folder = find_inbox_subfolder(inbox, folder_name)
        if folder is None:
            log(f"  NOT FOUND - creating empty sheet")
            report_lines.append(f"- {folder_name}: NOT FOUND (empty sheet created)")
            continue

        try:
            rows, skipped = read_folder_emails(folder)
        except Exception as exc:
            log(f"  ERROR reading: {exc}")
            report_lines.append(f"- {folder_name}: ERROR reading ({exc})")
            continue

        for row in rows:
            ws.append(list(row))

        total_rows += len(rows)
        total_skipped += skipped

        if not rows:
            log("  Empty - only headers written")
            report_lines.append(f"- {folder_name}: EMPTY (headers only)")
        else:
            note = f"- {folder_name}: {len(rows)} email(s)"
            if skipped:
                note += f", {skipped} skipped"
            report_lines.append(note)
            log(f"  {len(rows)} email(s) written" + (f", {skipped} skipped" if skipped else ""))

    log("Writing Excel file...")
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    out_dir   = os.path.join(downloads, "BE")
    os.makedirs(out_dir, exist_ok=True)
    out_path  = resolve_output_path(out_dir, "be_emails")
    wb.save(out_path)

    log(f"Saved to: {out_path}")
    log("BE task complete.")

    summary  = "BE export complete.\n\n"
    summary += "\n".join(report_lines)
    summary += f"\n\nTotal: {total_rows} email(s) exported"
    if total_skipped:
        summary += f", {total_skipped} skipped"
    summary += f"\nSaved to:\n{out_path}"
    messagebox.showinfo("BE Complete", summary)
    return True, out_dir

# ── GUI ───────────────────────────────────────────────────────────────────────

class OutlookExporterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Outlook -> Excel")
        self.root.geometry("700x520")

        username = get_display_name()

        # Welcome message
        welcome_label = tk.Label(
            root,
            text=f"Welcome back {username}",
            font=("Arial", 18, "bold"),
            fg="#2C3E50"
        )
        welcome_label.pack(pady=20)

        # Buttons frame
        buttons_frame = tk.Frame(root)
        buttons_frame.pack(pady=10)

        self.pgts_button = tk.Button(
            buttons_frame,
            text="PGTS",
            command=self.start_pgts,
            font=("Arial", 12, "bold"),
            bg="#5DADE2",
            fg="white",
            width=15,
            height=1,
            relief="raised",
            borderwidth=3,
            cursor="hand2"
        )
        self.pgts_button.grid(row=0, column=0, padx=10)

        self.be_button = tk.Button(
            buttons_frame,
            text="BE",
            command=self.start_be,
            font=("Arial", 12, "bold"),
            bg="#5DADE2",
            fg="white",
            width=15,
            height=1,
            relief="raised",
            borderwidth=3,
            cursor="hand2"
        )
        self.be_button.grid(row=0, column=1, padx=10)

        # Log output area
        log_label = tk.Label(root, text="Log Output:", font=("Arial", 10, "bold"))
        log_label.pack(pady=(20, 5))

        self.log_text = scrolledtext.ScrolledText(
            root, width=75, height=18, state="disabled")
        self.log_text.pack(pady=10)

    def log(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.root.update()

    def clear_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state="disabled")

    def disable_buttons(self):
        self.pgts_button.config(state="disabled")
        self.be_button.config(state="disabled")

    def enable_buttons(self):
        self.pgts_button.config(state="normal", text="PGTS")
        self.be_button.config(state="normal", text="BE")

    def start_pgts(self):
        self.disable_buttons()
        self.pgts_button.config(text="Processing...")
        self.clear_log()
        try:
            success, folder_path = run_pgts_task(self.log)
            if success and folder_path:
                os.startfile(folder_path)
        except Exception as exc:
            self.log(f"ERROR: {exc}")
            messagebox.showerror("Error", f"An error occurred: {exc}")
        finally:
            self.enable_buttons()

    def start_be(self):
        self.disable_buttons()
        self.be_button.config(text="Processing...")
        self.clear_log()
        try:
            success, folder_path = run_be_task(self.log)
            if success and folder_path:
                os.startfile(folder_path)
        except Exception as exc:
            self.log(f"ERROR: {exc}")
            messagebox.showerror("Error", f"An error occurred: {exc}")
        finally:
            self.enable_buttons()

if __name__ == "__main__":
    root = tk.Tk()
    app = OutlookExporterGUI(root)
    root.mainloop()
    
