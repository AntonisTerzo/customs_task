"""
Reads emails from the 'PGTS' subfolder inside Archive (personal mailbox only)
and writes Subject, Date, and Time to an Excel file saved at:
    %USERPROFILE%\Downloads\PGTS\outlook_emails.xlsx

Requirements:
    pip install pywin32 openpyxl
"""

import os
import tkinter as tk
from tkinter import messagebox
import win32com.client
from openpyxl import Workbook


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

    # Find Archive folder at mailbox root level
    archive_folder = None
    for i in range(store_root.Folders.Count):
        f = store_root.Folders.Item(i + 1)
        if f.Name.strip().lower() == "archive":
            archive_folder = f
            break

    if archive_folder is None:
        return None, "The 'Archive' folder was not found in your personal mailbox."

    # Find PGTS inside Archive
    for i in range(archive_folder.Folders.Count):
        f = archive_folder.Folders.Item(i + 1)
        if f.Name.strip().upper() == "PGTS":
            return f, None

    return None, "The 'PGTS' subfolder was not found inside Archive."


def run_task():
    btn.config(state=tk.DISABLED, text="Running...")
    root.update()

    # Connect to Outlook
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mapi = outlook.GetNamespace("MAPI")
        mapi.Logon()
    except Exception as exc:
        messagebox.showerror("Error", f"Could not connect to Outlook:\n{exc}")
        btn.config(state=tk.NORMAL, text="Run the task")
        return

    # Locate Archive/PGTS
    pgts_folder, err = get_pgts_folder(mapi)
    if pgts_folder is None:
        messagebox.showerror("Folder Not Found", err)
        btn.config(state=tk.NORMAL, text="Run the task")
        return

    # Read emails
    rows = []
    try:
        items = pgts_folder.Items
        items.Sort("[ReceivedTime]", True)
        for i in range(items.Count):
            try:
                item = items.Item(i + 1)
                if item.Class == 43:          # olMailItem = 43
                    subject  = item.Subject or "(no subject)"
                    received = item.ReceivedTime
                    date_str = received.strftime("%Y-%m-%d")
                    time_str = received.strftime("%H:%M:%S")
                    rows.append((subject, date_str, time_str))
            except Exception:
                continue
    except Exception as exc:
        messagebox.showerror("Error", f"Failed to read emails:\n{exc}")
        btn.config(state=tk.NORMAL, text="Run the task")
        return

    # Empty folder check
    if not rows:
        messagebox.showwarning(
            "Empty Folder",
            "The Archive/PGTS folder was found but contains no emails."
        )
        btn.config(state=tk.NORMAL, text="Run the task")
        return

    # Save to Excel
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    out_dir   = os.path.join(downloads, "PGTS")
    os.makedirs(out_dir, exist_ok=True)
    out_path  = os.path.join(out_dir, "outlook_emails.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "Emails"
    ws.append(["Subject", "Date", "Time"])
    for row in rows:
        ws.append(list(row))
    wb.save(out_path)

    messagebox.showinfo("Done", f"{len(rows)} email(s) exported to:\n{out_path}")
    btn.config(state=tk.NORMAL, text="Run the task")


# ── GUI ───────────────────────────────────────────────────────────────────────
root = tk.Tk()
root.title("Outlook → Excel")
root.resizable(False, False)
root.geometry("300x120")

lbl = tk.Label(root, text="Export Archive/PGTS emails to Excel", pady=10)
lbl.pack()

btn = tk.Button(root, text="Run the task", command=run_task,
                width=20, height=2, bg="#0078D4", fg="white",
                font=("Arial", 10, "bold"), relief=tk.FLAT, cursor="hand2")
btn.pack(pady=5)

root.mainloop()