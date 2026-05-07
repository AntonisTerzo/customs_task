import os
import tkinter as tk
from tkinter import messagebox
import win32com.client
from openpyxl import Workbook

def get_test_folder(inbox):
    for i in range(inbox.Folders.Count):
        folder = inbox.Folders.Item(i + 1)
        if folder.Name.strip().lower() == "test":
            return folder
    return None

def run_task():
    btn.config(state=tk.DISABLED, text="Running...")
    root.update()

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mapi = outlook.GetNamespace("MAPI")
        mapi.Logon()
    except Exception as exc:
        messagebox.showerror("Error", f"Could not connect to Outlook:\n{exc}")
        btn.config(state=tk.NORMAL, text="Run the task")
        return

    try:
        inbox = mapi.GetDefaultFolder(6)
    except Exception as exc:
        messagebox.showerror("Error", f"Could not open Inbox:\n{exc}")
        btn.config(state=tk.NORMAL, text="Run the task")
        return

    test_folder = get_test_folder(inbox)
    if test_folder is None:
        messagebox.showerror("Error", "No subfolder named 'test' found inside your Inbox.")
        btn.config(state=tk.NORMAL, text="Run the task")
        return

    rows = []
    try:
        items = test_folder.Items
        items.Sort("[ReceivedTime]", True)
        count = items.Count
        for i in range(count):
            try:
                item = items.Item(i + 1)
                # olMailItem = 43
                if item.Class == 43:
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

    if not rows:
        messagebox.showwarning("No emails", "The 'test' folder exists but contains no emails.")
        btn.config(state=tk.NORMAL, text="Run the task")
        return

    # Save to Excel
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    out_dir   = os.path.join(downloads, "test folder")
    os.makedirs(out_dir, exist_ok=True)
    out_path  = os.path.join(out_dir, "outlook_emails.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "Emails"
    ws.append(["Subject", "Date", "Time"])
    for row in rows:
        ws.append(list(row))
    wb.save(out_path)

    messagebox.showinfo("Done", f"✅ {len(rows)} email(s) exported to:\n{out_path}")
    btn.config(state=tk.NORMAL, text="Run the task")

# ── GUI ───────────────────────────────────────────────────────────────────────
root = tk.Tk()
root.title("Outlook → Excel")
root.resizable(False, False)
root.geometry("300x120")

lbl = tk.Label(root, text="Export 'test' folder emails to Excel", pady=10)
lbl.pack()

btn = tk.Button(root, text="Run the task", command=run_task,
                width=20, height=2, bg="#0078D4", fg="white",
                font=("Arial", 10, "bold"), relief=tk.FLAT, cursor="hand2")
btn.pack(pady=5)

root.mainloop()