import re
import sys
import threading
import tkinter as tk
from pathlib import Path

import win32com.client
from openpyxl import Workbook
from pypdf import PdfReader, PdfWriter

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
TARGET_FOLDER = "test"
OUTPUT_DIR    = Path.home() / "Downloads" / "customs_task"
EXCEL_OUTPUT  = OUTPUT_DIR / "invoice_summary.xlsx"

# Gross/Net Weight: label then number on same line or anywhere after
PATTERN_GW = re.compile(r"gross\s*weight[\s:]*([0-9]+(?:[.,][0-9]+)?)", re.IGNORECASE)
PATTERN_NW = re.compile(r"net\s*weight[\s:]*([0-9]+(?:[.,][0-9]+)?)",   re.IGNORECASE)

# Total Value: same line first, then fall back to next line
PATTERN_TV_INLINE   = re.compile(r"total\s*value[^\n]*?([0-9]+(?:[.,][0-9]+)?)", re.IGNORECASE)
PATTERN_TV_NEXTLINE = re.compile(r"total\s*value[^\n]*\n\s*([0-9]+(?:[.,][0-9]+)?)", re.IGNORECASE)

# ---------------------------------------------------------------------------
# Outlook
# ---------------------------------------------------------------------------
def get_test_folder():
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns      = outlook.GetNamespace("MAPI")
    inbox   = ns.GetDefaultFolder(6)
    for folder in inbox.Folders:
        if folder.Name.lower() == TARGET_FOLDER.lower():
            return folder
    raise RuntimeError(f"Folder '{TARGET_FOLDER}' not found inside Inbox.")

def download_invoice_pdfs(mail, dest_dir: Path) -> list[Path]:
    """Save PDF attachments that contain the word 'Invoice' to dest_dir."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    saved = []
    for attachment in mail.Attachments:
        name = attachment.FileName
        if not name.lower().endswith(".pdf"):
            continue
        path = dest_dir / name
        attachment.SaveAsFile(str(path))
        # Check if it's actually an invoice
        text = extract_text(path)
        if "invoice" in text.lower():
            saved.append(path)
        else:
            path.unlink()  # Not an invoice — delete and skip
    return saved

# ---------------------------------------------------------------------------
# PDF helpers
# ---------------------------------------------------------------------------
def extract_text(pdf_path: Path) -> str:
    reader = PdfReader(str(pdf_path))
    return "\n".join(page.extract_text() or "" for page in reader.pages)

def parse_simple(pattern: re.Pattern, text: str, field: str, pdf_name: str):
    match = pattern.search(text)
    if not match:
        return None, f"{pdf_name}: '{field}' not found"
    try:
        return float(match.group(1).replace(",", ".")), None
    except ValueError:
        return None, f"{pdf_name}: could not parse '{field}' value '{match.group(1)}'"
    
def parse_total_value(text: str, pdf_name: str):
    # Try same line first
    match = PATTERN_TV_INLINE.search(text)
    if match:
        try:
            return float(match.group(1).replace(",", ".")), None
        except ValueError:
            pass
    # Fall back to next line
    match = PATTERN_TV_NEXTLINE.search(text)
    if match:
        try:
            return float(match.group(1).replace(",", ".")), None
        except ValueError:
            pass
    return None, f"{pdf_name}: 'total_value' not found"

def extract_fields(pdf_path: Path) -> tuple[dict, list[str]]:
    text  = extract_text(pdf_path)
    flags = []

    gw, err = parse_simple(PATTERN_GW, text, "gross_weight", pdf_path.name)
    if err: flags.append(err)

    nw, err = parse_simple(PATTERN_NW, text, "net_weight", pdf_path.name)
    if err: flags.append(err)

    tv, err = parse_total_value(text, pdf_path.name)
    if err: flags.append(err)

    return {"gross_weight": gw, "net_weight": nw, "total_value": tv}, flags

def merge_pdfs(pdf_paths: list[Path], output_path: Path):
    writer = PdfWriter()
    for p in pdf_paths:
        for page in PdfReader(str(p)).pages:
            writer.add_page(page)
    with open(output_path, "wb") as f:
        writer.write(f)

# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------
def write_excel(rows: list[dict]):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice Summary"
    ws.append(["Email Subject", "Sender", "Date",
               "Gross Weight (sum)", "Net Weight (sum)", "Total Value (sum)", "Flags"])
    for row in rows:
        ws.append([
            row["subject"],
            row["sender"],
            row["date"],
            row["gross_weight"] if row["gross_weight"] is not None else "N/A",
            row["net_weight"]   if row["net_weight"]   is not None else "N/A",
            row["total_value"]  if row["total_value"]  is not None else "N/A",
            "; ".join(row["flags"]),
        ])
    wb.save(str(EXCEL_OUTPUT))

# ---------------------------------------------------------------------------
# Core task
# ---------------------------------------------------------------------------
def run_task(log):
    log("Connecting to Outlook...")
    try:
        folder = get_test_folder()
    except RuntimeError as e:
        log(f"ERROR: {e}")
        return

    messages = folder.Items
    log(f"Found {messages.Count} email(s) in '{TARGET_FOLDER}'")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    rows = []

    for mail in messages:
        try:
            subject = mail.Subject or "no_subject"
            sender  = mail.SenderName or ""
            date    = str(mail.ReceivedTime)[:10]
        except Exception as e:
            log(f"WARNING: Could not read mail properties: {e}")
            continue

        safe_subject = re.sub(r'[\\/*?:"<>|]', "_", subject)[:80]
        mail_dir     = OUTPUT_DIR / safe_subject

        log(f"\nProcessing: {subject[:60]}")

        pdf_paths = download_invoice_pdfs(mail, mail_dir)
        if not pdf_paths:
            log("  No invoice PDFs found, skipping.")
            continue

        log(f"  {len(pdf_paths)} invoice(s) downloaded")

        total_gw = total_nw = total_tv = 0.0
        all_flags = []

        for pdf_path in pdf_paths:
            fields, flags = extract_fields(pdf_path)
            all_flags.extend(flags)
            for flag in flags:
                log(f"  WARNING: {flag}")

            if fields["gross_weight"] is not None: total_gw += fields["gross_weight"]
            if fields["net_weight"]   is not None: total_nw += fields["net_weight"]
            if fields["total_value"]  is not None: total_tv += fields["total_value"]

        if len(pdf_paths) > 1:
            merged_path = OUTPUT_DIR / f"merged_{safe_subject}.pdf"
            merge_pdfs(pdf_paths, merged_path)
            log(f"  Merged PDF → {merged_path.name}")

        rows.append({
            "subject":      subject,
            "sender":       sender,
            "date":         date,
            "gross_weight": round(total_gw, 3) if total_gw else None,
            "net_weight":   round(total_nw, 3) if total_nw else None,
            "total_value":  round(total_tv, 3) if total_tv else None,
            "flags":        all_flags,
        })

    if not rows:
        log("\nNo invoice emails processed. Nothing exported.")
        return

    write_excel(rows)
    log(f"\nDone. Processed {len(rows)} email(s).")
    log(f"Output folder: {OUTPUT_DIR}")
    log(f"Excel: {EXCEL_OUTPUT}")

# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
def build_ui():
    root = tk.Tk()
    root.title("Customs Task")
    root.geometry("560x400")
    root.resizable(False, False)
    root.configure(bg="#1e1e2e")

    tk.Label(
        root, text="Customs Invoice Extractor",
        bg="#1e1e2e", fg="#ffffff",
        font=("Segoe UI", 14, "bold")
    ).pack(pady=(24, 4))

    tk.Label(
        root, text="Downloads invoices from Outlook → extracts weights & values → exports to Excel",
        bg="#1e1e2e", fg="#9999bb",
        font=("Segoe UI", 9), wraplength=480
    ).pack(pady=(0, 16))

    log_box = tk.Text(
        root, height=12, width=64,
        bg="#13131f", fg="#ccccdd",
        font=("Consolas", 9),
        relief="flat", state="disabled",
        padx=8, pady=6
    )
    log_box.pack(padx=20)

    def log(msg):
        log_box.configure(state="normal")
        log_box.insert("end", msg + "\n")
        log_box.see("end")
        log_box.configure(state="disabled")

    def on_click():
        btn.configure(state="disabled", text="Running...")
        log("─" * 48)
        def worker():
            run_task(log)
            btn.configure(state="normal", text="Run Customs Task")
        threading.Thread(target=worker, daemon=True).start()

    btn = tk.Button(
        root, text="Run Customs Task",
        bg="#2563eb", fg="#ffffff",
        activebackground="#1d4ed8", activeforeground="#ffffff",
        font=("Segoe UI", 11, "bold"),
        relief="flat", cursor="hand2",
        padx=20, pady=10,
        command=on_click
    )
    btn.pack(pady=16)

    root.mainloop()

if __name__ == "__main__":
    build_ui()