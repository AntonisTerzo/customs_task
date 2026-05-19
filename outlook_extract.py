import os
import re
import sys
from pathlib import Path

import win32com.client
from pypdf import PdfReader, PdfWriter
from openpyxl import Workbook

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
TARGET_FOLDER = "test"
DOWNLOADS_DIR = Path.home() / "Downloads"
OUTPUT_DIR    = DOWNLOADS_DIR / "invoices"   # All downloads + merged PDFs go here
EXCEL_OUTPUT  = DOWNLOADS_DIR / "invoice_summary.xlsx"

# Regex patterns: match label then capture the first number that follows
# Handles formats like:  "Gross Weight: 120.5 kg"  or  "Gross Weight 120.5"
PATTERNS = {
    "gross_weight": re.compile(
        r"gross\s*weight[\s:]*([0-9]+(?:[.,][0-9]+)?)", re.IGNORECASE
    ),
    "net_weight": re.compile(
        r"net\s*weight[\s:]*([0-9]+(?:[.,][0-9]+)?)", re.IGNORECASE
    ),
    "total_value": re.compile(
        r"total\s*value[\s:]*([0-9]+(?:[.,][0-9]+)?)", re.IGNORECASE
    ),
}

# ---------------------------------------------------------------------------
# Outlook helpers
# ---------------------------------------------------------------------------
def get_test_folder():
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns      = outlook.GetNamespace("MAPI")
    inbox   = ns.GetDefaultFolder(6)  # 6 = olFolderInbox

    for folder in inbox.Folders:
        if folder.Name.lower() == TARGET_FOLDER.lower():
            return folder

    print(f"[ERROR] Folder '{TARGET_FOLDER}' not found inside Inbox. Exiting.")
    sys.exit(1)

def download_pdfs(mail, dest_dir: Path) -> list[Path]:
    """Save all PDF attachments of a mail item to dest_dir. Returns saved paths."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    saved = []
    for attachment in mail.Attachments:
        name = attachment.FileName
        if name.lower().endswith(".pdf"):
            path = dest_dir / name
            attachment.SaveAsFile(str(path))
            saved.append(path)
    return saved

# ---------------------------------------------------------------------------
# PDF extraction
# ---------------------------------------------------------------------------
def extract_text_from_pdf(pdf_path: Path) -> str:
    reader = PdfReader(str(pdf_path))
    return "\n".join(page.extract_text() or "" for page in reader.pages)

def parse_value(text: str, field: str, pdf_path: Path):
    """
    Search for a labeled value in text.
    Returns float on success, None on failure (and prints a warning).
    """
    match = PATTERNS[field].search(text)
    if not match:
        print(f"  [WARNING] '{field}' not found in: {pdf_path.name}")
        return None
    raw = match.group(1).replace(",", ".")   # normalise European decimals
    try:
        return float(raw)
    except ValueError:
        print(f"  [WARNING] Could not parse '{field}' value '{raw}' in: {pdf_path.name}")
        return None

def extract_fields(pdf_path: Path) -> dict:
    text = extract_text_from_pdf(pdf_path)
    return {
        "gross_weight": parse_value(text, "gross_weight", pdf_path),
        "net_weight":   parse_value(text, "net_weight",   pdf_path),
        "total_value":  parse_value(text, "total_value",  pdf_path),
    }

# ---------------------------------------------------------------------------
# PDF merging
# ---------------------------------------------------------------------------
def merge_pdfs(pdf_paths: list[Path], output_path: Path):
    writer = PdfWriter()
    for p in pdf_paths:
        reader = PdfReader(str(p))
        for page in reader.pages:
            writer.add_page(page)
    with open(output_path, "wb") as f:
        writer.write(f)

# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------
def write_excel(rows: list[dict], output_path: Path):
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
            row.get("flags", ""),
        ])

    wb.save(str(output_path))
    print(f"\n[OK] Excel saved → {output_path.resolve()}")

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    folder = get_test_folder()
    messages = folder.Items
    print(f"[INFO] Found {messages.Count} email(s) in '{TARGET_FOLDER}'")

    rows = []

    for mail in messages:
        try:
            subject = mail.Subject or "no_subject"
            sender  = mail.SenderName or ""
            date    = str(mail.ReceivedTime)[:10]
        except Exception as e:
            print(f"[WARNING] Could not read mail properties: {e}")
            continue

        # Sanitise subject for use as folder name
        safe_subject = re.sub(r'[\\/*?:"<>|]', "_", subject)[:80]
        mail_dir     = OUTPUT_DIR / safe_subject

        print(f"\n--- Processing: {subject[:60]}")

        pdf_paths = download_pdfs(mail, mail_dir)
        if not pdf_paths:
            print("  [INFO] No PDF attachments found, skipping.")
            continue

        print(f"  [INFO] {len(pdf_paths)} PDF(s) downloaded → {mail_dir}")

        # Extract and accumulate values
        total_gw = total_nw = total_tv = 0.0
        missing_fields = []

        for pdf_path in pdf_paths:
            fields = extract_fields(pdf_path)

            if fields["gross_weight"] is None:
                missing_fields.append(f"{pdf_path.name}: gross_weight missing")
            else:
                total_gw += fields["gross_weight"]

            if fields["net_weight"] is None:
                missing_fields.append(f"{pdf_path.name}: net_weight missing")
            else:
                total_nw += fields["net_weight"]

            if fields["total_value"] is None:
                missing_fields.append(f"{pdf_path.name}: total_value missing")
            else:
                total_tv += fields["total_value"]

        # Merge PDFs (only if more than one)
        if len(pdf_paths) > 1:
            merged_path = mail_dir / f"merged_{safe_subject}.pdf"
            merge_pdfs(pdf_paths, merged_path)
            print(f"  [OK] Merged PDF → {merged_path.name}")

        rows.append({
            "subject":      subject,
            "sender":       sender,
            "date":         date,
            "gross_weight": round(total_gw, 3) if total_gw else (None if any("gross_weight" in f for f in missing_fields) else 0.0),
            "net_weight":   round(total_nw, 3) if total_nw else (None if any("net_weight"   in f for f in missing_fields) else 0.0),
            "total_value":  round(total_tv, 3) if total_tv else (None if any("total_value"  in f for f in missing_fields) else 0.0),
            "flags":        "\n".join(missing_fields),
        })

    if not rows:
        print("\n[INFO] No emails with PDF invoices found. Nothing to export.")
        return

    write_excel(rows, EXCEL_OUTPUT)
    print(f"[INFO] Done. Processed {len(rows)} email(s).")

if __name__ == "__main__":
    main()