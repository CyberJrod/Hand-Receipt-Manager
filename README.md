# ⚠️ Disclaimer (Not Government Endorsed)

This software is an independent tool and is **not affiliated with, endorsed by, or sponsored by** the U.S. Department of Defense (DoD), the U.S. Army, or any other government agency. Use of this software does **not** imply compliance with or substitution for official DoD/Army policies, forms, or procedures. **You are responsible** for verifying accuracy and complying with all applicable regulations and for using official systems where required.

---

# Hand Receipt Manager (DA Form 2062)

A desktop Python application to track inventory (e.g., IP phones, laptops), issue/return equipment, and generate **DA Form 2062** hand receipts using a **flattened** PDF template. Supports barcode scanning (keyboard wedge), bulk serial entry, managed dropdown lists, and a soft-delete **Recycle Bin**.

---

## 📌 Features

- **Inventory Management**
  - Track **Model, Category, Box #, Serial** (Asset Tag stored in DB only).
  - Bulk add serials (comma-separated **or** one-per-line).
  - Managed dropdowns for **Model / Category / Box** stored in `inventory_lists.json`.
  - Import/Export CSV.

- **Issuing & Returning**
  - Separate **Issue** and **Return** tabs with validation.
  - Custodian metadata: **Issued by (From)** and **Contact** stored per custodian.

- **Issued Items Overview**
  - See all items currently issued by custodian.
  - Edit custodian’s **Issued by** and **Contact**.
  - Generate DA 2062 **on demand** for any custodian.

- **DA 2062 Generation**
  - Uses `DA2062_flat.pdf` as a template.
  - Auto-groups by model.
  - **10 serials per row** (first line: Model + 4 S/N, wrapped line: 6 S/N).
  - Automatic pagination with page indicator `1/N`, `2/N`, etc.
  - Calibration tab for fine-tuning overlay positions.

- **Recycle Bin**
  - Delete is **allowed only for On Hand** items.
  - Issued items cannot be deleted.
  - Restore or permanently purge deleted items.

---

## ⚙️ Requirements

- **Python**: 3.10+
- **Dependencies**:
  ```bash
  python -m pip install --upgrade pypdf reportlab pandas openpyxl cryptography
  ```
- **Template**: `DA2062_flat.pdf` (flattened version of DA Form 2062) placed in the same folder as the script.

---

## 🚀 Quick Start

1. Clone or download the repository.
2. Place `DA2062_flat.pdf` next to `hand_receipt_manager.py`.
3. Install dependencies:
   ```bash
   python -m pip install --upgrade pypdf reportlab pandas openpyxl cryptography
   ```
4. Run the app:
   ```bash
   python hand_receipt_manager.py
   ```

On first run, the following files are created:
- `inventory.db` — SQLite database
- `inventory_lists.json` — Model/Category/Box dropdown values
- `da2062_layout.json` — Calibration settings

---

## 📖 Usage Guide

### Inventory Tab
- Add items by selecting **Model**, **Category**, and optional **Box #**.
- Paste/scan multiple serial numbers.
- Use **Add Model / Add Category / Add Box** to update dropdown lists.
- Delete moves On Hand items to the **Recycle Bin**.

### Issue Tab
- Fill out **From**, **To (Person/Unit)**, and **Contact**.
- Scan/enter serials, validate, then issue.
- Issued items are tracked under the custodian.

### Return Tab
- Scan/enter serials to validate and mark them returned.

### Issued Items Tab
- Shows custodians with their issued counts, **Issued by**, and **Contact**.
- Generate DA 2062 for the selected custodian.
- File name format:  
  ```
  DA2062_{Custodian}_{YYYYMMDD}.pdf
  ```

### Recycle Bin Tab
- Soft-deleted items are listed.
- Restore or permanently delete.
- Only **On Hand** items can be deleted.

### Calibration Tab
- Fine-tune X/Y positions and font sizes for overlay fields.
- Reset or save calibration settings.

---

## 📦 Import/Export

- **Export CSV**: Save current inventory (excludes deleted).
- **Import CSV**: Add/update items. Revives soft-deleted serials if matched.

CSV required columns:
- `Model`, `Category`, `Serial Number`  
Optional: `Box #`, `Asset Tag #`

---

## 🖨️ DA 2062 Generation

- Grouped by model.
- Each row can contain up to 10 serials:
  - Line 1: Model + 4 serials
  - Line 2: 6 more serials (wrapped)
- Page 2+ does not repeat To/From headers.
- Contact information is displayed below the To field.

---

## 🔒 Security & Privacy

- All data stays local in `inventory.db`.
- Exported CSVs and PDFs may contain sensitive equipment data—handle accordingly.

---

## 🛠️ Building an EXE (Optional)

You can bundle the app with **PyInstaller**:

```bash
python -m pip install pyinstaller
pyinstaller ^
  --onefile ^
  --name HandReceiptManager ^
  --add-data "DA2062_flat.pdf;." ^
  hand_receipt_manager.py
```

This creates `dist/HandReceiptManager.exe`.

---

## 📂 Files Created

- `inventory.db` — SQLite database
- `inventory_lists.json` — Dropdown values
- `da2062_layout.json` — Calibration settings
- `DA2062_flat.pdf` — Flattened DA Form 2062 template (required)

---

## 🧰 Troubleshooting

- **Missing modules** → Reinstall requirements:
  ```bash
  python -m pip install --upgrade pypdf reportlab pandas openpyxl cryptography
  ```
- **Template not found** → Ensure `DA2062_flat.pdf` is next to the script.
- **Encrypted template** → You’ll be prompted for a password.
- **Cannot delete issued items** → Only On Hand items can be deleted.

---

## 📜 License

Choose a license for your repo (e.g., MIT, Apache 2.0, Proprietary). Until then, assume internal use only.

---
=======
Application Guide
1) Inventory Tab

Add Item(s)

Choose Model, Category, optional Box (from dropdowns or type new).

Add new list values via Add Model / Add Category / Add Box. These persist in inventory_lists.json.

Paste/scan Serial Numbers (comma-separated or one-per-line).

Click Add to Inventory to create multiple entries at once.

Table

Shows current (non-deleted) items.

“Status/Custodian” shows On Hand or Issued to <name>.

Bottom Bar

Delete Selected (to Recycle Bin):

Works only for On Hand items.

If any selected items are Issued, they’ll be listed and skipped.

Confirm and (optionally) provide a reason. The item is soft-deleted (recoverable).

Export CSV: Writes a clean snapshot (excludes deleted).

Import CSV: Adds items; also revives matching soft-deleted serials.

Refresh: Reloads tables.

Right side shows counts: On Hand and Currently Issued.

CSV Format (Import/Export)

Columns (case-sensitive expected on import):

Model, Category, Box #, Serial Number, Asset Tag #, Status, Custodian, Updated At

Minimum required for import: Model, Category, Serial Number.

2) Issue Tab

Fields:

From (issuing unit/person)

To (Person/Unit)

To (Contact Information) (e.g., phone/email)

Scan/Enter Serials: comma-separated or one-per-line.

Validate On-Hand checks availability.

Issue updates items to Issued and records custodian metadata:

Saves/updates custodian’s Issued by (From) and Contact.

The Issue action does not generate a 2062; you can do that later from Issued Items.

3) Return Tab

Paste/scan serials to return.

Validate Issued shows what’s currently out and to whom.

Mark Returned flips status back to On Hand.

4) Issued Items Tab

Left: Custodians (To:) table shows custodians and counts, plus Issued by and Contact columns.

Select a custodian to see their items on the right.

Edit Selected Metadata lets you update the Issued by (From) and Contact for that custodian.

Generate 2062 (Selected Custodian):

Prompts for save location and generates a PDF using the flat template.

File name pattern: DA2062_{Custodian}_{YYYYMMDD}.pdf (custodian sanitized to filename-safe).

Header on page 1:

From (Issued by) and To (custodian)

Contact (beneath To)

Items grouped by model with 10 serials per row:

Line 1: Model - S/N: <4 serials>

(Wrapped) Line 2: S/N: <6 serials>

Page 2+ does not repeat To/From.

5) Recycle Bin Tab (Soft Delete)

Shows soft-deleted items with reason and deletion time.

Restore Selected: returns items to Inventory.

Permanently Delete Selected: irreversibly removes from DB.

Deletion rule: Only “On Hand” items can be deleted. Issued items are blocked.

6) Calibration Tab

Fine-tune where text lands on the PDF.

General Rules

Lower Y values → move down

Higher Y values → move up

X values move left/right

Key Settings

item_desc_x: X position of Item Description (column C).

qty_auth_x: X position of Quantity Authorized (column G).

item_start_y_first: First row baseline Y on page 1.

item_start_y_next: First row baseline Y on pages 2+.

line_spacing: Vertical distance between rows.

second_line_offset: Drop from the first line to the wrapped (second) line within a row.

Header controls for From/To and page fraction (1/N).

Save writes to da2062_layout.json. Reset Defaults restores sane defaults.

Barcode Scanners

Most barcode scanners operate as keyboard wedge devices:

Ensure focus is in the serials input box.

Each scan types the serial and often presses Enter. That’s fine—use one per line or commas.

Files & Persistence

inventory.db — SQLite DB for everything.

inventory_lists.json — values for Model/Category/Box dropdowns.

da2062_layout.json — calibration settings.

DA2062_flat.pdf — your flattened 2062 template (required).

Back these up if you migrate machines.

Troubleshooting

Missing modules
Install dependencies:

python -m pip install --upgrade pypdf reportlab pandas openpyxl cryptography


Template not found
Place DA2062_flat.pdf next to the script (or EXE). Filename must match exactly.

Encrypted template
You’ll be prompted for a password once; if decryption fails, verify the correct file/password.

“table has no column named …”
The app migrates the DB on startup. If you copied an old DB, re-run the app and let it migrate, or remove inventory.db to start fresh (this deletes data).

Cannot delete Issued items
By design. Only On Hand can be soft-deleted. Use Return first if necessary.

Security & Privacy

Inventory data stays local in inventory.db unless you export/share it.

If your environment has sensitive device data, treat the DB and exports as controlled information.
