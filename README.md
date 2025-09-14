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

## 📖 Application Guide

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
