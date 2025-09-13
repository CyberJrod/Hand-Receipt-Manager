\# Hand Receipt Manager (DA Form 2062 Automation)



A Python desktop application for managing Army equipment inventory, issuing and returning equipment, and generating official \*\*DA Form 2062 (DEC 2023)\*\* hand receipts in PDF format.



---



\## ✨ Features



\- \*\*Inventory Management\*\*

&nbsp; - Add/remove items with model, category, serial, asset tag, and box number

&nbsp; - Import/export inventory via CSV or Excel

&nbsp; - Prevents duplicate serial numbers



\- \*\*Issuing \& Returning\*\*

&nbsp; - Issue equipment by scanning or typing serials

&nbsp; - Return and validate items against current records

&nbsp; - Tracks custodian (who items are issued to) and issuing authority



\- \*\*Issued Items Dashboard\*\*

&nbsp; - View all equipment currently issued by custodian

&nbsp; - Store metadata per custodian (contact info, issued by)

&nbsp; - Generate official \*\*DA 2062 hand receipts\*\* directly from the issued list



\- \*\*DA 2062 PDF Generation\*\*

&nbsp; - Overlays data onto the official \*\*DA 2062 flat PDF template\*\*

&nbsp; - Supports multi-page receipts with correct pagination (1/3, 2/3, …)

&nbsp; - Groups equipment by model, with up to 10 serials per row (4 on line 1, 6 on line 2)

&nbsp; - Customizable layout calibration for precise alignment



\- \*\*Calibration\*\*

&nbsp; - Built-in calibration tab to adjust text positioning and spacing

&nbsp; - Saved in a JSON config file for consistency across sessions



---



\## 🛠 Tech Stack



\- Python 3.11+

\- Tkinter (GUI)

\- SQLite (local database)

\- ReportLab \& PyPDF (PDF generation/overlay)

\- Pandas \& OpenPyXL (CSV/Excel import/export)



---



\## 🚀 Getting Started



\### Prerequisites

\- \[Python 3.11+](https://www.python.org/downloads/)

\- Pip (comes with Python)



\### Installation

Clone this repository and set up a virtual environment:



```bash

git clone https://github.com/YOUR-USERNAME/hand-receipt-manager.git

cd hand-receipt-manager



python -m venv .venv

source .venv/bin/activate   # On Linux/Mac

.venv\\Scripts\\activate      # On Windows



pip install -r requirements.txt

Running the App

bash

Copy code

python src/hand\_receipt\_manager.py

📂 Repository Structure

bash

Copy code

hand-receipt-manager/

├─ src/

│  └─ hand\_receipt\_manager.py    # Main application

├─ templates/

│  └─ DA2062\_flat.pdf            # Flattened template (required)

├─ config/

│  └─ da2062\_layout.json         # Calibration settings

├─ requirements.txt

├─ README.md

├─ .gitignore

└─ LICENSE

📑 Notes

Template PDF: Only the flattened template (DA2062\_flat.pdf) is required. Place it in the templates/ directory.



Database: The app creates inventory.db automatically. This file should not be committed to Git.



Imports: CSV/Excel must include at least: Model, Category, Serial Number.





