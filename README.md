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
