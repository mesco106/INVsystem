INVenezuela is a Python-based automation tool that generates professional Excel invoices from a predefined template.
It preserves all formatting, images, logos, and formulas by using Excel COM automation (pywin32) and applies business rules commonly required for invoicing in Venezuela.

The system reads product data from a CSV file, fills a clean invoice template, and applies:

Line items across multiple pages

Sequential invoice numbering

Randomized control numbering

Realistic invoice & expiration dates (weekdays only)

All outputs are generated from a reusable template

✨ Key Features

📄 Template-driven invoices
Uses a clean Excel template (template.xlsx) that remains untouched.

🧾 Multi-page invoices
Automatically distributes products across pages (9 items per page).

🔢 Invoice numbering

Sequential per page

User-defined starting number

Safe limit (max 50 pages)

🔴 Control numbering

Always 6 digits (leading zeros preserved)

Random increments (1–11)

Never repeated more than 3 times per date

Always formatted in red

📅 Invoice & expiration dates

User-defined date range

Weekdays only (Mon–Fri)

Random but realistic distribution

Dates never go backwards

Expiration = invoice date + 30 days (adjusted to weekday)

🖼 Preserves Excel images and formatting

Logos, shapes, and layout remain intact

No formula loss or drawing corruption

🧩 Clean modular architecture

Easy to extend and maintain

Each concern handled in its own module

INVENEZUELA-MAIN/
├─ data/
│  ├─ template/
│  │  └─ template.xlsx          # Blank invoice template (tracked)
│  ├─ processed/
│  │  └─ inputTable.csv         # Input product data (tracked)
│  └─ invoice/
│     ├─ .gitkeep               # Keeps folder in repo
│     └─ invoice_*.xlsx         # Generated invoices (ignored)
│
├─ src/
│  ├─ main.py                   # Program entry point
│  ├─ products.py               # CSV → product objects
│  ├─ invoicing.py              # Line-item placement
│  ├─ invoiceNumbering.py       # Invoice number logic
│  ├─ controlNumbering.py       # Control number logic
│  ├─ invoiceDating.py          # Invoice & expiration dates
│  ├─ prompts.py                # User prompts (CLI)
│
├─ requirements.txt
├─ README.md
└─ .gitignore

⚙️ Requirements

Windows

Microsoft Excel installed

Python 3.10+

Python dependency
pywin32>=306


Install with:

pip install -r requirements.txt

