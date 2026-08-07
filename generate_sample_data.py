"""Generate a sample Excel file of contacts for the Email Automation app.

Run:  python generate_sample_data.py
Creates:  sample_contacts.xlsx  (requires pandas + openpyxl)
"""
import pandas as pd

COLUMNS = ["Name", "Email", "Company", "Code"]

ROWS = [
    ["Alice Johnson",    "alice@example.com",      "Acme Corp", "ALICE10"],
    ["Bob Smith",        "bob@example.com",        "Globex",    "BOB20"],
    ["Carol White",      "carol@example.com",      "Initech",   "CAROL30"],
    ["David Brown",      "david@example.com",      "Umbrella",  "DAVID40"],
    ["Eve Davis",        "eve@example.com",        "Stark",     "EVE50"],
    ["Frank Miller",     "frank@example.com",      "Wayne",     "FRANK60"],
    ["Grace Lee",        "grace@example.com",      "Oscorp",    "GRACE70"],
    ["Henry Wilson",     "henry@example.com",      "Hooli",     "HENRY80"],
    ["Ivy Garcia",       "ivy@example.com",        "Aperture",  "IVY90"],
    ["Jack Taylor",      "jack@example.com",       "Cyberdyne", "JACK100"],
]

df = pd.DataFrame(ROWS, columns=COLUMNS)
df.to_excel("sample_contacts.xlsx", index=False)
print(f"✅ Created sample_contacts.xlsx with {len(df)} contacts.")
