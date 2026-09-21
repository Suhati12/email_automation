
# 📧 Email Automation

A real-world bulk email sender with a polished Tkinter GUI. Send personalized
campaign emails to a list of contacts pulled from an Excel file, with
attachments and reusable email templates.

Built with **Python + Tkinter** using `smtplib` (built-in) and `pandas` (Excel/CSV reading).

## ✨ Features

- 🚀 **Bulk Email** — send a campaign to hundreds of contacts in one click
- 📎 **Attachments** — add multiple files to every email
- 👥 **Excel Contacts** — load `.xlsx`, `.xls`, or `.csv` contact lists
- 📝 **Email Templates** — reusable `.txt` templates with `{{placeholder}}` tokens
- 🧩 **Personalization** — tokens like `{{Name}}`, `{{Company}}`, `{{Code}}` are
  auto-filled per contact
- 🔌 **Test Connection** — verify SMTP settings before sending
- 📋 **Live Activity Log** — see per-recipient success/failure and a final summary

## 🛠 Technologies

| Purpose | Library |
|---|---|
| Sending email | `smtplib` (built-in) |
| Email message & attachments | `email` (built-in) |
| Reading Excel/CSV contacts | `pandas` + `openpyxl` |
| GUI | `tkinter` / `ttk` |

## 🚀 Setup & Run

### 1. Create a virtual environment

```powershell
cd DAA/email_automation
python -m venv venv
venv\Scripts\activate
```

### 2. Install dependencies

```powershell
pip install pandas openpyxl
```

### 3. (Optional) Generate sample contacts

```powershell
python generate_sample_data.py
```

> Creates `sample_contacts.xlsx` with 10 demo contacts.

### 4. Launch the app

```powershell
python email_automation.py
```

## 📖 How to Use

1. **⚙ SMTP Settings** — enter your SMTP server (e.g. `smtp.gmail.com`),
   port (`587` for STARTTLS, `465` for SSL), sender email, and app password.
   Click **Test SMTP Connection** to verify.
2. **👥 Contacts** — click **Load Excel/CSV** and pick your contacts file.
   The table previews the first rows. The file must contain an `Email` column.
3. **✉ Compose & Send**
   - Pick a **Template** from the dropdown (loads subject + body).
   - The body can use placeholders like `{{Name}}`, `{{Company}}`, `{{Code}}`.
   - Optional: add **Attachments**.
   - Click **🚀 Send Bulk Email**. The send button disables while sending and
     re-enables when the campaign finishes.

## 📁 Files

```
email_automation/
├── email_automation.py      # Main Tkinter app
├── generate_sample_data.py  # Creates sample_contacts.xlsx
├── sample_contacts.xlsx     # Generated sample contacts (10 rows)
├── email_templates/
│   ├── welcome.txt          # Welcome email template
│   └── promotion.txt        # Promo email template
└── README.md
```

## 💡 Gmail Tip

For Gmail, enable **2-Step Verification** and create an **App Password**
(Google Account → Security → App passwords). Use that 16-character password
here — not your normal account password.
</content>
