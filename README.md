# 📧 HR Outlook Email Automation (Python)

This project automates sending personalized emails through **Microsoft Outlook** using **Python**, based on employee data stored in an **Excel file**.  
It is especially useful for **HR teams** to send bulk, customized emails with optional attachments.

---

## 🚀 Features

- Reads employee data from Excel sheets
- Merges employee details using **Employee ID**
- Sends **personalized HTML emails**
- Supports **file attachments**
- Automatically skips missing email addresses
- Secure handling of sensitive data using `.env`
- Works directly with **Microsoft Outlook**

---

## 🛠️ Technologies Used

- Python
- Pandas
- Microsoft Outlook (via `win32com`)
- Excel (.xlsx)
- dotenv (`.env` file)

---

## 📁 Project Structure

HR Project/
│
├── hr-outlookAttatch.py # Main Python script
├── README.md # Project documentation
├── .gitignore # Ignored files (env, excel, etc.)
└── .env # Environment variables (NOT uploaded)


---

## 📊 Excel File Requirements

The Excel file (e.g. `HR-email.xlsx`) must contain the following sheets:

### 1️⃣ Sheet: `email id`
Required columns:
- `Empl ID`
- `Email`

### 2️⃣ Sheet: `emp-info`
Required columns:
- `Empl ID`
- `Name`
- `Title`
- `Subject`
- `Mail-body`
- `Information /related to`
- `Attachment` (optional)

> **Note:** Excel file is ignored in GitHub for data privacy.

---

## ⚙️ Setup Instructions

### 1️⃣ Clone the repository
```bash
git clone https://github.com/your-username/your-repo-name.git
cd your-repo-name
