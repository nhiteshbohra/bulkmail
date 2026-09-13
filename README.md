# 📧 AutoMail PRO - Enterprise Bulk Email Dispatcher
> **Made with ❤️ by Hitesh Bohra**

[![Electron](https://img.shields.io/badge/Electron-44.3.0-47848F?style=for-the-badge&logo=electron&logoColor=white)](https://electronjs.org/)
[![Node.js](https://img.shields.io/badge/Node.js-v18%2B-339933?style=for-the-badge&logo=node.js&logoColor=white)](https://nodejs.org/)
[![Nodemailer](https://img.shields.io/badge/Nodemailer-10.0.9-007ACC?style=for-the-badge&logo=nodemailer&logoColor=white)](https://nodemailer.com/)
[![SheetJS](https://img.shields.io/badge/SheetJS-0.20.3-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://sheetjs.com/)
[![License: ISC](https://img.shields.io/badge/License-ISC-blue.svg?style=for-the-badge)](https://opensource.org/licenses/ISC)

**AutoMail PRO** is a feature-rich, high-performance desktop application built with **Electron**, **Nodemailer**, and **SheetJS (XLSX)** that automates sending personalized bulk emails directly from Excel (`.xlsx`, `.xls`) or CSV spreadsheets. 

Equipped with **Windows Task Scheduler integration**, **Shield Quota MX domain verification**, **Randomized Jitter Delays**, **Live Email Preview**, and **Campaign History Tracking**, AutoMail PRO ensures maximum deliverability and workflow efficiency.

---

## 🖼️ Application Preview

![AutoMail Excel Desktop Interface](./image.png)

---

## ✨ Features & Highlights

### 🚀 Automation & Scheduling
- ⏰ **Fire-and-Forget Multi-Day Batch Scheduler**: Automatically split large campaigns into safe daily batches (e.g. max 500 emails/day on weekdays) integrated with **Windows Task Scheduler (`schtasks`)**—emails send on schedule even if the desktop app is closed!
- 🎲 **Randomized Jitter Delay (Anti-Spam Safeguard)**: Configure randomized delay ranges (e.g. 2s to 6s randomly per email) to simulate human sending patterns and bypass spam filters.
- 📩 **Send Test Mail to Myself**: Dispatch a quick test email to your inbox to verify formatting, attachments, and SMTP deliverability before launching full campaigns.

### 🛡️ Deliverability & Validation
- 🛡️ **Shield Quota Pre-Flight Email Verification**: Automatically scan recipient email addresses for invalid syntax and inactive/dead domain MX records before dispatching, saving your daily Gmail/SMTP sending quota.
- 👥 **Multi-Recipient Cell Splitting**: Automatically split cells containing multiple comma- or semicolon-separated email addresses into individual recipients.
- 🔗 **Clickable HTML Signature Links**: Automatic markdown link parsing (e.g., `[LinkedIn](url)`), email addresses, and phone numbers in signatures into clickable HTML hyperlinks.

### 📄 Composition, Templates & Preview
- 🔍 **Live Rendered Email Preview Modal**: Inspect exactly how emails will look for any row in your spreadsheet with `{VARIABLE}` interpolation, signature rendering, and attachment verification.
- 💾 **Reusable Email Templates System**: Save, load, and delete email mapping presets and signature templates stored locally (`templates.json`).
- 📎 **Global Master Attachments**: Attach master files/images to all outgoing emails alongside row-specific attachment file paths.

### 📊 Monitoring, Logs & Analytics
- 📜 **Past Campaign History Slide-Over Drawer**: Slide-over panel to view, search, inspect recipient tables, and re-export reports for all historical campaign logs (`campaign_log_*.json`).
- 📊 **Real-Time Execution Dashboard**: Live metrics showing Total, Sent, Failed, Quota Saved, and Pending counters with search/filtering (`All`, `Sent`, `Failed`, `Invalid`, `Pending`).
- 📁 **Campaign Execution Export**: Export detailed execution reports (status, response, timestamp, error logs) into Excel (`.xlsx`) or CSV files.

---

## 🛠️ Prerequisites

Before running the application, make sure you have the following installed:

- **Node.js**: `v18.0.0` or higher ([Download Node.js](https://nodejs.org/))
- **npm**: `v9.0.0` or higher (comes bundled with Node.js)
- **Windows OS**: Recommended for native `schtasks` background scheduling

---

## 🚀 Quick Start Guide

### 1. Clone the Repository

```bash
git clone https://github.com/nhiteshbohra/bulkmail.git
cd bulkmail
```

### 2. Install Dependencies

```bash
npm install
```

### 3. Generate Sample Excel File (Optional)

To test the application immediately with sample data, run:

```bash
node create_sample_excel.js
```

This creates `sample_contacts.xlsx` in the project folder with pre-filled test columns.

### 4. Launch the Application

```bash
npm start
```

---

## 📋 Excel / CSV Format Guide

AutoMail PRO automatically detects sheet headers or lets you map column names manually in the UI.

### Recommended Column Structure

| EMAIL ADDRESS | TOPIC | BODY | CC | ATTACHMENT PATH |
| :--- | :--- | :--- | :--- | :--- |
| `john.doe@example.com` | `Monthly Report - {FIRST_NAME}` | `Dear {FIRST_NAME},\n\nPlease see attached report for {COMPANY}.` | `manager@example.com` | `C:\Documents\Report.pdf` |
| `alice.smith@example.com` | `Welcome to the Team!` | `Hello Alice,\n\nWelcome aboard!` | | `C:\Documents\Onboarding.pdf` |

> 💡 **Pro Tips:**
> - Use placeholders like `{FIRST_NAME}`, `{COMPANY}`, etc., in your subject or body—they will be interpolated per recipient!
> - Multiple email addresses in one cell can be separated with commas (`,`) or semicolons (`;`).
> - The `ATTACHMENT PATH` field is optional. Leave empty if no individual attachment is required.

---

## ⚙️ SMTP Setup Instructions

### 1. Gmail / Google Workspace Setup
- **Host**: `smtp.gmail.com`
- **Port**: `587` (TLS) or `465` (SSL)
- **User**: Your Gmail address (`your.email@gmail.com`)
- **Password**: **Google App Password** *(Required if 2-Factor Authentication is enabled)*.
  > 🔑 **How to get a Gmail App Password:**
  > 1. Go to your [Google Account Security Settings](https://myaccount.google.com/security).
  > 2. Enable **2-Step Verification**.
  > 3. Search for **App Passwords** in the search bar.
  > 4. Create an App Password for **Mail** and paste the 16-character password into AutoMail PRO.

### 2. Outlook / Office 365 Setup
- **Host**: `smtp.office365.com`
- **Port**: `587`
- **User**: Your Outlook email address
- **Password**: Your account password or App Password

### 3. Custom SMTP
- Enter your SMTP server host, port, username, password, and SSL/TLS toggle as provided by your email host.

---

## 📁 Project Structure

```text
bulkmail/
├── main.js                  # Electron main process (IPC handlers, file parsing, SMTP & schedule tasks)
├── sender.js                # Standalone background mail dispatcher (Windows Scheduled Task entrypoint)
├── preload.js               # Context bridge bindings for secure IPC communication
├── emailValidator.js        # Email syntax and MX domain validation module
├── create_sample_excel.js   # Helper script to generate sample Excel contacts
├── package.json             # App metadata, dependencies, and npm scripts
├── run_schedule.bat         # Launcher script generated for Windows Task Scheduler
└── src/
    ├── index.html           # UI structure, dynamic forms, modals & slide-over drawer
    ├── styles.css           # Glassmorphism dark mode UI styling & animations
    └── renderer.js          # Application logic, preview renderer, template manager & log drawer
```

---

## 🛠️ Built With

- **[Electron](https://electronjs.org/)** - Desktop GUI Framework
- **[Node.js](https://nodejs.org/)** - JavaScript Runtime Environment
- **[Nodemailer](https://nodemailer.com/)** - SMTP Mail Dispatcher
- **[SheetJS / XLSX](https://sheetjs.com/)** - Excel & CSV Parser

---

## 📜 License

This project is licensed under the [ISC License](LICENSE).

---

## 👤 Author

Developed by **[Hitesh Bohra](https://github.com/nhiteshbohra)**.

⭐ **Star** this repository if you find it helpful!
