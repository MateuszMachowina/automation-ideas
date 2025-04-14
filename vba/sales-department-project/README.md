# Excel Automation Toolkit (VBA) for Sales Department

A growing collection of Excel VBA macros for automating sales workflows, email generation, data imports and more.

---

## 🔧 Features

### 📧 Invoice Email Draft Generator

- Generates Outlook draft emails based on rows in the `sales-april-2025` sheet.
- Populates each message with customer details, product info, and invoice amounts.
- Optionally attaches a PDF matching the invoice reference (must exist in the same folder).
- Emails are saved as drafts by default (can be modified to display or send directly).

### 📥 Archive Importer

- Prompts the user to select another Excel file.
- Lists sheets in that file and lets the user choose one via a numbered menu.
- Copies selected columns (A, D, F–J) from row 2 onward into the `archive` sheet.
- Records each import in the `logs` sheet with:
  - Action name
  - Date and time (EU format)
  - Source file name
  - Status (success/failed)

---

## 🧬 Project Structure

### 🧠 Modules
- `Module1`: Email generation and automation
- `Module2`: Data import and logging

### 📁 Files
- **`main-table-sales-april-2025.xlsm`**  
  Primary workbook that contains all macros and serves as the control center.
  
- **`secondary-table-sales-march-2025.xlsx`**  
  Example data source file used for importing past sales into the archive.

### 📑 Sheets
- **`sales-*month*-*year*`** – Monthly sales data (e.g., `sales-april-2025`)
- **`archive`** – Central repository where past sales are appended
- **`logs`** – Action log tracking each macro execution (with timestamp and status)

---

## 💡 Requirements

- Microsoft Excel with macros enabled
- Microsoft Outlook (for email-related features)
- Consistent sheet and column formatting for reliable macro behavior
