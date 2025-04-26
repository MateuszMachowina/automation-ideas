# ⚙️ Excel Automation Toolkit for Sales Department

Collection of Excel VBA macros for automating sales workflows, email generation, data imports, invoice generation and more.

## 🔧 Features

### 📜 Automated Invoice PDF Generator
- Effortlessly generates professional invoice PDFs for this month's sales, including customer details, product information and amounts.
- Exports each invoice as a neatly formatted PDF into the dedicated `Invoices` folder, making them easy to access and manage.
- Each invoice features the company's logo (stored as `logo.png` in the same folder as the Excel file), customer information and a thank-you note for a personal touch.

### 📧 Invoice Email Draft Generator
- Automatically generates Outlook draft emails using data from the `sales-april-2025` sheet.
- Populates each email with customer details, product info, invoice reference, and amounts.
- Attaches a corresponding PDF invoice (saved in the `Invoices` subfolder, based on the invoice reference).
- Emails are saved as drafts by default (optional configuration to display or send directly).
- Includes a professional signature, "Mateusz | Matt Games," with a friendly thank-you message for a polished communication.

### 📥 Archive Importer

- Prompts the user to select another Excel file.
- Lists sheets in that file and lets the user choose many via a numbered menu.
- It is advised to choose `2, 1` to first import the `archive` sheet and then the `sales` sheet for a full overview of past periods, including latest sales data.
- Copies selected columns A–J from row 2 onward into the `archive` sheet.
- Records each import in the `logs` sheet with:
  - Action name
  - Date and time (EU format)
  - Source file name
  - Target file name
  - Status (success/failed)

## 🧬 Related Projects Across Different Technologies

Explore connected projects developed for the Sales and Controlling departments, all working with the same invoice data ecosystem:

- **Sales Department**  
  ➔ [**Excel Exchange Rates Converter (Python)**](https://github.com/MateuszMachowina/python-apps/tree/main/Excel%20Exchange%20Rates%20Converter)  
  A Python-based tool that directly processes the `main-table-sales-april-2025.xlsm` file, fetching the latest currency exchange rates via an API and updating amounts accordingly. This ensures accurate international invoicing without manual calculations.

- **Sales Department**  
  ➔ [**Customer Overdue Payment Notifier (Power Automate)**](https://github.com/MateuszMachowina/power-automate/tree/main/Customer-Overdue-Payment-Notifier)  
  A Power Automate flow that monitors overdue invoices based on exported `.xlsx` versions of the sales data (macros and Power Automate do not cooperate directly). It automatically sends personalized notification emails to customers, helping improve payment collection processes.

- **Controlling Department (Customers' Side)**  
  ➔ [**Invoice OCR to Excel (UiPath)**](https://github.com/MateuszMachowina/ui-path/tree/main/Invoice_OCR_to_Excel)  
  A UiPath automation built for the company's customers. It processes received invoice PDFs—originally generated using the Excel macro toolkit from this project—by extracting data through OCR and transferring it into structured Excel spreadsheets. This streamlines financial reporting and simplifies auditing tasks for finance controllers.

## 🧩 Project Structure

### 🧠 Modules
- `Module1`: Email generation and automation
- `Module2`: Data import and logging
- `Module3`: Invoice PDF generation (creates invoices based on the current month's sales, stores them in the `Invoices` subfolder)
  
### 🗃️ Files
- **`main-table-sales-april-2025.xlsm`**  
  Primary workbook that contains all macros and serves as the control center.
  
- **`secondary-table-sales-march-2025.xlsx`**  
  Example data source file used for importing past sales into the archive.

### 📑 Sheets
- **`sales-*month*-*year*`** – Monthly sales data (e.g., `sales-april-2025`)
- **`archive`** – Central repository where past sales are appended
- **`logs`** – Action log tracking each macro execution (with timestamp and status)

### 📂 Subfolders
- **`Invoices`** – Subfolder where generated invoice PDFs are stored. The invoices are saved with names corresponding to their invoice reference numbers.
  
### 📄 Assets
- **`logo.png`** – The company logo image, placed in the same folder as the workbook, used for invoice branding.

## 💡 Requirements

- Microsoft Excel with macros enabled
- Microsoft Outlook (for email-related features)
- Consistent sheet and column formatting for reliable macro behavior
