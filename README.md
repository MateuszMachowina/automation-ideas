# Office Automation

This repository contains VBA macros and Python scripts designed for task automation. The goal is to speed up processes, save time and improve work efficiency.

## 📂 Contents

- **VBA Macros**  
  - `export-sql-queries-from-ms-access` - Exports all SQL queries from an Access database to a text file on your desktop.
  - `sales-department-project` - These scripts automate tasks including generating professional invoice PDFs with customer details, creating Outlook email drafts with attached invoices and importing data from external Excel files into an archive, all while logging each action for traceability.
  
- **Python Scripts**  
  - `excel-exchange-rates-converter` - Converts EUR invoice amounts in Excel to another currency using live rates and saves the result as a new file.
  - `python-markets` - Checks current exchange rates in real time using an API.

## 📝 How to Use

### 1. **VBA Macros**

- Open the relevant Excel/Access/Word file.
- Press `ALT + F11` to open the VBA editor.
- Paste or load the macro code into the appropriate module.
- Run the macro.

### 2. **Python Scripts**

#### Before Running a Script:

- Ensure **Python** is installed on your machine.
- Install any required libraries (refer to `requirements.txt` if available).
- Ensure a **`.env`** file is present in the same directory as the script (if environment variables are used).

#### **Running the Script:**

1. **Open Command Prompt**

2. **Navigate to the script folder**  
   *(If your script is on another drive, switch first by typing the drive letter)*

   ```bash
   E:
   cd "Path\To\Your\Script"
   ```

3. **Install required libraries**  
   *(if `requirements.txt` is available)*:

   ```bash
   pip install -r requirements.txt
   ```

4. **Run the script**:

   ```bash
   python script_name.py
   ```
