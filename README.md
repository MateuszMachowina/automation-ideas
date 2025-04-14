# Office Automation

This repository contains VBA macros and Python scripts designed for task automation. The goal is to speed up processes, save time, and improve work efficiency.

## 📂 Contents

- **VBA Macros**  
  Automate tasks in Microsoft Office programs like Excel, Access, and Word.
  
- **Python Scripts**  
  Automate various tasks such as file manipulation, data analysis, and web scraping.

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
   *(If your script is on another drive, switch first by typing the drive letter, e.g.:)*

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
