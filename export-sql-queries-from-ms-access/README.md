# 📤 Export Queries to Folder with File Name

This VBA script is designed to export all SQL queries from an Access database to a text file on your desktop. The script generates a file containing the names and SQL definitions of all queries (excluding system queries). 

## ✨ Features

- Exports queries from an Access database.
- Saves the queries in a `.txt` file on the user's Desktop.
- Organizes the file with a name based on the Access file's name.
- Creates a folder on the Desktop (`AccessExport`) if it does not exist.

## 🛠️ Requirements

- Microsoft Access (VBA support).

## ▶️ How to Use

1. Open your Access database.
2. Press `Alt + F11` to open the VBA editor.
3. Insert a new module and paste the code into the module.
4. Run the `ExportQueriesToFolderWithFileName` subroutine.

## 📝 Output

- The script creates a folder on your desktop called `AccessExport`.
- A `.txt` file is generated containing all the query names and their corresponding SQL statements.

Example file path: `C:\Users\<YourUsername>\Desktop\AccessExport\queries_<DatabaseName>.txt`

## 📄 Notes

- The script ignores system queries (those starting with a `~`).
- The file is named based on the Access database name (without extension).

