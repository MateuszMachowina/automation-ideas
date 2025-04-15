Attribute VB_Name = "Module2"
Sub Import_Last_Month_To_Archive()

    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet
    Dim logWs As Worksheet
    Dim sourceFile As Variant
    Dim lastRowSource As Long
    Dim lastColSource As Long
    Dim lastRowTarget As Long
    Dim lastRowLog As Long
    Dim i As Long
    Dim fDialog As FileDialog
    Dim sheetList As String
    Dim sheetIndexes As Variant
    Dim inputStr As String
    Dim index As Variant
    Dim sheetName As String
    Dim sourceFileName As String
    Dim successFlag As Boolean

    successFlag = False
    On Error GoTo cleanup

    ' Prompt the user to select a source Excel file using a file picker dialog
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .Title = "Select source Excel file"
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm", 1
        If .Show <> -1 Then GoTo cleanup
        sourceFile = .SelectedItems(1)
    End With

    ' Extract just the filename from the full path
    sourceFileName = Mid(sourceFile, InStrRev(sourceFile, "\") + 1)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set sourceWb = Workbooks.Open(sourceFile, ReadOnly:=True)

    ' Build a list of available sheets in the source file and ask the user to select the sheets to import
    ' It is advised to select 2, 1 in order to import old archive and last months data
    sheetList = "Available sheets:" & vbCrLf
    For i = 1 To sourceWb.Sheets.Count
        sheetList = sheetList & i & ". " & sourceWb.Sheets(i).Name & vbCrLf
    Next i

    ' Prompt the user to input sheet numbers they wish to import (comma-separated)
    inputStr = InputBox(sheetList & vbCrLf & "Enter sheet numbers to import, separated by commas (e.g., 2, 1):", "Select Sheets")
    If inputStr = "" Then GoTo cleanup

    ' Split the input into individual sheet indexes
    sheetIndexes = Split(inputStr, ",")

    ' Set the target sheet where data will be imported
    Set targetWs = ThisWorkbook.Sheets("archive")

    ' Loop through each selected sheet and copy data to the target sheet
    For Each index In sheetIndexes
        index = Trim(index)
        If IsNumeric(index) Then
            i = CLng(index)
            If i >= 1 And i <= sourceWb.Sheets.Count Then
                Set sourceWs = sourceWb.Sheets(i)
                sheetName = sourceWs.Name

                ' Get the last row and column in the source sheet
                lastRowSource = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).Row
                lastColSource = sourceWs.Cells(1, sourceWs.Columns.Count).End(xlToLeft).Column
                lastRowTarget = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row + 1

                ' Copy all data from source sheet (starting from row 2) to the target sheet
                sourceWs.Range(sourceWs.Cells(2, 1), sourceWs.Cells(lastRowSource, lastColSource)).Copy _
                    Destination:=targetWs.Cells(lastRowTarget, 1)

                successFlag = True
            End If
        End If
    Next index

    ' Display success or failure message based on whether sheets were successfully imported
    If successFlag Then
        MsgBox "Data imported successfully!", vbInformation
    Else
        MsgBox "No valid sheets were imported.", vbExclamation
    End If

cleanup:
    On Error Resume Next
    If Not sourceWb Is Nothing Then sourceWb.Close False
    Call RowsHeight
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Log the operation with details about the source file, target file, and success/failure status
    Set logWs = ThisWorkbook.Sheets("logs")
    lastRowLog = logWs.Cells(logWs.Rows.Count, "A").End(xlUp).Row + 1
    logWs.Cells(lastRowLog, 1).Value = "macro archived" ' Operation type
    logWs.Cells(lastRowLog, 2).Value = Format(Now, "dd.mm.yyyy HH:MM") ' Timestamp of operation
    logWs.Cells(lastRowLog, 3).Value = sourceFileName   ' Source file name
    logWs.Cells(lastRowLog, 4).Value = ThisWorkbook.Name ' Target file (main file) name
    logWs.Cells(lastRowLog, 5).Value = IIf(successFlag, "success", "failed") ' Status of the operation

End Sub

' Subroutine to adjust row height and alignment in the target sheet
Sub RowsHeight()
    With ThisWorkbook.Sheets("archive")
        .UsedRange.Rows.RowHeight = 15
        .UsedRange.HorizontalAlignment = xlLeft
        .Rows(1).HorizontalAlignment = xlCenter
    End With
End Sub


