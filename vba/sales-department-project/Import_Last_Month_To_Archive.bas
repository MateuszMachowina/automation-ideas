Attribute VB_Name = "Module2"
Sub Import_Last_Month_To_Archive()

    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet
    Dim logWs As Worksheet
    Dim sourceFile As Variant
    Dim lastRowSource As Long
    Dim lastRowTarget As Long
    Dim lastRowLog As Long
    Dim i As Long
    Dim fDialog As FileDialog
    Dim sheetList As String
    Dim sheetIndex As Variant
    Dim sheetName As String
    Dim sourceFileName As String
    Dim successFlag As Boolean

    successFlag = False ' Assume failure unless complete
    On Error GoTo cleanup

    ' Ask user to select file
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .Title = "Select source Excel file"
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm", 1
        If .Show <> -1 Then GoTo cleanup
        sourceFile = .SelectedItems(1)
    End With

    ' Get just the filename (no path)
    sourceFileName = Mid(sourceFile, InStrRev(sourceFile, "\") + 1)

    ' Open source workbook
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set sourceWb = Workbooks.Open(sourceFile, ReadOnly:=True)

    ' Build numbered sheet list
    sheetList = "Choose a sheet to import from:" & vbCrLf
    For i = 1 To sourceWb.Sheets.Count
        sheetList = sheetList & i & ". " & sourceWb.Sheets(i).Name & vbCrLf
    Next i

    ' Get user selection as number
    sheetIndex = Application.InputBox(sheetList, "Select Sheet by Number", Type:=1)

    ' Validate input
    If sheetIndex = False Or Not IsNumeric(sheetIndex) _
        Or sheetIndex < 1 Or sheetIndex > sourceWb.Sheets.Count Then
        MsgBox "Invalid selection.", vbExclamation
        GoTo cleanup
    End If

    ' Get the selected sheet
    sheetName = sourceWb.Sheets(sheetIndex).Name
    Set sourceWs = sourceWb.Sheets(sheetName)

    ' Get last row of data in source sheet
    lastRowSource = sourceWs.Cells(sourceWs.Rows.Count, "A").End(xlUp).Row

    ' Get target sheet
    Set targetWs = ThisWorkbook.Sheets("archive")
    lastRowTarget = targetWs.Cells(targetWs.Rows.Count, "A").End(xlUp).Row + 1

    ' Copy from A, D, F, G, H, I, J starting from row 2
    For i = 2 To lastRowSource
        targetWs.Cells(lastRowTarget, 1).Value = sourceWs.Cells(i, 1).Value ' A › A
        targetWs.Cells(lastRowTarget, 2).Value = sourceWs.Cells(i, 4).Value ' D › B
        targetWs.Cells(lastRowTarget, 3).Value = sourceWs.Cells(i, 6).Value ' F › C
        targetWs.Cells(lastRowTarget, 4).Value = sourceWs.Cells(i, 7).Value ' G › D
        targetWs.Cells(lastRowTarget, 5).Value = sourceWs.Cells(i, 8).Value ' H › E
        targetWs.Cells(lastRowTarget, 6).Value = sourceWs.Cells(i, 9).Value ' I › F
        targetWs.Cells(lastRowTarget, 7).Value = sourceWs.Cells(i, 10).Value ' J › G
        lastRowTarget = lastRowTarget + 1
    Next i

    successFlag = True
    MsgBox "Data imported successfully from '" & sheetName & "'!", vbInformation

cleanup:
    On Error Resume Next
    If Not sourceWb Is Nothing Then sourceWb.Close False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' === Logging ===
    Set logWs = ThisWorkbook.Sheets("logs")
    lastRowLog = logWs.Cells(logWs.Rows.Count, "A").End(xlUp).Row + 1
    logWs.Cells(lastRowLog, 1).Value = "macro archived"
    logWs.Cells(lastRowLog, 2).Value = Format(Now, "dd.mm.yyyy HH:MM")
    logWs.Cells(lastRowLog, 3).Value = sourceFileName
    logWs.Cells(lastRowLog, 4).Value = IIf(successFlag, "success", "failed")

End Sub

