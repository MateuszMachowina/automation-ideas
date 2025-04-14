Attribute VB_Name = "Module3"
Sub Generate_Invoices_PDF()

    ' Declare variables for worksheets, data fields, paths, and counters
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim invoiceRef As String
    Dim customerName As String
    Dim customerEmail As String
    Dim dueDate As String
    Dim productName As String
    Dim invoiceNet As Double
    Dim invoiceGross As Double
    Dim currencyType As String
    Dim pdfPath As String
    Dim invoiceSheet As Worksheet
    Dim invoiceFileName As String
    Dim pdfFolder As String
    Dim sheet As Worksheet
    Dim logoPath As String

    ' Define path to the logo image (must exist in workbook directory)
    logoPath = ThisWorkbook.Path & "\logo.png"

    ' Find the first worksheet that matches "sales-*" pattern
    For Each sheet In ThisWorkbook.Sheets
        If sheet.Name Like "sales-*" Then
            Set ws = sheet
            Exit For
        End If
    Next sheet

    ' Display error and exit if no sales sheet found
    If ws Is Nothing Then
        MsgBox "No sales sheet found!", vbCritical
        Exit Sub
    End If

    ' Find the last row of data in the sales sheet (based on column A)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Define folder path to store generated PDFs
    pdfFolder = ThisWorkbook.Path & "\Invoices\"
    
    ' Create the folder if it doesn't already exist
    If Dir(pdfFolder, vbDirectory) = "" Then MkDir pdfFolder

    ' Loop through each row of data starting from row 2 (assuming headers in row 1)
    For i = 2 To lastRow

        ' Extract invoice data from the sales sheet
        invoiceRef = ws.Cells(i, 6).Value
        customerName = ws.Cells(i, 3).Value
        customerEmail = ws.Cells(i, 5).Value
        dueDate = ws.Cells(i, 2).Value
        productName = ws.Cells(i, 7).Value
        invoiceNet = ws.Cells(i, 8).Value
        invoiceGross = ws.Cells(i, 9).Value
        currencyType = ws.Cells(i, 10).Value

        ' Create a new worksheet for the invoice
        Set invoiceSheet = ThisWorkbook.Sheets.Add
        invoiceSheet.Name = "Invoice_" & invoiceRef

        ' Insert logo image if the file exists
        If Dir(logoPath) <> "" Then
            With invoiceSheet.Pictures.Insert(logoPath)
                .ShapeRange.LockAspectRatio = msoFalse
                .Left = 160
                .Top = 0
                .Width = 100
                .Height = 100
            End With
        End If

        ' Populate the invoice sheet with relevant details
        With invoiceSheet
            .Cells(2, 1).Value = "Invoice Reference: " & invoiceRef
            .Cells(4, 1).Value = "Invoice Date: " & ws.Cells(i, 1).Value
            .Cells(5, 1).Value = "Due Date: " & dueDate
            .Cells(7, 1).Value = "Customer Name: " & customerName
            .Cells(8, 1).Value = "Customer Email: " & customerEmail

            ' Set up product and pricing table
            .Cells(12, 1).Value = "Product Name"
            .Cells(12, 2).Value = "Invoice Net"
            .Cells(12, 3).Value = "Invoice Gross"
            .Cells(12, 4).Value = "Currency"

            .Cells(13, 1).Value = productName
            .Cells(13, 2).Value = Format(invoiceNet, "#,##0.00") & " " & currencyType
            .Cells(13, 3).Value = Format(invoiceGross, "#,##0.00") & " " & currencyType
            .Cells(13, 4).Value = currencyType

            ' Add thank-you message and signature
            .Cells(17, 4).Value = "Thank you for your order!"
            .Cells(18, 4).Value = "Mateusz"
            With Cells(18, 4).Font
                .Name = "Baguet Script"
                .Size = 12
                .ThemeColor = xlThemeColorLight1
            End With

            ' Format the invoice table with borders
            Call FormatTableBorders(.Range("A12:D13"))

            ' Auto-resize the columns for better appearance
            .Columns("A:D").AutoFit
        End With

        ' Export the invoice as a PDF to the designated folder
        invoiceFileName = pdfFolder & invoiceRef & ".pdf"
        invoiceSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=invoiceFileName

        ' Delete the temporary invoice sheet after export to keep workbook clean
        Application.DisplayAlerts = False
        invoiceSheet.Delete
        Application.DisplayAlerts = True
    Next i

    ' Notify the user once all invoices have been generated
    MsgBox "Invoices generated successfully!", vbInformation

End Sub

' Subroutine to apply consistent formatting and borders to invoice tables
Sub FormatTableBorders(tblRange As Range)
    With tblRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin

        ' Remove diagonal lines if any
        .Item(xlDiagonalDown).LineStyle = xlNone
        .Item(xlDiagonalUp).LineStyle = xlNone

        ' Add internal and edge borders
        .Item(xlInsideHorizontal).LineStyle = xlContinuous
        .Item(xlInsideVertical).LineStyle = xlContinuous
        .Item(xlEdgeLeft).LineStyle = xlContinuous
        .Item(xlEdgeTop).LineStyle = xlContinuous
        .Item(xlEdgeBottom).LineStyle = xlContinuous
        .Item(xlEdgeRight).LineStyle = xlContinuous
    End With

    ' Make the header row bold and center-align text
    tblRange.Rows(1).Font.Bold = True
    tblRange.HorizontalAlignment = xlCenter
End Sub

