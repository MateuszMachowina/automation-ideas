Attribute VB_Name = "Module1"

Sub Generate_Invoice_Emails()
    Dim ws As Worksheet
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim lastRow As Long
    Dim i As Long
    Dim emailBody As String
    Dim subjectLine As String
    Dim filePath As String
    
    Set ws = ThisWorkbook.Sheets("sales-april-2025")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    
    For i = 2 To lastRow
        ' Build subject and body
        subjectLine = "Invoice " & ws.Cells(i, 6).Value & " - " & ws.Cells(i, 7).Value
        
        emailBody = "Dear " & ws.Cells(i, 3).Value & "," & vbCrLf & vbCrLf & _
            "Please find your invoice details below:" & vbCrLf & vbCrLf & _
            "Invoice Reference: " & ws.Cells(i, 6).Value & vbCrLf & _
            "Product: " & ws.Cells(i, 7).Value & vbCrLf & _
            "Invoice Date: " & ws.Cells(i, 1).Value & vbCrLf & _
            "Due Date: " & ws.Cells(i, 2).Value & vbCrLf & _
            "Net Amount: $" & Format(ws.Cells(i, 8).Value, "#,##0.00") & vbCrLf & _
            "Gross Amount: $" & Format(ws.Cells(i, 9).Value, "#,##0.00") & vbCrLf & vbCrLf & _
            "Thank you for your business!" & vbCrLf & _
            "Best regards," & vbCrLf & _
            "Your Company Name"
        
        ' Create the draft email
        Set outlookMail = outlookApp.CreateItem(0)
        With outlookMail
            .To = ws.Cells(i, 5).Value
            .Subject = subjectLine
            .Body = emailBody
            
            ' Try to attach PDF named like the invoice reference
            ' Attachment has to be in the same folder as this excel file
            filePath = ThisWorkbook.Path & "\" & ws.Cells(i, 6).Value & ".pdf"
            If Dir(filePath) <> "" Then
                .Attachments.Add filePath
            End If
            
            .Save   ' Saves the email as a draft in Outlook
            '.Display ' Shows the email in a new window for review (but doesn't send)
            '.Send    ' Sends the email immediately
        End With
    Next i
    
    MsgBox "Draft emails generated successfully.", vbInformation
End Sub


