Attribute VB_Name = "ModSendEmail"
Option Explicit

Sub SendEmail()


Dim OutlookApp As Outlook.Application
Dim MailItem As Outlook.MailItem
Dim Subjects As String
Dim RecipientEmail As String
Dim EmailBody As String
Dim InvoiceNumber As String


InvoiceNumber = Sheets("FindInvoice").Range("N1").Value
Subjects = "Invoice #" & InvoiceNumber
RecipientEmail = "Recipient Email"
EmailBody = "Please find attached your invoice " & InvoiceNumber

Set OutlookApp = New Outlook.Application
Set MailItem = OutlookApp.CreateItem(olMailItem)

With MailItem
    .To = RecipientEmail
    .Subject = Subjects
    .Body = EmailBody
    .Attachments.Add InvoiceNumber
    .SentOnBehalfOfName = "Email_other_account"
'    .SendUsingAccount = OutlookApp.Session.Accounts.Item(1)
'    .Display
    .Send
End With

Set MailItem = Nothing
Set OutlookApp = Nothing
End Sub
'========================================================
' Project: Invoice Automation System
' Author: Mariano Ferrer
' Role: Excel VBA Developer
' Date: 2026
' Description:
' Excel VBA system that automates invoice generation,
' PDF export, printing and email sending.
'
' GitHub: https://github.com/cracka07
'========================================================
