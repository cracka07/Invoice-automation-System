Attribute VB_Name = "ModFindInvoice"
Option Explicit

Sub FindInvoice()

    
    Dim wsInvoice As Worksheet
    Dim wsPt As Worksheet
    Dim Invoice As Integer
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim pi As PivotItem
    
    Set wsInvoice = ThisWorkbook.Worksheets("FindInvoice")
    Set wsPt = ThisWorkbook.Worksheets("Pt-Invoice Analysis")
    
    Invoice = Application.InputBox("Please, enter the invoice number to search", "TechNova Solutions")
    wsInvoice.Range("E4").Value = Invoice
    
    On Error GoTo ErrorHandler
    
    Set pt = wsPt.PivotTables("ptDetail")
    Set pf = pt.PivotFields("Consecutive")
    
    pf.ClearAllFilters
    
    For Each pi In pf.PivotItems
                    If pi.Name = Invoice Then
                          pi.Visible = True
                    Else
                            pi.Visible = False
                    End If
    Next pi
    Exit Sub
ErrorHandler:
MsgBox "Invoice " & Invoice & " not found", vbExclamation, "TechNova Solutions"
Range("E4").ClearContents

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
