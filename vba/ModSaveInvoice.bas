Attribute VB_Name = "ModSaveInvoice"
Option Explicit

Sub SaveInvoice()

    Dim SheetName As Worksheet
    Dim TargetRange As Range
    Dim NewRow As Long
    Dim InvoiceRow As Integer
    Dim InvoiceNumber As Integer
    Dim Path As String
    Dim Resp As Integer
    Dim i As Long
    Dim j As Long
    
    InvoiceRow = Application.WorksheetFunction.CountA(Range("tblInvoice[CODE]"))
    InvoiceNumber = ThisWorkbook.Sheets("Invoice").Range("E4").Value
    
    If InvoiceRow = 0 Or Range("btCustomer").Value = "" Then
            MsgBox "Please enter a customer and product code", vbExclamation, "TechNova Solutions"
    Exit Sub
    End If
    
    ThisWorkbook.ActiveSheet.PrintOut Copies:=1
    
         
        With Application.FileDialog(msoFileDialogFolderPicker)
            .InitialFileName = Application.DefaultFilePath & "  "
            .Title = "TechNova Solutions - Select Folder"
            .Show
            If .SelectedItems.Count = 0 Then
            Else
                Path = .SelectedItems(1)
                  
                 MsgBox "Generating PDF for Invoice #" & InvoiceNumber & ". Click OK to continue...", _
                 vbInformation, "TechNova Solutions"
                    
                    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
                    Path & "\" & "Invoice -" & InvoiceNumber & ".pdf", Quality:=xlQualityStandard, _
                    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        
            End If
        End With
    
    
    
    With ThisWorkbook.Sheets("Invoice Detail")
                
               For i = 1 To InvoiceRow

                    Set TargetRange = .Range("A1").CurrentRegion
                     NewRow = TargetRange.Rows.Count + 1
                    .Cells(NewRow, 1).Value = Date
                    .Cells(NewRow, 2).Value = InvoiceNumber
                    .Cells(NewRow, 3).Value = Range("btCustomer").Value

                    For j = 1 To 4
                            .Cells(NewRow, j + 3) = ThisWorkbook.Sheets("Invoice").Cells(12 + i, j + 1).Value
                    Next j
                    .Cells(NewRow, 8).Value = Path & "\" & "Invoice -" & InvoiceNumber & ".pdf"
               Next i
                
    End With
    
    MsgBox "Invoice saved successfully", vbInformation, "TechNova Solutions"

 ThisWorkbook.Worksheets("Pt-Invoice Analysis").PivotTables("ptDetail").PivotCache.Refresh

Resp = MsgBox("Do you want to delete the data?", vbYesNo + vbQuestion, "TechNova Solutions")

If Resp = vbYes Then
    
    With ThisWorkbook.Sheets("Invoice")
           .Range("btCustomer").ClearContents
           .Range("B13:B20").ClearContents
           .Range("D13:D20").ClearContents
    End With
Else
End If

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
