Attribute VB_Name = "ModFunction"
Option Explicit

Function NumberToText(num As Double, rate1 As String, rate2 As String, Optional Cents As Byte, Optional Denomination As String) As String
Dim IntegerPart As Long
Dim DecimalPart As Double
Dim Text As String

IntegerPart = Int(num)
DecimalPart = Int(Round((num - IntegerPart) * 100))

Text = cNumber(IntegerPart)


If IntegerPart = 1 Then
    Text = Text + " " + rate1
   
Else
    If (IntegerPart Mod 1000000) = 0 Then
        Text = Text + " De"
    End If
    Text = Text + " " + rate2
    
End If


If Cents = 1 Then
    
    If DecimalPart <> 0 Then
        Text = Text & " Con " & cNumber(DecimalPart)
        If DecimalPart = 1 Then
            Text = Text & " Centavo"
        Else
            Text = Text & " Cents"
        End If
    End If
    
ElseIf Cents = 0 Then
    
    If DecimalPart <> 0 Then
        Text = Text
        If DecimalPart = 1 Then
            Text = Text & DecimalPart & "/100"
        Else
            Text = Text & " " & DecimalPart & "/100"
        End If
    End If
End If
NumberToText = VBA.UCase(Text) & " " & Denomination


End Function

Function cNumber(ByVal num As Long) As String
Dim Text As String

Dim cUnits, cTens, cHundreds
Dim nUnits, nTens, nHundreds As Byte

Dim Thousands As Long
Dim Millions As Long

cUnits = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen", "Twenty", "Twenty One", "Twenty two", "Twenty three", "Twenty four", "Twenty five", "Twenty six", "Twenty seven", "Twenty eight", "Twenty nine")
cTens = Array("", "Ten", "Twenty", "Thirty", "Fourty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety", "One Hundred")
cHundreds = Array("", "One Hundred", "Two Hundred", "Three Hundred", "Four Hundred", "Five Hundred", "Six Hundred", "Seven Hundred", "Eight Hundred", "Nine Hundred")

Millions = num \ 1000000
Thousands = (num \ 1000) Mod 1000
nHundreds = (num \ 100) Mod 10
nTens = (num \ 10) Mod 10
nUnits = num Mod 10



If Millions = 1 Then
    Text = "Un Millón" + IIf(num Mod 1000000 <> 0, " " + cNumber(num Mod 1000000), "")
    cNumber = Text
    Exit Function
ElseIf Millions >= 2 And Millions <= 999 Then
    Text = cNumber(num \ 1000000) + " Millones" + IIf(num Mod 1000000 <> 0, " " + cNumber(num Mod 1000000), "")
    cNumber = Text
    Exit Function
    
        

    
ElseIf Thousands = 1 Then
    Text = "Thousand" + IIf(num Mod 1000 <> 0, " " + cNumber(num Mod 1000), "")
    cNumber = Text
    Exit Function
ElseIf Thousands >= 2 And Thousands <= 999 Then
    Text = cNumber(num \ 1000) + " Thousand" + IIf(num Mod 1000 <> 0, " " + cNumber(num Mod 1000), "")
    cNumber = Text
    Exit Function
    
End If


If num = 100 Then
    Text = cTens(10)
    cNumber = Text
    Exit Function
ElseIf num = 0 Then
    Text = "Zero"
    cNumber = Text
    Exit Function
End If


If nHundreds <> 0 Then
    Text = cHundreds(nHundreds)
End If

If nTens <> 0 Then
    If nTens = 1 Or nTens = 2 Then
        If nHundreds <> 0 Then
            Text = Text + " "
        End If
        Text = Text + cUnits(num Mod 100)
        cNumber = Text
        Exit Function
    Else
        
        If nHundreds <> 0 Then
            Text = Text + " "
        End If
        
        Text = Text + cTens(nTens)
    End If
End If


If nUnits <> 0 Then
    If nTens <> 0 Then
        Text = Text + " y "
    ElseIf nHundreds <> 0 Then
        Text = Text + " "
    End If
    Text = Text + cUnits(nUnits)
End If

cNumber = Text

End Function
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
