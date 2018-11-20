' Copy Range into Array •Copy according to criteria to other worksheet •fill up last row in specific col in other worksheet
' Worksheets 1=source 2,3,4=target
Option Explicit

Sub DinoDiet()
Dim r As Range
Dim counter As Long
Dim Dinos As Variant

Tabelle2.UsedRange.ClearContents
Tabelle3.UsedRange.ClearContents
Tabelle4.UsedRange.ClearContents

Tabelle1.Activate
Set r = ThisWorkbook.Worksheets(1).Range("E1:E" & Cells(Rows.Count, 1).End(xlUp).Row)

Dinos = r.Value
With Application
    .ScreenUpdating = False
End With

For counter = 1 To 29
    Debug.Print "Counter=" & counter & " " & Dinos(counter, 1) & " " & Tabelle1.Range("A" & counter).Value
    If Dinos(counter, 1) = "omnivore" Then
        Debug.Print "xlUp before " & Tabelle4.Cells(Tabelle4.Rows.Count, 1).End(xlUp).Row
        Tabelle1.Range("A" & counter).EntireRow.Copy
        Tabelle4.Range("A" & Tabelle4.Cells(Tabelle4.Rows.Count, 1).End(xlUp).Row).Offset(1, 0).Rows.PasteSpecial xlPasteAll
        Debug.Print "xlUp after " & Tabelle4.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        ElseIf Dinos(counter, 1) = "carnivore" Then
        Debug.Print "xlUp before " & Tabelle2.Cells(Rows.Count, 1).End(xlUp).Row
        Tabelle1.Range("A" & counter).EntireRow.Copy
        Tabelle2.Range("A" & Tabelle2.Cells(Tabelle2.Rows.Count, 1).End(xlUp).Row).Offset(1, 0).Rows.PasteSpecial xlPasteAll
        Debug.Print "xlUp after " & Tabelle2.Cells(Rows.Count, 1).End(xlUp).Row
        
        ElseIf Dinos(counter, 1) = "herbivore" Then
        Debug.Print "xlUp before " & Tabelle3.Cells(Rows.Count, 1).End(xlUp).Row
        Tabelle1.Range("A" & counter).EntireRow.Copy
        Tabelle3.Range("A" & Tabelle3.Cells(Tabelle3.Rows.Count, 1).End(xlUp).Row).Offset(1, 0).Rows.PasteSpecial xlPasteAll
        Debug.Print "xlUp after " & Tabelle3.Cells(Rows.Count, 1).End(xlUp).Row

    End If
Next counter
Erase Dinos

End Sub

Sub SelectLastRow()

Dim counter As Integer


'For counter = 1 To 10
'Tabelle1.Range("A" & counter).EntireRow.Copy
'Tabelle5.Range("A" & Tabelle4.Cells(Rows.Count, 1).End(xlUp).Row).Offset(counter, 0).Rows.PasteSpecial xlPasteAll
'Next counter
Tabelle5.Activate
Debug.Print Tabelle5.Range("M" & Cells(Tabelle5.Rows.Count, 13).End(xlUp).Row).Select
Debug.Print Tabelle5.Range("M" & Cells(Tabelle5.Rows.Count, 13)).Select
For counter = 1 To 3
Tabelle5.Range("M" & Cells(Rows.Count, 13).End(xlUp).Row).Offset(1, 0).Select
Tabelle5.Range("A" & Tabelle5.Cells(Tabelle3.Rows.Count, 1).End(xlUp).Row).Offset(1, 0).Select
Tabelle5.Range("M" & Cells(Rows.Count, 13).End(xlUp).Row).Offset(1, 0).Select
Tabelle5.Range("M" & Cells(Rows.Count, 13).End(xlUp).Row).Offset(1, 0).Value = counter
Next counter
Tabelle5.Calculate

End Sub
