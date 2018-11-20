' Open Workbook and write entry in lastrow via offset
Sub test()

'WB öffnen, letzte Zeile beschreiben, nochmal neue letzte Zeile beschreiben, speichern
Application.DisplayAlerts = False

Dim lastWrittenRow As Long
Dim newWkb As Workbook
Set newWkb = Workbooks.Open("C:\Users\Apr17\Documents\VBA\SavedWB.xlsm")

Dim newWksh As Worksheet
Set newWksh = newWkb.Worksheets("Tabelle1")

Debug.Print newWkb.Worksheets("Tabelle1").Name
'Debug.Print newWkb.newWksh.Name

Debug.Print "Value row before: " & newWkb.Worksheets("Tabelle1").Range("A" & lastWrittenRow).Rows.Value
Debug.Print "Value Offset before: " & newWkb.Worksheets("Tabelle1").Range("A" & lastWrittenRow).Offset(1, 0).Rows.Value

newWkb.Worksheets("Tabelle1").Range("A" & lastWrittenRow).Offset(1, 0).Rows.Value = "neuer Wert"

Debug.Print "Value row after: " & newWkb.Worksheets("Tabelle1").Range("A" & lastWrittenRow).Rows.Value
Debug.Print "Value Offset after: " & newWkb.Worksheets("Tabelle1").Range("A" & lastWrittenRow).Offset(1, 0).Rows.Value


newWkb.Worksheets("Tabelle1").Range("A" & lastWrittenRow).Offset(1, 0).Rows.Value = "neuer Wert2"

Debug.Print "Value row after: " & newWkb.Worksheets("Tabelle1").Range("A" & lastWrittenRow).Rows.Value
Debug.Print "Value Offset after: " & newWkb.Worksheets("Tabelle1").Range("A" & lastWrittenRow).Offset(1, 0).Rows.Value

newWkb.Worksheets("Tabelle1").Range("A" & lastWrittenRow).Offset(1, 0).Rows.Select
    ActiveWorkbook.SaveAs "C:\Users\Apr17\Documents\VBA\SavedWB.xlsm"

End Sub

Function lastWrittenRow() As Integer
lastWrittenRow = ActiveWorkbook.Worksheets("Tabelle1").Cells(ActiveWorkbook.Worksheets("Tabelle1").Rows.Count, 1).End(xlUp).Row
End Function
