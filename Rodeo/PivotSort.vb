Option Explicit

Sub FilterAllPivot()

Dim pT As PivotTable
Dim ws As Worksheet
Dim pvtField As PivotField

For Each ws In ThisWorkbook.Worksheets
    For Each pT In ws.PivotTables
        Debug.Print vbNewLine & "********  " & pT.Name & vbTab & ws.Name & "  ********"

            pT.PivotFields("Outer Container Type").AutoSort xlDescending, "Anzahl von Scannable ID", pT.PivotColumnAxis.PivotLines(1), 1

    Next pT
Next ws

End Sub
