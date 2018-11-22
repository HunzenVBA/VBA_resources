Option Explicit

Sub CopyAndFormatTable()

    Dim tbl As ListObject
    Dim rng As Range

    Tabelle2.UsedRange.ClearContents
    Tabelle3.UsedRange.ClearContents

    Debug.Print "Tabelle1 letzte Zeile: " & Tabelle1.Cells(Rows.Count, 1).End(xlUp).Rows.Row
    
    Tabelle1.UsedRange.Copy 'Destination:=Tabelle2.Range("A1")
    Tabelle2.Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Debug.Print "Tabelle2 letzte Zeile: " & Tabelle2.Cells(Rows.Count, 1).End(xlUp).Rows.Row

    Set rng = Tabelle2.Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.TableStyle = "TableStyleMedium15"

    Debug.Print "Tabelle3 letzte Zeile: " & Tabelle3.Cells(Rows.Count, 1).End(xlUp).Rows.Row
End Sub

Sub rangeproblem()
    Dim rng As Range
    Dim tbl As ListObject

    Set rng = Tabelle2.Range(Tabelle2.Range("A1"), Tabelle2.Range("A1").SpecialCells(xlLastCell))
    rng.Select
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.TableStyle = "TableStyleMedium15"
End Sub

Sub deleteqt()
    Dim qt As QueryTable
    For Each qt In Tabelle2.QueryTables
        Debug.Print "qt in Tabelle2: " & qt.Name
        qt.Delete
    Next qt
End Sub
