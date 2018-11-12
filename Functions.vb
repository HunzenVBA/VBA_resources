Option Explicit

Public StartTimer As Double
Public EndTimer As Double

Function StartTime()
StartTimer = Timer
End Function

Function EndTime()
EndTimer = Timer
End Function

Function PrintSecondsElapsed()
Dim SecondsElapsed As Double
SecondsElapsed = Round(EndTimer - StartTimer, 2)
Debug.Print "=================  code finished in " & SecondsElapsed & " seconds ================="
End Function

Sub TimerTest()
StartTimer = Timer
Application.Wait (Now + 0.00001)
EndTimer = Timer
Debug.Print Round(EndTimer - StartTimer, 2)
End Sub

Sub Seitennamen()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print ws.Index & " " & ws.Name
    Next ws
End Sub

Sub time() 'Datumswerte konvertieren Format
    Dim intZeileWS1 As Integer
    Dim intletzteZeileWS1 As Integer
    intletzteZeileWS1 = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
    Worksheets(1).Cells(1, 21).Value = "Time and Date"
    For intZeileWS1 = 2 To intletzteZeileWS1
        Worksheets(1).Cells(intZeileWS1, 21).Value = Format(((Left(Worksheets(1).Cells(intZeileWS1, 11).Value, 10)) / 86400) + 25569, "dd.MM.yyyy hh:mm")
        Worksheets(1).Cells(intZeileWS1, 19).Value = Format(Worksheets(1).Cells(intZeileWS1, 21).Value, "hh:nn")
        Worksheets(1).Cells(intZeileWS1, 20).Value = Format(Worksheets(1).Cells(intZeileWS1, 21).Value, "dd.MM.yyyy")
    Next intZeileWS1
End Sub

Sub EnableEvents()
Application.EnableEvents = True
Application.DisplayAlerts = True
End Sub
