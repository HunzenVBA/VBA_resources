Option Explicit

'************** Subs in this Module ****************

'************** /Subs in this Module ****************

Sub RodeoAddQueryTotal()
Application.ScreenUpdating = False
    Dim importWS As Worksheet
    Dim lastrow As Long
    Dim lastcol As Long
    Dim i As Long
    Dim DelRange As Range
    Dim URL As String
    Dim qtRodeoTotal As QueryTable
    Dim ImportWbk As Workbook
    Dim counter As String

    'assign importWS = Importsheet
'    Set ImportWbk = Workbooks.Open(FileName:=strRodeoHistoryFile, UpdateLinks:=False)
    Set ImportWbk = Workbooks(strRodeoHistoryFileName)
    ImportWbk.Worksheets.Add ImportWbk.Worksheets(1)
    Set importWS = ActiveSheet
'    Set importWS = ThisWorkbook.Worksheets("RodeoTotal")
    'Initialize error handling
        On Error GoTo Whoa
    'Löschen alter Daten auf dem Rodeo Tabellenblatt
'        importWS.UsedRange.Delete xlUp
'        importWS.Activate

    'Web Query
'**********************************************************************************************************************************
'   Rodeo link with parameters as shown
'   URL = "https://tiny.amazon.com/17n4oyxa8/rodeamazDTM2Item"  'dwelling time < 1h
    URL = "https://tiny.amazon.com/ejyw3yjp/rodeamazDTM2Item" 'dwell time <30min

    Debug.Print "URL = " & URL
    Set qtRodeoTotal = importWS.QueryTables.Add(Connection:="URL;" & URL, Destination:=Range("A1"))
        With qtRodeoTotal  'Datei und Zielort auswählen
'        Spaltenbreiten bleiben aktuell erhalten
            .Name = "qTRodeoTotal"
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = True
            .PreserveFormatting = False
            .RefreshOnFileOpen = False
            .BackgroundQuery = False 'geändert als Query lange dauerte und delay hatte
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = False
            .RefreshPeriod = 0
            .WebSelectionType = xlEntirePage
            .WebFormatting = xlWebFormattingNone
            .WebPreFormattedTextToColumns = True
            .WebConsecutiveDelimitersAsOne = True
            .WebSingleBlockTextImport = False
            .WebDisableDateRecognition = True
            .WebDisableRedirections = False
            .Refresh BackgroundQuery:=False
'           .PreserveColumnInfo = True '::: liefert im Moment noch Fehler
        End With
        'Sortieren nach Spalte dwelling time
        'importWS.Range("DwellTime").Sort Key1:=Range("P1"), Order1:=xlDescending, Header:=xlYes

'    Delete unused rows
'**********************************************************************************************************************************

    lastrow = importWS.Range("A1").CurrentRegion.Rows.Count 'Anzahl der bis zur letzten beschriebenen Zeile
    lastcol = importWS.Range("A1").CurrentRegion.Columns.Count 'Anzahl der bis zur letzten beschriebenen Spalte
    importWS.Rows(lastrow & ":" & importWS.Rows.Count).Delete 'Zeilen ab der letzten geschriebenen Zeile löschen um Blattgröße zu minimieren
    importWS.Cells(lastrow + 2, 1).Value = URL
    importWS.Cells(lastrow + 3, 1).Value = Format(Now, "DD.MM.YYYY HH:MM") 'Zeitstempel Werte eintragen
    qtDeleteInAllWbks
    counter = Format(Now, "DD.MM_HH.mm.ss")
    importWS.Name = "RodeoTotal" & counter
    Call fDeleteColumns(importWS)
    importWS.Columns.AutoFit
    Call fRodeoColumnsWidth(importWS)
'    ImportWbk.Save

    Exit Sub

Whoa:
    MsgBox Err.Description

    Exit Sub
End Sub

Sub UpdateRodeoTotal()
'    Runtime Start
    StartTime = Timer
    currProcedureName = "UpdateRodeoTotal"
    Debug.Print "============================ Beginn Sub " & currProcedureName & " ============================"
    Debug.Print Now
    With qtRodeoTotal
        ThisWorkbook.Worksheets("4RodeoTotalImport").QueryTables(1).Refresh
    End With
'    Timer end and print
'************************************************************
    SecondsElapsed = Round(Timer - StartTime, 2)
'    Notify user in seconds
    Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds"
End Sub
