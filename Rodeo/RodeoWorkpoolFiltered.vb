Option Explicit

'************** Subs in this Module ****************

'************** /Subs in this Module ****************

Sub RodeoAddWorkpool()
StartTime = Timer
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
    Set ImportWbk = Workbooks(strRodeoWorkpoolFileName)
    ImportWbk.Worksheets.Add ImportWbk.Worksheets(1)
    Set importWS = ImportWbk.Worksheets(1)
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
    URL = "https://tiny.amazon.com/md8cl8ex/rodeamazDTM2Item" 'dwell time <30min

    Debug.Print "URL = " & URL
    Set qtRodeoTotal = importWS.QueryTables.Add(Connection:="URL;" & URL, Destination:=Range("A1"))
        With qtRodeoTotal  'Datei und Zielort auswählen
'        Spaltenbreiten bleiben aktuell erhalten
            .Name = "qTRodeoWorkpool"
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
'    qtDeleteInAllWbks
    counter = Format(Now, "DD.MM_HH.mm.ss")
    importWS.Name = "RodeoWkpool" & counter
    Call fDeleteColumns(importWS)
    importWS.Columns.AutoFit
'    ImportWbk.Save
SecondsElapsed = Round(Timer - StartTime, 2)
    Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds"
    Exit Sub
Whoa:
    MsgBox Err.Description
    Exit Sub
End Sub

Sub UpdateRodeoWorkpool()
StartTime = Timer
    Dim importWS As Worksheet
    Dim qt As QueryTable
    Dim ImportWbk As Workbook
    Dim counter As String
    Set ImportWbk = Workbooks(strRodeoWorkpoolFileName)
    Set importWS = ImportWbk.Worksheets(1)
    Set qt = importWS.QueryTables(1)

    With qt
        .Refresh
    End With
    counter = Format(Now, "DD.MM_HH.mm.ss")
    importWS.Name = "RodeoWkpool" & counter

    Call fDeleteColumns(importWS)
    importWS.Columns.AutoFit
    Call fRodeoColumnsWidth(importWS)
    SecondsElapsed = Round(Timer - StartTime, 2)
    Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds"
End Sub
