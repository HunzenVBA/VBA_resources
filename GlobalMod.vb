Option Explicit

Sub testfiles()
    'FileDirectories
    If Dir(strYardFile) = "" Then
      MsgBox "File" & strYardFile & "does not exist"
    End If
    If Dir(strUnprocessedFile) = "" Then
      MsgBox "File" & strUnprocessedFile & "does not exist"
    End If
    If Dir(strDMdata) = "" Then
      MsgBox "File" & strDMdata & "does not exist"
    End If
    If Dir(strDockmasterFile) = "" Then
      MsgBox "File" & strDockmasterFile & "does not exist"
    End If
    If IsError(Workbooks(strYardHistory)) Then
      MsgBox "File" & strYardHistory & "does not exist"
    End If
'    'FileNames
        If IsError(Workbooks(strYardFileName)) Then
      MsgBox "File" & strYardFileName & "does not exist"
    End If
End Sub

Sub SichernUnprocessed()
    Dim PW As String
    Dim Blatt As Integer
    Dim ws As Worksheet
    Application.ScreenUpdating = False

    PW = "4878"

    For Blatt = 1 To Workbooks(strUnprocessedFileName).Worksheets.Count             'Wbk muss vorher geöffnet sein
        Worksheets(Blatt).Protect (PW), _
        UserInterFaceOnly:=True
    Next
    'Test Protection
    For Each ws In Workbooks(strUnprocessedFileName).Worksheets                     'Wbk muss vorher geöffnet sein
        Debug.Print ws.Index & " " & ws.Name & " Protected: " & ws.ProtectContents
    Next ws
    Application.ScreenUpdating = True
    'Sichern der Datenblätter, um ungewollten Änderungen Vorzubeugen (Ausgenommen Datenblätter 1,5 und 8)
    'Wird von jeder Sub zuletzt aufgerufen
    'Passwort ist DTM2
End Sub
Sub EntsichernUnprocessed()
    'If ThisWorkbook.Name <> "Brückenübersicht.xlsm" Then
    'MsgBox ("Bitte die Datei 'Brückenübersicht.xlsm' benutzen!")
    'Exit Sub
    'End If
    Dim intBlatt As Integer
    Dim intSichtbar As Integer
    Dim PW As String
    Dim ws As Worksheet

    PW = "4878"

    For intBlatt = 1 To Workbooks(strUnprocessedFileName).Worksheets.Count              'Wbk muss vorher geöffnet sein
        Worksheets(intBlatt).Unprotect (PW)
    Next
    'Test Protection
    For Each ws In Workbooks(strUnprocessedFileName).Worksheets                         'Wbk muss vorher geöffnet sein
        Debug.Print ws.Index & " " & ws.Name & " Protected: " & ws.ProtectContents
    Next ws
    'Entsichern der Datenblätter zur Bearbeitung (Sub wird von jeder anderen Sub zu beginn aufgerufen)
End Sub

Sub AllSheetsVisible()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    ws.Visible = xlSheetVisible
Next ws
End Sub

Function HideSheet(ws As Worksheet)
    ws.Visible = xlSheetHidden
End Function
Function HideSheetVery(ws As Worksheet)
    ws.Visible = xlSheetVeryHidden
End Function
' ========================================================================================================
Function fCopyAndPasteWorksheet(ws As Worksheet)
    With ws
        .Copy After:=Worksheets(ws.Index)
    End With
End Function

Function fDeleteEmptyWorksheetsInThisWorkbook()

    For Each ws In ThisWorkbook.Worksheets
        If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
            ws.Delete
        Else
        End If
    Next ws

End Function

Function fDeleteEmptyRows(Optional AllWorksheets As Boolean = False)
    Dim ws As Worksheet
    Dim LastWrittenRow As Long
    Dim LastRowOnSheet As Long
    Dim EmptyRowsRange As Range

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
    If AllWorksheets = True Then
        For Each ws In ThisWorkbook.Worksheets
            Set ws = ThisWorkbook.ActiveSheet
            LastRowOnSheet = ws.Rows.Count
            LastWrittenRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            Set EmptyRowsRange = ws.Range("A" & LastWrittenRow + 1 & ":A" & LastRowOnSheet)

            With EmptyRowsRange
                .EntireRow.Delete Shift:=xlUp
            End With
            ActiveSheet.UsedRange.SpecialCells (xlCellTypeLastCell)
        Next ws
    Else
            Set ws = ThisWorkbook.ActiveSheet
            LastRowOnSheet = ws.Rows.Count
            LastWrittenRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            Set EmptyRowsRange = ws.Range("A" & LastWrittenRow + 1 & ":A" & LastRowOnSheet)

            With EmptyRowsRange
                .EntireRow.Delete Shift:=xlUp
            End With
            ActiveSheet.UsedRange.SpecialCells (xlCellTypeLastCell)
    End If

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
End Function
Function PopulateFullColumn(Col As String)
    ThisWorkbook.ActiveSheet.Range(Col & ":" & Col).Value = "DummyValue"
End Function
Function fLastWrittenRow(ws As Worksheet, Column As Integer) As Integer
    fLastWrittenRow = ws.Cells(Rows.Count, Column).End(xlUp).Row
End Function
Function fStartTimer()
    StartTimerVar = Now
End Function
Function fEndTimer()
    EndTimerVar = Now
End Function
Function fPrintSubRuntime()
    Dim PrintTimeVar As Date
    PrintTimeVar = EndTimerVar - StartTimerVar
    Debug.Print "sub(s) finished in " & Format(PrintTimeVar, "s") & " seconds."
End Function
' Retrieve current time and ws Name as a string
Function fZeitstempelDesWorksheet(ws As Worksheet) As String
    fZeitstempelDesWorksheet = Format(Now, "hh:mm:ss")
End Function

Function fListAllModules(Optional AllWorkbooks As Boolean = False)
    Call GetFunctionAndSubNames 'in a seperate module
End Function
'
Function fRefreshZeitstempel()
    Dim ws                  As Worksheet
    Dim Zeitstempels        As Dictionary
    Dim i                   As Integer
    Dim j                   As Integer
    Dim k                   As Integer
    Set Zeitstempels = New Dictionary

        Zeitstempels.RemoveAll

        For Each ws In ThisWorkbook.Worksheets
            Zeitstempels.Add ws.Index, ws.Name & " " & fZeitstempelDesWorksheet(ws)
        Next ws

        For i = 0 To Zeitstempels.Count - 1 Step 1
'            ControlSheet.Range("A" & (i + 3)) = Zeitstempels.Keys(i)                'careful Range(0) not possible
            ControlSheet.Range("B" & (i + 3)) = Zeitstempels.Items(i)
        Next i
End Function

Function WriteZeitstempelToControlSheet()
    Dim collZeitstempel(1 To 4, 1 To 2)         As String
    Dim i                                       As Integer
    Dim r                                       As Range
    Dim Dimension1                              As Long
    Dim Dimension2                              As Long

    Dimension1 = UBound(collZeitstempel, 1)
    Dimension2 = UBound(collZeitstempel, 2)
        collZeitstempel(1, 1) = ControlSheet.Name
        collZeitstempel(2, 1) = ImportSheet.Name
        collZeitstempel(3, 1) = ImportSheet2.Name
        collZeitstempel(4, 1) = ""

        collZeitstempel(1, 2) = ZeitstempelWS1
        collZeitstempel(2, 2) = ZeitstempelWS2
        collZeitstempel(3, 2) = ZeitstempelWS3
        collZeitstempel(4, 2) = ZeitstempelWS4

'        If Len(Join(collZeitstempel( , 2)) = 0 Then
'            MsgBox "Zeitstempel berechnen!"
'        End If

'        For i = LBound(collZeitstempel, 1) To UBound(collZeitstempel, 1) 'ws Names
'            ControlSheet.Range("D" & i + 4).Value = collZeitstempel(i, 1)
'        Next i
'
'        For i = LBound(collZeitstempel, 1) To UBound(collZeitstempel, 1) 'Zeitstempel
'            ControlSheet.Range("E" & i + 4).Value = collZeitstempel(i, 2)
'        Next i

        Set r = ControlSheet.Range("A3", ControlSheet.Range("A3").Offset(Dimension1 - 1, Dimension2 - 1))
        r.Value = collZeitstempel
End Function

Function fEnableFormula(ws As Worksheet)
    ws.Cells.FormulaHidden = False
End Function
