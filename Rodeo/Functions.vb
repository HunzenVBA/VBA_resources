Option Explicit



'************** Subs in this Module ****************



'************** /Subs in this Module ****************

Sub runStacker()

    Dim StartTimeAll        As Double
    Dim SecondsElapsedAll   As Double

    StartTimeAll = Timer

'    Namensausgabe des Subs im Direktfenster, zur Info/Debugging
    currProcedureName = "NameWorksheets"
    Debug.Print "============================ Beginn Sub " & currProcedureName & " ============================"
    Debug.Print Now
    Call NameWorksheets

    currProcedureName = "PrintWSnames"
    Debug.Print "============================ Beginn Sub " & currProcedureName & " ============================"
    Debug.Print Now
    Call PrintWSnames

    currProcedureName = "UpdateRodeoTotal"
    Debug.Print "============================ Beginn Sub " & currProcedureName & " ============================"
    Debug.Print Now
    Call UpdateRodeoTotal

    currProcedureName = "DeleteEmptyRows"
    Debug.Print "============================ Beginn Sub " & currProcedureName & " ============================"
    Debug.Print Now
    Call DeleteEmptyRows

'    currProcedureName = "DeleteEmptyRows"
'    Debug.Print "============================ Beginn Sub " & currProcedureName & " ============================"
'    Debug.Print Now
'    Call DeleteEmptyRows



'Determine how many seconds code took to run
  SecondsElapsedAll = Round(Timer - StartTimeAll, 1)

'Notify user in seconds
  Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds"
End Sub

Sub NameWorksheets()

    StartTime = Timer

'*****************************************************************************
'** Benennung der WS nach Position
    ThisWorkbook.Worksheets(1).Name = "1RodeoImport"
    ThisWorkbook.Worksheets(2).Name = "2RodeoFormatted"
    ThisWorkbook.Worksheets(3).Name = "3RodeoLU_Anzahl_Rec"
    ThisWorkbook.Worksheets(4).Name = "4RodeoTotalImport"
    ThisWorkbook.Worksheets(5).Name = "5RodeoTotalFormatted"
    ThisWorkbook.Worksheets(6).Name = "6RodeoCaseRec+General"
    ThisWorkbook.Worksheets(7).Name = "7Scanpunkte"
    ThisWorkbook.Worksheets(8).Name = "8"
'*****************************************************************************
'** Zuweisung ws-Variablen anhand der Benennung
    Set ws1 = ThisWorkbook.Worksheets("1RodeoImport")
    Set ws2 = ThisWorkbook.Worksheets("2RodeoFormatted")
    Set ws3 = ThisWorkbook.Worksheets("3RodeoLU_Anzahl_Rec")
    Set ws4 = ThisWorkbook.Worksheets("4RodeoTotalImport")
    Set ws5 = ThisWorkbook.Worksheets("5RodeoTotalFormatted")
    Set ws6 = ThisWorkbook.Worksheets("6RodeoCaseRec+General")
    Set ws7 = ThisWorkbook.Worksheets("7Scanpunkte")
    Set ws8 = ThisWorkbook.Worksheets("8")

    'Determine how many seconds code took to run
    SecondsElapsed = Round(Timer - StartTime, 1)

    'Notify user in seconds
    Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds"

End Sub

Sub PrintWSnames()
    Debug.Print "WS Nr.|Name"
    For Each ws In ThisWorkbook.Worksheets
'    Debug.Print ws.Index & ": "; ws.Name
    Next ws
End Sub

Sub PrintWorksheets() 'Sub für Nummerierung der Sheets im Direktfenster

    Dim lngSheets As Long 'Sheets sind alle Register
    Dim lngWorksheets As Long 'Worksheets sind nur Tabellenblätter
    Dim lngCharts As Long 'Charts sind nur die Diagrammblätter
    Dim i As Long
    lngSheets = ThisWorkbook.Sheets.Count
    lngWorksheets = ThisWorkbook.Worksheets.Count
    lngCharts = ThisWorkbook.Charts.Count

    Debug.Print "WS Nr.|Name"
    For i = 1 To lngWorksheets
    Debug.Print ThisWorkbook.Worksheets(i).Index; " " & ThisWorkbook.Worksheets(i).Name
    Next i

End Sub

'Sub CalculateRunTime_Seconds()
''PURPOSE: Determine how many seconds it took for code to completely run
'
'Dim StartTime As Double
'Dim SecondsElapsed As Double
'
''Remember time when macro starts
'  StartTime = Timer
'
''*****************************
''Insert Your Code Here...
''*****************************
'
''Determine how many seconds code took to run
'  SecondsElapsed = Round(Timer - StartTime, 2)
'
''Notify user in seconds
'  Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
'
'End Sub


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


Function getColName(colNumber As Integer) As String
'return column name when passed column number
    getColName = Cells(1, colNumber).Value
End Function
    Public Function CountUnique(rng As Range) As Integer
        Dim dict As Scripting.Dictionary
        Dim cell As Range
        Set dict = New Scripting.Dictionary
        For Each cell In rng.Cells
             If Not dict.Exists(cell.Value) Then
                dict.Add cell.Value, 0
            End If
        Next
        CountUnique = dict.Count
    End Function


Sub fDeleteAllQueries()
Dim qt As QueryTable
Dim ws As Worksheet

For Each qt In ws.QueryTables
qt.Delete
Next
End Sub


Function IsWorkBookOpen(FileName As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function
Sub test()
    Dim wb As Workbook
    Set wb = GetWorkbook("C:\Users\dick\Dropbox\Excel\Hoops.xls")
    If Not wb Is Nothing Then
        Debug.Print wb.Name
    End If
End Sub

Public Function GetWorkbook(ByVal sFullName As String) As Workbook
    Dim sFile As String
    Dim wbReturn As Workbook
    sFile = Dir(sFullName)
    On Error Resume Next
        Set wbReturn = Workbooks(sFile)
        If wbReturn Is Nothing Then
            Set wbReturn = Workbooks.Open(sFullName)
        End If
    On Error GoTo 0
    Set GetWorkbook = wbReturn
End Function
Function fWorkbookIsOpen(WorkbookName As String) As Boolean
On Error Resume Next
fWorkbookIsOpen = Workbooks(WorkbookName).Name = WorkbookName
End Function
