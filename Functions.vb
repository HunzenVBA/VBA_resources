Option Explicit

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
    fLastWrittenRow = ControlSheet.Cells(Rows.Count, Column).End(xlUp).Row
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
