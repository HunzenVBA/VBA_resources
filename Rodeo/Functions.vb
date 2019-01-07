Option Explicit



'************** Subs in this Module ****************



'************** /Subs in this Module ****************




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
Function PopulateFullColumn(col As String)
    ThisWorkbook.ActiveSheet.Range(col & ":" & col).Value = "DummyValue"
End Function
Function fLastWrittenRow(ws As Worksheet, Column As Long) As Long
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


Sub fDeleteAllQueries(ws As Worksheet)
Dim qt As QueryTable

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

Function fCreateUniqueCSXDict(inputarray As Variant) As Dictionary
    Dim result As Dictionary
    Dim counter As Long
    Dim uniqueRow As Long

    Set result = New Dictionary

    counter = 0
        For uniqueRow = LBound(inputarray, 1) To UBound(inputarray, 1)
            If Not result.Exists(inputarray(uniqueRow, 1)) Then
                counter = counter + 1
                globalcounter = globalcounter + 1
                'Add to Dictionary
                result.Add inputarray(uniqueRow, 1), globalcounter
            End If
        Next uniqueRow
    Set fCreateUniqueCSXDict = result
End Function

Function fCompareIDsbetweenDicts(collOfDicts As Variant) As Dictionary
    Dim resultDict As Dictionary
    Dim currIteminColl As Dictionary
    Dim currKey As Variant
    Dim cIDsbetweenDicts As Long
    Dim collAllUniqeCSX As Collection

    Set resultDict = New Dictionary
    Set collAllUniqeCSX = New Collection

    For Each currIteminColl In collOfDicts          'curritem = Dictionary
        For Each currKey In currIteminColl.Keys         'currkey = csx
            If Not resultDict.Exists(currKey) Then
            cIDsbetweenDicts = cIDsbetweenDicts + 1
                resultDict.Add currKey, cIDsbetweenDicts
            End If
        Next currKey
        Set collAllUniqeCSX = fJoinDictionaries
    Next currIteminColl

    Set fCompareIDsbetweenDicts = resultDict
End Function

Function fAddUniqueCSXcounterToACollection(DictOfcsx As Dictionary) As Long
    Dim resultColl As Collection
    Set resultColl = New Collection
    resultColl.Add DictOfcsx.Count
    Set fAddUniqueCSXcounterToACollection = resultColl
End Function


Function mergeArrays(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    Dim holdarr As Variant
    Dim ub1 As Long
    Dim ub2 As Long
    Dim bi As Long
    Dim i As Long
    Dim newind As Long
        ub1 = UBound(arr1) + 1
        ub2 = UBound(arr2) + 1
        bi = IIf(ub1 >= ub2, ub1, ub2)
        ReDim holdarr(ub1 + ub2 - 1)
        For i = 0 To bi
            If i < ub1 Then
                holdarr(newind) = arr1(i)
                newind = newind + 1
            End If
            If i < ub2 Then
                holdarr(newind) = arr2(i)
                newind = newind + 1
            End If
        Next i
        mergeArrays = holdarr
End Function

Function fJoin2Dictionaries(dict1 As Dictionary, dict2 As Dictionary) As Dictionary
    Dim result As Dictionary
    Dim counter As Long
    Dim uniqueRow As Long
    Dim dict1key As Variant
    Dim dict2key As Variant
    Set result = New Dictionary
    Debug.Print dict1.Items()(1), dict1.Keys()(1)
    counter = 0
        For Each dict1key In dict1.Keys
            If dict2.Exists(dict1key) Then
                counter = counter + 1
                result.Add dict1key, counter
            End If
        Next dict1key
    Set fJoinDictionaries = result
End Function

Function fJoinDictionariesTestData(collOfDicts As Collection) As Dictionary
    Dim result As Dictionary
    Dim counter As Long
    Dim uniqueRow As Long
    Dim DictInColl As Dictionary
    Dim holdDict As Dictionary
    Dim dict1key As Variant
    Dim dict2key As Variant
    Dim cDict As Integer
    Dim currSheet As Worksheet

    Dim wbkCsxByDict As Workbook

    Set wbkCsxByDict = Workbooks("CsxByDict.xlsm")
    Set currSheet = Worksheets("outputTestOnSliceCSX")
    wbkCsxByDict.Worksheets("outputTestOnSliceCSX").Range("A2:F" & fLastWrittenRow(currSheet, 1)).ClearContents

    Set result = New Dictionary
    Set holdDict = New Dictionary
'    Debug.Print dict1.Items()(1), dict1.Keys()(1)
    counter = 0
    cDict = 0
    For Each DictInColl In collOfDicts
        If cDict < 1 Then                   'für den ersten Durchgang Holddict mit DictinColl füllen
        Set holdDict = DictInColl           'first Dict is a hold dict. All csx of this will remain, but only if they can be found in all other Dicts in Coll
        End If
        For Each dict1key In holdDict.Keys

        Debug.Print dict1key
            If Not DictInColl.Exists(dict1key) Then       'must not exist in result Dict
                    counter = counter + 1
                    holdDict.Remove (dict1key)
            End If
        Next dict1key
    cDict = cDict + 1
    wbkCsxByDict.Worksheets("outputTestOnSliceCSX").Cells(2, cDict).Resize(DictInColl.Count, 1) = Application.Transpose(DictInColl.Keys)
    wbkCsxByDict.Worksheets("outputTestOnSliceCSX").Cells(2, cDict + 3).Resize(holdDict.Count, 1) = Application.Transpose(holdDict.Keys)
    Next DictInColl
    Set result = holdDict
    Set fJoinDictionaries = result                                      'output result as function value
    Call fSortColumnsIndividually(wbkCsxByDict.Worksheets("outputTestOnSliceCSX"))
End Function

Function fSortColumnsIndividually(ws As Worksheet)
Dim intColumn As Long
    For intColumn = 1 To 8
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Range(Cells(1, intColumn), Cells(fLastWrittenRow(ws, intColumn), intColumn)), xlSortOnValues, xlAscending
            .SetRange ws.Range(Cells(1, intColumn), Cells(fLastWrittenRow(ws, intColumn), intColumn))
            .Header = xlYes
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    Next intColumn
End Function

Function fAutofitWS(ws As Worksheet)
With ws
 .Columns.AutoFit
End With
End Function
Function fJoinDictionaries(collOfDicts As Collection, collOfDictNames As Collection) As Dictionary
    Dim result As Dictionary
    Dim counter As Long
    Dim uniqueRow As Long
    Dim DictInColl As Dictionary
    Dim holdDict As Dictionary
    Dim dict1key As Variant
    Dim dict2key As Variant
    Dim cDict As Integer
    Dim currSheet As Worksheet

    Dim wbkCsxByDict As Workbook

    Set wbkCsxByDict = Workbooks("CsxByDict.xlsm")
    Set currSheet = wbkCsxByDict.Worksheets("Slice1minCSX")
    wbkCsxByDict.Worksheets("Slice1minCSX").Range("A2:BB" & fLastWrittenRow(currSheet, 1)).ClearContents

    Set result = New Dictionary
    Set holdDict = New Dictionary
'    Debug.Print dict1.Items()(1), dict1.Keys()(1)
    counter = 0
    cDict = 0
    For Each DictInColl In collOfDicts
        If cDict < 1 Then                   'für den ersten Durchgang Holddict mit DictinColl füllen
        Set holdDict = DictInColl           'first Dict is a hold dict. All csx of this will remain, but only if they can be found in all other Dicts in Coll
        End If
        For Each dict1key In holdDict.Keys

'        Debug.Print dict1key
            If Not DictInColl.Exists(dict1key) Then       'must not exist in result Dict
                    counter = counter + 1
                    holdDict.Remove (dict1key)
            End If
        Next dict1key
    cDict = cDict + 1
'    wbkCsxByDict.Worksheets("Slice1minCSX").Cells(2, cDict).Resize(DictInColl.Count, 1) = Application.Transpose(DictInColl.Keys)
    wbkCsxByDict.Worksheets("Slice1minCSX").Cells(1, cDict).Value2 = collOfDictNames(cDict)
    wbkCsxByDict.Worksheets("Slice1minCSX").Cells(2, cDict).Resize(holdDict.Count, 1) = Application.Transpose(holdDict.Keys)
    Next DictInColl
    Set result = holdDict
    Set fJoinDictionaries = result                                      'output result as function value
'    Call fSortColumnsIndividually(wbkCsxByDict.Worksheets("Slice1minCSX"))         'deactivated for the moment
End Function

Function fJoinDictionariesSlice(collOfDicts As Collection, collOfDictNames As Collection) As Dictionary
    Dim result As Dictionary
    Dim counter As Long
    Dim uniqueRow As Long
    Dim DictInColl As Dictionary
    Dim holdDict As Dictionary
    Dim dict1key As Variant
    Dim dict2key As Variant
    Dim cDict As Integer
    Dim currSheet As Worksheet

    Dim wbkCsxByDict As Workbook

    Set wbkCsxByDict = Workbooks("CsxByDict.xlsm")
    Set currSheet = wbkCsxByDict.Worksheets("SliceCSX")
    wbkCsxByDict.Worksheets("SliceCSX").Range("A2:BB" & fLastWrittenRow(currSheet, 1)).ClearContents

    Set result = New Dictionary
    Set holdDict = New Dictionary
'    Debug.Print dict1.Items()(1), dict1.Keys()(1)
    counter = 0
    cDict = 0
    For Each DictInColl In collOfDicts
        If cDict < 1 Then                   'für den ersten Durchgang Holddict mit DictinColl füllen
        Set holdDict = DictInColl           'first Dict is a hold dict. All csx of this will remain, but only if they can be found in all other Dicts in Coll
        End If
        For Each dict1key In holdDict.Keys

'        Debug.Print dict1key
            If Not DictInColl.Exists(dict1key) Then       'must not exist in result Dict
                    counter = counter + 1
                    holdDict.Remove (dict1key)
            End If
        Next dict1key
    cDict = cDict + 1
'    wbkCsxByDict.Worksheets("SliceCSX").Cells(2, cDict).Resize(DictInColl.Count, 1) = Application.Transpose(DictInColl.Keys)
    wbkCsxByDict.Worksheets("SliceCSX").Cells(1, cDict).Value2 = collOfDictNames(cDict)
    wbkCsxByDict.Worksheets("SliceCSX").Cells(2, cDict).Resize(holdDict.Count, 1) = Application.Transpose(holdDict.Keys)
    Next DictInColl
    Set result = holdDict
    Set fJoinDictionaries = result                                      'output result as function value
'    Call fSortColumnsIndividually(wbkCsxByDict.Worksheets("SliceCSX"))         'deactivated for the moment
End Function
Function fDeleteColumns(ws As Worksheet)
'Delete Status and Process Path columns = because useless
ws.Range("M1").EntireColumn.Delete Shift:=xlLeft        'Status ist immer Crossdock
ws.Range("H1").EntireColumn.Delete Shift:=xlLeft        'Pick Priority ist immer Min
ws.Range("G1").EntireColumn.Delete Shift:=xlLeft        'Process Path ist immer leer
ws.Range("C1").EntireColumn.Delete Shift:=xlLeft        'Next Destination ist immer dasselbe wie Destination Warehouse
End Function

Function fRodeoColumnsWidth(ws As Worksheet)
'Use on filtered/deleted columns version of Rodeo
ws.Columns("A").ColumnWidth = 45        'Transfer Request ID
ws.Columns("B").ColumnWidth = 12        'Destination
ws.Columns("G").ColumnWidth = 20        'Scannable ID
ws.Columns("H").ColumnWidth = 20        'Outer Container Type
ws.Columns("i").ColumnWidth = 13        'Container Type
ws.Columns("L").ColumnWidth = 10        'Dwell Time
End Function


Function fConvertTimestampToDate(inputString As String) As Date
Dim formatted As String
Dim result As Date
If Left(inputString, 7) <> "Tabelle" Then
formatted = Replace(inputString, ".", ":")
result = TimeValue(formatted)
fConvertTimestampToDate = result
End If
End Function

Function fWriteCSXData(Location As String, csxID As String, csxLastTimeStamp) As clsCsx

Dim csx As clsCsx
Set csx = New clsCsx

csx.csxID = csxID
csx.Location = Location
csx.LastTimestamp = csxLastTimeStamp
fWriteCSXData = csx
End Function

Function fDeleteRowsInArray(inputarray As Variant, SelectCondition As Variant) As Variant

Const COMPARE_COL As Long = 1
Dim a, aNew(), nr As Long, nc As Long
Dim r As Long, col As Long, rNew As Long
Dim tmp As Variant

    nr = UBound(inputarray, 1)
    nc = UBound(inputarray, 2)

    ReDim aNew(1 To nr, 1 To nc)
    rNew = 0

    For r = 1 To nr
        tmp = inputarray(r, COMPARE_COL)
        If tmp = SelectCondition Then
            rNew = rNew + 1
            For col = 1 To nc
                aNew(rNew, col) = inputarray(r, col)
            Next col
        End If
    Next r

    fDeleteRowsInArray = aNew
End Function

Function fWriteDictionariesToWS(ws As Worksheet, collOfDicts As Collection)
Dim dictCsxUpdatedLastTimestamp As Dictionary
Dim dictCsxUpdatedLastLocation As Dictionary
Dim dictCsxUpdatedOutCont As Dictionary
Dim dictCsxUpdatedDwell As Dictionary

Set dictCsxUpdatedLastTimestamp = collOfDicts(1)
Set dictCsxUpdatedLastLocation = collOfDicts(2)
Set dictCsxUpdatedOutCont = collOfDicts(3)
Set dictCsxUpdatedDwell = collOfDicts(4)

    ws.Cells(2, 4).Resize(dictCsxUpdatedLastTimestamp.Count, 1) = Application.Transpose(dictCsxUpdatedLastTimestamp.Keys)
    ws.Cells(2, 5).Resize(dictCsxUpdatedLastTimestamp.Count, 1) = Application.Transpose(dictCsxUpdatedLastLocation.Items)
    ws.Cells(2, 6).Resize(dictCsxUpdatedLastTimestamp.Count, 1) = Application.Transpose(dictCsxUpdatedOutCont.Items)
    ws.Cells(2, 7).Resize(dictCsxUpdatedLastTimestamp.Count, 1) = Application.Transpose(dictCsxUpdatedLastTimestamp.Items)
    ws.Cells(2, 8).Resize(dictCsxUpdatedLastTimestamp.Count, 1) = Application.Transpose(dictCsxUpdatedDwell.Items)

    ws.Cells(2, 7).Value2 = "LastTimestamp"
'    ws.Cells(1, 6).Value2 = "OuterContainer"
'    ws.Cells(1, 5).Value2 = "OuterScannable"
'    ws.Cells(1, 4).Value2 = "ScannableID"
'    ws.Cells(1, 8).Value2 = "Dwell Time (hours)"
    ws.Cells(2, 9).Value2 = "Process"

End Function

Function fWriteLocationMapping(outputws As Worksheet, dictOutContIDs As Dictionary) As Dictionary
Dim currentrow As Long
Dim lastrow As Long
Dim strOwner As String
Dim strOutContID As String
Dim wsMapping As Worksheet
Dim key As Variant
Dim result As Dictionary
Dim strResultAdress As Range
Dim rngSearchRange As Range
Dim collOutput As Collection
Dim collMissingOwner As Collection
Dim dictMissingOwners As Dictionary

Set result = New Dictionary
'Dim outputws As Worksheet

Set collOutput = New Collection
Set collMissingOwner = New Collection
Set dictMissingOwners = New Dictionary

Set wsMapping = ThisWorkbook.Worksheets("LocationMapping")
Set rngSearchRange = wsMapping.Range("A1:A" & fLastWrittenRow(wsMapping, 1))
'Set outputws = ThisWorkbook.Worksheets("AllCsx")

lastrow = dictOutContIDs.Count


    For Each key In dictOutContIDs.Keys
        currentrow = currentrow + 1
        strOutContID = dictOutContIDs(key)
        If strOutContID <> "" Then
            strOwner = WorksheetFunction.VLookup(strOutContID, wsMapping.Range("A1:C" & fLastWrittenRow(wsMapping, 1)), 3)
            result.Add currentrow, strOwner
            collOutput.Add strOutContID
            collOutput.Add strOwner

            Set strResultAdress = rngSearchRange.Find(strOutContID)
            If Not strResultAdress Is Nothing Then
                    collOutput.Add strResultAdress.Address
            Else
                    Debug.Print strOutContID
                    collMissingOwner.Add strOutContID
                    If dictMissingOwners.Exists(strOutContID) Then
                    Else
                        dictMissingOwners.Add strOutContID, strOutContID & "_missing"
                    End If
            End If
        End If
    Next key

    Set fWriteLocationMapping = result
    outputws.Cells(3, 9).Resize(result.Count, 1) = Application.Transpose(result.Items)
End Function
