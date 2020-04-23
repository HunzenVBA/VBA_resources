Option Explicit

Sub CopycsxAndTimestamp()
Application.ScreenUpdating = False
StartTime = Timer
    Dim cTimestamp As Integer
    Dim importWS As Worksheet
    Dim lastrow As Long
    Dim currentrow As Long
    Dim lrow As Long
    Dim dataSetWorkpool As String
    Dim dataSetcsx As String
    Dim dataSetoutScan As String
    Dim dataSetoutCont As String
    Dim uniqueRow As Long
'    Dim collUniqueDicts As Dictionary
    Dim collUniqueCounter As Collection
    Dim collImportWSnames As Collection
    Dim collRuntimes As Collection
    Dim collWorkpool As Collection
    Dim collUniqeCSXCounter As Collection
    Dim ImportWbk As Workbook
    Dim counter As Long
    Dim csxWbk As Workbook
    Dim slice
    Dim app As Application

    Dim testarray As Variant
    Dim timeStampscsx As Variant
    Dim csxDict As Dictionary
    Dim outScanDict As Dictionary
    Dim outContDict As Dictionary
    Dim workpoolDict As Dictionary
    Dim csxBetweenDicts As Dictionary

    Set csxDict = New Dictionary
    Set outScanDict = New Dictionary
    Set outContDict = New Dictionary
    Set workpoolDict = New Dictionary
    Set csxBetweenDicts = New Dictionary
    Set collUniqueCounter = New Collection
    Set collImportWSnames = New Collection
    Set collUniqueDicts = New Collection
    Set collRuntimes = New Collection
    Set collWorkpool = New Collection
    Set collUniqeCSXCounter = New Collection
    Set app = Application
'    Set csxWbk = Workbooks.Open(FileName:=strcsxStampsFile, UpdateLinks:=False)
    Set ImportWbk = Workbooks(strRodeoHistoryFileName)
    For Each importWS In ImportWbk.Worksheets
        timeStampscsx = Right(importWS.Name, 9)
        lastrow = fLastWrittenRow(importWS, 1)
        'Original Data
        ReDim csxData(1 To lastrow, 1)
        ReDim OuterScannableData(1 To lastrow, 1)
        ReDim OuterContainerData(1 To lastrow, 1)
        ReDim WorkpoolData(1 To lastrow, 1)
        'Filtered
        ReDim csxDataFiltered(1 To lastrow, 1)
        ReDim OuterScannableDataFiltered(1 To lastrow, 1)
        'Unique csx Arrays
        ReDim csxDataUniqe(1 To lastrow, 1)
        ReDim OuterContainerDataUnique(1 To lastrow, 1)
        ReDim OuterScannableDataUnique(1 To lastrow, 1)
        'Fill arrays with values'
        csxData = importWS.Range("I1:I" & lastrow).Value2
        OuterScannableData = importWS.Range("J1:J" & lastrow).Value2
        OuterContainerData = importWS.Range("K1:K" & lastrow).Value2
        WorkpoolData = importWS.Range("O1:O" & lastrow).Value2
        'ReDim Filtered Arrays'
        ReDim csxDataFiltered(1 To 1, 1 To 1)
        ReDim OuterScannableDataFiltered(1 To 1, 1 To 1)
        ReDim OuterContainerDataFiltered(1 To 1, 1 To 1)
        ReDim WorkpoolDataFiltered(1 To 1, 1 To 1)
        'ReDim Unique Arrays'
        ReDim csxDataUniqe(1 To 1, 1 To 1)
        ReDim OuterScannableDataUnique(1 To 1, 1 To 1)
        ReDim OuterContainerDataUnique(1 To 1, 1 To 1)
        counter = 0
            For lrow = 1 To lastrow
                    'watch-variables
                    dataSetWorkpool = WorkpoolData(lrow, 1)
                    dataSetcsx = csxData(lrow, 1)
                    dataSetoutScan = OuterScannableData(lrow, 1)
                    dataSetoutCont = OuterContainerData(lrow, 1)
'                If WorkpoolData(lrow, 1) <> "Palletized" And WorkpoolData(lrow, 1) <> "Loaded" And WorkpoolData(lrow, 1) <> "TransshipSorted" Then
                If WorkpoolData(lrow, 1) <> "Palletized" And WorkpoolData(lrow, 1) <> "Loaded" Then
                    'Resize arrays by value=counter on each hit of conditions
                    counter = counter + 1
                    ReDim Preserve csxDataFiltered(1 To 1, 1 To counter)
                    ReDim Preserve OuterScannableDataFiltered(1 To 1, 1 To counter)
                    ReDim Preserve OuterContainerDataFiltered(1 To 1, 1 To counter)
                    ReDim Preserve WorkpoolDataFiltered(1 To 1, 1 To counter)
                    'Fill new row of array with value on hit conditions
                    'Transpoe array so you can ReDim last dimension later
                    csxDataFiltered(1, counter) = csxData(lrow, 1)
                    OuterScannableDataFiltered(1, counter) = OuterScannableData(lrow, 1)
                    OuterContainerDataFiltered(1, counter) = OuterContainerData(lrow, 1)
                    WorkpoolDataFiltered(1, counter) = WorkpoolData(lrow, 1)
                End If
            Next lrow
            counter = 0
            Set csxDict = fCreateUniqueCSXDict(csxDataFiltered)     'get a Dict of unique values
'            Set collUniqeCSXCounter = fAddUniqueCSXcounterToACollection(csxDict)
                'Collections to track data on each repeat step
                collUniqueCounter.Add csxDict.Count
                collImportWSnames.Add importWS.Name
                SecondsElapsed = Round(Timer - StartTime, 0)
                collRuntimes.Add SecondsElapsed
                collUniqueDicts.Add csxDict, timeStampscsx

                Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds"
                Erase csxDataFiltered
'                Erase OuterScannableDataFiltered
'                Erase OuterContainerDataFiltered
'                Erase WorkpoolDataFiltered
'            csxDict.RemoveAll
            Set csxDict = New Dictionary
            cTimestamp = cTimestamp + 1
        Next importWS
    Set csxBetweenDicts = fJoinDictionaries(collUniqueDicts)
    For Each ws In csxWbk.Worksheets
        ws.Cells.Columns.AutoFit
    Next ws
End Sub
