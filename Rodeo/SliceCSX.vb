Option Explicit

Sub TestOnSliceCSX()
Application.ScreenUpdating = False
StartTime = Timer
    Dim cTimestamp As Integer
    Dim importWS As Worksheet
    Dim lastrow As Long
    Dim currentrow As Long
    Dim lrow As Long
    Dim dataSetcsx As String
    Dim uniqueRow As Long
'    Dim collUniqueDicts As Dictionary
    Dim collUniqueCounter As Collection
    Dim collImportWSnames As Collection
    Dim collRuntimes As Collection
    Dim collUniqeCSXCounter As Collection
    Dim ImportWbk As Workbook
    Dim counter As Long
    Dim csxWbk As Workbook
    Dim slice
    Dim app As Application
    Dim testarray As Variant
    Dim timeStampscsx As String
    Dim csxDict As Dictionary
    Dim csxBetweenDicts As Dictionary

    Set csxDict = New Dictionary
    Set csxBetweenDicts = New Dictionary
    Set collUniqueCounter = New Collection
    Set collImportWSnames = New Collection
    Set collUniqueDicts = New Collection
    Set collRuntimes = New Collection
    Set collUniqeCSXCounter = New Collection
    Set app = Application
'    Set csxWbk = Workbooks.Open(FileName:=strcsxStampsFile, UpdateLinks:=False)
    Set ImportWbk = Workbooks(strRodeoHistoryFileName)
    For Each importWS In ImportWbk.Worksheets
        timeStampscsx = Right(importWS.Name, 9)
        lastrow = fLastWrittenRow(importWS, 1)
        'Original Data
        ReDim csxData(1 To lastrow, 1)
        'Unique csx Arrays
        ReDim csxDataUniqe(1 To lastrow, 1)
        'Fill arrays with values'
        csxData = importWS.Range("i2:i" & lastrow).Value2
        'ReDim Unique Arrays'
        ReDim csxDataUniqe(1 To 1, 1 To 1)
        counter = 0
            Set csxDict = fCreateUniqueCSXDict(csxData)     'get a Dict of unique values
'            Set collUniqeCSXCounter = fAddUniqueCSXcounterToACollection(csxDict)
                'Collections to track data on each repeat step
                collUniqueCounter.Add csxDict.Count
                collImportWSnames.Add importWS.Name
                SecondsElapsed = Round(Timer - StartTime, 0)
                collRuntimes.Add SecondsElapsed
                collUniqueDicts.Add csxDict, timeStampscsx

                Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds"
                Erase csxData
'                Erase OuterScannableDataFiltered
'                Erase OuterContainerDataFiltered
'                Erase WorkpoolDataFiltered
'            csxDict.RemoveAll
            Set csxDict = New Dictionary
            cTimestamp = cTimestamp + 1
        Next importWS
    Set csxBetweenDicts = fJoinDictionaries(collUniqueDicts, collImportWSnames)
'    For Each ws In csxWbk.Worksheets
'        ws.Cells.Columns.AutoFit
'    Next ws
End Sub
