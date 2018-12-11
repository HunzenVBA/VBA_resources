Option Explicit

Sub BuildCSXdict()
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
    Dim timeStampscsx As Date
    Dim csxDict As Dictionary
    Dim csxBetweenDicts As Dictionary
    Dim csx As clsCsx
    Dim collCsx As Collection
    Dim wbkcsxObj As Workbook
    Dim dictCsxUpdated As Dictionary


    Set dictCsxUpdated = New Dictionary
    Set collCsx = New Collection
    Set csxDict = New Dictionary
    Set csxBetweenDicts = New Dictionary
    Set collUniqueCounter = New Collection
    Set collImportWSnames = New Collection
    Set collUniqueDicts = New Collection
    Set collRuntimes = New Collection
    Set collUniqeCSXCounter = New Collection
    Set app = Application
    Set wbkcsxObj = Workbooks("AllCsxObjects.xlsm")
'    Set csxWbk = Workbooks.Open(FileName:=strcsxStampsFile, UpdateLinks:=False)
    Set ImportWbk = Workbooks("RodeoImport10minTestData.xlsm")
    For Each importWS In ImportWbk.Worksheets
        timeStampscsx = fConvertTimestampToDate(Right(importWS.Name, 8))
        lastrow = fLastWrittenRow(importWS, 2)
        'Original Data
        ReDim csxData(1 To lastrow, 1)
        ReDim OuterScannableData(1 To lastrow, 1)
        ReDim OuterContainerData(1 To lastrow, 1)
        'Fill arrays with values'
        csxData = importWS.Range("f1:f" & lastrow).Value2
        OuterScannableData = importWS.Range("g1:g" & lastrow).Value2
        OuterContainerData = importWS.Range("j1:j" & lastrow).Value2

        For currentrow = 1 To lastrow
            Set csx = New clsCsx
            Set csxDict = fCreateUniqueCSXDict(csxData)     'get a Dict of unique values

            'Timestamp
                If cTimestamp < 1 Then
                    csx.LastTimestamp = timeStampscsx
                    csx.Location = OuterScannableData(currentrow, 1)
                    csx.csxID = csxData(currentrow, 1)
                    collCsx.Add csx
                    dictCsxUpdated.Add csx.csxID, csx.LastTimestamp
                End If
                If csx.LastTimestamp < timeStampscsx Then           'neuere timestamp 'nur noch Loc und ID adden

                    If dictCsxUpdated.Exists(csx.csxID) Then
                        csx.LastTimestamp = timeStampscsx
                        csx.Location = OuterScannableData(currentrow, 1)
                        csx.csxID = csxData(currentrow, 1)
                    End If
                    dictCsxUpdated.Item(csx.csxID) = csx.LastTimestamp
                    collCsx.Add csx
                End If

        Next currentrow

        counter = 0

'            Set collUniqeCSXCounter = fAddUniqueCSXcounterToACollection(csxDict)
                'WriteActualData of CSX

                'Collections to track data on each repeat step
                collImportWSnames.Add importWS.Name
                SecondsElapsed = Round(Timer - StartTime, 0)
                collRuntimes.Add SecondsElapsed
'                collUniqueDicts.Add csxDict, timeStampscsx

                Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds"
                Erase csxData

'            csxDict.RemoveAll
            Set csxDict = New Dictionary
'            Set collCsx = New Collection
            cTimestamp = cTimestamp + 1
        Next importWS
'    Set csxBetweenDicts = fJoinDictionaries(collUniqueDicts, collImportWSnames)

'    wbkcsxObj.Worksheets("csx").Cells(1, cDict).Value2 = collOfDictNames(cDict)
    wbkcsxObj.Worksheets("csx").UsedRange.ClearContents
    For Each csx In collCsx
        wbkcsxObj.Worksheets("csx").Range("A" & fLastWrittenRow(wbkcsxObj.Worksheets("csx"), 1)).Offset(1, 0).Value2 = csx.csxID
        wbkcsxObj.Worksheets("csx").Range("B" & fLastWrittenRow(wbkcsxObj.Worksheets("csx"), 2)).Offset(1, 0).Value2 = csx.Location
        wbkcsxObj.Worksheets("csx").Range("C" & fLastWrittenRow(wbkcsxObj.Worksheets("csx"), 3)).Offset(1, 0).Value2 = csx.LastTimestamp
        wbkcsxObj.Worksheets("csx").Cells(2, 4).Resize(dictCsxUpdated.Count, 1) = Application.Transpose(dictCsxUpdated.Keys)
        wbkcsxObj.Worksheets("csx").Cells(2, 5).Resize(dictCsxUpdated.Count, 1) = Application.Transpose(dictCsxUpdated.Items)
    Next csx

End Sub
