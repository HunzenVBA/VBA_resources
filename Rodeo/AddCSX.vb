Option Explicit

Sub BuildCSXdictTest()
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
    Dim dictCsxUpdatedLastTimestamp As Dictionary
    Dim dictCsxUpdatedLastLocation As Dictionary
    Dim dictCsxUpdatedOutCont As Dictionary
    Dim dictCsxUpdatedDwell As Dictionary
    Dim tempTimstampDict As Date
    Dim csxKey As Variant
    Dim csxDwell As Variant



    Set dictCsxUpdatedLastTimestamp = New Dictionary
    Set dictCsxUpdatedLastLocation = New Dictionary
    Set dictCsxUpdatedOutCont = New Dictionary
    Set dictCsxUpdatedDwell = New Dictionary
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
        ReDim DwellData(1 To lastrow, 1)
        'Fill arrays with values'
        csxData = importWS.Range("f1:f" & lastrow).Value2
        OuterScannableData = importWS.Range("g1:g" & lastrow).Value2
        OuterContainerData = importWS.Range("h1:h" & lastrow).Value2
        DwellData = importWS.Range("l1:l" & lastrow).Value2

        For currentrow = 1 To lastrow
            Set csx = New clsCsx
            Set csxDict = fCreateUniqueCSXDict(csxData)     'get a Dict of unique values
            csx.csxID = csxData(currentrow, 1)
            csx.DwellTime = DwellData(currentrow, 1)
            csxKey = csx.csxID
            csxDwell = csx.DwellTime

            'Timestamp
                If cTimestamp < 1 Then
                    csx.LastTimestamp = timeStampscsx
                    csx.Location = OuterScannableData(currentrow, 1)
                    csx.OutContainer = OuterContainerData(currentrow, 1)
                    csx.csxID = csxData(currentrow, 1)
                    csx.DwellTime = DwellData(currentrow, 1)
                    collCsx.Add csx
                    dictCsxUpdatedLastTimestamp.Add csx.csxID, csx.LastTimestamp
                    dictCsxUpdatedLastLocation.Add csx.csxID, csx.Location
                    dictCsxUpdatedOutCont.Add csx.csxID, csx.OutContainer
                    dictCsxUpdatedDwell.Add csx.csxID, csx.DwellTime
                    tempTimstampDict = csx.LastTimestamp
                End If
                If csx.LastTimestamp < timeStampscsx Then           'neuere timestamp 'nur noch Loc und ID adden
                For Each csxKey In dictCsxUpdatedLastTimestamp
                    If dictCsxUpdatedLastTimestamp.Exists(csx.csxID) Then
                        csx.LastTimestamp = timeStampscsx
                        csx.Location = OuterScannableData(currentrow, 1)
                        csx.OutContainer = OuterContainerData(currentrow, 1)
                        csx.csxID = csxData(currentrow, 1)
                        dictCsxUpdatedLastTimestamp.Item(csx.csxID) = csx.LastTimestamp
                        dictCsxUpdatedLastLocation.Item(csx.csxID) = csx.Location
                        dictCsxUpdatedOutCont(csx.csxID) = csx.OutContainer
                        dictCsxUpdatedDwell.Item(csx.csxID) = csx.DwellTime
                        collCsx.Add csx
                        tempTimstampDict = csx.LastTimestamp
                    End If
                Next csxKey
                End If
        Next currentrow
        counter = 0

                'Collections to track data on each repeat step
                collImportWSnames.Add importWS.Name
                SecondsElapsed = Round(Timer - StartTime, 0)
                collRuntimes.Add SecondsElapsed
                Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds"
                Erase csxData
            Set csxDict = New Dictionary
            cTimestamp = cTimestamp + 1
        Next importWS
    wbkcsxObj.Worksheets("csx").UsedRange.ClearContents

        wbkcsxObj.Worksheets("csx").Cells(2, 4).Resize(dictCsxUpdatedLastTimestamp.Count, 1) = Application.Transpose(dictCsxUpdatedLastTimestamp.Keys)
        wbkcsxObj.Worksheets("csx").Cells(2, 5).Resize(dictCsxUpdatedLastTimestamp.Count, 1) = Application.Transpose(dictCsxUpdatedLastLocation.Items)
        wbkcsxObj.Worksheets("csx").Cells(2, 6).Resize(dictCsxUpdatedOutCont.Count, 1) = Application.Transpose(dictCsxUpdatedOutCont.Items)
        wbkcsxObj.Worksheets("csx").Cells(2, 7).Resize(dictCsxUpdatedLastTimestamp.Count, 1) = Application.Transpose(dictCsxUpdatedLastTimestamp.Items)
        wbkcsxObj.Worksheets("csx").Cells(2, 8).Resize(dictCsxUpdatedLastTimestamp.Count, 1) = Application.Transpose(dictCsxUpdatedDwell.Items)

        wbkcsxObj.Worksheets("csx").Cells(2, 7).Value2 = "LastTimestamp"

End Sub
