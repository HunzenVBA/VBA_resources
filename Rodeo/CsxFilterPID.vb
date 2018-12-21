Option Explicit



Sub BuildCSXdictPIDFilter()
Application.ScreenUpdating = False
StartTime = Timer
    Dim cRow As Long
    Dim currrow As Long
    Dim cTimestamp As Integer
    Dim importWS As Worksheet
    Dim outputWS As Worksheet
    Dim lastrow As Long
    Dim currentrow As Long
    Dim lrow As Long
    Dim dataSetcsx As String
    Dim uniqueRow As Long
'    Dim collUniqueDicts As Dictionary
    Dim collOutputDicts As Collection
    Dim collUniqueCounter As Collection
    Dim collImportWSnames As Collection
    Dim collRuntimes As Collection
    Dim collUniqeCSXCounter As Collection
    Dim ImportWbk As Workbook
    Dim counter As Long
    Dim csxWbk As Workbook
    Dim slice
    Dim app As Application
    Dim testArray As Variant
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
    Dim importWSname As String


    Set dictCsxUpdatedLastTimestamp = New Dictionary
    Set dictCsxUpdatedLastLocation = New Dictionary
    Set dictCsxUpdatedOutCont = New Dictionary
    Set dictCsxUpdatedDwell = New Dictionary
    Set collCsx = New Collection
    Set collOutputDicts = New Collection
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
    Set ImportWbk = Workbooks("RodeoImport10min.xlsm")
    For Each importWS In ImportWbk.Worksheets
    If Left(importWS.Name, 7) <> "Tabelle" Then
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

        testArray = fDeleteRowsInArray(OuterScannableData, "ws-rcv-pid-1")

        wbkcsxObj.Worksheets("Filtered").Range("A1:A" & UBound(testArray)).Value2 = testArray

        For currentrow = 1 To lastrow
            Set csx = New clsCsx
            Set csxDict = fCreateUniqueCSXDict(csxData)     'get a Dict of unique values
            csx.csxID = csxData(currentrow, 1)
            csx.DwellTime = DwellData(currentrow, 1)
            If Left(OuterScannableData(currentrow, 1), 10) = "ws-rcv-pid" Then
                cRow = cRow + 1
                csxKey = csx.csxID
                csxDwell = csx.DwellTime

                'Timestamp
                    If cTimestamp < 1 Then
                        If Not dictCsxUpdatedLastTimestamp.Exists(csx.csxID) Then

                            csx.LastTimestamp = timeStampscsx
                            csx.Location = OuterScannableData(currentrow, 1)
                            csx.OutContainer = OuterContainerData(currentrow, 1)
                            csx.csxID = csxData(currentrow, 1)
                            csx.DwellTime = DwellData(currentrow, 1)
                            dictCsxUpdatedLastTimestamp.Add csx.csxID, csx.LastTimestamp
                            dictCsxUpdatedLastLocation.Add csx.csxID, csx.Location
                            dictCsxUpdatedOutCont.Add csx.csxID, csx.OutContainer
                            dictCsxUpdatedDwell.Add csx.csxID, csx.DwellTime
                            collCsx.Add csx

                            tempTimstampDict = csx.LastTimestamp
                            End If
                    End If
                    If csx.LastTimestamp < timeStampscsx Then           'neuere timestamp 'nur noch Loc und ID adden
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
                    End If
            End If
            SecondsElapsed = Round(Timer - StartTime, 0)
            collRuntimes.Add SecondsElapsed             'hinzufügen runtime pro Zeile
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
            importWSname = importWS.Name
        End If
        Next importWS

    Set outputWS = wbkcsxObj.Worksheets("csx")
    outputWS.UsedRange.ClearContents
    collOutputDicts.Add dictCsxUpdatedLastTimestamp
    collOutputDicts.Add dictCsxUpdatedLastLocation
    collOutputDicts.Add dictCsxUpdatedOutCont
    collOutputDicts.Add dictCsxUpdatedDwell

    Call fWriteDictionariesToWS(outputWS, collOutputDicts)
        wbkcsxObj.Worksheets("RuntimeBuildCsxDict").Range("A2:A" & fLastWrittenRow(wbkcsxObj.Worksheets("RuntimeBuildCsxDict"), 1)).ClearContents

        For currrow = 2 To collRuntimes.Count
        wbkcsxObj.Worksheets("RuntimeBuildCsxDict").Cells(currrow, 1).Value2 = collRuntimes(currrow)
        Next currrow

End Sub
