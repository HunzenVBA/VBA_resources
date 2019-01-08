Option Explicit

Sub BuildCSXdictPIDFilter()
Application.ScreenUpdating = False
StartTime = Timer
    Dim strCSXID As String
    Dim cRow As Long
    Dim currrow As Long
    Dim cTimestamp As Integer
    Dim importWS As Worksheet
    Dim outputws As Worksheet
    Dim lastrow As Long
    Dim currentrow As Long
    Dim lrow As Long
    Dim dataSetcsx As String
    Dim uniqueRow As Long
    Dim countImportWS As Integer
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
    Dim testarray As Variant
    Dim TimeStampImportWS As Date
    Dim csxDict As Dictionary
    Dim csxBetweenDicts As Dictionary
    Dim csx As clsCsx
    Dim collCsx As Collection
    Dim wbkcsxObj As Workbook
    Dim dictCsxUpdatedLastTimestamp As Dictionary
    Dim dictCsxUpdatedLastLocation As Dictionary
    Dim dictCsxUpdatedOutCont As Dictionary
    Dim dictCsxUpdatedDwell As Dictionary
    Dim currentCSXTimeStamp As Date
    Dim csxKey As Variant
    Dim csxDwell As Variant
    Dim importWSname As String
    Dim maxcountImportWS As Integer

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
    'Insert filter for count of Import Sheets
    maxcountImportWS = 3        'Select this via UI later
    For countImportWS = maxcountImportWS To 1 Step -1
'        For Each importWS In ImportWbk.Worksheets       'fängt immer bei worksheets(1) an, also ganz links
        Set importWS = ImportWbk.Worksheets(countImportWS)
        If Left(importWS.Name, 7) <> "Tabelle" Then
            TimeStampImportWS = fConvertTimestampToDate(Right(importWS.Name, 8))
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

            testarray = fDeleteRowsInArray(OuterScannableData, "ws-rcv-pid-1")          'Filterkriterium

            wbkcsxObj.Worksheets("Filtered").Range("A1:A" & UBound(testarray)).Value2 = testarray

'            For currentrow = 1 To lastrow
            For currentrow = 1 To lastrow
                Set csx = New clsCsx
                Set csxDict = fCreateUniqueCSXDict(csxData)     'get a Dict of unique values
                csx.csxID = csxData(currentrow, 1)
                csx.DwellTime = DwellData(currentrow, 1)
                csx.Location = OuterScannableData(currentrow, 1)
                csx.OutContainer = OuterContainerData(currentrow, 1)
                strCSXID = csx.csxID
                If csx.OutContainer <> "PALLET" Then
                                            'Location filter here
                    cRow = cRow + 1
                    csxKey = csx.csxID
                    csxDwell = csx.DwellTime

                    'Timestamp
                        If cTimestamp < 1 Then
                            If Not dictCsxUpdatedLastTimestamp.Exists(csx.csxID) Then
'                            Stop
                                csx.LastTimestamp = TimeStampImportWS                       'Populate initially dictionary
                                dictCsxUpdatedLastTimestamp.Add csx.csxID, csx.LastTimestamp
                                dictCsxUpdatedLastLocation.Add csx.csxID, csx.Location
                                dictCsxUpdatedOutCont.Add csx.csxID, csx.OutContainer
                                dictCsxUpdatedDwell.Add csx.csxID, csx.DwellTime
                                collCsx.Add csx
                                currentCSXTimeStamp = csx.LastTimestamp
                            End If
                        End If
                        If dictCsxUpdatedLastTimestamp.Exists(csx.csxID) Then
'                            Stop
                            csx.LastTimestamp = TimeStampImportWS
                            dictCsxUpdatedLastTimestamp.Item(csx.csxID) = csx.LastTimestamp
                            dictCsxUpdatedLastLocation.Item(csx.csxID) = csx.Location
                            dictCsxUpdatedOutCont(csx.csxID) = csx.OutContainer
                            dictCsxUpdatedDwell.Item(csx.csxID) = csx.DwellTime
                            collCsx.Add csx
                            currentCSXTimeStamp = csx.LastTimestamp
                        End If

                SecondsElapsed = Round(Timer - StartTime, 0)
                collRuntimes.Add SecondsElapsed             'hinzufügen runtime pro Zeile
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
                importWSname = importWS.Name
            End If      'This If clause is to sort out any "Tabelle..." sheets that return error anyways
'        Next importWS
'    Loop    'counter for amount of importWS
    Next countImportWS

    Set outputws = wbkcsxObj.Worksheets("Allcsx")
    outputws.UsedRange.ClearContents
    collOutputDicts.Add dictCsxUpdatedLastTimestamp
    collOutputDicts.Add dictCsxUpdatedLastLocation
    collOutputDicts.Add dictCsxUpdatedOutCont
    collOutputDicts.Add dictCsxUpdatedDwell

    Call fWriteDictionariesToWS(outputws, collOutputDicts)          'Write all data to output WS
    wbkcsxObj.Worksheets("RuntimeBuildCsxDict").Range("A2:A" & fLastWrittenRow(wbkcsxObj.Worksheets("RuntimeBuildCsxDict"), 1)).ClearContents

    For currrow = 2 To collRuntimes.Count
        wbkcsxObj.Worksheets("RuntimeBuildCsxDict").Cells(currrow, 1).Value2 = collRuntimes(currrow)
    Next currrow
    'LocationMapping
    fWriteLocationMapping outputws, dictCsxUpdatedLastLocation
'    fWriteLocationMapping outputws, dictCsxUpdatedOutCont


End Sub
