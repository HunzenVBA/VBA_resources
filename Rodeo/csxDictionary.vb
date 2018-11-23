Option Explicit

Sub CopycsxAndTimestamp()

    Dim importWS As Worksheet
    Dim lastrow As Long
    Dim currentrow As Long
    Dim lrow As Long
    Dim dataSetWorkpool As String
    Dim dataSetcsx As String
    Dim dataSetoutCont As String
    Dim dataSetoutScan As String
    Dim uniqueRow As Long


    Dim ImportWbk As Workbook
    Dim counter As Long
    Dim csxWbk As Workbook

    Dim slice
    Dim app As Application

'    Dim csxData(1 To 1) As Variant
'    Dim csxDataFiltered(1 To 1) As Variant
'    Dim OuterScannableData(1 To 1) As Variant
'    Dim OuterScannableDataFiltered(1 To 1) As Variant
'    Dim OuterContainerData(1 To 1) As Variant
'    Dim WorkpoolData(1 To 1) As Variant
    Dim testarray As Variant
    Dim timeStampscsx As Variant
    Dim csxDict As Dictionary
    Dim outScanDict As Dictionary
    Dim outContDict As Dictionary

    Set csxDict = New Dictionary
    Set outScanDict = New Dictionary
    Set outContDict = New Dictionary


'    'check if files are open
'    On Error Resume Next
'    Set ImportWbk = Workbooks(strRodeoHistoryFileName)
'    Set csxWbk = Workbooks(strcsxStampsFileName)
'    On Error GoTo 0

'    If ImportWbk Is Nothing Then
'        Set ImportWbk = Workbooks.Open(FileName:=strRodeoHistoryFile, UpdateLinks:=False)
'    Else
'        ImportWbk.Close SaveChanges:=False
'    End If
'    If csxWbk Is Nothing Then
'        Set csxWbk = Workbooks.Open(FileName:=strcsxStampsFile, UpdateLinks:=False)
'    Else
'        csxWbk.Close SaveChanges:=False
'    End If

'    Set ImportWbk = Workbooks.Open(FileName:=strRodeoHistoryFile, UpdateLinks:=False)
    Set app = Application
    Set csxWbk = Workbooks.Open(FileName:=strcsxStampsFile, UpdateLinks:=False)
    Set ImportWbk = Workbooks(strRodeoHistoryFileName)
    For Each importWS In ImportWbk.Worksheets
        timeStampscsx = Right(importWS.Name, 9)
        lastrow = fLastWrittenRow(importWS, 1)

        ReDim csxData(1 To lastrow, 1)
        ReDim OuterScannableData(1 To lastrow, 1)
        ReDim OuterContainerData(1 To lastrow, 1)
        ReDim WorkpoolData(1 To lastrow, 1)
        ReDim csxDataFiltered(1 To lastrow, 1)
        ReDim OuterScannableDataFiltered(1 To lastrow, 1)

        ReDim csxDataUniqe(1 To lastrow, 1)
        ReDim OuterContainerDataUnique(1 To lastrow, 1)
        ReDim OuterScannableDataUnique(1 To lastrow, 1)

        'Fill arrays with values'
        csxData = importWS.Range("I1:I" & lastrow).Value2
        OuterScannableData = importWS.Range("J1:J" & lastrow).Value2
        OuterContainerData = importWS.Range("K1:K" & lastrow).Value2
        WorkpoolData = importWS.Range("O1:O" & lastrow).Value2

        ReDim csxDataFiltered(1 To 1, 1 To 1)
        ReDim OuterScannableDataFiltered(1 To 1, 1 To 1)
        ReDim OuterContainerDataFiltered(1 To 1, 1 To 1)
        ReDim WorkpoolDataFiltered(1 To 1, 1 To 1)

        ReDim csxDataUniqe(1 To 1, 1 To 1)
        ReDim OuterScannableDataUnique(1 To 1, 1 To 1)
        ReDim OuterContainerDataUnique(1 To 1, 1 To 1)
        counter = 0
            For lrow = 1 To lastrow
                    'watch-variables
                    dataSetWorkpool = WorkpoolData(lrow, 1)
                    dataSetcsx = csxData(lrow, 1)
                    dataSetoutScan = OuterContainerData(lrow, 1)
                    dataSetoutCont = OuterContainerData(lrow, 1)
                If WorkpoolData(lrow, 1) <> "Palletized" And WorkpoolData(lrow, 1) <> "Loaded" And WorkpoolData(lrow, 1) <> "TransshipSorted" Then
                    'Resize arrays by value=counter on each hit of conditions
                    counter = counter + 1
                    ReDim Preserve csxDataFiltered(1 To 1, 1 To counter)
                    ReDim Preserve OuterScannableDataFiltered(1 To 1, 1 To counter)
                    ReDim Preserve OuterContainerDataFiltered(1 To 1, 1 To counter)
                    ReDim Preserve WorkpoolDataFiltered(1 To 1, 1 To counter)

                    'Fill new row of array with value on hit conditions
                    csxDataFiltered(1, counter) = csxData(lrow, 1)
                    OuterScannableDataFiltered(1, counter) = OuterScannableData(lrow, 1)
                    OuterContainerDataFiltered(1, counter) = OuterContainerData(lrow, 1)
                    WorkpoolDataFiltered(1, counter) = WorkpoolData(lrow, 1)
                End If
            Next lrow
            counter = 0
            For uniqueRow = LBound(csxDataFiltered, 2) To UBound(csxDataFiltered, 2)
                If Not csxDict.Exists(csxDataFiltered(1, uniqueRow)) Then
                    counter = counter + 1
                    dataSetcsx = csxDataFiltered(1, counter)
                    dataSetoutScan = OuterScannableDataFiltered(1, counter)
                    dataSetoutCont = OuterContainerDataFiltered(1, counter)
                    Debug.Print dataSetcsx
                    csxDict.Add csxDataFiltered(1, uniqueRow), OuterScannableDataFiltered(1, uniqueRow)
                    outScanDict.Add csxDataFiltered(1, uniqueRow), OuterScannableDataFiltered(1, uniqueRow)
                    outContDict.Add csxDataFiltered(1, uniqueRow), OuterContainerDataFiltered(1, uniqueRow)

                        ReDim Preserve OuterContainerDataUnique(1 To 1, counter)
                        ReDim Preserve csxDataUniqe(1 To 1, counter)

                    OuterContainerDataUnique(1, counter) = OuterContainerDataFiltered(uniqueRow, 1)
                    csxDataUniqe(1, counter) = csxDataFiltered(1, uniqueRow)
                End If
            Next uniqueRow

            With csxDict
                csxWbk.Worksheets("FilteredUnique").Cells.Clear
                csxWbk.Worksheets("FilteredUnique").Cells(1, 1).Resize(.Count, 1) = Application.Transpose(.Keys)
                csxWbk.Worksheets("FilteredUnique").Cells(1, 2).Resize(.Count, 1) = Application.Transpose(.Items)
            End With

        Erase csxDataFiltered
        Erase OuterScannableDataFiltered
        Erase OuterContainerDataFiltered
        Erase WorkpoolDataFiltered
'            Erase timeStampscsx
    Next importWS
End Sub
