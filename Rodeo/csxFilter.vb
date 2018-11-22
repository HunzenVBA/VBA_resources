Option Explicit

Sub CopycsxAndTimestamp()

    Dim importWS As Worksheet
    Dim lastrow As Long
    Dim currentrow As Long
    Dim lrow As Long
    Dim dataSetWorkpool As String
    Dim dataSetcsx As String
    Dim wscounter As Integer


    Dim ImportWbk As Workbook
    Dim counter As Integer
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
        ReDim csxDataFiltered(1 To lastrow, 1)
        ReDim OuterScannableData(1 To lastrow, 1)
        ReDim OuterScannableDataFiltered(1 To lastrow, 1)
        ReDim OuterContainerData(1 To lastrow, 1)
        ReDim WorkpoolData(1 To lastrow, 1)


        csxData = importWS.Range("I1:I" & lastrow).Value2
        OuterScannableData = importWS.Range("J1:J" & lastrow).Value2
        OuterContainerData = importWS.Range("K1:K" & lastrow).Value2
        WorkpoolData = importWS.Range("O1:O" & lastrow).Value2


        ReDim csxDataFiltered(1 To 1, 1 To 1)
        ReDim OuterScannableDataFiltered(1 To 1, 1 To 1)
        ReDim OuterContainerDataFiltered(1 To 1, 1 To 1)
        ReDim WorkpoolDataFiltered(1 To 1, 1 To 1)
            For counter = 1 To lastrow
                    'watch-variables
                    dataSetWorkpool = WorkpoolData(counter, 1)
                    dataSetcsx = csxData(counter, 1)
                If WorkpoolData(counter, 1) <> "Palletized" And WorkpoolData(counter, 1) <> "Loaded" And WorkpoolData(counter, 1) <> "TransshipSorted" Then
                    'Resize arrays by value=counter on each hit of conditions
                    ReDim Preserve csxDataFiltered(1 To 1, 1 To counter)
                    ReDim Preserve OuterScannableDataFiltered(1 To 1, 1 To counter)
                    ReDim Preserve OuterContainerDataFiltered(1 To 1, 1 To counter)
                    ReDim Preserve WorkpoolDataFiltered(1 To 1, 1 To counter)
                    
                    'Fill new row of array with value on hit conditions

                    csxDataFiltered(1, counter) = csxData(counter, 1)
                    OuterScannableDataFiltered(1, counter) = OuterScannableData(counter, 1)
                    OuterContainerDataFiltered(1, counter) = OuterContainerData(counter, 1)
                    WorkpoolDataFiltered(1, counter) = WorkpoolData(counter, 1)
                End If
            Next wscounter
        csxWbk.Worksheets(4).Range("A1:A" & counter).Value2 = app.Transpose(csxDataFiltered)
        csxWbk.Worksheets(4).Range("B1:B" & counter).Value2 = app.Transpose(OuterScannableDataFiltered)
        csxWbk.Worksheets(4).Range("C1:C" & lastrow).Value2 = app.Transpose(OuterContainerDataFiltered)
        csxWbk.Worksheets(4).Range("D1:D" & lastrow).Value2 = app.Transpose(WorkpoolDataFiltered)
        csxWbk.Worksheets(4).Range("e1:e" & lastrow).Value2 = timeStampscsx

        Erase csxDataFiltered
        Erase OuterScannableDataFiltered
        Erase OuterContainerDataFiltered
        Erase WorkpoolDataFiltered
'            Erase timeStampscsx
    Next importWS
End Sub
