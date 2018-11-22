Option Explicit

Sub CopycsxAndTimestamp()

    Dim importWS As Worksheet
    Dim lastrow As Long
    Dim currentrow As Long

    Dim ImportWbk As Workbook
    Dim counter As Integer
    Dim csxWbk As Workbook

    Dim csxData As Variant
    Dim csxDataFiltered As Variant
    Dim OuterScannableData As Variant
    Dim OuterScannableDataFiltered As Variant
    Dim OuterContainerData As Variant
    Dim WorkpoolData As Variant
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
    Set csxWbk = Workbooks.Open(FileName:=strcsxStampsFile, UpdateLinks:=False)

    Set ImportWbk = Workbooks(strRodeoHistoryFileName)

    For Each importWS In ImportWbk.Worksheets
        For counter = 1 To ImportWbk.Worksheets.Count
            timeStampscsx = Right(importWS.Name, 9)
            lastrow = fLastWrittenRow(importWS, 1)

            csxData = importWS.Range("I1:I" & lastrow).Value2
            OuterScannableData = importWS.Range("J1:J" & lastrow).Value2
            OuterContainerData = importWS.Range("K1:K" & lastrow).Value2
            WorkpoolData = importWS.Range("O1:O" & lastrow).Value2
            csxDataFiltered = csxData
            OuterScannableDataFiltered = OuterScannableData

            ReDim csxDataFiltered(0)
            ReDim OuterScannableDataFiltered(0)

        If WorkpoolData(counter, 1) <> "Palletized" Or WorkpoolData(counter, 1) <> "Loaded" Or WorkpoolData(counter, 1) <> "TransshipSorted" Then
            ReDim Preserve csxDataFiltered(0 To counter)
            ReDim Preserve OuterScannableDataFiltered(0 To counter)
            csxDataFiltered(counter) = csxData(counter)
            OuterScannableDataFiltered(counter) = OuterScannableData(counter)
        End If
            csxWbk.Worksheets(counter).Range("A1:A" & lastrow).Value2 = csxDataFiltered
            csxWbk.Worksheets(counter).Range("B1:B" & lastrow).Value2 = OuterScannableDataFiltered
            csxWbk.Worksheets(counter).Range("C1:C" & lastrow).Value2 = OuterContainerData
            csxWbk.Worksheets(counter).Range("D1:D" & lastrow).Value2 = WorkpoolData
            csxWbk.Worksheets(counter).Range("e1:e" & lastrow).Value2 = timeStampscsx

            Erase csxData
            Erase OuterScannableData
            Erase OuterContainerData
            Erase WorkpoolData
'            Erase timeStampscsx
        Next counter
    Next importWS
End Sub
