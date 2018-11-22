Option Explicit

Sub CopycsxAndTimestamp()

Application.ScreenUpdating = False
    Dim importWS As Worksheet
    Dim lastrow As Long
    Dim currentrow As Long

    Dim ImportWbk As Workbook
    Dim counter As Long
    Dim csxWbk As Workbook

    Dim csxData As Variant
    Dim OuterScannableData As Variant
    Dim OuterContainerData As Variant
    Dim WorkpoolData As Variant
    Dim timeStampscsx As Variant

    Dim arDataset As Variant
    Dim dicDataset As Variant

    Dim lRow As Long
    Dim lCol As Long

    Dim dicRow As Long

    Dim csxDict As Dictionary

    Set csxDict = New Dictionary
    csxDict.RemoveAll
'    'check if files are open
    On Error Resume Next
    Set ImportWbk = Workbooks(strRodeoPath & "RodeoOhneLoadedundPalletized.xlsx")
    Set csxWbk = Workbooks(strcsxStampsFileName)
    On Error GoTo 0

    If ImportWbk Is Nothing Then
        Set ImportWbk = Workbooks.Open(FileName:=strRodeoPath & "RodeoOhneLoadedundPalletized.xlsx", UpdateLinks:=False)
    Else
        ImportWbk.Close SaveChanges:=False
    End If
    If csxWbk Is Nothing Then
        Set csxWbk = Workbooks.Open(FileName:=strcsxStampsFile, UpdateLinks:=False)
    Else
        csxWbk.Close SaveChanges:=False
    End If

    Set ImportWbk = Workbooks.Open(FileName:=strRodeoPath & "RodeoOhneLoadedundPalletized.xlsx", UpdateLinks:=False)
    Set csxWbk = Workbooks.Open(FileName:=strcsxStampsFile, UpdateLinks:=False)

    For Each importWS In ImportWbk.Worksheets
            timeStampscsx = Right(importWS.Name, 9)
            lastrow = fLastWrittenRow(importWS, 1)

            csxData = importWS.Range("h1:h" & lastrow).Value2
            OuterScannableData = importWS.Range("i1:i" & lastrow).Value2
            OuterContainerData = importWS.Range("K1:K" & lastrow).Value2
            WorkpoolData = importWS.Range("O1:O" & lastrow).Value2

'            importWS.Range("A1:A" & lastrow).Value2 = csxData
'            csxWbk.Worksheets(counter).Range("B1:B" & lastrow).Value2 = OuterScannableData
'            csxWbk.Worksheets(counter).Range("C1:C" & lastrow).Value2 = OuterContainerData
'            csxWbk.Worksheets(counter).Range("D1:D" & lastrow).Value2 = WorkpoolData
'            csxWbk.Worksheets(counter).Range("e1:e" & lastrow).Value2 = timeStampscsx

            'If csx doesnt exist in range, then add to dict
            For lRow = LBound(csxData, 1) To UBound(csxData, 1)
                If Not csxDict.Exists(csxData(lRow, 1)) Then
                    arDataset = csxData(lRow, 1)
                    Debug.Print csxData(lRow, 1) & "  " & timeStampscsx
                    csxDict.Add csxData(lRow, 1), OuterScannableData(lRow, 1)
                    dicDataset = csxDict(csxData(lRow, 1))
                Else
                End If
            Next lRow

            Debug.Print "Dictionary unique csx count: " & csxDict.Count

            Erase csxData
            Erase OuterScannableData
            Erase OuterContainerData
            Erase WorkpoolData
'            Erase timeStampscsx
    Next importWS


    With csxDict
        csxWbk.Worksheets(3).Cells(2, 1).Resize(.Count, 1) = Application.Transpose(.Keys)
        csxWbk.Worksheets(3).Cells(2, 2).Resize(.Count, 1) = Application.Transpose(.Items)
    End With

    csxWbk.Worksheets(3).Columns.AutoFit


End Sub
