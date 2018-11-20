Option Explicit

Sub CopycsxAndTimestamp()

    Dim importWS As Worksheet
    Dim timeStampscsx As String
    Dim csxID As String
    Dim lastrow As Long
    Dim currentrow As Long

    Dim ImportWbk As Workbook
    Dim counter As String
    Dim csxWbk As Workbook

    Dim csxData As Variant
    Dim OuterScannableData As Variant
    Dim OuterContainerData As Variant
    Dim WorkpoolData As Variant

    'check if files are open
    If IsWorkBookOpen(strcsxStampsFile) Then
    Workbooks(strcsxStampsFileName).Close
    If IsWorkBookOpen(strRodeoHistoryFile) Then
    Workbooks(strRodeoHistoryFile).Close

    Set csxWbk = Workbooks.Open(FileName:=strcsxStampsFile, UpdateLinks:=False)
    Set ImportWbk = Workbooks.Open(FileName:=strRodeoHistoryFile, UpdateLinks:=False)


'
'    For Each importWS In ImportWbk.Worksheets
'        ImportWbk.Activate
'        Debug.Print importWS.Name
'        Debug.Print ImportWbk.Name
'        Debug.Print csxWbk.Name
'        lastrow = fLastWrittenRow(importWS, 1)
'        Debug.Print lastrow
'        importWS.Range("Z3").Value2 = "Hello"
'    Next importWS
'

    For Each importWS In ImportWbk.Worksheets
        timeStampscsx = Right(importWS.Name, 9)
        lastrow = fLastWrittenRow(importWS, 1)

        csxData = importWS.Range("I1:I" & lastrow).Value2
        OuterScannableData = importWS.Range("J1:J" & lastrow).Value2
        OuterContainerData = importWS.Range("K1:K" & lastrow).Value2
        WorkpoolData = importWS.Range("O1:O" & lastrow).Value2

        csxWbk.Worksheets(1).Range("A1:A" & lastrow).Value2 = csxData
        csxWbk.Worksheets(1).Range("B1:B" & lastrow).Value2 = OuterScannableData
        csxWbk.Worksheets(1).Range("C1:C" & lastrow).Value2 = OuterContainerData
        csxWbk.Worksheets(1).Range("D1:D" & lastrow).Value2 = WorkpoolData

    Next importWS


End Sub
