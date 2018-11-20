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

    Dim csxData As Range
    Dim OuterScannableData As Range
    Dim OuterContainerData As Range
    Dim WorkpoolData As Range


    Set ImportWbk = Workbooks.Open(Filename:=strRodeoHistoryFile, UpdateLinks:=False)
    Set csxWbk = Workbooks.Open(Filename:=strcsxStampsFile, UpdateLinks:=False)

    For Each importWS In ImportWbk.Worksheets
        timeStampscsx = Right(importWS.Name, 9)
        lastrow = fLastWrittenRow(importWS, 1)

        Set csxData = ImportWbk.importWS.Range("I1:I" & lastrow)
        Set OuterScannableData = ImportWbk.importWS.Range("")
        Set OuterContainerData = ImportWbk.importWS.Range("")
        Set WorkpoolData = ImportWbk.importWS.Range("")

        For currentrow = 2 To lastrow
                importWS.Range("I1:I" & lastrow).Copy
        Next currentrow
    Next importWS


End Sub
