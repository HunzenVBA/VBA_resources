'check if files are open
On Error Resume Next
Set ImportWbk = Workbooks(strRodeoHistoryFileName)
Set csxWbk = Workbooks(strcsxStampsFileName)
On Error GoTo 0

If ImportWbk Is Nothing Then
    Set ImportWbk = Workbooks.Open(FileName:=strRodeoHistoryFile, UpdateLinks:=False)
Else
    ImportWbk.Close SaveChanges:=False
End If
If csxWbk Is Nothing Then
    Set csxWbk = Workbooks.Open(FileName:=strcsxStampsFile, UpdateLinks:=False)
Else
    csxWbk.Close SaveChanges:=False
End If

Set ImportWbk = Workbooks.Open(FileName:=strRodeoHistoryFile, UpdateLinks:=False)
