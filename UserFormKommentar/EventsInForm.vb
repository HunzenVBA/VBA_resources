Option Explicit


Sub bxSchlag1_Change()
'Call Entsichern
Dim arrA
Dim arrB
Dim arrC
Dim arrD
Dim counterA As Integer
Dim counterB As Integer
Dim counterC As Integer
Dim counterD As Integer
Dim intZeileWS2A
Dim intZeileWS2B
Dim intZeileWS2C
Dim intZeileWS2D
Dim intletzteZeileWS2A
Dim intletzteZeileWS2B
Dim intletzteZeileWS2C
Dim intletzteZeileWS2D
intletzteZeileWS2A = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
intletzteZeileWS2B = Worksheets(2).Cells(Rows.Count, 2).End(xlUp).Row
intletzteZeileWS2C = Worksheets(2).Cells(Rows.Count, 3).End(xlUp).Row
intletzteZeileWS2D = Worksheets(2).Cells(Rows.Count, 4).End(xlUp).Row
For intZeileWS2C = 2 To intletzteZeileWS2C
    If Worksheets(2).Cells(intZeileWS2C, 2).Value <> "" Then intZeileWS2B = intZeileWS2C
    If Worksheets(2).Cells(intZeileWS2B, 2).Value = UserForm3.bxSchlag1.Value _
    And Worksheets(2).Cells(intZeileWS2C, 3).Value <> "" Then _
    counterC = counterC + 1
Next intZeileWS2C
ReDim arrC(counterC)
'test
intZeileWS2C = intletzteZeileWS2C
counterC = 1
For intZeileWS2C = 2 To intletzteZeileWS2C
    If Worksheets(2).Cells(intZeileWS2C, 2).Value <> "" Then intZeileWS2B = intZeileWS2C
    If Worksheets(2).Cells(intZeileWS2B, 2).Value = UserForm3.bxSchlag1.Value _
    And Worksheets(2).Cells(intZeileWS2C, 3).Value <> "" Then
    arrC(counterC) = Worksheets(2).Cells(intZeileWS2C, 3).Value
    counterC = counterC + 1
    End If
Next intZeileWS2C
UserForm3.bxSchlag2.List = arrC
Call CheckIfOk
End Sub
'Sub bxSchlag2_Change()()
'Call Worksheets(1).Entsichern
'UserForm3.txtOk.Value= ""
'If UserForm3.bxSchlag2.Value <> "" And UserForm3.txtLogin.Value <> "" _
'Then UserForm3.txtOk.Value= "OK"

Sub bxSchlag2_Change()
'Call Entsichern
If UserForm3.bxSchlag1.Value <> "" _
And UserForm3.bxSchlag2.Value <> "" _
And UserForm3.bxDepartment.Value <> "" _
And UserForm3.txtLogin.Value <> "" Then
UserForm3.txtOk.Value = "OK"
ElseIf UserForm3.bxSchlag1.Value <> "" _
And UserForm3.bxDepartment.Value <> "" _
And UserForm3.txtLogin.Value <> "" _
And UserForm3.bxSchlag2.ListCount <= 1 Then
UserForm3.txtOk.Value = "OK"
End If
Call CheckIfOk
End Sub

Sub bxDepartment_Change()
'Call Entsichern
Dim arrA
Dim arrB
Dim arrC
Dim counterA As Integer
Dim counterB As Integer
Dim counterC As Integer
Dim intZeileWS2A
Dim intZeileWS2B
Dim intZeileWS2C
Dim intletzteZeileWS2A
Dim intletzteZeileWS2B
Dim intletzteZeileWS2C
intletzteZeileWS2A = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
intletzteZeileWS2B = Worksheets(2).Cells(Rows.Count, 2).End(xlUp).Row
intletzteZeileWS2C = Worksheets(2).Cells(Rows.Count, 3).End(xlUp).Row
For intZeileWS2B = 2 To intletzteZeileWS2B
    If Worksheets(2).Cells(intZeileWS2B, 1).Value <> "" Then intZeileWS2A = intZeileWS2B
    If Worksheets(2).Cells(intZeileWS2A, 1).Value = UserForm3.bxDepartment.Value _
    And Worksheets(2).Cells(intZeileWS2B, 2).Value <> "" Then _
    counterB = counterB + 1
Next intZeileWS2B
ReDim arrB(counterB)
counterB = 1
For intZeileWS2B = 2 To intletzteZeileWS2B
    If Worksheets(2).Cells(intZeileWS2B, 1).Value <> "" Then intZeileWS2A = intZeileWS2B
    If Worksheets(2).Cells(intZeileWS2A, 1).Value = UserForm3.bxDepartment.Value _
    And Worksheets(2).Cells(intZeileWS2B, 2).Value <> "" Then
    arrB(counterB) = Worksheets(2).Cells(intZeileWS2B, 2).Value
    counterB = counterB + 1
    End If
Next intZeileWS2B
UserForm3.bxSchlag1.List = arrB
Call CheckIfOk
End Sub

'Sub bxDepartment_Change()()
'If ComboBox3.Value = "DTM2.IB" Then Worksheets(1).Cells(2, 1).Value = "DTM2.IB"
'If ComboBox3.Value = "DTM2.OB" Then Worksheets(1).Cells(2, 1).Value = "DTM2.OB"
'Call UserForm3.bxSchlag1_füllen
'End Sub

Sub btnComment_Click()
'Call Entsichern
If UserForm3.txtLogin.Value = "" Then UserForm3.txtOk.Value = "no"
If UserForm3.txtLogin.Value <> "" Then UserForm3.txtOk.Value = "OK"
If UserForm3.txtOk.Value = "" Then
MsgBox "Falsche Einstellungen vorgenommen."
ElseIf UserForm3.txtOk.Value = "OK" Then
Dim addedString As String
If UserForm3.bxSchlag1.Value = "DA" Then
addedString = " - " & FindSchlagwort4Code(UserForm3.bxSchlag1.Value)
ElseIf UserForm3.bxSchlag1.Value = "EmptyTotes" Or UserForm3.bxSchlag1.Value = "EmptyPallets" Or UserForm3.bxSchlag1.Value = "PalletSleeves" Then
addedString = " - " & FindSchlagwort4Code(UserForm3.bxSchlag1.Value)
Else
addedString = ""
End If


UserForm3.txtComment.Value = (UserForm3.bxSchlag1.Value & UserForm3.bxSchlag2.Value & " - " & UserForm3.txtDate.Value _
& " - " & UserForm3.txtLogin.Value & addedString)
UserForm3.txtComment.SetFocus
UserForm3.txtComment.SelStart = 0
UserForm3.txtComment.SelLength = Len(UserForm3.txtComment.Value)
'UserForm1.Show
End If
With Worksheets(1)
.Cells(4, 4).Value = ""
.Cells(4, 5).Value = ""
.ComboBox1.Clear
.ComboBox2.Clear
.ComboBox3.Clear
.ComboBox1.Value = ""
.ComboBox2.Value = ""
.ComboBox3.Value = ""
End With
Call Starten
'Call Sichern
End Sub

Private Sub UserForm_Initialize()
Call Starten
End Sub
