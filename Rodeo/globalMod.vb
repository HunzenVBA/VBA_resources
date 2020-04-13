Option Explicit

Sub runStacker()

    Dim StartTimeAll        As Double
    Dim SecondsElapsedAll   As Double

    StartTimeAll = Timer

'    Namensausgabe des Subs im Direktfenster, zur Info/Debugging
    currProcedureName = "NameWorksheets"
    Debug.Print "============================ Beginn Sub " & currProcedureName & " ============================"
    Debug.Print Now
    Call NameWorksheets

    currProcedureName = "PrintWSnames"
    Debug.Print "============================ Beginn Sub " & currProcedureName & " ============================"
    Debug.Print Now
    Call PrintWSnames

    currProcedureName = "UpdateRodeoTotal"
    Debug.Print "============================ Beginn Sub " & currProcedureName & " ============================"
    Debug.Print Now
    Call UpdateRodeoTotal

    currProcedureName = "DeleteEmptyRows"
    Debug.Print "============================ Beginn Sub " & currProcedureName & " ============================"
    Debug.Print Now
    Call DeleteEmptyRows

'    currProcedureName = "DeleteEmptyRows"
'    Debug.Print "============================ Beginn Sub " & currProcedureName & " ============================"
'    Debug.Print Now
'    Call DeleteEmptyRows



'Determine how many seconds code took to run
  SecondsElapsedAll = Round(Timer - StartTimeAll, 1)

'Notify user in seconds
  Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds"
End Sub

Sub RepeatImportTotal()

    Dim lAnzahl As String
    Dim i As Long
    Dim tStop As Date
    tStop = TimeValue(Now)

    'Prüfen ob eine Zahl eingegeben wurde
    Do While tStop < TimeValue("22:00:00")
    tStop = TimeValue(Now)
        i = i + 1
                StartTimeAll = Timer
                currProcedureName = "Rodeo10minDwell"
                Debug.Print "============================ Beginn Sub " & currProcedureName & " ============================"
                Debug.Print Now
                Call Rodeo10minDwell
                'Determine how many seconds code took to run
                SecondsElapsedAll = Round(Timer - StartTimeAll, 1)
                'Notify user in seconds
                'Debug.Print "This code ran successfully in " & SecondsElapsedAll & " seconds"
                Application.Wait (Now + TimeValue("0:01:00"))      '10 seconds delay between queries
                Debug.Print "Makro Start Nr.: " & i
    Loop


End Sub

Sub RepeatImportFiltered()

    Dim lAnzahl As String
    Dim i As Long

Anf:
    lAnzahl = InputBox("Wie oft soll das Makro laufen ?", , 3)

    If lAnzahl = "" Then Exit Sub

    'Prüfen ob eine Zahl eingegeben wurde
    If IsNumeric(lAnzahl) Then
        For i = 1 To CLng(lAnzahl)
                StartTimeAll = Timer
                currProcedureName = "RodeoAddWorkpool"
                Debug.Print "============================ Beginn Sub " & currProcedureName & " ============================"
                Debug.Print Now
                Call RodeoAddWorkpool
                'Determine how many seconds code took to run
                SecondsElapsedAll = Round(Timer - StartTimeAll, 1)
                'Notify user in seconds
                'Debug.Print "This code ran successfully in " & SecondsElapsedAll & " seconds"
                Debug.Print "Makro Start Nr.: " & i
                Application.Wait (Now + TimeValue("0:01:00"))      '10 seconds delay between queries
        Next i
    Else
    MsgBox "Bitte ein Zahl eingeben !", vbInformation
    GoTo Anf
    End If

End Sub

Sub RepeatImport1min()

    Dim lAnzahl As String
    Dim i As Long

Anf:
    lAnzahl = InputBox("Wie oft soll das Makro laufen ?", , 3)

    If lAnzahl = "" Then Exit Sub

    'Prüfen ob eine Zahl eingegeben wurde
    If IsNumeric(lAnzahl) Then
        For i = 1 To CLng(lAnzahl)
                StartTimeAll = Timer
                currProcedureName = "Rodeo1minDwell"
                Debug.Print "============================ Beginn Sub " & currProcedureName & " ============================"
                Debug.Print Now
                Call Rodeo1minDwell
                'Determine how many seconds code took to run
                SecondsElapsedAll = Round(Timer - StartTimeAll, 1)
                'Notify user in seconds
                'Debug.Print "This code ran successfully in " & SecondsElapsedAll & " seconds"
                Debug.Print "Makro Start Nr.: " & i
                Application.Wait (Now + TimeValue("0:00:08"))      '10 seconds delay between queries
        Next i
    Else
    MsgBox "Bitte ein Zahl eingeben !", vbInformation
    GoTo Anf
    End If

End Sub
