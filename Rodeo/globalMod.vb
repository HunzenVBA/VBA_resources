Option Explicit

Sub RepeatImport()

    Dim lAnzahl As String
    Dim i As Long

Anf:
    lAnzahl = InputBox("Wie oft soll das Makro laufen ?", , 3)

    If lAnzahl = "" Then Exit Sub

    'Prüfen ob eine Zahl eingegeben wurde
    If IsNumeric(lAnzahl) Then
        For i = 1 To CLng(lAnzahl)
                StartTimeAll = Timer
                currProcedureName = "RodeoAddQueryTotal"
                Debug.Print "============================ Beginn Sub " & currProcedureName & " ============================"
                Debug.Print Now
                Call RodeoAddQueryTotal
                'Determine how many seconds code took to run
                SecondsElapsedAll = Round(Timer - StartTimeAll, 1)
                'Notify user in seconds
                Debug.Print "This code ran successfully in " & SecondsElapsedAll & " seconds"
                Debug.Print "Makro Start Nr.: " & i
        Next i
    Else
    MsgBox "Bitte ein Zahl eingeben !", vbInformation
    GoTo Anf
    End If

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
                currProcedureName = "RodeoAddQueryTotal"
                Debug.Print "============================ Beginn Sub " & currProcedureName & " ============================"
                Debug.Print Now
                Call RodeoAddWorkpool
                'Determine how many seconds code took to run
                SecondsElapsedAll = Round(Timer - StartTimeAll, 1)
                'Notify user in seconds
                Debug.Print "This code ran successfully in " & SecondsElapsedAll & " seconds"
                Debug.Print "Makro Start Nr.: " & i
                Application.Wait (10)       '10 seconds delay between queries
        Next i
    Else
    MsgBox "Bitte ein Zahl eingeben !", vbInformation
    GoTo Anf
    End If

End Sub
