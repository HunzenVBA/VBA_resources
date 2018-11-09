Sub Abgleichen()
    Dim bolamTor As Boolean
    Dim intamTor As Integer
    Dim intZeileWS1 As Integer
    Dim intZeileWS2 As Integer
    Dim intletzteZeileWS1 As Integer
    Dim intletzteZeileWS2 As Integer
    Dim bolVergeben As Boolean
    
    intletzteZeileWS1 = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
    intletzteZeileWS2 = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
    For intZeileWS2 = 2 To intletzteZeileWS2
    Worksheets(2).Cells(intZeileWS2, 1).Interior.Color = vbWhite
    Next intZeileWS2
    For intZeileWS1 = 2 To intletzteZeileWS1
        intletzteZeileWS2 = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
        bolVergeben = False
        For intZeileWS2 = 2 To intletzteZeileWS2                                'loop over rows WS2
            If Worksheets(2).Cells(intZeileWS2, 5).Value = Worksheets(1).Cells(intZeileWS1, 5).Value Then
                bolVergeben = True
                Worksheets(2).Cells(intZeileWS2, 7).Value = Worksheets(1).Cells(intZeileWS1, 8).Value
                Worksheets(2).Cells(intZeileWS2, 2).Value = Worksheets(1).Cells(intZeileWS1, 3).Value
                Worksheets(2).Cells(intZeileWS2, 5).Value = Worksheets(1).Cells(intZeileWS1, 5).Value
                If Worksheets(2).Cells(intZeileWS2, 1).Value <> "Empty" Then
                    Worksheets(2).Cells(intZeileWS2, 1).Value = Worksheets(1).Cells(intZeileWS1, 12).Value
                End If
                If Worksheets(2).Cells(intZeileWS2, 1).Value = "FULL" Then
                    Worksheets(2).Cells(intZeileWS2, 1).Value = "Offen"
                End If
                'Worksheets(2).Cells(intZeileWS2, 3).Value = Format(Worksheets(1).Cells(intZeileWS1, 19).Value, "hh:mm")
                'Worksheets(2).Cells(intZeileWS2, 4).Value = Format(Worksheets(1).Cells(intZeileWS1, 20).Value, "dd.MM.yyyy")
                Worksheets(2).Cells(intZeileWS2, 8).Value = Worksheets(1).Cells(intZeileWS1, 16).Value
                Worksheets(2).Cells(intZeileWS2, 13).Value = ""
                If Weekday(Worksheets(2).Cells(intZeileWS2, 4).Value, 2) = 1 Then
                    Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 46
                ElseIf Weekday(Worksheets(2).Cells(intZeileWS2, 4).Value, 2) = 2 Then
                    Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 26
                ElseIf Weekday(Worksheets(2).Cells(intZeileWS2, 4).Value, 2) = 3 Then
                    Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 4
                ElseIf Weekday(Worksheets(2).Cells(intZeileWS2, 4).Value, 2) = 4 Then
                    Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 33
                ElseIf Weekday(Worksheets(2).Cells(intZeileWS2, 4).Value, 2) = 5 Then
                    Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 6
                ElseIf Weekday(Worksheets(2).Cells(intZeileWS2, 4).Value, 2) = 6 Then
                    Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 13
                Else:
                    Worksheets(2).Cells(intZeileWS2, 4).Interior.Color = vbWhite
                End If
                If Worksheets(2).Cells(intZeileWS2, 4).Value = "" Then
                    Worksheets(2).Cells(intZeileWS2, 4).Interior.Color = vbWhite
                End If
                If Worksheets(2).Cells(intZeileWS2, 2).Value = "MENDEN" _
                Or Worksheets(2).Cells(intZeileWS2, 2).Value = "K&N" Then
                    Worksheets(2).Cells(intZeileWS2, 2).Interior.Color = vbBlue
                End If
            End If
        Next intZeileWS2
        If bolVergeben = False Then
                Worksheets(2).Cells(intZeileWS2, 13).Value = "NEW"
                Worksheets(2).Cells(intZeileWS2, 7).Value = Worksheets(1).Cells(intZeileWS1, 8).Value
                Worksheets(2).Cells(intZeileWS2, 2).Value = Worksheets(1).Cells(intZeileWS1, 3).Value
                Worksheets(2).Cells(intZeileWS2, 5).Value = Worksheets(1).Cells(intZeileWS1, 5).Value
                If Worksheets(2).Cells(intZeileWS2, 12).Value = "" Then
                    Worksheets(2).Cells(intZeileWS2, 1).Value = Worksheets(1).Cells(intZeileWS1, 12).Value
                End If
                If Worksheets(2).Cells(intZeileWS2, 1).Value = "FULL" Then
                    Worksheets(2).Cells(intZeileWS2, 1).Value = "Offen"
                End If
                Worksheets(2).Cells(intZeileWS2, 8).Value = Worksheets(1).Cells(intZeileWS1, 16).Value
                Worksheets(2).Cells(intZeileWS2, 3).Value = Format(Worksheets(1).Cells(intZeileWS1, 19).Value, "hh:mm")
                Worksheets(2).Cells(intZeileWS2, 4).Value = Format(Worksheets(1).Cells(intZeileWS1, 20).Value, "dd.MM.yyyy")
                Worksheets(2).Cells(intZeileWS2, 6).Value = 300
'                Worksheets(2).Cells(intZeileWS2, 31).FormulaR1C1 = "=VLOOKUP(ROUNDUP(RC[-28]*24,0),R2C24:R24C25,2,TRUE)"
                
                
            If Weekday(Worksheets(2).Cells(intZeileWS2, 4).Value, 2) = 1 Then
                Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 46
            ElseIf Weekday(Worksheets(2).Cells(intZeileWS2, 4).Value, 2) = 2 Then
                Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 26
            ElseIf Weekday(Worksheets(2).Cells(intZeileWS2, 4).Value, 2) = 3 Then
                Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 4
            ElseIf Weekday(Worksheets(2).Cells(intZeileWS2, 4).Value, 2) = 4 Then
                Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 33
            ElseIf Weekday(Worksheets(2).Cells(intZeileWS2, 4).Value, 2) = 5 Then
                Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 6
            ElseIf Weekday(Worksheets(2).Cells(intZeileWS2, 4).Value, 2) = 6 Then
                Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 13
            Else:
                Worksheets(2).Cells(intZeileWS2, 4).Interior.Color = vbWhite
            End If
            If Worksheets(2).Cells(intZeileWS2, 4).Value = "" _
            Then Worksheets(2).Cells(intZeileWS2, 4).Interior.Color = vbWhite
            If Worksheets(2).Cells(intZeileWS2, 2).Value = "MENDEN" _
            Or Worksheets(2).Cells(intZeileWS2, 2).Value = "K&N" _
            Then Worksheets(2).Cells(intZeileWS2, 2).Interior.Color = vbBlue
        End If
    Next intZeileWS1
    intletzteZeileWS2 = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
    For intZeileWS2 = intletzteZeileWS2 To 2 Step -1
        bolVergeben = False
        For intZeileWS1 = 2 To intletzteZeileWS1
            If Worksheets(2).Cells(intZeileWS2, 8).Value = Worksheets(1).Cells(intZeileWS1, 16).Value _
            And Worksheets(2).Cells(intZeileWS2, 5).Value = Worksheets(1).Cells(intZeileWS1, 5).Value _
            Then bolVergeben = True
        Next intZeileWS1
'        If bolVergeben = False _
'        Then Worksheets(2).Rows(intZeileWS2).Delete
    Next intZeileWS2
    For intZeileWS2 = 2 To intletzteZeileWS2
        intletzteZeileWS2 = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
        If Left(Worksheets(2).Cells(intZeileWS2, 7).Value, 3) = "IB0" _
        And Worksheets(2).Cells(intZeileWS2, 1).Value = "Offen" _
            Then Worksheets(2).Cells(intZeileWS2, 1).Value = "An Tor: " & CStr(CInt(Right(Worksheets(2).Cells(intZeileWS2, 7).Value, 2)))
    Next intZeileWS2
    bolamTor = False
    Call Sortieren
    For intZeileWS2 = intletzteZeileWS2 To 2 Step -1
    Worksheets(2).Cells(intZeileWS2, 1).Interior.Color = vbWhite
        If Left(Worksheets(2).Cells(intZeileWS2, 7).Value, 2) = "IB" And bolamTor = False Then
            bolamTor = True
            intamTor = intZeileWS2
        End If
    Next intZeileWS2
    For intZeileWS2 = 2 To intamTor
        If Worksheets(2).Cells(intZeileWS2, 1).Value = "Offen" _
        And (Left(Worksheets(2).Cells(intZeileWS2, 12).Value, 4) = "Mend" _
        Or Left(Worksheets(2).Cells(intZeileWS2, 12).Value, 4) = "Parc" _
        Or Left(Worksheets(2).Cells(intZeileWS2, 12).Value, 3) = "K&N" _
        Or Worksheets(2).Cells(intZeileWS2, 12).Value = "") _
        Then Worksheets(2).Cells(intZeileWS2, 1).Interior.Color = vbRed
    Next intZeileWS2                'Abschnitt hinzugefügt: es werden alle offenen mit Rot markiert
        For intZeileWS2 = 2 To intletzteZeileWS2
        If Worksheets(2).Cells(intZeileWS2, 1).Value = "Offen" _
        Then Worksheets(2).Cells(intZeileWS2, 1).Interior.Color = vbRed
    Next intZeileWS2
    'Do While Left(Worksheets(2).Cells(intZeileWS2, 7).Value, 2) <> "IB" And intZeileWS2 <= intletzteZeileWS2 And bolamTor = True
        'If Worksheets(2).Cells(intZeileWS2, 1).Value = "FULL" Then Worksheets(2).Cells(intZeileWS2, 1).Interior.Color = vbRed
        'intZeileWS2 = intZeileWS2 + 1
    'Loop
End Sub

Sub Sortieren()
    Worksheets(2).Activate
'    Columns("A:AF").Select
    ActiveWorkbook.Worksheets(2).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(2).Sort.SortFields.Add Key:=Range("D:D" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    ActiveWorkbook.Worksheets(2).Sort.SortFields.Add Key:=Range("C:C" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(2).Sort.SortFields.Add Key:=Range("G:G" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(2).Sort
        .SetRange Range("A2:M" & fLastWrittenRow(shtYardView, 1))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


Sub Torenachoben()
    Dim bolFinish As Boolean
    Dim intZeileWS2 As Integer
    Dim intletzteZeileWS2 As Integer
    Dim intletztesTor As Integer
    Dim intSchleifenzähler As Integer
    Dim intcounter As Integer
    Dim intmin As Integer
    Dim inzZeilemin As Integer
    intletzteZeileWS2 = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
    intletztesTor = 2
    intSchleifenzähler = 0
    intcounter = 0
    intmin = 200
    bolFinish = False
    Do While bolFinish = False And intSchleifenzähler < 200
        bolFinish = True
        For intZeileWS2 = intletztesTor To intletzteZeileWS2
            If Left(Worksheets(2).Cells(intZeileWS2, 7).Value, 3) = "IB0" Then
                Rows(intletztesTor & ":" & intletztesTor).Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Rows(intZeileWS2 + 1).Copy Destination:=Worksheets(2).Cells(intletztesTor, 1)
                Rows(intZeileWS2 + 1).Delete
                bolFinish = False
                intletztesTor = intletztesTor + 1
                intcounter = intcounter + 1
            End If
        Next intZeileWS2
        intSchleifenzähler = intSchleifenzähler + 1
    Loop
    'For intZeileWS2 = 2 To intcounter + 1
        'If CInt(Left(Worksheets(2).Cells(intZeileWS2, 7).Value, 3)) <= intmin Then
            'intmin = CInt(Left(Worksheets(2).Cells(intZeileWS2, 7).Value, 3))
            'Zeilemin = intZeileWS2
        'End If
    'Next intZeileWS2
    If intcounter > 0 Then
    Range("A1:M" & intcounter + 1).Select
        ActiveWorkbook.Worksheets(2).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(2).Sort.SortFields.Add Key:=Range( _
            "A2:A" & intcounter + 1), SortOn:=xlSortOnValues, Order:=xlDescending, CustomOrder:= _
            "An Tor: 94,An Tor: 93,An Tor: 92,An Tor: 91,An Tor: 90,An Tor: 89,An Tor: 88,An Tor: 87,An Tor: 86,An Tor: 85,An Tor: 84,An Tor: 83,An Tor: 82,An Tor: 81,An Tor: 80,An Tor: 79,An Tor: 78,An Tor: 77,An Tor: 76" _
            , DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Yardübersicht").Sort
            .SetRange Range("A1:M" & intcounter + 1)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
End Sub
