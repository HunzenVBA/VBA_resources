Sub Abgleichen()
    Dim bolamTor As Boolean
    Dim intamTor As Integer
    Dim intZeileWS1 As Integer
    Dim intZeileWS2 As Integer
    Dim intletzteZeileWS1 As Integer
    Dim intletzteZeileWS2 As Integer
    Dim bolVergeben As Boolean

	'variables WS2
	Set cStatusWS2 = Worksheets(2).Cells(intZeileWS2, 1 ).Value				'status = Offen, An Tor
	Set cCarrierWS2 = Worksheets(2).Cells(intZeileWS2, 2 ).Value				'Carriername
	Set cTimeWS2 = Worksheets(2).Cells(intZeileWS2, 3 ).Value				'Time of Arrival
	Set cDateWS2 = Worksheets(2).Cells(intZeileWS2, 4 ).Value				'Date of Arrival
	Set cDockNoWS2 = Worksheets(2).Cells(intZeileWS2, 5 ).Value				'Docknumber
	Set cLocationWS2 = Worksheets(2).Cells(intZeileWS2, 7 ).Value			'Location = IB Parcel, IB123, PSIBYARD
	Set cISAWS2 = Worksheets(2).Cells(intZeileWS2, 8 ).Value					'ISA

	'variables WS1
	Set cOperatorWS1 = Worksheets(1).Cells(intZeileWS1, 3 ).Value			'Carriername
	Set cOwnerIDWS1 = Worksheets(1).Cells(intZeileWS1, 5 ).Value			'Docknumber
	Set cLocationCodeWS1 = Worksheets(1).Cells(intZeileWS1, 8 ).Value		'LocationCode = IB123, PSIBYARD, IB Parcel


    intletzteZeileWS1 = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
    intletzteZeileWS2 = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
    For intZeileWS2 = 2 To intletzteZeileWS2								'Loop over all written rows in WS2
    Worksheets(2).Cells(intZeileWS2, 1).Interior.Color = vbWhite			'color all rows of the first Colum in white
    Next intZeileWS2
    For intZeileWS1 = 2 To intletzteZeileWS1									'loop over all written rows in WS1
        intletzteZeileWS2 = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
        bolVergeben = False														'set Vergeben to FALSE
        For intZeileWS2 = 2 To intletzteZeileWS2                                'loop over rows WS2
            If cDockNoWS2 = cOwnerIDWS1 Then
                bolVergeben = True												'set Vergeben to TRUE if Docknumber exists in WS2 already
                cLocationWS2 = cLocationCodeWS1									'write Location in WS2 as LocationCode from WS1
                cCarrierWS2 = cOperatorWS1										'write Carriername in WS2 as Operator from WS1
                cDockNoWS2 = cOwnerIDWS1										'write Docknumber in WS2 as LocationCode from WS1 = Silly
                If cStatusWS2 <> "Empty" Then									'if status in Ws2 is not EMPTY
                    cStatusWS2 = Worksheets(1).Cells(intZeileWS1, 12).Value		'set status in WS2 as StatusCode from WS1
                End If
                If cStatusWS2 = "FULL" Then										'if status in WS2 is FULL then set status in WS2 as Offen
                    cStatusWS2 = "Offen"
                End If
                'cTimeWS2 = Format(Worksheets(1).Cells(intZeileWS1, 19).Value, "hh:mm")
                'cDateWS2 = Format(Worksheets(1).Cells(intZeileWS1, 20).Value, "dd.MM.yyyy")
                Worksheets(2).Cells(intZeileWS2, 8).Value = Worksheets(1).Cells(intZeileWS1, 16).Value	'write ISA from WS1
                Worksheets(2).Cells(intZeileWS2, 13).Value = ""											'Set  Neue Brücke = empty
                If Weekday(cDateWS2, 2) = 1 Then									'Color the date according to weekday
                    Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 46
                ElseIf Weekday(cDateWS2, 2) = 2 Then
                    Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 26
                ElseIf Weekday(cDateWS2, 2) = 3 Then
                    Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 4
                ElseIf Weekday(cDateWS2, 2) = 4 Then
                    Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 33
                ElseIf Weekday(cDateWS2, 2) = 5 Then
                    Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 6
                ElseIf Weekday(cDateWS2, 2) = 6 Then
                    Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 13
                Else:
                    Worksheets(2).Cells(intZeileWS2, 4).Interior.Color = vbWhite
                End If
                If cDateWS2 = "" Then												'if Date is empty color in white
                    Worksheets(2).Cells(intZeileWS2, 4).Interior.Color = vbWhite
                End If
                If cCarrierWS2 = "MENDEN" _
                Or cCarrierWS2 = "K&N" Then
                    Worksheets(2).Cells(intZeileWS2, 2).Interior.Color = vbBlue		'use blue for MENDEN and K&N
                End If
            End If
        Next intZeileWS2															'at this point every match of OwnerIDs have been found
        If bolVergeben = False Then													'following lines are only dealing with new Docknumbers
                Worksheets(2).Cells(intZeileWS2, 13).Value = "NEW"					'Set Neue Brücke = "NEW"
                cLocationWS2 = cLocationCodeWS1
                cCarrierWS2 = cOperatorWS1
                cDockNoWS2 = cOwnerIDWS1
                If Worksheets(2).Cells(intZeileWS2, 12).Value = "" Then				'if Comment is empty write status from StatusCode in WS1
                    cStatusWS2 = Worksheets(1).Cells(intZeileWS1, 12).Value
                End If
                If cStatusWS2 = "FULL" Then
                    cStatusWS2 = "Offen"
                End If
                Worksheets(2).Cells(intZeileWS2, 8).Value = Worksheets(1).Cells(intZeileWS1, 16).Value
                cTimeWS2 = Format(Worksheets(1).Cells(intZeileWS1, 19).Value, "hh:mm")
                cDateWS2 = Format(Worksheets(1).Cells(intZeileWS1, 20).Value, "dd.MM.yyyy")
                Worksheets(2).Cells(intZeileWS2, 6).Value = 300
'                Worksheets(2).Cells(intZeileWS2, 31).FormulaR1C1 = "=VLOOKUP(ROUNDUP(RC[-28]*24,0),R2C24:R24C25,2,TRUE)"


            If Weekday(cDateWS2, 2) = 1 Then
                Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 46
            ElseIf Weekday(cDateWS2, 2) = 2 Then
                Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 26
            ElseIf Weekday(cDateWS2, 2) = 3 Then
                Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 4
            ElseIf Weekday(cDateWS2, 2) = 4 Then
                Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 33
            ElseIf Weekday(cDateWS2, 2) = 5 Then
                Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 6
            ElseIf Weekday(cDateWS2, 2) = 6 Then
                Worksheets(2).Cells(intZeileWS2, 4).Interior.ColorIndex = 13
            Else:
                Worksheets(2).Cells(intZeileWS2, 4).Interior.Color = vbWhite
            End If
            If cDateWS2 = "" _
            Then Worksheets(2).Cells(intZeileWS2, 4).Interior.Color = vbWhite
            If cCarrierWS2 = "MENDEN" _
            Or cCarrierWS2 = "K&N" _
            Then Worksheets(2).Cells(intZeileWS2, 2).Interior.Color = vbBlue
        End If
    Next intZeileWS1
    intletzteZeileWS2 = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row							'neubestimmung letzte Zeile WS2
    For intZeileWS2 = intletzteZeileWS2 To 2 Step -1												'loop over WS2
        bolVergeben = False
        For intZeileWS1 = 2 To intletzteZeileWS1														'loop over WS1
            If Worksheets(2).Cells(intZeileWS2, 8).Value = Worksheets(1).Cells(intZeileWS1, 16).Value _	'if ISA from WS2 matches with ISA fom WS1
            And cDockNoWS2 = cOwnerIDWS1 Then																'and Docknumber = owner
				bolVergeben = True
        Next intZeileWS1
        If bolVergeben = False Then
			Worksheets(2).Rows(intZeileWS2).Delete
    Next intZeileWS2																				'abgeschlossen: in WS2 gibt es keine mit WS1 nicht übereinstimmende ISA
	intletzteZeileWS2 = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
    For intZeileWS2 = 2 To intletzteZeileWS2
        If Left(cLocationWS2, 3) = "IB0" _
        And cStatusWS2 = "Offen" _																	'falls Location = IBO und Status offen, dann Tornummer in schreiben
            Then cStatusWS2 = "An Tor: " & CStr(CInt(Right(cLocationWS2, 2)))
    Next intZeileWS2
    bolamTor = False
    Call Sortieren
    For intZeileWS2 = intletzteZeileWS2 To 2 Step -1
    Worksheets(2).Cells(intZeileWS2, 1).Interior.Color = vbWhite
        If Left(cLocationWS2, 2) = "IB" And bolamTor = False Then
            bolamTor = True
            intamTor = intZeileWS2																	'intZeile = Anzahl Brücken am Tor = intamTor
        End If
    Next intZeileWS2
    For intZeileWS2 = 2 To intamTor
        If cStatusWS2 = "Offen" _
        And (Left(Worksheets(2).Cells(intZeileWS2, 12).Value, 4) = "Mend" _
        Or Left(Worksheets(2).Cells(intZeileWS2, 12).Value, 4) = "Parc" _
        Or Left(Worksheets(2).Cells(intZeileWS2, 12).Value, 3) = "K&N" _
        Or Worksheets(2).Cells(intZeileWS2, 12).Value = "") _
        Then Worksheets(2).Cells(intZeileWS2, 1).Interior.Color = vbRed
    Next intZeileWS2                'Abschnitt hinzugefügt: es werden alle offenen mit Rot markiert
        For intZeileWS2 = 2 To intletzteZeileWS2
        If cStatusWS2 = "Offen" _
        Then Worksheets(2).Cells(intZeileWS2, 1).Interior.Color = vbRed
    Next intZeileWS2
    'Do While Left(cLocationWS2, 2) <> "IB" And intZeileWS2 <= intletzteZeileWS2 And bolamTor = True
        'If cStatusWS2 = "FULL" Then Worksheets(2).Cells(intZeileWS2, 1).Interior.Color = vbRed
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
            If Left(cLocationWS2, 3) = "IB0" Then
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
        'If CInt(Left(cLocationWS2, 3)) <= intmin Then
            'intmin = CInt(Left(cLocationWS2, 3))
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
