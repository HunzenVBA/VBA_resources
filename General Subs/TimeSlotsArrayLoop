Option Explicit

Sub TimeSlotsMenden()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
    
    StartTimer = Timer
    Dim arrYardViewCopy As Variant
    Dim arrTimeSlots As Variant
    Dim intTabelle As Integer
    Dim intTimeSlot As Integer
    Dim counter As Long
    Dim rngYardViewCopy As Range
    Dim currDate As String
    Dim currTimeSlot As String
    Dim intAbstandDatum As Integer
    Dim intAbstandTimeSlot As Integer
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    i = 0
    j = 0
    counter = 0
    
    Set rngYardViewCopy = shtYardViewCopy.Range("A1:AE" & fLastWrittenRow(shtYardViewCopy, 31))
    arrYardViewCopy = rngYardViewCopy.Value2
    '31 = last Column = AE = Timeslot of each delivery
    '1 = Offen
    '2 = CarrierName
    '4 = Date
    
    'tabelle1
    
    intAbstandDatum = 0
        
            For intAbstandDatum = 1 To 86 Step 17                           'loop over Tables
            k = 0
                For k = 1 To 12                                             'loop over Timeslots
                counter = 0
                currDate = shtTimeslots.Cells(intAbstandDatum, 1).Value2
                currTimeSlot = shtTimeslots.Cells(1 + k + intAbstandDatum, 1).Value2
                    
                    For i = 1 To UBound(arrYardViewCopy, 1)                  'loop over array
                        If arrYardViewCopy(i, 1) <> "" Then
                            Debug.Print arrYardViewCopy(i, 1) & vbTab & arrYardViewCopy(i, 2) & vbTab & arrYardViewCopy(i, 4)
                            If currTimeSlot = arrYardViewCopy(i, 31) And arrYardViewCopy(i, 1) = "Offen" And arrYardViewCopy(i, 4) = currDate Then
        '                                    And (arrYardViewCopy(i, 2) = "Menden" Or arrYardViewCopy(i, 2) = "K&N") Then
'                                Stop
                                counter = counter + 1
                                shtTimeslots.Cells(intAbstandDatum + k + 1, 12).Value2 = counter
                            End If
                        End If
                    Next i
                Next k
            Next intAbstandDatum
    EndTimer = Timer
    PrintSecondsElapsed
End Sub


