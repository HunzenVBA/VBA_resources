Sub testdynamicarray()
Dim counter As Long
'Dim testarray(1 To 1) As Variant
ReDim testarray(1 To 5)

For counter = 1 To 10

    ReDim Preserve testarray(1 To counter)
    testarray(counter) = counter
Next counter
