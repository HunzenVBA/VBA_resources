Sub CreatePTandPivotCache()

Dim pC As PivotCache
Dim pT As PivotTable

    Set pC = ThisWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=shtImportSomethingxyz.Name & "!" & shtImportSomethingxyz.Range("A1").CurrentRegion.Address, _
    Version:=xlPivotTableVersion15)

    Set pT = pC.CreatePivotTable( _
    TableDestination:=ActiveCell, _
    TableName:="ImportPivot")

    Debug.Print ThisWorkbook.PivotCaches.Count
    Debug.Print pC.MemoryUsed, pC.RecordCount, pC.Version
End Sub

Sub AddPTandPivotCache()

Dim ws As Worksheet
Dim pC As PivotCache
Dim pT As PivotTable

    Set pC = ThisWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=shtImportSomethingxyz.Name & "!" & shtImportSomethingxyz.Range("A1").CurrentRegion.Address, _
    Version:=xlPivotTableVersion15)

    Set pT = ws.PivotTables.add(_
    PivotCache:= pC, _
    TableDestination:=ActiveCell, _
    TableName:= "ImportPivot2")

    Debug.Print ThisWorkbook.PivotCaches.Count
    Debug.Print pC.MemoryUsed, pC.RecordCount, pC.Version
End Sub

Sub CreatePTUsingExistingPivotCache()

Dim pC As PivotCache
Dim pT As PivotTable
Dim pvtField as PivotField

  If ThisWorkbook.PivotCaches.Count = 0 THen
  Set pC = ThisWorkbook.PivotCaches.Create( _
  SourceType:=xlDatabase, _
  SourceData:=shtImportSomethingxyz.Name & "!" & shtImportSomethingxyz.Range("A1").CurrentRegion.Address, _
  Version:=xlPivotTableVersion15)
  Else
    Set pC = ThisWorkbook.PivotCaches(1)
    End If

    Set pT = pC.CreatePivotTable( _
    TableDestination:=ActiveCell, _
    TableName:="ImportPivot")

    Debug.Print ThisWorkbook.PivotCaches.Count
    Debug.Print pC.MemoryUsed, pC.RecordCount, pC.Version

    Set pvtField = pt.PivotFields("Genre")
    pvtField.Orientation = xlRowField
End Sub
