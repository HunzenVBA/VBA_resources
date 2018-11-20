'' Based on:
'' Displaying a List of All VBA Procedures in an Excel 2007 Workbook
''     from the Ribbon (June 2009)
'' by Frank Rice, Microsoft Corporation
'' http://msdn.microsoft.com/en-us/library/dd890502(office.11).aspx#

'' set a reference to the Microsoft Visual Basic for Applications Extensibility 5.3 Library

Sub GetProcedures()
  ' Declare variables to access the Excel workbook.
  Dim app As Excel.Application
  Dim wb As Excel.Workbook
  Dim wsOutput As Excel.Worksheet
  Dim sOutput() As String
  Dim sFileName As String

  ' Declare variables to access the macros in the workbook.
  Dim vbProj As VBIDE.VBProject
  Dim vbComp As VBIDE.VBComponent
  Dim vbMod As VBIDE.CodeModule

  ' Declare other miscellaneous variables.
  Dim iRow As Long
  Dim iCol As Long
  Dim iLine As Integer
  Dim sProcName As String
  Dim pk As vbext_ProcKind

  Set app = Excel.Application

  ' create new workbook for output
  Set wsOutput = app.Workbooks.Add.Worksheets(1)

  'For Each wb In app.Workbooks
  For Each vbProj In app.VBE.VBProjects

    ' Get the project details in the workbook.
    On Error Resume Next
    sFileName = vbProj.Filename
    If Err.Number <> 0 Then sFileName = "file not saved"
    On Error GoTo 0

    ' initialize output array
    ReDim sOutput(1 To 2)
    sOutput(1) = sFileName
    sOutput(2) = vbProj.Name
    iRow = 0

    ' check for protected project
    On Error Resume Next
    Set vbComp = vbProj.VBComponents(1)
    On Error GoTo 0

    If Not vbComp Is Nothing Then
      ' Iterate through each component in the project.
      For Each vbComp In vbProj.VBComponents

        ' Find the code module for the project.
        Set vbMod = vbComp.CodeModule

        ' Scan through the code module, looking for procedures.
        iLine = 1
        Do While iLine < vbMod.CountOfLines
          sProcName = vbMod.ProcOfLine(iLine, pk)
          If sProcName <> "" Then
            iRow = iRow + 1
            ReDim Preserve sOutput(1 To 2 + iRow)
            sOutput(2 + iRow) = vbComp.Name & ": " & sProcName
            iLine = iLine + vbMod.ProcCountLines(sProcName, pk)
          Else
            ' This line has no procedure, so go to the next line.
            iLine = iLine + 1
          End If
        Loop

        ' clean up
        Set vbMod = Nothing
        Set vbComp = Nothing

      Next
    Else
      ReDim Preserve sOutput(1 To 3)
      sOutput(3) = "Project protected"
    End If

    If UBound(sOutput) = 2 Then
      ReDim Preserve sOutput(1 To 3)
      sOutput(3) = "No code in project"
    End If

    ' define output location and dump output
    If Len(wsOutput.Range("A1").Value) = 0 Then
      iCol = 1
    Else
      iCol = wsOutput.Cells(1, wsOutput.Columns.Count).End(xlToLeft).Column + 1
    End If
    wsOutput.Cells(1, iCol).Resize(UBound(sOutput) + 1 - LBound(sOutput)).Value = _
        WorksheetFunction.Transpose(sOutput)

    ' clean up
    Set vbProj = Nothing
  Next

  ' clean up
  wsOutput.UsedRange.Columns.AutoFit
End Sub

