' Purpose   :       Prints all subs and functions in a project
' Prerequisites:    Microsoft Visual Basic for Applications Extensibility 5.3 library
'                   CreateLogFile
' How to run:       Run GetFunctionAndSubNames, set a parameter to blnWithParentInfo
'                   If ComponentTypeToString(vbext_ct_StdModule) = "Code Module" Then
'
' Used:             ComponentTypeToString from -> http://www.cpearson.com/excel/vbe.aspx
'---------------------------------------------------------------------------------------
 
Option Explicit
Public STR_ERROR_REPORT                 As String
Private strSubsInfo As String
Public Sub GetFunctionAndSubNames()
 
    Dim item            As Variant
    strSubsInfo = ""
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        
        If ComponentTypeToString(vbext_ct_StdModule) = "Code Module" Then
            ListProcedures item.Name, False
'            Debug.Print item.CodeModule.Lines(1, item.CodeModule.CountOfLines)
        End If
        
    Next item
    CreateLogFile strSubsInfo
End Sub
 
Private Sub ListProcedures(strName As String, Optional blnWithParentInfo = True)
 
    'Microsoft Visual Basic for Applications Extensibility 5.3 library
 
    Dim vbProj          As VBIDE.VBProject
    Dim vbComp          As VBIDE.VBComponent
    Dim CodeMod         As VBIDE.CodeModule
    Dim LineNum         As Long
    Dim ProcName        As String
    Dim ProcKind        As VBIDE.vbext_ProcKind
    Dim subName         As String
    Dim wb              As Workbook
    
    ThisWorkbook.Activate
    Set vbProj = ActiveWorkbook.VBProject
    Set vbComp = vbProj.VBComponents(strName)
    Set CodeMod = vbComp.CodeModule
    blnWithParentInfo = True
 
    With CodeMod
        LineNum = .CountOfDeclarationLines + 1
        
        Do Until LineNum >= .CountOfLines
            ProcName = .ProcOfLine(LineNum, ProcKind)
 
            If blnWithParentInfo Then
                strSubsInfo = strSubsInfo & IIf(strSubsInfo = vbNullString, vbNullString, vbCrLf) & strName & "." & ProcName
            Else
                strSubsInfo = strSubsInfo & IIf(strSubsInfo = vbNullString, vbNullString, vbCrLf) & ProcName
            End If
 
            LineNum = .ProcStartLine(ProcName, ProcKind) + .ProcCountLines(ProcName, ProcKind) + 1
        Loop
    End With
End Sub
 
Function ComponentTypeToString(ComponentType As VBIDE.vbext_ComponentType) As String
    'ComponentTypeToString from http://www.cpearson.com/excel/vbe.aspx
    Select Case ComponentType
    
        Case vbext_ct_ActiveXDesigner
            ComponentTypeToString = "ActiveX Designer"
            
        Case vbext_ct_ClassModule
            ComponentTypeToString = "Class Module"
            
        Case vbext_ct_Document
            ComponentTypeToString = "Document Module"
            
        Case vbext_ct_MSForm
            ComponentTypeToString = "UserForm"
            
        Case vbext_ct_StdModule
            ComponentTypeToString = "Code Module"
            
        Case Else
            ComponentTypeToString = "Unknown Type: " & CStr(ComponentType)
            
    End Select
End Function

' export to notepad export txt export string string to txt string to notepad
Sub CreateLogFile(Optional str_print As String)

    On Error GoTo CreateLogFile_Error

    Dim fs                      As Object
    Dim obj_text                As Object
    Dim str_filename            As String
    Dim str_new_file            As String
    Dim str_shell               As String

    str_new_file = ThisWorkbook.Name & " "

    str_filename = ThisWorkbook.Path & str_new_file & TimeString
    If Dir(ThisWorkbook.Path & str_new_file, vbDirectory) = vbNullString Then MkDir ThisWorkbook.Path & str_new_file

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set obj_text = fs.CreateTextFile(str_filename, True)

    If Len(STR_ERROR_REPORT) > 1 Then
        obj_text.writeline (STR_ERROR_REPORT)
    Else
        obj_text.writeline (str_print)
    End If
    
    obj_text.Close

    str_shell = "C:\WINDOWS\notepad.exe "
    str_shell = str_shell & str_filename
    Call Shell(str_shell)

    On Error GoTo CreateLogFile_Error
    Exit Sub

CreateLogFile_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateLogFile of Sub mod_TDD_Export"
    
End Sub

Public Function codify_time(Optional b_make_str As Boolean = False) As String

    On Error GoTo codify_Error
    
    Dim dbl_01                  As Variant
    Dim dbl_02                  As Variant
    Dim dbl_now                 As Double
    
    dbl_now = Round(Now(), 8)
    
    dbl_01 = Split(CStr(dbl_now), ",")(0)
    dbl_02 = Split(CStr(dbl_now), ",")(1)
    
    codify_time = Hex(dbl_01) & "_" & Hex(dbl_02)
    
    If b_make_str Then codify_time = "\" & codify_time & ".txt"
    
    On Error GoTo 0
    Exit Function

codify_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure codify of Function TDD_Export"

End Function

Public Function TimeString() As String
    TimeString = Format(Now, "hh_mm_ss")
End Function

