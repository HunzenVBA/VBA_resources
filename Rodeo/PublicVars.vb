Option Explicit

'Public Const strRodeoHistoryFile As String = "C:\Users\denizku\Documents\Rodeo\RodeoImportHistory.xlsx"
'Public Const strcsxStampsFile As String = "C:\Users\denizku\Documents\Rodeo\csxStamps.xlsx"

'Home'
Public Const strRodeoHistoryFile As String = "C:\Users\Apr17\Documents\VBA\Rodeo\RodeoImportHistory.xlsx"
Public Const strcsxStampsFile As String = "C:\Users\Apr17\Documents\VBA\Rodeo\csxStamps.xlsx"



Public ws                           As Worksheet
Public ws1                          As Worksheet
Public ws2                          As Worksheet
Public ws3                          As Worksheet
Public ws4                          As Worksheet
Public ws5                          As Worksheet
Public ws6                          As Worksheet
Public ws7                          As Worksheet
Public qtRodeoRecLane               As QueryTable
Public qts                          As QueryTables
Public StartTime                    As Double
Public SecondsElapsed               As Double
Public currProcedureName            As String
