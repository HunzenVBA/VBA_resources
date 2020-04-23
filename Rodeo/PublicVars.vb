Option Explicit

'Work
Public Const strRodeoHistoryFile As String = "C:\Users\denizku\Documents\Rodeo\RodeoImportHistory.xlsm"
Public Const strcsxStampsFile As String = "C:\Users\denizku\Documents\Rodeo\csxStamps.xlsx"
Public Const strRodeoHistoryFileName As String = "RodeoImportHistory.xlsm"
Public Const strcsxStampsFileName As String = "csxStamps.xlsx"
Public Const strRodeoPath As String = "C:\Users\denizku\Documents\Rodeo\"
Public Const strRodeoFiltered1 As String = "C:\Users\denizku\Documents\Rodeo\RodeoOhneLoadedundPalletized.xlsx"
Public Const strRodeoWorkpoolFileName As String = "RodeoWorkpoolFiltered.xlsx"

'Home'
'Public Const strRodeoHistoryFile As String = "C:\Users\Apr17\Documents\VBA\Rodeo\RodeoImportHistory.xlsx"
'Public Const strcsxStampsFile As String = "C:\Users\Apr17\Documents\VBA\Rodeo\csxStamps.xlsx"



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
Public StartTimeAll                 As Double
Public SecondsElapsedAll            As Double
Public EndTimerVar                  As Date
Public time                         As Date
Public StartTimerVar                As Date
Public globalcounter                As Long
Public collAllUniqeCSX              As Collection
Public collUniqueDicts              As Collection
