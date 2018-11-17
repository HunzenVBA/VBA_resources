Option Explicit

Public time                         As Date
Public StartTimerVar                As Date
Public EndTimerVar                  As Date
Public qt                           As QueryTable
'Public Zeitstempel As String
Public ZeitstempelWS1               As String
Public ZeitstempelWS2               As String
Public ZeitstempelWS3               As String
Public ZeitstempelWS4               As String
Public ws                           As Worksheet
Public ws1                          As Worksheet
Public ws2                          As Worksheet
Public ws3                          As Worksheet
Public ws4                          As Worksheet
Public ws5                          As Worksheet
Public ws6                          As Worksheet
Public ws7                          As Worksheet
Public qT                           As QueryTable
Public qts                          As QueryTables
Public StartTime                    As Double
Public SecondsElapsed               As Double
Public currProcedureName            As String

''Directories
''work
'Public Const strYardFile As String = "C:\Users\Apr17\Documents\VBA\unprocessedbrckenbersicht\Brueckenuebersicht_Dummy _neu.xlsm"
'Public Const strUnprocessedFile As String = "C:\Users\Apr17\Documents\VBA\unprocessedbrckenbersicht\Unprocessed Unitbasis - Kopie _neu.xlsm"
'Public Const strDMdata As String = "C:\Users\Apr17\Documents\VBA\unprocessedbrckenbersicht\DM-data.xlsx"
'Public Const strDMdataSicherung As String = "C:\Users\Apr17\Documents\VBA\unprocessedbrckenbersicht\DM-data_Sicherung.xlsx"
'Public Const strDockmasterFile As String = "C:\Users\Apr17\Documents\VBA\unprocessedbrckenbersicht\Dockmaster.xlsm"
'
''work
'Public Const strYardFile As String = "C:\Users\denizku\Documents\Inbound\Unprocessed\mail 01112018\Brueckenuebersicht_Dummy _neu.xlsm"
'Public Const strUnprocessedFile As String = "C:\Users\denizku\Documents\Inbound\Unprocessed\mail 01112018\Unprocessed Unitbasis - Kopie _neu.xlsm"
'Public Const strDMdata As String = "\\ant\dept-eu\DTM2\DTM2-Inbound\extern\Flowmanagement\Daten\DM-data.xlsx"
'Public Const strDMdataSicherung As String = "\\ant\dept-eu\DTM2\DTM2-Inbound\extern\Flowmanagement\Daten\DM-data_Sicherung.xlsx"
'Public Const strDockmasterFile As String = "\\ant\dept-eu\DTM2\DTM2-Inbound\extern\#New Volume Count\Dockmaster.xlsm"

'home2
Public Const strYardFile As String = "C:\Users\Apr17\Documents\VBA\Inbound\SwapFiles\Brueckenuebersicht.xlsm"
Public Const strUnprocessedFile As String = "C:\Users\Apr17\Documents\VBA\Inbound\SwapFiles\Unprocessed Unitbasis.xlsm"
Public Const strDMdata As String = "C:\Users\Apr17\Documents\VBA\Inbound\SwapFiles\DM-data.xlsx"
Public Const strDMdataSicherung As String = "C:\Users\Apr17\Documents\VBA\Inbound\SwapFiles\DM-data_Sicherung.xlsx"
Public Const strDockmasterFile As String = "C:\Users\Apr17\Documents\VBA\Inbound\SwapFiles\Dockmaster.xlsm"
Public Const strYardFileName as string = "Brueckenuebersicht"
