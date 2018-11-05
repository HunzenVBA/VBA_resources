Option Explicit

Public time                         As Date
Public StartTimerVar                As Date
Public EndTimerVar                  As Date
Public ws                           As Worksheet
Public qt                           As QueryTable
'Public Zeitstempel As String
Public i                            As Integer
Public ZeitstempelWS1               As String
Public ZeitstempelWS2               As String
Public ZeitstempelWS3               As String
Public ZeitstempelWS4               As String

'Directories
'work
Public Const strYardFile As String = "C:\Users\denizku\Documents\Inbound\Unprocessed\mail511218\Brueckenuebersicht.xlsm"
Public Const strUnprocessedFile As String = "C:\Users\denizku\Documents\Inbound\Unprocessed\mail511218\Unprocessed Unitbasis.xlsm"
Public Const strDMdata As String = "\\ant\dept-eu\DTM2\DTM2-Inbound\extern\Flowmanagement\Daten\DM-data.xlsx"
Public Const strDMdataSicherung As String = "\\ant\dept-eu\DTM2\DTM2-Inbound\extern\Flowmanagement\Daten\DM-data_Sicherung.xlsx"
Public Const strDockmasterFile As String = "\\ant\dept-eu\DTM2\DTM2-Inbound\extern\#New Volume Count\Dockmaster.xlsm"

'home2
Public Const strYardFile As String = "C:\Users\Apr17\Documents\VBA\unprocessed2\Brueckenuebersicht.xlsm"
Public Const strUnprocessedFile As String = "C:\Users\Apr17\Documents\VBA\unprocessed2\Unprocessed Unitbasis.xlsm"
Public Const strDMdata As String = "C:\Users\Apr17\Documents\VBA\unprocessed2\DM-data.xlsx"
Public Const strDMdataSicherung As String = "C:\Users\Apr17\Documents\VBA\unprocessed2\DM-data_Sicherung.xlsx"
Public Const strDockmasterFile As String = "C:\Users\Apr17\Documents\VBA\unprocessed2\Dockmaster.xlsm"
