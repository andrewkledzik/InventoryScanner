Attribute VB_Name = "Globals"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Const IS_CHECKSUM_ON As Boolean = False

'High level list of the folders which need to be scanned
Public currentScanDetails As ScanDetails


'Executes when the Ribbon Button is pressed
Public Sub ScanSingleFolder(control As IRibbonControl)

    ScanInventoryAndGetFileFolderList

End Sub
