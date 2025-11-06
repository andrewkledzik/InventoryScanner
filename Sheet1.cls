VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


'Begin the multi folder scan
Private Sub cmdMultiFolderScan_Click()

    Debug.Print "Started At: " & Time()
    ScanAllFoldersAndGetFileFolderList
    Debug.Print "Started At: " & Time()

End Sub
