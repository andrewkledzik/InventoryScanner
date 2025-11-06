VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScanDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public dicFoldersToScan                 As Dictionary

Private pIsMultiScan                    As Boolean
Private pDeleteOriginalFolder           As Boolean


'Single folder or batch folder scanning
Public Property Get IsMultiScan() As Boolean
    IsMultiScan = pIsMultiScan
End Property
Public Property Let IsMultiScan(value As Boolean)
    pIsMultiScan = value
End Property












'Initialise the class
Private Sub Class_Initialize()

    Set dicFoldersToScan = New Dictionary          'Folders to scan

End Sub
