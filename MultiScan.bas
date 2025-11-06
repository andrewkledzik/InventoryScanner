Attribute VB_Name = "MultiScan"
Option Explicit

Dim wMultiScan As Worksheet
Dim strInventoryWbName As String

'1. Get range starting in A2
'2. Must be at least 1 value
'3. If folder does not exist, report back status


Sub ScanAllFoldersAndGetFileFolderList()


'Instantiate the class
Set currentScanDetails = New ScanDetails

'Get the folders
SetAndPopulateMultiFolderDictionary


RunMultiScan


Application.StatusBar = ""


End Sub



Sub RunMultiScan()

Dim lDictionaryItem As Long
Dim folderToScan As String
Dim fso As New Scripting.FileSystemObject


If currentScanDetails.dicFoldersToScan.Count > 0 Then

Application.ScreenUpdating = False

setUpWorksheet
iColumn = 0

    For lDictionaryItem = 1 To currentScanDetails.dicFoldersToScan.Count
            
        folderToScan = currentScanDetails.dicFoldersToScan.Item(lDictionaryItem + 1)
        
        'Check for the existance of the folder
        If fso.FolderExists(folderToScan) Then
            GetFileList fso.GetFolder(folderToScan), fso
            SetFolderScanResultStatus lDictionaryItem - 1, "Complete"
        Else
            SetFolderScanResultStatus lDictionaryItem - 1, "Folder Not Found"
        End If
        
    Next
    
    'Tidy up the worksheet
    formatInventorySheet

Else

End If

Application.ScreenUpdating = True

End Sub


'Get the range which the user has populated.
'1. Must begin in A2
'2. Must be continuous
'3. It will stop at the first blank row, no need for over-kill
Sub SetAndPopulateMultiFolderDictionary()

Dim lOffset As Long
Dim rOffset As Range

Set wMultiScan = ThisWorkbook.Sheets("MultiFolderScan")
Set rOffset = wMultiScan.Range("A2")

currentScanDetails.dicFoldersToScan.RemoveAll

Do Until IsEmpty(rOffset.value)

    'Debug.Print rOffset.value
    currentScanDetails.dicFoldersToScan.Add rOffset.Row, rOffset.value
    Set rOffset = rOffset.Offset(1, 0)

Loop

currentScanDetails.IsMultiScan = True

'If IsObject(wMultiScan) Then Set wMultiScan = Nothing

End Sub

'Does a simple replace on the worksheet name
Function GetCleanWorksheetName(ByVal sDirtyName As String) As String

sDirtyName = Replace(sDirtyName, ":", "")
sDirtyName = Replace(sDirtyName, "\", "_")
sDirtyName = Replace(sDirtyName, "/", "_")
sDirtyName = Replace(sDirtyName, "?", "_")
sDirtyName = Replace(sDirtyName, "*", "_")

GetCleanWorksheetName = Trim(Left(sDirtyName, 32))

End Function

'Now that we have scanned the folder, were we successfull
Sub SetFolderScanResultStatus(lRowNumber As Long, sResult As String)

Dim rCellToUpdate As Range


Set rCellToUpdate = wMultiScan.Range("B" & currentScanDetails.dicFoldersToScan.Keys(lRowNumber))


'ActiveCell.value = "ar"

Select Case sResult

    Case "Complete"
        'With wMultiScan.Cells(currentScanDetails.dicFoldersToScan.Keys(lRowNumber), 2)
        With rCellToUpdate
            .value = "a"
            .Font.Name = "Webdings"
            .Font.Size = 11
        End With
    
    Case "Folder Not Found"
        
        wMultiScan.Cells(currentScanDetails.dicFoldersToScan.Keys(lRowNumber), 2).value = sResult
        rCellToUpdate.value = sResult
        
    Case "Fail"
        
        rCellToUpdate.value = sResult

    Case Else


End Select

End Sub

