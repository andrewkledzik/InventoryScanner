Attribute VB_Name = "Inventory"
Option Explicit

Public Const VERSION_NUMBER As String = "1.0"

Public iColumn As Long
Dim strInventoryWbName As String

'Inputs:    A high level folder name
'Outputs:   File properties and Information
'           Written to a worksheet
'           Each folder name with related file counts and subfolder counts

Sub GetFileList(folder As Scripting.folder, fso As Scripting.FileSystemObject)

    Dim startRange As Range
    Set startRange = ThisWorkbook.Sheets(strInventoryWbName).Range("A2")

    Dim subFolder As Scripting.folder
    Dim file As Scripting.file
    
    'TODO:
    'Write the high level folder to the inventory with counts

    startRange.Offset(iColumn, 0).value = folder.Path
    startRange.Offset(iColumn, 2).value = Round(folder.Size / 1024)
    startRange.Offset(iColumn, 7).value = folder.files.Count
    startRange.Offset(iColumn, 8).value = folder.SubFolders.Count

    iColumn = iColumn + 1
    
    Application.StatusBar = "Scanning Folder: " & Left(folder.Path, 237)
    For Each file In folder.files
        
        startRange.Offset(iColumn, 0).value = file.ParentFolder
        startRange.Offset(iColumn, 1).value = file.Name
        startRange.Offset(iColumn, 2).value = Round(file.Size / 1024)
        startRange.Offset(iColumn, 3).value = file.Type
        startRange.Offset(iColumn, 4).value = file.DateCreated
        startRange.Offset(iColumn, 5).value = file.DateLastAccessed
        startRange.Offset(iColumn, 6).value = file.DateLastModified
        If IS_CHECKSUM_ON Then _
            startRange.Offset(iColumn, 9).value = CRCFileContent(file.Path)
        
        iColumn = iColumn + 1
        
    Next file
    
    For Each subFolder In folder.SubFolders
      'Debug.Print subFolder.Name
      GetFileList subFolder, fso
    Next subFolder
  
End Sub

Sub ScanInventoryAndGetFileFolderList()

    Dim fso As New Scripting.FileSystemObject
    Dim folderPicker As FileDialog
    Dim folderToScan As String

    'Open up a dialog box and allow the user to pick a foler
    Set folderPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With folderPicker
        .Title = "Select the folder you want to scan then click on Scan Inventory"
        .AllowMultiSelect = False
        .InitialFileName = "C:\"
        .ButtonName = "Scan Inventory"
        .Show
    End With
    
    If folderPicker.SelectedItems.Count > 0 Then
        'Scan the selected Folder
        folderToScan = folderPicker.SelectedItems(1)
        
        'Display a confirmation Ok box.
        If MsgBox("You have selected the following Folder to scan:" & vbCrLf & vbCrLf & folderToScan & vbCrLf & vbCrLf & _
        "Select Yes to confirm and continue or No to cancel or pick again", vbYesNo + vbQuestion, "Confirm selection") = vbYes Then
        
            'Create a new worksheet
            Debug.Print "Started At: " & Time()
            
            setUpWorksheet
            
            iColumn = 0
            Application.ScreenUpdating = False
            ThisWorkbook.Sheets(strInventoryWbName).Select
            GetFileList fso.GetFolder(folderToScan), fso
            
            'Tidy up the worksheet
            formatInventorySheet
            Application.ScreenUpdating = True
            
            Debug.Print "Finished At: " & Time()
            
        End If
        
    Else
    
        'Nothing Selected
        MsgBox "Whoops!" & vbCrLf & vbCrLf & "You didn't select a folder." & vbCrLf & _
        "To run the inventory, please select a folder then click on Scan Inventory.", vbInformation, "No folder selected"
    
    End If

Application.StatusBar = ""
If IsObject(fso) Then Set fso = Nothing
If IsObject(folderPicker) Then Set folderPicker = Nothing

End Sub

Sub setUpWorksheet(Optional wsWorksheetName As String = vbNullString)

'Add a new worksheet
'Call it inventory
'Add some field headers

Dim wsInventory As Worksheet
Dim headersStartRange As Range

  
    'Add a new sheet
    Set wsInventory = ThisWorkbook.Sheets.Add(After:=Sheets(Worksheets.Count))
    
    If wsWorksheetName = vbNullString Then wsWorksheetName = "Inventory_" & Format(Now(), "yyyydd_hhmmss")
    
    wsInventory.Name = wsWorksheetName
    strInventoryWbName = wsInventory.Name
    
    Set headersStartRange = wsInventory.Range("A1")
    
    'Paste in the headers
    headersStartRange.Offset(0, 0).value = "Folder Name"
    headersStartRange.Offset(0, 1).value = "File Name"
    headersStartRange.Offset(0, 2).value = "Size"
    headersStartRange.Offset(0, 3).value = "File Type"
    headersStartRange.Offset(0, 4).value = "Date Created"
    headersStartRange.Offset(0, 5).value = "Date Last Accessed"
    headersStartRange.Offset(0, 6).value = "Date Last Modified"
    headersStartRange.Offset(0, 7).value = "Number Of Files"
    headersStartRange.Offset(0, 8).value = "Number Of Subfolders"
    If IS_CHECKSUM_ON Then _
        headersStartRange.Offset(0, 9).value = "Check Sum"
    
    'Make the header row bold and freeze the top pane
    wsInventory.Range("A1").EntireRow.Font.Bold = True
    wsInventory.Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    ActiveWindow.Zoom = 80

'Kill objects
If IsObject(wsInventory) Then Set wsInventory = Nothing
If IsObject(headersStartRange) Then Set headersStartRange = Nothing


End Sub

Sub formatInventorySheet()

Dim ws As Worksheet

'AutoFit all the columns
Set ws = ThisWorkbook.Sheets(strInventoryWbName)
ws.Range("C:I").EntireColumn.AutoFit

'Don't make the first 2 too l
ws.Columns("A").ColumnWidth = 25
ws.Columns("B").ColumnWidth = 50

'Add an auto filter
ws.Range("A1").AutoFilter

End Sub

