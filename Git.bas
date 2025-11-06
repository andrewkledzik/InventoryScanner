Attribute VB_Name = "Git"
Option Explicit

Sub ExportAllVbaModules()
    '----------------------------------------------------------
    ' Purpose: Export all VBA modules, class modules, and forms
    '          to a specified folder for version control.
    '
    ' Usage:   Run this macro, and it will create (if needed)
    '          a folder next to the workbook and export all
    '          VBA components as text files (.bas, .cls, .frm)
    '
    ' Author:  Andrew
    ' Date:    06-Nov-2025
    '----------------------------------------------------------
    
    Dim vbComp As Object
    Dim exportPath As String
    Dim fileName As String
    Dim fs As Object
    
    ' Create export folder (same location as workbook)
    exportPath = ThisWorkbook.Path & "\src\"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(exportPath) Then
        fs.CreateFolder exportPath
    End If
    
    ' Loop through all VBA components
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        
        Select Case vbComp.Type
            Case 1, 2, 3, 100
                ' vbext_ct_StdModule = 1
                ' vbext_ct_ClassModule = 2
                ' vbext_ct_MSForm = 3
                ' vbext_ct_Document = 100 (Sheet/Workbook)
                
                fileName = exportPath & vbComp.Name
                
                Select Case vbComp.Type
                    Case 1: fileName = fileName & ".bas"
                    Case 2: fileName = fileName & ".cls"
                    Case 3: fileName = fileName & ".frm"
                    Case 100: fileName = fileName & ".cls"
                End Select

                vbComp.Export fileName
                Debug.Print "Exported: " & fileName
        End Select
    Next vbComp
    
    MsgBox "VBA modules exported to: " & exportPath, vbInformation, "Export Complete"
    
End Sub


