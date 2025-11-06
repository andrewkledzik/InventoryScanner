Attribute VB_Name = "CRC"
Option Explicit

'http://www.vbforums.com/showthread.php?412922-How-find-CRC32-checksums-of-files

Private pInititialized As Boolean
Private pTable(0 To 255) As Long

Public Sub CRCInit(Optional ByVal Poly As Long = &HEDB88320)

Dim CRC As Long
Dim i As Integer
Dim j As Integer

For i = 0 To 255
    CRC = i
    For j = 0 To 7
        If CRC And &H1 Then
            'CRC = (CRC >>> 1) ^ Poly
            CRC = ((CRC And &HFFFFFFFE) \ &H2 And &H7FFFFFFF) Xor Poly
        Else
            'CRC = (CRC >>> 1)
            CRC = CRC \ &H2 And &H7FFFFFFF
        End If
    Next j
    pTable(i) = CRC
Next i
pInititialized = True

End Sub

Public Function CRCFileContent(Path As String) As Long

Dim Buffer() As Byte
Dim BufferSize As Long
Dim CRC As Long
Dim FileNr As Integer
Dim Length As Long
Dim i As Long

If Not pInititialized Then CRCInit

BufferSize = &H1000 '4 KB
ReDim Buffer(1 To BufferSize)

FileNr = FreeFile
Open Path For Binary As #FileNr
Length = LOF(FileNr)

CRC = &HFFFFFFFF
    
Do While Length
    
    If Length < BufferSize Then
        BufferSize = Length
        ReDim Buffer(1 To Length)
    End If
    
    Get #FileNr, , Buffer
    
    For i = 1 To BufferSize
        CRC = ((CRC And &HFFFFFF00) \ &H100) And &HFFFFFF Xor pTable(Buffer(i) Xor CRC And &HFF&)
    Next i
    
    Length = Length - BufferSize
    
Loop

CRCFileContent = Not CRC

Close #FileNr

End Function

Sub TEST_Hex()

'To get the hex checksum and it's perfect!
Debug.Print Hex(CRCFileContent("F:\_4.0.4\ExecutionFlow Template_v4.0.4_v5_CopiedToServer.xlsm"))

End Sub
