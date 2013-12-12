Attribute VB_Name = "LOC"
Option Explicit
' NFS5 Locale File
' version 1.00

' Need For Speed 5 (Porsche 2000 / Porsche Unleashed) *.loc

'=> Valid Header?
Function LOC_validHeader(FilePath As String) As Boolean
If jrBytes.get_Header_Len4(FilePath) = "LOCH" Then
LOC_validHeader = True
Else
LOC_validHeader = False
End If
End Function
'=> Chuncks
Function LOC_get_Chuncks(FilePath As String)
LOC_get_Chuncks = get_4Bytes(FilePath, 13)
End Function
'=> First Chunck Offset
Function LOC_get_FirstChunckOffset(FilePath As String)
LOC_get_FirstChunckOffset = get_4Bytes(FilePath, 17)
End Function


