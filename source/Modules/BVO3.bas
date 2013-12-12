Attribute VB_Name = "BVO3"
Option Explicit

'version 1.0
'
'Driv3r
'
'Vehicle Config. file (BVO3) *.BO3

'=> Get Header
Function BO3_get_Header(BO3FilePath As String)
BO3_get_Header = get_Header_Len4(BO3FilePath)
End Function
'=> Valid Header?
Function BO3_checkIfValid(BO3FilePath As String) As Boolean
If BO3_get_Header(BO3FilePath) = "BVO3" Then
BO3_checkIfValid = True
Else
BO3_checkIfValid = False
End If
End Function

'=> First Chunck position
Function BO3_get_FirstChunckPosition(BO3FilePath As String)
BO3_get_FirstChunckPosition = get_4Bytes(BO3FilePath, 21)
End Function
