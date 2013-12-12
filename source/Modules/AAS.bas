Attribute VB_Name = "AAS"
Option Explicit

' version 1.0
'
' Return to Castle Wolfenstein SBWL (*.aas)

'=> Get Header
Function AAS_get_Header(AASFilePath As String)
AAS_get_Header = get_Header_Len4(AASFilePath)
End Function
'=> Valid Header?
Function AAS_checkIfValid(AASFilePath As String) As Boolean
If AAS_get_Header(AASFilePath) = "EAAS" Then
AAS_checkIfValid = True
Else
AAS_checkIfValid = False
End If
End Function
