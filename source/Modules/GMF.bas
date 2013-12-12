Attribute VB_Name = "GMF"
' version 1.0

' John Deere American Builder Deluxe (*.gmf *.gma *.gms)


'=> Get header (GMF_GMA)
Function GMF_GMA_get_Header(GMF_GMAFilePath As String)
' Print result
GMF_GMA_get_Header = get_Header_Len4(GMF_GMAFilePath)
End Function

'=> Valid? (GMF_GMA)
Public Function GMF_GMA_CheckIfValid(GMF_GMAFilePath As String) As Boolean
' Dims
Dim strTMP As String: strTMP = get_Header_Len3(GMF_GMAFilePath)
' Check
If strTMP = "GMA" Then
GMF_GMA_CheckIfValid = True
End If
End Function

'=> Valid? (GMF_GMI)
Public Function GMF_GMA_CheckIfValidGMI(GMF_GMAFilePath As String) As Boolean
' Dims
Dim strTMP As String: strTMP = get_Header_Len3(GMF_GMAFilePath)
' Check
If strTMP = "GMI" Then
GMF_GMA_CheckIfValidGMI = True
End If
End Function


