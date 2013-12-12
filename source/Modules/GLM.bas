Attribute VB_Name = "GLM"
Option Explicit

' version 1.0
'
' Return to Castle Wolfenstein SBWL (*.glm)

'=> Get Header
Function GLM_get_Header(GLMFilePath As String)
GLM_get_Header = get_Header_Len4(GLMFilePath)
End Function
'=> Valid Header?
Function GLM_checkIfValid(GLMFilePath As String) As Boolean
If GLM_get_Header(GLMFilePath) = "2LGM" Then
GLM_checkIfValid = True
Else
GLM_checkIfValid = False
End If
End Function

