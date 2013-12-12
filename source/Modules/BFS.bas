Attribute VB_Name = "BFS"
Option Explicit

' version 1.0
'
' FlatOut 1
' FlatOut 2

'=> Get Header
Function BFS_get_Header(BFSFilePath As String)
BFS_get_Header = get_Header_Len4(BFSFilePath)
End Function
'=> Valid Header?
Function BFS_checkIfValid(BFSFilePath As String) As Boolean
If BFS_get_Header(BFSFilePath) = "bfs1" Then
BFS_checkIfValid = True
Else
BFS_checkIfValid = False
End If
End Function
