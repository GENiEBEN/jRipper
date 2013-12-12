Attribute VB_Name = "BIK"
Option Explicit
Dim fso As New FileSystemObject
' version 1.0
'
' BiK Player (BINK Movies)

' Play BIK Movie (using 3rd party tool)
Function BIK_play(ByVal BIK_FilePath As String, ByVal switches As String) As Long
If fso.FileExists(BIK_FilePath) = True Then
    Shell (App.Path & "\bin\binkplay.dll " & Chr(34) & BIK_FilePath & Chr(34) & switches), vbNormalFocus
Else
    MsgBox "Invalid Path", vbOKOnly, "jRipper " & IAPPV
End If
End Function

'Get Number of Frames
Function BIK_Frames(ByVal BIK_FilePath As String)
BIK_Frames = get_4Bytes(BIK_FilePath, 9)
End Function

' Video Width
Function BIK_VideoWidth(ByVal BIK_FilePath As String)
BIK_VideoWidth = get_4Bytes(BIK_FilePath, 21)
End Function

' Video Height
Function BIK_VideoHeight(ByVal BIK_FilePath As String)
BIK_VideoHeight = get_4Bytes(BIK_FilePath, 25)
End Function

' Frames-per-Second
Function BIK_VideoFPS(ByVal BIK_FilePath As String)
BIK_VideoFPS = get_4Bytes(BIK_FilePath, 29)
End Function

' Largest Frame Size
Function BIK_LargestFrameSize(ByVal BIK_FilePath As String)
BIK_LargestFrameSize = get_4Bytes(BIK_FilePath, 13)
End Function

