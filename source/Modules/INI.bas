Attribute VB_Name = "INI"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal Filename$)

Function ReadINI(ByVal INIFileLoc As String, ByVal Section As String, ByVal Key As String)
    
    Dim RetVal As String, Worked As Integer
    RetVal = String$(255, 0)
    Worked = GetPrivateProfileString(Section, Key, "", RetVal, Len(RetVal), INIFileLoc)
    
    If Worked = 0 Then
        ReadINI = ""
    Else
        ReadINI = Left(RetVal, InStr(RetVal, Chr(0)) - 1)
    End If
End Function
Function AddINI(ByVal INIFileLoc As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)
    WritePrivateProfileString Section, Key, Value, INIFileLoc
End Function

Public Sub LoadSections(path As String, Combo As ComboBox)
    Dim f As Integer, str As String
    f = FreeFile
    Open path For Input As #f
    str = Input(LOF(f), #f)
    Close #f
    Do: DoEvents
        If InStr(str, "[") = 0 Then Exit Do
        str = Mid(str, InStr(str, "[") + 1)
        Combo.AddItem Mid(str, 1, InStr(str, "]") - 1)
        str = Mid(str, InStr(str, "]") + 1)
    Loop
End Sub

