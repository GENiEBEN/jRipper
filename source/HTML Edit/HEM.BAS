Attribute VB_Name = "hem"
Public fMainForm As frmMain
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Public asp As String
Public php As String
Public jsp As String
Public js As String
Public cgi As String
Public css As String
Public lc As String

Const cmPrnPromptDlg As Integer = &H0
Const cmPrnDefaultPrn As Integer = &H1
Const cmPrnHDC As Integer = &H2
Const cmPrnRichFonts As Integer = &H4
Const cmPrnColor As Integer = &H8
Const cmPrnPageNums As Integer = &H10
Const cmPrnDateTime As Integer = &H20
Const cmPrnBorderThin As Integer = &H40
Const cmPrnBorderThick As Integer = &H80
Const cmPrnBorderDouble As Integer = &H100
Const cmPrnSelection = &H200

Public inifile As String
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public splashshow As Boolean
Public color As String


Sub Install(ByVal DllName As String)

    On Error GoTo 2
    FileCopy App.Path & "\" & DllName, GetWindowsSysDir & DllName
2

    Shell "regsvr32 " & DllName & " /s"

End Sub
Public Sub HHShowContents(Filename As String, lhWnd As Long)
    HTMLHelp lhWnd, Filename & "", HH_DISPLAY_TOC, 0
End Sub
Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = writeprivateprofilestring(sSection, sKeyName, sNewString, sFileName)
End Function


Function ReadINI(Section As String, KeyName As String, Filename As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName$, "", sRet, Len(sRet), Filename))
End Function

Sub Main()
Install App.Path & "\bin\wwtabs.ocx"
Install App.Path & "\bin\cbox.ocx"

    Set fMainForm = New frmMain
    inifile = App.Path & "\HTML Edit.ini"

End Sub
