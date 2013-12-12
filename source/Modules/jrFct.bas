Attribute VB_Name = "jRFct"
Option Explicit

Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_NOTOPMOST = -2
Public Const LB_ITEMFROMPOINT = &H1A9
Const SW_SHOWNORMAL = 1
Const STARTF_USESHOWWINDOW = &H1&

Public Declare Function Setwindowpos Lib "user32" Alias "SetWindowPos" (ByVal HWND As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWND As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



'=> Intelligent Application Version
Public Function IAPPV()
Dim V1 As Long: V1 = App.Major
Dim V2 As Long: V2 = App.Minor
Dim V3 As Long: V3 = App.Revision
If V3 > 99 And V3 < 1000 Then
V2 = Strings.Left(V3, 1)
V3 = Strings.Right(V3, Len(V3) - 2)
IAPPV = V1 & "." & "0" & V2 & "." & "0" & V3
Exit Function
ElseIf V3 > 999 And V3 < 9001 Then
V2 = Strings.Left(V3, 2)
V3 = Strings.Right(V3, Len(V3) - 1)
IAPPV = V1 & "." & V2 & "." & "00" & V3
Exit Function
End If
End Function

' ==> Get extension from a filename (advanced, not the Split method)
Function get_ExtensionFromFileName(FilePath)
Dim sEXT_tmp: Dim sEXT_tmp2: Dim c2: Dim sEXT
    For sEXT_tmp = 1 To Len(FilePath)
        sEXT_tmp2 = Strings.Mid(FilePath, sEXT_tmp, 1)
        If sEXT_tmp2 = "." Then
        c2 = Val(c2) + 1
        End If
    Next sEXT_tmp
' Return Extension
If c2 = 0 Then
get_ExtensionFromFileName = "###"
Else
sEXT = StrConv(Split(FilePath, ".")(c2), vbLowerCase)
get_ExtensionFromFileName = sEXT
End If
End Function

' ==> Get extension from a filename (advanced, not the Split method) => modified for Commandos 2 MOC
Function get_ExtensionFromFileName_prv(FilePath)
Dim sEXT_tmp: Dim sEXT_tmp2: Dim c2: Dim sEXT
    For sEXT_tmp = 1 To Len(FilePath)
        sEXT_tmp2 = Strings.Mid(FilePath, sEXT_tmp, 1)
        If sEXT_tmp2 = "." Then
        c2 = Val(c2) + 1
        End If
    Next sEXT_tmp
' Return Extension
If c2 = 0 Then
get_ExtensionFromFileName_prv = "DIR"
Else
sEXT = StrConv(Split(FilePath, ".")(c2), vbLowerCase)
get_ExtensionFromFileName_prv = sEXT
End If
End Function

'==> remove extension from a filename (not the usual Split function)
Function get_FileNameWithoutExtension(FilePath)
Dim sEXT: sEXT = get_ExtensionFromFileName(FilePath)
Dim sFN: sFN = Replace(FilePath, "." & sEXT, "", , 1)
get_FileNameWithoutExtension = sFN
End Function

'==> Set Window Position
Public Function SetWinPos(iPos As Integer, lHWnd As Long) As Boolean
    Dim lwinpos As Long
    iPos = 1
    Select Case iPos
        Case 1
            lwinpos = HWND_TOPMOST
        End Select
    If Setwindowpos(lHWnd, lwinpos, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE) Then
        SetWinPos = True
    End If
End Function

Function StartDoc(DocName As String, Form As Form) As Long
    StartDoc = ShellExecute(Form.HWND, "Open", DocName, _
    "", App.Path, STARTF_USESHOWWINDOW)
End Function

Function OpenWWW(webadress, Form As Form)
ShellExecute Form.HWND, vbNullString, webadress, vbNullString, "C:\", SW_SHOWNORMAL
End Function
