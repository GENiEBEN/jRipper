Attribute VB_Name = "jrOFN"
Option Explicit

Public OFN As OPENFILENAME
Public Const SW_RESTORE = 9
Public Const WM_SETTEXT = &HC
Public Const BM_CLICK = &HF5 ' Top MS secret? Not included into API32.txt
Public Const OFN_ALLOWMULTISELECT As Long = &H200
Public Const OFN_CREATEPROMPT As Long = &H2000
Public Const OFN_ENABLEHOOK As Long = &H20
Public Const OFN_ENABLETEMPLATE As Long = &H40
Public Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Public Const OFN_EXPLORER As Long = &H80000
Public Const OFN_EXTENSIONDIFFERENT As Long = &H400
Public Const OFN_FILEMUSTEXIST As Long = &H1000
Public Const OFN_HIDEREADONLY As Long = &H4
Public Const OFN_LONGNAMES As Long = &H200000
Public Const OFN_NOCHANGEDIR As Long = &H8
Public Const OFN_NODEREFERENCELINKS As Long = &H100000
Public Const OFN_NOLONGNAMES As Long = &H40000
Public Const OFN_NONETWORKBUTTON As Long = &H20000
Public Const OFN_NOREADONLYRETURN As Long = &H8000& 'see comments
Public Const OFN_NOTESTFILECREATE As Long = &H10000
Public Const OFN_NOVALIDATE As Long = &H100
Public Const OFN_OVERWRITEPROMPT As Long = &H2
Public Const OFN_PATHMUSTEXIST As Long = &H800
Public Const OFN_READONLY As Long = &H1
Public Const OFN_SHAREAWARE As Long = &H4000
Public Const OFN_SHAREFALLTHROUGH As Long = 2
Public Const OFN_SHAREWARN As Long = 0
Public Const OFN_SHARENOWARN As Long = 1
Public Const OFN_SHOWHELP As Long = &H10
Public Const OFS_MAXPATHNAME As Long = 260
Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_CREATEPROMPT _
             Or OFN_NODEREFERENCELINKS
Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_OVERWRITEPROMPT _
             Or OFN_HIDEREADONLY

Public Type OPENFILENAME
  nStructSize       As Long
  hWndOwner         As Long
  hInstance         As Long
  sFilter           As String
  sCustomFilter     As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  sFile             As String
  nMaxFile          As Long
  sFileTitle        As String
  nMaxTitle         As Long
  sInitialDir       As String
  sDialogTitle      As String
  flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  sDefFileExt       As String
  nCustData         As Long
  fnHook            As Long
  sTemplateName     As String
End Type

Public Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function EnumChildWindows& Lib "user32" (ByVal hParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long)
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

Dim bFound As Boolean
Dim hButton As Long
Dim sName As String

Public Sub ActivateWindow(h As Long)
 If h Then
    If IsIconic(h) Then
        Call ShowWindow(h, SW_RESTORE)
    End If
    Call SetForegroundWindow(h)
 Else
    Exit Sub
 End If
End Sub

Function EnumWinProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
  If bFound Then
     EnumWinProc = 0
     Exit Function
  Else
     Call EnumChildWindows(hWnd, AddressOf EnumChildWinProc, 0)
     EnumWinProc = 1
  End If
End Function

Function EnumChildWinProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
  If GetWndText(hWnd) = sName Then
     EnumChildWinProc = 0
     bFound = True
     hButton = hWnd
     Exit Function
  Else
     EnumChildWinProc = 1
  End If
End Function

Private Function GetWndText(hWnd As Long) As String
  Dim k As Long, sName As String
  sName = Space$(128)
  k = GetWindowText(hWnd, sName, 128)
  If k > 0 Then sName = Left$(sName, k) Else sName = "No name"
  GetWndText = sName
End Function

Public Function GetButtonHandle(ByVal sCaption As String) As Long
  sName = sCaption
  Call EnumWindows(AddressOf EnumWinProc, 0)
  GetButtonHandle = hButton
End Function

Public Function GetName(fzpath)
On Error Resume Next
Dim fSlash As Long
Dim fStart As Long
fSlash = 1
Do Until fSlash = 0
fStart = fSlash
fSlash = fSlash + 1
fSlash = InStr(fSlash, fzpath, "\", vbTextCompare)
Loop
GetName = Right(fzpath, (Len(fzpath) - fStart))
End Function

Public Function GetPathFrom(nPath)
GetPathFrom = Left(nPath, Len(nPath) - Len(GetName(nPath)))
End Function

' Required for LaVolpe button
Public Function lv_TimerCallBack(ByVal hWnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tgtButton As Butt
CopyMemory tgtButton, GetProp(hWnd, "lv_ClassID"), &H4
Call tgtButton.TimerUpdate(GetProp(hWnd, "lv_TimerID"))  ' fire the button's event
CopyMemory tgtButton, 0&, &H4                            ' erase this instance
End Function

