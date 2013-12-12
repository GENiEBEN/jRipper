Attribute VB_Name = "jRFct"
Option Explicit

Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_NOTOPMOST = -2
Public Const LB_ITEMFROMPOINT = &H1A9
Const SW_SHOWNORMAL = 1
Const STARTF_USESHOWWINDOW = &H1&

Public Declare Function Setwindowpos Lib "user32" Alias "SetWindowPos" (ByVal HWND As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWND As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const MAX_PATH = 256&
Private Declare Function GetLongPathName Lib "kernel32" Alias _
    "GetLongPathNameA" (ByVal lpszShortPath As String, _
    ByVal lpszLongPath As String, ByVal cchBuffer As Long) _
    As Long
Private Declare Function LoadLibrary Lib "kernel32" _
  Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" _
  (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" _
  (ByVal hLibModule As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal HWND As Long, ByVal lpString As String) As Long
Public Const SRCCOPY = &HCC0020
Private Const INVALID_HANDLE_VALUE = -1

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
   (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindClose Lib "kernel32" _
   (ByVal hFindFile As Long) As Long
Enum FlagConstants
  cdlOFNAllowMultiselect = &H200 'The user can select more than one file atrun time by pressing the SHIFT key and using the UP ARROW and DOWN ARROW keys to select the desired files. When this is done, the FileName property returns a string containing the names of all selected files. The names in the string are delimited by spaces.
  cdlOFNCreatePrompt = &H2000 ' Specifies that the dialog box prompts the user to create a file that doesn't currently exist. This flag automatically sets the cdlOFNPathMustExist and cdlOFNFileMustExist flags.
  cdlOFNExplorer = &H80000 ' Use the Explorer-like Open A File dialog box template. Works with Windows 95 and Windows NT 4.0.
  cdlOFNExtensionDifferent = &H400 ' Indicates that the extension of the returned filename is different from the extension specified by the DefaultExt property. This flag isn't set if the DefaultExt property is Null, if the extensions match, or if the file has no extension. This flag value can be checked upon closing the dialog box.
  cdlOFNFileMustExist = &H1000 ' Specifies that the user can enter only names of existing files in the File Name text box. If this flag is set and the user enters an invalid filename, a warning is displayed. This flag automatically sets the cdlOFNPathMustExist flag.
  cdlOFNHelpButton = &H10 ' Causes the dialog box to display the Help button.
  cdlOFNHideReadOnly = &H4 'Hides the Read Onlycheck box.
  cdlOFNLongNames = &H200000 ' Use long filenames.
  cdlOFNNoChangeDir = &H8 'Forces the dialog box to set the current directory to what it was when the dialog box was opened.
  cdlOFNNoDereferenceLinks = &H100000 ' Do not dereference shell links (also known as shortcuts). By default, choosing a shell link causes it to be dereferenced by the shell.
  cdlOFNNoLongNames = &H40000 ' No long file names.
  cdlOFNNoReadOnlyReturn = &H8000 ' Specifies that the returned file won't have the Read Only attribute set and won't be in a write-protected directory.
  cdlOFNNoValidate = &H100 ' Specifies that the common dialog box allows invalid characters in the returned filename.
  cdlOFNOverwritePrompt = &H2 'Causes the Save As dialog box to generate a message box if the selected file already exists. The user must confirm whether to overwrite the file.
  cdlOFNPathMustExist = &H800 ' Specifies that the user can enter only valid paths. If this flag is set and the user enters an invalid path, a warning message is displayed.
  cdlOFNReadOnly = &H1 'Causes the Read Only check box to be initially checked when the dialog box is created. This flag also indicates the state of the Read Only check box when the dialog box is closed.
  cdlOFNShareAware = &H4000 ' Specifies that sharing violation errors will be ignored.
End Enum
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

'=> Convert 8.3 to LongFileName (255chr)
Public Function GetLongFileName(ByVal FullPath As String) _
   As String

'*****************************************
'USAGE: Convert short (8.3) file name to long file name
'INPUT: FULL PATH OF A SHORT FILE NAME
'RETURNS: LONG FILE NAME:
'EXAMPLE: dim sLongFile as String
'         sLongFile = GetLongFileName("C:\MyShor~1.txt")
'NOTES: ONLY WORKS ON WIN 98 and WIN 2000.  WILL RETURN
'       EMPTY STRING ELSEWHERE
'***********************************************************

    Dim lLen As Long
    Dim sBuffer As String
    
    'Function only available on 98 and 2000/XP
    'so we check to see if it's available before proceeding
    
    If Not APIFunctionPresent("GetLongPathNameA", "kernel32") _
       Then Exit Function
    
    sBuffer = String$(MAX_PATH, 0)
    lLen = GetLongPathName(FullPath, sBuffer, Len(sBuffer))
    If lLen > 0 And Err.Number = 0 Then
        GetLongFileName = Left$(sBuffer, lLen)
    End If
End Function

Private Function APIFunctionPresent(ByVal FunctionName _
   As String, ByVal DllName As String) As Boolean

   'http://www.freevbcode.com/ShowCode.Asp?ID=429

    Dim lHandle As Long
    Dim lAddr  As Long

    lHandle = LoadLibrary(DllName)
    If lHandle <> 0 Then
        lAddr = GetProcAddress(lHandle, FunctionName)
        FreeLibrary lHandle
    End If
    
    APIFunctionPresent = (lAddr <> 0)

End Function

