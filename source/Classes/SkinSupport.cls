VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SkinSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lpPt As POINTAPI) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Const STRETCH_HALFTONE  As Long = &H4&

Private Type POINTAPI
    x  As Long
    Y  As Long
End Type
Dim picW As Long
Dim picH As Long


Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Const RGN_OR = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
        ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long


Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

    
Private Function Make_Region(picSkin As PictureBox, Red, Green, Blue) As Long

    Dim x As Long, Y As Long, StartLineX As Long
    Dim LineRegion As Long
    Dim TransparentColor As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean
    Dim hdc As Long
    Dim PicWidth As Long
    Dim PicHeight As Long
    
    hdc = picSkin.hdc
    PicWidth = picSkin.ScaleWidth
    PicHeight = picSkin.ScaleHeight
    
    InFirstRegion = True: InLine = False
    x = Y = StartLineX = 0

     TransparentColor = RGB(Red, Green, Blue)
    
    For Y = 0 To PicHeight - 1
        For x = 0 To PicWidth - 1

            If GetPixel(hdc, x, Y) = TransparentColor Or x = PicWidth Then

                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, Y, x, Y + 1)
                    
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR

                        DeleteObject LineRegion
                    End If
                End If
            Else

                If Not InLine Then
                    InLine = True
                    StartLineX = x
                End If
            End If
        Next
    Next
     Make_Region = FullRegion
End Function

Function Set_Background(Form As Form, Picture As PictureBox, Optional UsingTransparentColor As Boolean, Optional TransparentColor1RED = "255", Optional TransparentColor2GREEN = "0", Optional TransparentColor3BLUE = "255")
    If UsingTransparentColor = True Then
    RegionMini = Make_Region(Picture, TransparentColor1RED, TransparentColor2GREEN, TransparentColor3BLUE)
    SetWindowRgn Form.hWnd, RegionMini, True
    Picture.Top = 0
    Picture.Left = 0
    Form.Width = Picture.Width
    Form.Height = Picture.Height
    Else
    Form.Picture = Picture.Picture
    Picture.AutoSize = True
    Form.Width = Picture.Width
    Form.Height = Picture.Height
    Picture.Top = 0
    Picture.Left = 0
    End If
End Function

Function Drag_Form(MouseButton As MouseButtonConstants, Form As Form)

On Local Error Resume Next
    
    If MouseButton = vbLeftButton Then
        Call ReleaseCapture
        Call SendMessage(Form.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
    If MouseButton = vbMiddleButton Then
        Call ReleaseCapture
        Call SendMessage(Form.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
    If MouseButton = vbRightButton Then
        Call ReleaseCapture
        Call SendMessage(Form.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Function

Function Debug_WriteEXEmanifest(AppPath, AppEXEName, AppMajor, AppMinor, AppRevision, AppFileDescription, Optional AutoRestart As Boolean = True, Optional AutoRestartWindowStyle As VbAppWinStyle = vbNormalFocus, Optional CreateNew As Boolean) As Boolean

 InitCommonControls

On Error Resume Next
 Dim XML             As String
 Dim ManifestCheck   As String
 Dim strManifest     As String
 Dim FreeFileNo      As Integer

If AutoRestart = True Then CreateNew = False

XML = ("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?> " & vbCrLf & _
  "<assembly " & vbCrLf & "   xmlns=""urn:schemas-microsoft-com:asm.v1"" " & vbCrLf & _
  "   manifestVersion=""1.0"">" & vbCrLf & "<assemblyIdentity " & vbCrLf & _
  "    processorArchitecture=""x86"" " & vbCrLf & _
  "    version=""EXEVERSION""" & vbCrLf & "    type=""win32""" & vbCrLf & _
  "    name=""EXENAME""/>" & vbCrLf & _
  "    <description>EXEDESCRIBTION</description>" & vbCrLf & _
  "    <dependency>" & vbCrLf & "    <dependentAssembly>" & vbCrLf & _
  "    <assemblyIdentity" & vbCrLf & "         type=""win32""" & vbCrLf & _
  "         name=""Microsoft.Windows.Common-Controls""" & vbCrLf & _
  "         version=""6.0.0.0""" & vbCrLf & _
  "         publicKeyToken=""6595b64144ccf1df""" & vbCrLf & _
  "         language=""*""" & vbCrLf & _
  "         processorArchitecture=""x86""/>" & vbCrLf & _
  "    </dependentAssembly>" & vbCrLf & "    </dependency>" & vbCrLf & _
  "</assembly>" & vbCrLf & "")


strManifest = App.Path & "\" & AppEXEName & ".exe.manifest"


ManifestCheck = VBA.dir(strManifest, vbNormal + vbSystem + vbHidden + vbReadOnly + vbArchive)


If ManifestCheck = "" Or CreateNew = True Then

  XML = Replace(XML, "EXENAME", AppEXEName & ".exe")

  XML = Replace(XML, "EXEVERSION", AppMajor & "." & AppMinor & "." & AppRevision & ".0")

  XML = Replace(XML, "EXEDESCRIBTION", AppFileDescription)
  
  FreeFileNo = FreeFile
  

  If ManifestCheck <> "" Then
    SetAttr strManifest, vbNormal
    Kill strManifest
  End If
  

  Open strManifest For Binary As #(FreeFileNo)

     Put #(FreeFileNo), , XML
  Close #(FreeFileNo)
  

  SetAttr strManifest, vbHidden + vbSystem
  
  If ManifestCheck = "" Then
    Debug_WriteEXEmanifest = False
  Else
    Debug_WriteEXEmanifest = True
  End If
  

  If AutoRestart = True Then

    Shell AppPath & "\" & AppEXEName & ".exe", AutoRestartWindowStyle
    End
  End If
  
Else
  Debug_WriteEXEmanifest = True
End If

End Function

Function Check_isTransparent(ByVal hWnd As Long) As Boolean

On Error Resume Next
Dim Msg As Long
    Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
    If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
        Check_isTransparent = True
    Else
        Check_isTransparent = False
    End If
If Err Then
    Check_isTransparent = False
    End If
End Function

Function Make_Transparent(ByVal hWnd As Long, Percentage As Integer) As Long
Dim Msg As Long
On Error Resume Next
If Percentage < 0 Or Percentage > 255 Then
    Make_Transparent = 1
Else
    Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
    Msg = Msg Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, Msg
    SetLayeredWindowAttributes hWnd, 0, Percentage, LWA_ALPHA
    Make_Transparent = 0
End If
If Err Then
    Make_Transparent = 2
End If
End Function

Function Make_Opaque(ByVal hWnd As Long) As Long
Dim Msg As Long
On Error Resume Next
    Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
    Msg = Msg And Not WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, Msg
    SetLayeredWindowAttributes hWnd, 0, 0, LWA_ALPHA
    Make_Opaque = 0
If Err Then
    Make_Opaque = 2
End If
End Function

