VERSION 5.00
Object = "{F924C9A7-D9B7-4808-8A32-108A70944450}#1.0#0"; "HookMenu.ocx"
Begin VB.Form VP6_Playa 
   BackColor       =   &H004D483F&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VP6 Player 1.00"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   7530
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VP6_Player.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H004D483F&
      Height          =   1260
      Left            =   75
      TabIndex        =   5
      Top             =   30
      Width           =   7365
      Begin VB.Image Image2 
         Height          =   1095
         Left            =   45
         Picture         =   "VP6_Player.frx":F4D6
         Stretch         =   -1  'True
         Top             =   135
         Width           =   1110
      End
      Begin VB.Label About 
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BDB8AF&
         Height          =   1005
         Left            =   1245
         TabIndex        =   6
         Top             =   180
         Width           =   6030
      End
   End
   Begin jR_RC2.Butt WatchVP6 
      Height          =   390
      Left            =   6090
      TabIndex        =   2
      Top             =   1995
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   688
      Caption         =   "[2] WATCH"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin jR_RC2.Butt Convert 
      Height          =   390
      Left            =   4695
      TabIndex        =   3
      Top             =   1995
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   688
      Caption         =   "[1] CONVERT"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   375
      Top             =   570
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   12
      Bmp:1           =   "VP6_Player.frx":19F88
      Mask:1          =   16777215
      Key:1           =   "#mnu_FileOpen"
      Bmp:2           =   "VP6_Player.frx":1A082
      Mask:2          =   12632256
      Key:2           =   "#mnu_FileSave"
      Bmp:3           =   "VP6_Player.frx":1A5C4
      Key:3           =   "#mnu_File_Exit"
      Bmp:4           =   "VP6_Player.frx":1B32C
      Mask:4          =   526344
      Key:4           =   "#mnu_Tools_NiMP"
      Bmp:5           =   "VP6_Player.frx":1B754
      Key:5           =   "#mnu_Help_HelpContents"
      Bmp:6           =   "VP6_Player.frx":1BB7C
      Mask:6          =   16711935
      Key:6           =   "#mnu_File_Preferences"
      Bmp:7           =   "VP6_Player.frx":1BC8E
      Mask:7          =   16711935
      Key:7           =   "#mnu_Help_About"
      Bmp:8           =   "VP6_Player.frx":1BFE0
      Mask:8          =   16711935
      Key:8           =   "#mnu_Help_UsedDll"
      Bmp:9           =   "VP6_Player.frx":1C332
      Mask:9          =   16711935
      Key:9           =   "#mnu_Tools_MSNotepad"
      Bmp:10          =   "VP6_Player.frx":1C684
      Mask:10         =   16711935
      Key:10          =   "#mnu_Tools_MSPaint"
      Bmp:11          =   "VP6_Player.frx":1C9D6
      Key:11          =   "#mnu_Tools_MSRegistry"
      Bmp:12          =   "VP6_Player.frx":1CD28
      Mask:12         =   15065571
      Key:12          =   "#mnu_Tools_MSCalculator"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label fname 
      BackStyle       =   0  'Transparent
      Caption         =   "File Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BDB8AF&
      Height          =   270
      Left            =   390
      TabIndex        =   4
      Top             =   2490
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label switches 
      Height          =   1245
      Left            =   270
      TabIndex        =   1
      Top             =   2610
      Width           =   4155
   End
   Begin VB.Image Image1 
      Height          =   45
      Left            =   -60
      Picture         =   "VP6_Player.frx":1D07A
      Stretch         =   -1  'True
      Top             =   1845
      Width           =   7605
   End
   Begin VB.Label Path 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BDB8AF&
      Height          =   420
      Left            =   75
      TabIndex        =   0
      Top             =   1365
      Width           =   7365
   End
   Begin VB.Menu mnu_File 
      Caption         =   "File"
      Begin VB.Menu mnu_FileConvert 
         Caption         =   "Convert a File and Watch"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnu_FileConvertAll 
         Caption         =   "Convert all Files in Folder"
         Shortcut        =   ^{F2}
      End
   End
End
Attribute VB_Name = "VP6_Playa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fso As New FileSystemObject

Function convertFile()
' SCRIPT
switches.Caption = _
"@echo off" & vbNewLine & _
"TITLE " & Chr(34) & "jRipper - VP6 Convertor" & Chr(34) & vbNewLine & _
"echo ***********************************" & vbNewLine & _
"echo jRipper - VP6 Convertor" & vbNewLine & _
"echo ***********************************" & vbNewLine & _
"echo." & vbNewLine & _
"for %%f in (" & Me.fname.Caption & ") do (" & vbNewLine & _
"echo *Convert %%f " & vbNewLine & _
"start /wait eaconv.exe %%f" & vbNewLine & _
"del eaconv.exe" & vbNewLine & _
"del eaconv.cmd" & vbNewLine & _
")"
' Paths
Dim export As String: export = Replace(Path.Caption, fname.Caption, "eaconv.cmd")
Dim expor2 As String: expor2 = Replace(Path.Caption, fname.Caption, "eaconv.exe")
' Create script
fso.CreateTextFile export, True
TXT.WriteTextFile export, switches.Caption
' Copy convertor
fso.CopyFile App.Path & "\bin\eaconv.exe", expor2, True
' Launch Script
Shell (export), vbNormalFocus
' Enable WATCH button
WatchVP6.Enabled = True
Convert.Enabled = False
End Function

Function convertFolder()
' SCRIPT
switches.Caption = _
"@echo off" & vbNewLine & _
"TITLE " & Chr(34) & "jRipper - VP6 Convertor" & Chr(34) & vbNewLine & _
"echo ***********************************" & vbNewLine & _
"echo jRipper - VP6 Convertor" & vbNewLine & _
"echo ***********************************" & vbNewLine & _
"echo." & vbNewLine & _
"for %%f in (*.vp6) do (" & vbNewLine & _
"echo *Convert %%f " & vbNewLine & _
"start /wait eaconv.exe %%f" & vbNewLine & _
")"
' Paths
Dim export As String: export = Replace(Path.Caption, fname.Caption, "eaconv.cmd")
Dim expor2 As String: expor2 = Replace(Path.Caption, fname.Caption, "eaconv.exe")
' Create script
fso.CreateTextFile export, True
TXT.WriteTextFile export, switches.Caption
' Copy convertor
fso.CopyFile App.Path & "\bin\eaconv.exe", expor2, True
' Launch Script
Shell (export), vbNormalFocus
' Enable WATCH button
WatchVP6.Enabled = True
Convert.Enabled = False
End Function

Public Function OpenFN()
On Error GoTo errx
' Dims
Dim myFilters As String
Dim buff As String
Dim bpath As String
Dim f
Dim NOFilters As String: NOFilters = ReadINI(App.Path & "\bin\jr.ini", "VP6Player", "NumberOfFilters")
Dim filterformat As String: filterformat = ReadINI(App.Path & "\bin\jr.ini", "VP6Player", "ShowExtension")
filterformat = LCase(filterformat)
Dim fStore As String
Dim fname As String
Dim fExt As String
Dim tmp1 As String
Dim tmp2 As String
' Load Filters
For f = 1 To NOFilters
fStore = ReadINI(App.Path & "\bin\jr.ini", "VP6Player", f)
fname = Split(fStore, "$")(0)
fExt = Split(fStore, "$")(1)
If filterformat = "true" Then
myFilters = myFilters & fname & "(" & fExt & ")" & vbNullChar & fExt & vbNullChar
Else
myFilters = myFilters & fname & vbNullChar & fExt & vbNullChar
End If
Next f
myFilters = myFilters & vbNullChar & vbNullChar
' Show Dialog
With OFN
   .nStructSize = Len(OFN)
   .hWndOwner = jrMain.hWnd
   .sFilter = myFilters
   .nFilterIndex = 1
   .sFile = GetName(.sFileTitle) & Space$(1024) & vbNullChar & vbNullChar
   .nMaxFile = Len(.sFile)
   .sDefFileExt = "###" & vbNullChar & vbNullChar
   .sFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
   .nMaxTitle = Len(OFN.sFileTitle)
   .sInitialDir = GetPathFrom(bpath) & vbNullChar & vbNullChar
   .sDialogTitle = "Select a file"
   .flags = OFS_FILE_OPEN_FLAGS Or OFN_HIDEREADONLY
    If GetOpenFileName(OFN) Then
    buff = Replace(OFN.sFileTitle, vbNullChar, "")
    buff = Trim(buff)
    End If
' LoadFile
Me.Path.Caption = OFN.sFile
Me.fname.Caption = buff
End With
errx:
Exit Function
End Function

'=============================================================================================

Private Sub Convert_Click()
convertFile
End Sub
Private Sub Form_Load()
About.Caption = "Handling .VP6 files requires 3rd Party Tools" & vbNewLine & "Convertor : EA Electronics" & vbNewLine & "Codec : ON2"
End Sub

Private Sub mnu_FileConvert_Click()
OpenFN
convertFile
Shell ("C:\Program Files\Windows Media Player\wmplayer.exe " & Chr(34) & Replace(Me.Path.Caption, ".vp6", ".avi") & Chr(34)), vbMaximizedFocus
Unload Me
End Sub

Private Sub mnu_FileConvertAll_Click()
OpenFN
convertFolder
Shell (Environ("windir") & "\explorer.exe " & Chr(34) & Replace(Path.Caption, "\" & fname.Caption, "") & Chr(34)), vbNormalFocus
End Sub

Private Sub WatchVP6_Click()
Shell ("C:\Program Files\Windows Media Player\wmplayer.exe " & Chr(34) & Replace(Me.Path.Caption, ".vp6", ".avi") & Chr(34)), vbMaximizedFocus
Unload Me
End Sub
