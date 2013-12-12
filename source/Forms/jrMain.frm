VERSION 5.00
Object = "{F924C9A7-D9B7-4808-8A32-108A70944450}#1.0#0"; "HookMenu.ocx"
Begin VB.MDIForm jrMain 
   BackColor       =   &H004D483F&
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11025
   Icon            =   "jrMain.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ToolbarContainer 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   11025
      TabIndex        =   0
      Top             =   6810
      Visible         =   0   'False
      Width           =   11025
      Begin VB.Label FPath 
         Caption         =   "current file name"
         Height          =   435
         Left            =   2505
         TabIndex        =   2
         Top             =   105
         Width           =   1470
      End
      Begin VB.Label path 
         Caption         =   "current file path"
         Height          =   435
         Left            =   75
         TabIndex        =   1
         Top             =   105
         Width           =   2160
      End
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   75
      Top             =   225
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   14
      Bmp:1           =   "jrMain.frx":1CCA
      Mask:1          =   16777215
      Key:1           =   "#mnu_FileOpen"
      Bmp:2           =   "jrMain.frx":1DC4
      Mask:2          =   12632256
      Key:2           =   "#mnu_FileSave"
      Bmp:3           =   "jrMain.frx":2306
      Key:3           =   "#mnu_File_Exit"
      Bmp:4           =   "jrMain.frx":306E
      Mask:4          =   526344
      Key:4           =   "#mnu_Tools_NiMP"
      Bmp:5           =   "jrMain.frx":3496
      Key:5           =   "#mnu_Help_HelpContents"
      Bmp:6           =   "jrMain.frx":38BE
      Mask:6          =   16711935
      Key:6           =   "#mnu_File_Preferences"
      Bmp:7           =   "jrMain.frx":39D0
      Mask:7          =   16711935
      Key:7           =   "#mnu_Help_About"
      Bmp:8           =   "jrMain.frx":3D22
      Mask:8          =   16711935
      Key:8           =   "#mnu_Help_UsedDll"
      Bmp:9           =   "jrMain.frx":4074
      Mask:9          =   16711935
      Key:9           =   "#mnu_Tools_MSNotepad"
      Bmp:10          =   "jrMain.frx":43C6
      Mask:10         =   16711935
      Key:10          =   "#mnu_Tools_MSPaint"
      Bmp:11          =   "jrMain.frx":4718
      Key:11          =   "#mnu_Tools_MSRegistry"
      Bmp:12          =   "jrMain.frx":4A6A
      Mask:12         =   15065571
      Key:12          =   "#mnu_Tools_MSCalculator"
      Bmp:13          =   "jrMain.frx":4DBC
      Key:13          =   "#mnu_Tools_BIK"
      Bmp:14          =   "jrMain.frx":51E4
      Key:14          =   "#mnu_Tools_VP6"
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
   Begin VB.Menu mnu_File 
      Caption         =   "File"
      Begin VB.Menu mnu_FileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_FileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_File_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Preferences 
         Caption         =   "Preferences"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu_File_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Exit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnu_Tools 
      Caption         =   "Tools"
      Begin VB.Menu mnu_Tools_Add 
         Caption         =   "Add New"
      End
      Begin VB.Menu mnu_Tools_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Tools_VP6 
         Caption         =   "VP6 Player 1.00"
      End
      Begin VB.Menu mnu_Tools_Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Tools_MSNotepad 
         Caption         =   "Microsoft Notepad"
      End
      Begin VB.Menu mnu_Tools_MSPaint 
         Caption         =   "Microsoft Paint"
      End
      Begin VB.Menu mnu_Tools_MSRegistry 
         Caption         =   "Microsoft Registry"
      End
      Begin VB.Menu mnu_Tools_MSCalculator 
         Caption         =   "Microsoft Calculator"
      End
      Begin VB.Menu mnu_Tools_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Tools_NiMP 
         Caption         =   "NiMP 2.7.7"
      End
   End
   Begin VB.Menu mnu_Modding 
      Caption         =   "Modding Tools"
      Begin VB.Menu mnu_Modding_NFSMW 
         Caption         =   "NFS Most Wanted"
         Begin VB.Menu mnu_Modding_NFSMW_MenuTweak 
            Caption         =   "Menus Tweak 3.00"
         End
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "Help"
      Begin VB.Menu mnu_Help_HelpContents 
         Caption         =   "Help Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnu_Help_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Help_VersionHistory 
         Caption         =   "Version History"
      End
      Begin VB.Menu mnu_Help_UsedDll 
         Caption         =   "Used DLL/OCX"
      End
      Begin VB.Menu mnu_Help_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Help_About 
         Caption         =   "About"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "jrMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
Me.Caption = "jRipper " & IAPPV
End Sub

Private Sub mnu_File_Exit_Click()
Unload Me
End Sub

Public Function OpenFN()
' Dims
Dim myFilters As String
Dim buff As String
Dim bpath As String
Dim f
Dim NOFilters As String: NOFilters = ReadINI(App.path & "\bin\jr.ini", "Filters", "NumberOfFilters")
Dim filterformat As String: filterformat = ReadINI(App.path & "\bin\jr.ini", "Filters", "ShowExtension")
filterformat = LCase(filterformat)
Dim fStore As String
Dim fname As String
Dim fExt As String
Dim tmp1 As String
Dim tmp2 As String
' Load Filters
For f = 1 To NOFilters
fStore = ReadINI(App.path & "\bin\jr.ini", "Filters", f)
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
path.Caption = OFN.sFile
FPath.Caption = buff
a.LOADFILE (buff)
End With
End Function

Private Sub mnu_FileOpen_Click()
OpenFN
End Sub

Private Sub mnu_Help_About_Click()
AboutJR.Show
End Sub

Private Sub mnu_Help_HelpContents_Click()
MsgBox "No Help Contents yet :(", vbInformation, "Sorry"
End Sub

Private Sub mnu_Modding_NFSMW_MenuTweak_Click()
NFSMW_MT.Show
End Sub


Private Sub mnu_Tools_MSCalculator_Click()
Shell (Environ("windir") & "\system32\calc.exe"), vbNormalFocus
End Sub

Private Sub mnu_Tools_MSNotepad_Click()
Shell (Environ("windir") & "\notepad.exe"), vbNormalFocus
End Sub

Private Sub mnu_Tools_MSPaint_Click()
Shell (Environ("windir") & "\system32\mspaint.exe"), vbNormalFocus
End Sub

Private Sub mnu_Tools_MSRegistry_Click()
Shell (Environ("windir") & "\regedit.exe"), vbNormalFocus
End Sub

Private Sub mnu_Tools_NiMP_Click()
NIMP.Show
End Sub

Private Sub mnu_Tools_VP6_Click()
VP6_Playa.Show
End Sub
