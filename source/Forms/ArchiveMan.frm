VERSION 5.00
Object = "{F924C9A7-D9B7-4808-8A32-108A70944450}#1.0#0"; "HOOKMENU.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ArchiveMan 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Archive Manager 1.00"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ArchiveMan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   8910
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   210
      Top             =   225
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   4
      Bmp:1           =   "ArchiveMan.frx":08CA
      Mask:1          =   12632256
      Key:1           =   "#mnu_OPEN"
      Bmp:2           =   "ArchiveMan.frx":0E0C
      Key:2           =   "#mnu_EXIT"
      Bmp:3           =   "ArchiveMan.frx":1B74
      Key:3           =   "#mnu_HELP"
      Bmp:4           =   "ArchiveMan.frx":1F9C
      Mask:4          =   16711935
      Key:4           =   "#mnu_ABOUT"
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
   Begin MSComctlLib.ProgressBar Pb1 
      Height          =   300
      Left            =   165
      TabIndex        =   2
      Top             =   8190
      Visible         =   0   'False
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00ECECEC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   8220
      ScaleHeight     =   330
      ScaleWidth      =   540
      TabIndex        =   3
      Top             =   8175
      Width           =   540
      Begin VB.Label OpenPathText 
         BackStyle       =   0  'Transparent
         Caption         =   "Open"
         Height          =   195
         Left            =   45
         TabIndex        =   4
         Top             =   45
         Width           =   420
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6270
      Top             =   7425
   End
   Begin VB.CheckBox selall 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   195
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Check/Uncheck all"
      Top             =   30
      Width           =   195
   End
   Begin MSComctlLib.ListView Port 
      Height          =   7245
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   12779
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "typelist"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "col1"
         Text            =   "       Filename"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "col2"
         Text            =   "Size (bytes)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "col3"
         Text            =   "Offsets"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Type"
         Object.Width           =   1500
      EndProperty
   End
   Begin MSComctlLib.ImageList typelist2 
      Left            =   5595
      Top             =   7425
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":22EE
            Key             =   "rtcwsbwl"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":9130
            Key             =   "tr3tla"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":944A
            Key             =   "tr3aolc"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":9764
            Key             =   "s2ttb"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":9A7E
            Key             =   "c2nw"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":A6D0
            Key             =   "toca"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":AFAA
            Key             =   "cmr04"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":B884
            Key             =   "nfo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":C15E
            Key             =   "setts"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":DE38
            Key             =   "g17"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":E152
            Key             =   "carmageddon"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":FE2C
            Key             =   "carmageddon2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList typelist 
      Left            =   5010
      Top             =   7410
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":10706
            Key             =   "big"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":10A5A
            Key             =   "arc"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":1320C
            Key             =   "image"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":13560
            Key             =   "lng"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":138B4
            Key             =   "ini"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":13C08
            Key             =   "txt"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":13F5C
            Key             =   "p3d"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":142B0
            Key             =   "wav"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":14604
            Key             =   "unknown2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":14958
            Key             =   "pal"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":14CAC
            Key             =   "3d"
            Object.Tag             =   "3d"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":150B7
            Key             =   "3ds"
            Object.Tag             =   "3ds"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":151B1
            Key             =   "unknown"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":155DA
            Key             =   "bik"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":1719C
            Key             =   "tools"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":175EE
            Key             =   "s2ttb"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":17B88
            Key             =   "prisont"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":1A33A
            Key             =   "carma2"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":1AC14
            Key             =   "toca"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":1B4EE
            Key             =   "cossacks2"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":1BDC8
            Key             =   "cmr4"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":1C6A2
            Key             =   "g17"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":1C9BC
            Key             =   "rtcwsbwl"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":237FE
            Key             =   "tr3g"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ArchiveMan.frx":23B18
            Key             =   "nfo"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Path 
      Appearance      =   0  'Flat
      BackColor       =   &H00ECECEC&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2010
      TabIndex        =   5
      ToolTipText     =   "Adress of opened file. Enter a path and press OPEN to load that file."
      Top             =   8205
      Width           =   6195
   End
   Begin jR_RC2.Butt xa 
      Height          =   360
      Left            =   7650
      TabIndex        =   15
      Top             =   7725
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   635
      Caption         =   "Extract All"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   14737632
      Focus           =   0   'False
      LockHover       =   3
      cGradient       =   16777215
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   12632256
   End
   Begin jR_RC2.Butt x 
      Height          =   360
      Left            =   7650
      TabIndex        =   16
      Top             =   7335
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   635
      Caption         =   "Extract Sel"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   14737632
      Focus           =   0   'False
      LockHover       =   3
      cGradient       =   16777215
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   12632256
   End
   Begin VB.Label Label1 
      Caption         =   "hidden"
      Height          =   300
      Left            =   5550
      TabIndex        =   14
      Top             =   7575
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label info_LEN 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   180
      TabIndex        =   13
      Top             =   7260
      Width           =   2685
   End
   Begin VB.Label info_NOF 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   180
      TabIndex        =   12
      Top             =   7470
      Width           =   2685
   End
   Begin VB.Label info_MORE1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   180
      TabIndex        =   11
      Top             =   7650
      Width           =   2685
   End
   Begin VB.Label info_MORE2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   180
      TabIndex        =   10
      Top             =   7860
      Width           =   2685
   End
   Begin VB.Label info_MORE3 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3090
      TabIndex        =   9
      Top             =   7470
      Width           =   2685
   End
   Begin VB.Label info_MORE4 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3090
      TabIndex        =   8
      Top             =   7650
      Width           =   2685
   End
   Begin VB.Label info_MORE5 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3090
      TabIndex        =   7
      Top             =   7860
      Width           =   2685
   End
   Begin VB.Label info_LOAD 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   8220
      Width           =   1770
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image6 
      Height          =   8160
      Left            =   8775
      Picture         =   "ArchiveMan.frx":243F2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   8160
      Left            =   0
      Picture         =   "ArchiveMan.frx":24450
      Stretch         =   -1  'True
      Top             =   0
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   465
      Left            =   135
      Picture         =   "ArchiveMan.frx":244AE
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   8640
   End
   Begin VB.Image Image3 
      Height          =   465
      Left            =   8775
      Picture         =   "ArchiveMan.frx":2456C
      Top             =   8160
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   465
      Left            =   0
      Picture         =   "ArchiveMan.frx":24912
      Top             =   8160
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   135
      Picture         =   "ArchiveMan.frx":24CB8
      Stretch         =   -1  'True
      Top             =   7230
      Width           =   8640
   End
   Begin VB.Menu mnu_File 
      Caption         =   "File"
      Begin VB.Menu mnu_OPEN 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_SEP 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Extract 
         Caption         =   "Extract"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_EXTRACTALL 
         Caption         =   "Extract All"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu mnu_SEP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_EXIT 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnu_HELP 
      Caption         =   "Help"
      Begin VB.Menu mnu_ABOUT 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "ArchiveMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject

Private Sub mnu_File_Click()
ArchiveMan.mnu_EXTRACTALL.Enabled = xa.Enabled
ArchiveMan.mnu_File_Extract.Enabled = x.Enabled
End Sub

Private Sub mnu_OPEN_Click()
jrMain.OpenFN
End Sub

Private Sub X_Click()
a.EXTRACTFILE Me.Path.Text, False
End Sub

Private Sub xa_Click()
a.EXTRACTFILE Me.Path.Text, True
End Sub

'========================================================================
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
' Set Path textbox backcolor to default
Path.BackColor = &HECECEC
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
' Set Path textbox backcolor to default
Path.BackColor = &HECECEC
End Sub

Private Sub mnu_ABOUT_Click()
MsgBox "File Format Decoder : codin" & vbNewLine & "Application Code       : codin" & vbNewLine & "Application Interface: codin" & vbNewLine & vbNewLine & "Skin inspired by Windows Vista" & vbNewLine & vbNewLine & "Contact me at : gbcrk@yahoo.com", vbInformation, "About Application"
End Sub

Private Sub mnu_EXIT_Click()
Set frm_PCK = Nothing
Unload Me
End Sub

Private Sub mnu_EXTRACTALL_Click()
xa_Click
End Sub

Private Sub mnu_File_Extract_Click()
X_Click
End Sub

Private Sub Path_Click()
' When clicking adress(path) bar, color will be white
Path.BackColor = vbWhite
End Sub

Private Sub Path_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
' when moving mouse on adress(path) bar, color will be changed
Path.BackColor = &HF8F8F8
End Sub

Private Sub OpenPathText_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
OpenPathText.FontBold = True
End Sub

Private Sub OpenPathText_Click()
' Convert to lowercase (so it wont be casesensitive)
Dim myPath As String
myPath = StrConv(Path.Text, vbLowerCase)
' Open file entered in Path
LOADFILE Path.Text
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
' Reset others
OpenPathText.FontBold = False
' Set Path textbox backcolor to default
Path.BackColor = &HECECEC
End Sub

Private Sub selall_Click()
Me.Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
Dim x
If Me.selall.Value = 1 Then
    For x = 1 To Me.Port.ListItems.Count
        If Port.ListItems(x).Checked = False Then
        Port.ListItems(x).Checked = True
        End If
    Next x
ElseIf Me.selall.Value = 0 Then
    For x = 1 To Me.Port.ListItems.Count
        If Port.ListItems(x).Checked = True Then
        Port.ListItems(x).Checked = False
        End If
    Next x
End If
Timer2.Enabled = False
End Sub

Function addinfo(Optional More1, Optional More2, Optional More3, Optional More4, Optional More5)
On Error Resume Next
Dim filepath As String: filepath = ArchiveMan.Path.Text
info_LEN.Caption = "Archive Size: " & Format(FileLen(filepath), "###,###") & " bytes"
info_LEN.ToolTipText = Format(FileLen(filepath) / 1024 / 1024, "###,###") & " MB | " & Format(FileLen(filepath) / 1000 / 1000, "###,###") & " MiB | " & Format(FileLen(filepath) / 1024, "###,###") & " KB | " & Format(FileLen(filepath) / 1000, "###,###") & " KiB"
info_NOF.Caption = "Files: " & ArchiveMan.Port.ListItems.Count
info_MORE1.Caption = More1
info_MORE2.Caption = More2
info_MORE3.Caption = More3
info_MORE4.Caption = More4
info_MORE5.Caption = More5
End Function
