VERSION 5.00
Begin VB.Form BIK_Playa 
   BackColor       =   &H004D483F&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BIK Player 1.00"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BIK_Player.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H004D483F&
      Height          =   1260
      Left            =   75
      TabIndex        =   13
      Top             =   30
      Width           =   5010
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
         TabIndex        =   14
         Top             =   180
         Width           =   3450
      End
      Begin VB.Image Image2 
         Height          =   1095
         Left            =   45
         Picture         =   "BIK_Player.frx":1BB2
         Top             =   135
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H004D483F&
      Height          =   3240
      Left            =   75
      TabIndex        =   3
      Top             =   1245
      Width           =   5010
      Begin VB.CheckBox bik_showstats2 
         BackColor       =   &H004D483F&
         Caption         =   "Show the PlayBack summary when finished"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   200
         TabIndex        =   12
         Top             =   2745
         Width           =   4380
      End
      Begin VB.CheckBox bik_noMT 
         BackColor       =   &H004D483F&
         Caption         =   "Don't use multi-threaded device reading"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   200
         TabIndex        =   11
         Top             =   2445
         Width           =   3915
      End
      Begin VB.CheckBox bik_minimizeJR 
         BackColor       =   &H004D483F&
         Caption         =   "Minimize jR after starting a movie"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   200
         TabIndex        =   10
         Top             =   345
         Width           =   3915
      End
      Begin VB.CheckBox bik_focuslost 
         BackColor       =   &H004D483F&
         Caption         =   "Don't pause movie when focus is lost"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   200
         TabIndex        =   9
         Top             =   2145
         Width           =   3915
      End
      Begin VB.CheckBox bik_blackbg 
         BackColor       =   &H004D483F&
         Caption         =   "Fill the PlayBack Window with a black background"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   200
         TabIndex        =   8
         Top             =   1845
         Width           =   4590
      End
      Begin VB.CheckBox bik_hidecursor 
         BackColor       =   &H004D483F&
         Caption         =   "Hide Cursor in PlayBack Window"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   200
         TabIndex        =   7
         Top             =   1560
         Width           =   3105
      End
      Begin VB.CheckBox bik_dontskip 
         BackColor       =   &H004D483F&
         Caption         =   "Never skip frames when falling behind"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   200
         TabIndex        =   6
         Top             =   1260
         Width           =   3930
      End
      Begin VB.CheckBox bik_runtimestats 
         BackColor       =   &H004D483F&
         Caption         =   "Show Run-time Playback Statistics"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   200
         TabIndex        =   5
         Top             =   945
         Width           =   3840
      End
      Begin VB.CheckBox bik_preload 
         BackColor       =   &H004D483F&
         Caption         =   "Preload Entire Video File into Memory"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   200
         TabIndex        =   4
         Top             =   645
         Width           =   3945
      End
   End
   Begin VB.CheckBox always 
      BackColor       =   &H004D483F&
      Caption         =   "Always use this settings"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5100
      Width           =   2415
   End
   Begin jR_RC2.Butt WatchBIK 
      Height          =   390
      Left            =   3675
      TabIndex        =   2
      Top             =   5025
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   688
      Caption         =   "Watch Movie"
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
   Begin VB.Label switches 
      Height          =   300
      Left            =   270
      TabIndex        =   15
      Top             =   5670
      Width           =   4155
   End
   Begin VB.Image Image1 
      Height          =   45
      Left            =   -15
      Picture         =   "BIK_Player.frx":5BD4
      Stretch         =   -1  'True
      Top             =   4875
      Width           =   5235
   End
   Begin VB.Label Path 
      BackStyle       =   0  'Transparent
      Caption         =   "Path"
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
      Left            =   90
      TabIndex        =   0
      Top             =   4560
      Width           =   5025
   End
End
Attribute VB_Name = "BIK_Playa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub always_Click()
Dim pathx As String: pathx = App.path & "\bin\jr.ini"
If always.Value = 1 Then
INI.AddINI pathx, "BIKPlayer", "SFilter", bik_minimizeJR.Value
INI.AddINI pathx, "BIKPlayer", "Filterx", Me.switches.Caption
INI.AddINI pathx, "BIKPlayer", "Always", always.Value
INI.AddINI pathx, "BIKPlayer", "preload", bik_preload.Value
INI.AddINI pathx, "BIKPlayer", "runtimestats", bik_runtimestats.Value
INI.AddINI pathx, "BIKPlayer", "dontskip", bik_dontskip.Value
INI.AddINI pathx, "BIKPlayer", "hidecursor", bik_hidecursor.Value
INI.AddINI pathx, "BIKPlayer", "blackbg", bik_blackbg.Value
INI.AddINI pathx, "BIKPlayer", "focuslost", bik_focuslost.Value
INI.AddINI pathx, "BIKPlayer", "noMT", bik_noMT.Value
INI.AddINI pathx, "BIKPlayer", "showstats2", bik_showstats2.Value
Else
bik_minimizeJR.Value = 0
Me.switches.Caption = ""
bik_preload.Value = 0
bik_runtimestats.Value = 0
bik_dontskip.Value = 0
bik_hidecursor.Value = 0
bik_blackbg.Value = 0
bik_focuslost.Value = 0
bik_noMT.Value = 0
bik_showstats2.Value = 0
End If
End Sub

Private Sub Form_Load()
About.Caption = "Handling .BIK files requires 3rd Party Tools" & vbNewLine & "RAD Tools binkplay.dll 1.8r" & vbNewLine & "(c) 1997-2006 BINK Technologies"
    BIK_Playa.bik_preload.Value = ReadINI(App.path & "\bin\jr.ini", "BIKPlayer", "preload")
    BIK_Playa.bik_runtimestats.Value = ReadINI(App.path & "\bin\jr.ini", "BIKPlayer", "runtimestats")
    BIK_Playa.bik_dontskip.Value = ReadINI(App.path & "\bin\jr.ini", "BIKPlayer", "dontskip")
    BIK_Playa.bik_hidecursor.Value = ReadINI(App.path & "\bin\jr.ini", "BIKPlayer", "hidecursor")
    BIK_Playa.bik_blackbg.Value = ReadINI(App.path & "\bin\jr.ini", "BIKPlayer", "blackbg")
    BIK_Playa.bik_focuslost.Value = ReadINI(App.path & "\bin\jr.ini", "BIKPlayer", "focuslost")
    BIK_Playa.bik_noMT.Value = ReadINI(App.path & "\bin\jr.ini", "BIKPlayer", "noMT")
    BIK_Playa.bik_showstats2.Value = ReadINI(App.path & "\bin\jr.ini", "BIKPlayer", "showstats2")
End Sub

Private Function addS(ByVal switch As String)
switches.Caption = switches.Caption & "/" & switch
End Function
Private Function remS(ByVal switch As String)
switches.Caption = Replace(switches.Caption, "/" & switch, "")
End Function

Private Sub bik_noMT_Click()
If bik_noMT.Value = 1 Then
addS "k"
Else
remS "k"
End If
End Sub

Private Sub bik_preload_Click()
If bik_preload.Value = 1 Then
addS "p"
Else
remS "p"
End If
End Sub

Private Sub bik_runtimestats_Click()
If bik_runtimestats.Value = 1 Then
addS "q"
Else
remS "q"
End If
End Sub

Private Sub bik_showstats2_Click()
If bik_showstats2.Value = 1 Then
addS "s"
Else
remS "s"
End If
End Sub

Private Sub bik_blackbg_Click()
If bik_blackbg.Value = 1 Then
addS "r"
Else
remS "r"
End If
End Sub

Private Sub bik_dontskip_Click()
If bik_dontskip.Value = 1 Then
addS "n"
Else
remS "n"
End If
End Sub


Private Sub bik_focuslost_Click()
If bik_focuslost.Value = 1 Then
addS "j"
Else
remS "j"
End If
End Sub

Private Sub bik_hidecursor_Click()
If bik_hidecursor.Value = 1 Then
addS "c"
Else
remS "c"
End If
End Sub

Private Sub WatchBIK_Click()
Dim pathx As String: pathx = App.path & "\bin\jr.ini"
If always.Value = 1 Then
INI.AddINI pathx, "BIKPlayer", "SFilter", bik_minimizeJR.Value
INI.AddINI pathx, "BIKPlayer", "Filterx", Me.switches.Caption
INI.AddINI pathx, "BIKPlayer", "preload", bik_preload.Value
INI.AddINI pathx, "BIKPlayer", "runtimestats", bik_runtimestats.Value
INI.AddINI pathx, "BIKPlayer", "dontskip", bik_dontskip.Value
INI.AddINI pathx, "BIKPlayer", "hidecursor", bik_hidecursor.Value
INI.AddINI pathx, "BIKPlayer", "blackbg", bik_blackbg.Value
INI.AddINI pathx, "BIKPlayer", "focuslost", bik_focuslost.Value
INI.AddINI pathx, "BIKPlayer", "noMT", bik_noMT.Value
INI.AddINI pathx, "BIKPlayer", "showstats2", bik_showstats2.Value
Else
INI.AddINI pathx, "BIKPlayer", "SFilter", "0"
INI.AddINI pathx, "BIKPlayer", "Filterx", ""
INI.AddINI pathx, "BIKPlayer", "preload", "0"
INI.AddINI pathx, "BIKPlayer", "runtimestats", "0"
INI.AddINI pathx, "BIKPlayer", "dontskip", "0"
INI.AddINI pathx, "BIKPlayer", "hidecursor", "0"
INI.AddINI pathx, "BIKPlayer", "blackbg", "0"
INI.AddINI pathx, "BIKPlayer", "focuslost", "0"
INI.AddINI pathx, "BIKPlayer", "noMT", "0"
INI.AddINI pathx, "BIKPlayer", "showstats2", "0"
End If
INI.AddINI pathx, "BIKPlayer", "Always", always.Value

If bik_minimizeJR.Value = 1 Then
jrMain.WindowState = vbMinimized
End If
BIK.BIK_play path.Caption, switches.Caption
End Sub
