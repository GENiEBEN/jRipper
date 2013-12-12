VERSION 5.00
Begin VB.Form BlackMirrorConfig 
   BorderStyle     =   0  'None
   Caption         =   "BM Config"
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "wndMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   6255
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   6180
      Left            =   0
      Picture         =   "wndMain.frx":0CCA
      ScaleHeight     =   6180
      ScaleWidth      =   6255
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin jR_RC2.Butt Butt1 
         Height          =   420
         Left            =   4875
         TabIndex        =   13
         Top             =   5595
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   741
         Caption         =   "APPLY"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   16777215
         Focus           =   0   'False
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   0
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   660
         Top             =   4080
      End
      Begin VB.TextBox Text1 
         Height          =   2205
         Left            =   1860
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "wndMain.frx":A431
         ToolTipText     =   "Take care when editing this shit!"
         Top             =   3270
         Width           =   4050
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "wndMain.frx":A604
         Left            =   1860
         List            =   "wndMain.frx":A606
         TabIndex        =   9
         Text            =   "DISABLED"
         Top             =   2925
         Width           =   4050
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "wndMain.frx":A608
         Left            =   1860
         List            =   "wndMain.frx":A60A
         TabIndex        =   7
         Text            =   "ENABLED"
         Top             =   2580
         Width           =   4050
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "wndMain.frx":A60C
         Left            =   1860
         List            =   "wndMain.frx":A60E
         TabIndex        =   3
         Text            =   "800x600x32 YUV"
         Top             =   2235
         Width           =   4050
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "wndMain.frx":A610
         Left            =   1860
         List            =   "wndMain.frx":A612
         TabIndex        =   2
         Text            =   "800x600x32"
         ToolTipText     =   "The 1024x768 resolution is obsolete."
         Top             =   1890
         Width           =   4050
      End
      Begin jR_RC2.Butt Butt2 
         Height          =   420
         Left            =   3675
         TabIndex        =   14
         Top             =   5595
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "CANCEL/EXIT"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   16777215
         Focus           =   0   'False
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   0
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Other game options :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   240
         TabIndex        =   11
         Top             =   3360
         Width           =   1680
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Normalize Sample :"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   2970
         Width           =   1680
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sound :"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   2625
         Width           =   1680
      End
      Begin VB.Label buff 
         Caption         =   "buffer"
         Height          =   225
         Left            =   780
         TabIndex        =   6
         Top             =   4755
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Visit http://www.genieben.net"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   300
         MouseIcon       =   "wndMain.frx":A614
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   5760
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cutscene Resolution :"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   1680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Game Resolution :"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   240
         TabIndex        =   1
         Top             =   1935
         Width           =   1965
      End
   End
End
Attribute VB_Name = "BlackMirrorConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Dim skinS As New SkinSupport
Dim fso As New FileSystemObject

Private Sub Butt1_Click()
buff.Caption = jrRegistry.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\The Adventure Company\The Black Mirror", "Path")
buff.Caption = buff.Caption & "\agds.cfg"
'
fso.DeleteFile buff.Caption, True
fso.CreateTextFile buff.Caption, True
'
Dim film_yuy2 As String
Dim film_switch As String
    Select Case Combo2.Text
    Case "640x480x32 RGB"
    film_yuy2 = "0"
    film_switch = "1"
    Case "800x600x32 RGB"
    film_yuy2 = "0"
    film_switch = "0"
    Case "640x480x32 YUV"
    film_yuy2 = "1"
    film_switch = "1"
    Case "800x600x32 YUV"
    film_yuy2 = "1"
    film_switch = "0"
    End Select
Dim sound As String
    Select Case Combo3.Text
    Case "ENABLED"
    sound = "1"
    Case "DISABLED"
    sound = "0"
    End Select
Dim normalize_sample As String
    Select Case Combo4.Text
    Case "ENABLED"
    normalize_sample = "1"
    Case "DISABLED"
    normalize_sample = "0"
    End Select
    
Dim Final As String
Final = "AGDS.CFG" & vbNewLine & vbNewLine & _
"charset=1252" & vbNewLine & _
"videomode=" & Me.Combo1.Text & vbNewLine & _
"film_yuy2=" & film_yuy2 & vbNewLine & _
"film_switch=" & film_switch & vbNewLine & _
"sound=" & sound & vbNewLine & _
"normalize_sample=" & normalize_sample & vbNewLine & vbNewLine & _
Text1.Text
'
Open buff.Caption For Output As #1
Print #1, (Final)
Close #1
'
MsgBox "Done!", vbInformation, "BM Config"
FadeOut Me, 0, 4
End
End Sub

Private Sub Butt2_Click()
FadeOut Me, 0, 4
Unload Me
End Sub

Private Sub Form_Load()
FadeIn Me, 255, 4
'
Combo1.AddItem "640x480x8"
Combo1.AddItem "640x480x16"
Combo1.AddItem "640x480x32"
Combo1.AddItem "800x600x8"
Combo1.AddItem "800x600x16"
Combo1.AddItem "800x600x32"
Combo1.AddItem "1024x768x8"
Combo1.AddItem "1024x768x16"
Combo1.AddItem "1024x768x32"
'
Combo2.AddItem "640x480x32 RGB" ' (0,1) (film_yuy2,film_switch)
Combo2.AddItem "800x600x32 RGB" ' (0,0) (film_yuy2,film_switch)
Combo2.AddItem "640x480x32 YUV" ' (1,1) (film_yuy2,film_switch)
Combo2.AddItem "800x600x32 YUV" ' (1,0) (film_yuy2,film_switch)
'
Combo3.AddItem "ENABLED" ' sound=1
Combo3.AddItem "DISABLED" ' sound=0
'
Combo4.AddItem "ENABLED" ' normalize_sample=1
Combo4.AddItem "DISABLED" ' normalize_sample=0

End Sub

Private Sub Form_Terminate()
FadeOut Me, 0, 4
End Sub

Private Sub Form_Unload(Cancel As Integer)
FadeOut Me, 0, 4
End Sub

Private Sub Label4_Click()
OpenWWW "http://www.genieben.net", Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
skinS.Drag_Form vbRightButton, Me
End Sub

Private Sub Timer1_Timer()
    If GetAsyncKeyState(vbKeyF1) Then
     If fso.FileExists(App.Path & "\GENiEBEN.nfo") Then
        Shell "C:\Windows\Notepad.exe " & App.Path & "\GENiEBEN.nfo", vbNormalFocus
     End If
    End If
    If GetAsyncKeyState(vbKeyEscape) Then
        FadeOut Me, 0, 4
        Unload Me
    End If
    If GetAsyncKeyState(vbKeyF2) Then
        OpenWWW "http://www.genieben.net", Me
    End If
End Sub
