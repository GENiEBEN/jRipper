VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form NIMP 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "NiMP"
   ClientHeight    =   6765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "wndMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   7770
   Begin VB.CheckBox selall 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Check/Uncheck all"
      Top             =   1230
      Width           =   195
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   165
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   600
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6765
      Left            =   0
      Picture         =   "wndMain.frx":199A
      ScaleHeight     =   6765
      ScaleWidth      =   7770
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   7770
      Begin jR_RC2.Butt Butt1 
         Height          =   525
         Left            =   6570
         TabIndex        =   11
         Top             =   525
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   926
         Caption         =   "APPLY"
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
         cBhover         =   16777215
         Focus           =   0   'False
         LockHover       =   3
         cGradient       =   16777215
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin MSComctlLib.ImageList typelist 
         Left            =   3135
         Top             =   -90
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "wndMain.frx":BE54
               Key             =   "bik"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "wndMain.frx":DA16
               Key             =   "exe"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "wndMain.frx":DDB0
               Key             =   "video"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "wndMain.frx":E3BA
               Key             =   "text"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "wndMain.frx":E754
               Key             =   "image"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView list 
         Height          =   5430
         Left            =   105
         TabIndex        =   1
         Top             =   1200
         Width           =   7560
         _ExtentX        =   13335
         _ExtentY        =   9578
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         OLEDropMode     =   1
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "typelist"
         SmallIcons      =   "typelist"
         ForeColor       =   0
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "col1"
            Text            =   "          Video Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "col2"
            Text            =   "File Name"
            Object.Width           =   7144
         EndProperty
      End
      Begin VB.PictureBox picMain 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   5445
         Left            =   1170
         ScaleHeight     =   5445
         ScaleWidth      =   5430
         TabIndex        =   7
         Top             =   1185
         Visible         =   0   'False
         Width           =   5430
         Begin VB.PictureBox picPar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000012&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3435
            Left            =   450
            Picture         =   "wndMain.frx":EAEE
            ScaleHeight     =   3405
            ScaleWidth      =   5445
            TabIndex        =   8
            Top             =   5160
            Visible         =   0   'False
            Width           =   5475
            Begin VB.Label lblFull 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "code:// codin"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   238
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   1935
               Left            =   570
               TabIndex        =   9
               Top             =   2610
               Width           =   3375
            End
         End
         Begin VB.Label lblMAb 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   26.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   2535
            Left            =   630
            TabIndex        =   10
            Top             =   1500
            Width           =   4215
         End
      End
      Begin VB.Timer tmrSc 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   5790
         Top             =   4905
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   330
         Top             =   495
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   1725
         Top             =   570
      End
      Begin jR_RC2.Butt Butt2 
         Height          =   525
         Left            =   5400
         TabIndex        =   12
         Top             =   525
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   926
         Caption         =   "EXIT"
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
         cBhover         =   16777215
         Focus           =   0   'False
         LockHover       =   3
         cGradient       =   16777215
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin jR_RC2.Butt Butt3 
         Height          =   525
         Left            =   4230
         TabIndex        =   13
         Top             =   525
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   926
         Caption         =   "ABOUT"
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
         cBhover         =   16777215
         Focus           =   0   'False
         LockHover       =   3
         cGradient       =   16777215
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin VB.Image Image3 
         Height          =   135
         Left            =   0
         Picture         =   "wndMain.frx":14518
         Stretch         =   -1  'True
         Top             =   6630
         Width           =   7770
      End
      Begin VB.Image Image2 
         Height          =   5565
         Left            =   7665
         Picture         =   "wndMain.frx":16FE6
         Stretch         =   -1  'True
         Top             =   1185
         Width           =   105
      End
      Begin VB.Image Image1 
         Height          =   5565
         Left            =   0
         Picture         =   "wndMain.frx":18C18
         Stretch         =   -1  'True
         Top             =   1185
         Width           =   105
      End
      Begin VB.Label ret 
         Caption         =   "return"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1635
         TabIndex        =   6
         Top             =   90
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   90
         Width           =   4485
      End
   End
   Begin VB.Label ret2 
      Caption         =   "return2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   2685
   End
End
Attribute VB_Name = "NIMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Sub Butt1_Click()
' is list empty?
If list.ListItems.Count = 0 Then
MsgBox "No videos to process. Select a game first", vbCritical, "NiMP 2.8.11"
Me.Combo1.Text = Combo1.list(1)
Me.Combo1.SetFocus
Exit Sub
End If
' is this game installed?
If Strings.Left(ret.Caption, 5) = "Error" Then
MsgBox "This game is not installed!", vbCritical, "NiMP 2.8.11"
Combo1.SetFocus
Exit Sub
Else
' select game and do it
Select Case Combo1.Text

    Case "Ed Edd n Eddy - Mis Ed-ventures"
    Apply2 1, "BIK"
    
    Case "FallOut"
    Apply_FALLOUT
    
    Case "Brigade E5 - New Jagged Union"
    Apply2 1, "BIK"
    
    Case "Jagged Alliance 2"
    Apply_JA2
    
    Case "Just Cause"
    Apply2 1, "BIK"
    
    Case "Ducati World Championship"
    Apply2 7, "AVI_IV50"

    Case "Reservoir Dogs"
    Apply2 1, "BIK"
    
    Case "Need for Speed - Underground 1"
    Apply2 6, "MAD"
    
    Case "GodFather, The"
    Apply2 5, "VP6"
    
    Case "Lord of the Rings: War of the Ring"
    Apply_LoTRWoR
    
    Case "Stalin Subway, The"
    Apply2 4, "AVI"

    Case "Street Hacker"
    Apply2 4, "AVI"
    
    Case "F.E.A.R. First Encounter Assault and Recon"
    Apply_FEAR 'this one replaces the video file entry in a 4GB archive
    
    Case "Grand Theft Auto III"
    Apply2 3, "MPG"
    
    Case "Grand Theft Auto - Vice City"
    Apply2 3, "MPG"
    
    Case "ChampionSheep Rally"
    Apply_CSR 'copy TITLE.DBC as LOGO.DBC
    
    Case "FarCry"
    Apply
    Apply3
    
    Case "Midnight Outlaw - 6 hours to Sun up"
    Apply2 3, "MPG"
    
    Case "FORD Street Racing *European*"
    Apply2 2, "WMV"
    
    Case "FORD Racing 3 *European*"
    Apply2 2, "WMV"
    
    Case "SWAT 4 - Stetchkov Syndicate"
    Apply2 1, "BIK" 'replace but dont delete
    
    Case "SWAT 4"
    Apply2 1, "BIK" 'replace but dont delete
    
    Case "Condemned - Criminal Origins"
    Apply_CondemnedCO 'this one replaces the video file entry in a 6GB archive
    
    Case "Rogue Trooper"
    Apply '1st folder
    Apply3 '2nd folder
    
    Case "Splinter Cell - Pandora Tomorrow [PAL]"
    Apply
    Apply3
    
    Case "F1 Challenge '99-'02"
    Apply
    Apply3
    
    Case "Syberia II"
    Apply2 1, "BIK" 'replace but dont delete
    
    Case "ToCA Race Driver 3"
    Apply2 1, "BIK" 'replace but dont delete
    
    Case "ToCA Race Driver 2"
    Apply2 1, "BIK" 'replace but dont delete
    
    Case "Colin McRae Rally 2"
    Apply2 1, "BIK" 'replace but dont delete
    
    Case "House of the Dead III"
    Apply
    MsgBox "Don't forget to press ENTER when the black screen shows up"
    
    Case ""
    MsgBox "No game selected!"
    
    Case Else 'this one is for all the games not listed above (generic function)
    Apply
    MsgBox "Done!", vbInformation, "NiMP 2.8.11"
    End
End Select
' refresh list content
CheckFiles
End If
' end
    MsgBox "Done!", vbInformation, "NiMP 2.8.11"
FadeOut Me, , 15
End
End Sub

Private Sub Butt2_Click()
FadeOut Me, , 15
Unload Me
End Sub

Private Sub Butt3_Click()
'
If Butt3.Caption = "ABOUT" Then
list.Visible = False
selall.Visible = False
picMain.Visible = True
Butt1.Enabled = False
lblFull.Caption = "code:> codin" & vbNewLine & "gfx:> codin && murdo"
Butt3.Caption = "Hide ABOUT"
Combo1.Enabled = False

    On Error Resume Next
    If GL.bClick = False Then Exit Sub
    If tmrSc.Enabled Then Exit Sub
    Call SHLabels(False)
    GL.bClick = False


ElseIf Butt3.Caption = "Hide ABOUT" Then
list.Visible = True
selall.Visible = True
picMain.Visible = False
Butt1.Enabled = True
Butt3.Caption = "ABOUT"
Combo1.Enabled = True
Combo1.SetFocus


    On Error Resume Next
    If GL.bClick = False Then Exit Sub
    Call SHLabels(True)
    GL.bClick = False
End If
End Sub

Private Sub Butt3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
GL.bClick = True
End Sub

Private Sub Combo1_Change()
LoadVideos Me.Combo1.Text
CheckFiles
' If there's any file checked, then check the 'Check/Uncheck all'
selall.Value = 0
Dim x
    For x = 1 To Me.list.ListItems.Count
        If list.ListItems(x).Checked = True Then
        Me.selall.Value = 1
        End If
    Next x
' refresh list (to solve a gfx bug when showing scroolbar)
list.Refresh
End Sub

Private Sub Combo1_Click()
LoadVideos Me.Combo1.Text
CheckFiles
' If there's any file checked, then check the 'Check/Uncheck all'
selall.Value = 0
Dim x
    For x = 1 To Me.list.ListItems.Count
        If list.ListItems(x).Checked = True Then
        Me.selall.Value = 1
        End If
    Next x
' refresh list (to solve a gfx bug when showing scroolbar)
list.Refresh
End Sub

Private Sub Form_Load()
' Fill combobox with games
LoadGames Me.Combo1
' smart version set TODO
''''''Dim it As String: it = Me.Combo1.ListCount
''''''Dim v1 As String: v1 = Strings.Left(it, 1)
''''''Dim v2 As String: v2 = Strings.Mid(it, 2, 1)

' Check if files are backuped or not
CheckFiles
' Set Form Caption
Label1.Caption = "NiMP 2.8.11"
Me.Caption = "NiMP 2.8.11"
' If there's any file checked, then check the 'Check/Uncheck all'
Dim x
    For x = 1 To Me.list.ListItems.Count
        If list.ListItems(x).Checked = True Then
        Me.selall.Value = 1
        End If
    Next x
' Do animation effect
FadeIn Me, 250, 15
' set focus
Me.Combo1.SetFocus
' Set tooltip for games combobox
Combo1.ToolTipText = "There are " & Combo1.ListCount & " games in NiMP database"
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
FormDrag Me
End Sub

Private Sub list_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim tmpStr As String: tmpStr = "n"
list.Sorted = True

If ColumnHeader.Text = "          Video Name" Then
    list.SortKey = 0
    list.SortOrder = lvwAscending
    list.Sorted = False
End If

If ColumnHeader.Text = "FileName" Then
list.SortKey = 1
list.Sorted = False
End If
End Sub

Private Sub list_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Item.Checked = False Then
selall.Value = 0
Me.Timer2.Enabled = False
End If
End Sub

Private Sub list_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
FormDrag Me
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
FormDrag Me
End Sub

Private Sub picPar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
FormDrag Me
End Sub

Private Sub selall_Click()
Me.Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    If GetAsyncKeyState(vbKeyF1) Then
        Shell "C:\Windows\Notepad.exe " & App.Path & "\GENiEBEN.nfo", vbNormalFocus
    End If
    If GetAsyncKeyState(vbKeyEscape) Then
        FadeOut Me, 0, 15
        End
    End If
    If GetAsyncKeyState(vbKeyF2) Then
    
    End If
    If GetAsyncKeyState(vbKeyF3) Then
    Butt1_Click
    End If
End Sub

Private Sub Timer2_Timer()
Dim x
If Me.selall.Value = 1 Then
    For x = 1 To Me.list.ListItems.Count
        If list.ListItems(x).Checked = False Then
        list.ListItems(x).Checked = True
        End If
    Next x
ElseIf Me.selall.Value = 0 Then
    For x = 1 To Me.list.ListItems.Count
        If list.ListItems(x).Checked = True Then
        list.ListItems(x).Checked = False
        End If
    Next x
End If
Timer2.Enabled = False
End Sub

Private Sub tmrSc_Timer()
    On Error Resume Next
    With picPar
     If .Top < picMain.Height - picMain.Height - .Height Then
      .Top = .Height - 1
      .Top = picMain.Height - 10
     Else
      .Top = .Top - 10
     End If
    End With

End Sub
