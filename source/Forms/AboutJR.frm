VERSION 5.00
Begin VB.Form AboutJR 
   BackColor       =   &H004D483F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " About jR"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AboutJR.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7590
   Begin VB.Frame Frame1 
      BackColor       =   &H004D483F&
      Height          =   3825
      Left            =   1980
      TabIndex        =   1
      Top             =   0
      Width           =   5520
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
         Height          =   3645
         Left            =   30
         ScaleHeight     =   3645
         ScaleWidth      =   5445
         TabIndex        =   2
         Top             =   135
         Width           =   5445
         Begin VB.Timer tmrSc 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   165
            Top             =   285
         End
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
            Height          =   5000
            Left            =   0
            ScaleHeight     =   4965
            ScaleWidth      =   5445
            TabIndex        =   3
            Top             =   3330
            Visible         =   0   'False
            Width           =   5475
            Begin VB.Image Image2 
               Height          =   2850
               Left            =   795
               Picture         =   "AboutJR.frx":1CCA
               Top             =   60
               Width           =   3825
            End
            Begin VB.Label lblFull 
               BackStyle       =   0  'Transparent
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
               Height          =   4965
               Left            =   105
               TabIndex        =   4
               Top             =   2985
               Width           =   5115
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
            TabIndex        =   5
            Top             =   120
            Width           =   4215
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H004D483F&
      Height          =   3825
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Contact me at gbcrk@yahoo.com or visit genieben.net"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BDB8AF&
         Height          =   795
         Left            =   135
         TabIndex        =   6
         Top             =   2910
         Width           =   1680
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3645
         Left            =   45
         Picture         =   "AboutJR.frx":2570C
         Stretch         =   -1  'True
         Top             =   135
         Width           =   1845
      End
   End
End
Attribute VB_Name = "AboutJR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Caption = "About jRipper " & IAPPV
lblFull.Caption = "Development: codin" & vbNewLine & "Graphics: codin" & vbNewLine & vbNewLine & "Thanks to: " & vbNewLine & "* Dave Perry (nodtveit@sover.net)" & vbNewLine & "* Arkadiy Olovyannikov" & vbNewLine & "* Racer_S" & vbNewLine & "* jTommy (http://jtommy.by.ru)"
    Call SHLabels_2(False)
    GL.bClick = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
GL.bClick = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub tmrSc_Timer()
    On Error Resume Next
    With Me.picPar
     If .Top < picMain.Height - picMain.Height - .Height Then
      .Top = .Height - 1
      .Top = picMain.Height - 10
     Else
      .Top = .Top - 10
     End If
    End With

End Sub

