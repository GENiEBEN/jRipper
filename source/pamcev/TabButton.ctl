VERSION 5.00
Begin VB.UserControl TabButton 
   Alignable       =   -1  'True
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   ScaleHeight     =   5205
   ScaleWidth      =   7500
   Begin VB.Image ImgTab2 
      Height          =   495
      Left            =   3645
      Top             =   -45
      Width           =   2775
   End
   Begin VB.Image ImgTab1 
      Height          =   495
      Left            =   780
      Top             =   -45
      Width           =   2775
   End
   Begin VB.Label LblTab2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TabText"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3870
      TabIndex        =   1
      Top             =   135
      Width           =   2295
   End
   Begin VB.Label LblTab1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TabText"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1020
      TabIndex        =   0
      Top             =   135
      Width           =   2295
   End
   Begin VB.Image ImgTabs2 
      Height          =   1440
      Left            =   0
      Picture         =   "TabButton.ctx":0000
      Top             =   0
      Width           =   7500
   End
   Begin VB.Image ImgTabs1 
      Height          =   1440
      Left            =   0
      Picture         =   "TabButton.ctx":BFC2
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "TabButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event TabClick1()
Event TabClick2()

Private Sub ImgTab1_Click()
TabClick1
End Sub

Private Sub ImgTab2_Click()
TabClick2
End Sub

Private Sub UserControl_Resize()
UserControl.Width = ImgTabs1.Width
UserControl.Height = ImgTabs1.Height
End Sub

Public Sub SetText(TabIndex As Integer, Text As String)
If TabIndex = 1 Then
LblTab1.Caption = Text
End If
If TabIndex = 2 Then
LblTab2.Caption = Text
End If
End Sub

Private Sub TabClick1()
ImgTabs1.Visible = False
ImgTabs2.Visible = True
RaiseEvent TabClick1
End Sub

Private Sub TabClick2()
ImgTabs1.Visible = True
ImgTabs2.Visible = False
RaiseEvent TabClick2
End Sub
