VERSION 5.00
Begin VB.Form wndGDP 
   BackColor       =   &H004D483F&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GDP Manager 1.00"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GDP_Manager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin jR_RC2.Butt Butt1 
      Height          =   525
      Left            =   4635
      TabIndex        =   1
      Top             =   1185
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   926
      Caption         =   "UnPACK"
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
      Height          =   330
      Left            =   180
      TabIndex        =   2
      Top             =   1305
      Width           =   4395
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
      Height          =   915
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   5520
   End
End
Attribute VB_Name = "wndGDP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Butt1_Click()
GDP_Unpack Me.Path.Caption, App.Path & "\bin\ExMachinaUnGDP.dll"
Unload Me
End Sub

Private Sub Form_Load()
About.Caption = "Handling .GDP files requires 3rd Party Tools" & vbNewLine & "ExMachina GDP Archives Extractor & Packer v1.1 by jTommy" & vbNewLine & "(c) 2006 jTommy, E-mail: jTommy@by.ru, WWW: http://jtommy.by.ru"
End Sub

Public Function GDP_Unpack(ByVal FilePath As String, ByVal ExMachinaUnGDP_Path As String)
Shell (ExMachinaUnGDP_Path & " " & Chr(34) & FilePath & Chr(34)), vbNormalFocus
End Function
