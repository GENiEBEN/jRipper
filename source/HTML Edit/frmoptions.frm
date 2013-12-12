VERSION 5.00
Begin VB.Form frmoptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmoptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkse 
      Caption         =   "Color coded script editor"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CheckBox chkcolor 
      Caption         =   "Color coded text"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2760
      Width           =   855
   End
End
Attribute VB_Name = "frmoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim inifile As String: inifile = App.Path & "\HTML Edit.ini"

If chkcolor.Value = Checked Then
WriteINI "Options", "Colored", True, inifile
ElseIf chkcolor.Value = Unchecked Then
WriteINI "Options", "Colored", False, inifile
End If

If chkse.Value = Checked Then
WriteINI "Options", "Scripts", True, inifile
ElseIf chkse.Value = Unchecked Then
WriteINI "Options", "Scripts", False, inifile
End If
Unload Me
End Sub

Private Sub Form_Load()
Dim s As String
Dim se As String
Dim inifile As String: inifile = App.Path & "\HTML Edit.ini"
color = ReadINI("Options", "Colored", inifile)
If color = True Then
chkcolor.Value = Checked
ElseIf color = False Then
chkcolor.Value = Unchecked
End If

se = ReadINI("Options", "Scripts", inifile)
If se = True Then
chkse.Value = Checked
ElseIf color = False Then
chkse.Value = Unchecked
End If

End Sub
