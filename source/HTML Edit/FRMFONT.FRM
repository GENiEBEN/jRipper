VERSION 5.00
Begin VB.Form frmfont 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Font"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2385
   Icon            =   "frmfont.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   2385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdinsert 
      Caption         =   "Insert"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
   Begin VB.CheckBox chkcolor 
      Caption         =   "Use selected font color?"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin VB.ComboBox cbcolor 
      Height          =   315
      ItemData        =   "frmfont.frx":030A
      Left            =   960
      List            =   "frmfont.frx":0338
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.CheckBox chksize 
      Caption         =   "Use selected font size?"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.ComboBox cbsize 
      Height          =   315
      ItemData        =   "frmfont.frx":03A2
      Left            =   960
      List            =   "frmfont.frx":03B5
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblcolor 
      AutoSize        =   -1  'True
      Caption         =   "Font Color:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   765
   End
   Begin VB.Label lblfsize 
      AutoSize        =   -1  'True
      Caption         =   "Font Size:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   705
   End
End
Attribute VB_Name = "frmfont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbcolor_Change()

End Sub

Private Sub cmdcancel_Click()
Unload Me

End Sub

Private Sub cmdinsert_Click()
Dim s As String
Dim c As String
Dim f As String

If cbsize.ListIndex = 0 Then
s = "+2"
ElseIf cbsize.ListIndex = 1 Then
s = "+3"
ElseIf cbsize.ListIndex = 2 Then
s = "+4"
ElseIf cbsize.ListIndex = 3 Then
s = "+5"
ElseIf cbsize.ListIndex = 4 Then
s = "+6"
End If

If cbcolor.ListIndex = 0 Then
c = "#FF3300"
ElseIf cbcolor.ListIndex = 1 Then
c = "#006600"
ElseIf cbcolor.ListIndex = 2 Then
c = "#0000FF"
ElseIf cbcolor.ListIndex = 3 Then
c = "#FF6600"
ElseIf cbcolor.ListIndex = 4 Then
c = "#000000"
ElseIf cbcolor.ListIndex = 5 Then
c = "#FFFFFF"
ElseIf cbcolor.ListIndex = 6 Then
c = "#FFFF00"
ElseIf cbcolor.ListIndex = 7 Then
c = "#FF99CC"
ElseIf cbcolor.ListIndex = 8 Then
c = "#990000"
ElseIf cbcolor.ListIndex = 9 Then
c = "#660099"
ElseIf cbcolor.ListIndex = 10 Then
c = "#999999"
ElseIf cbcolor.ListIndex = 11 Then
c = "#009999"
ElseIf cbcolor.ListIndex = 12 Then
c = "#87BAE0"
ElseIf cbcolor.ListIndex = 13 Then
c = "#000099"
End If

If chksize.Value = Checked And chkcolor.Value = Checked Then
f = "<font color=" & c & " size=" & s & "> </font>" & vbCrLf
frmMain.ActiveForm.CodeMax1.SelText = f
Unload Me
ElseIf chksize.Value = Checked Then
f = "<font size=" & s & "> </font>" & vbCrLf
frmMain.ActiveForm.CodeMax1.SelText = f
Unload Me
ElseIf chkcolor.Value = Checked Then
f = "<font color=" & c & "> </font>" & vbCrLf
frmMain.ActiveForm.CodeMax1.SelText = f
Unload Me
ElseIf chksize.Value = Unchecked And chkcolor.Value = Unchecked Then
MsgBox "can not insert font without a size or a color", vbOKOnly, "Font insert error"
End If

End Sub


