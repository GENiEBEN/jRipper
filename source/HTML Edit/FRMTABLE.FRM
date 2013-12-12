VERSION 5.00
Begin VB.Form frmtable 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insert Table"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4860
   Icon            =   "frmtable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkbg 
      Caption         =   "Use Background Color ?"
      Height          =   375
      Left            =   2760
      TabIndex        =   19
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CheckBox chkbc 
      Caption         =   "Use Border Color?"
      Height          =   255
      Left            =   3000
      TabIndex        =   18
      Top             =   480
      Width           =   1815
   End
   Begin VB.ComboBox cbbg 
      Height          =   315
      ItemData        =   "frmtable.frx":030A
      Left            =   3720
      List            =   "frmtable.frx":0341
      TabIndex        =   17
      Text            =   "Select"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cbbc 
      Height          =   315
      ItemData        =   "frmtable.frx":03B9
      Left            =   3720
      List            =   "frmtable.frx":03F0
      TabIndex        =   15
      Text            =   "Select"
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtborder 
      Height          =   285
      Left            =   3720
      TabIndex        =   13
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtcs 
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtcp 
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Text            =   "0"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtcol 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Text            =   "0"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtrow 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtwidth 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblbgcolor 
      AutoSize        =   -1  'True
      Caption         =   "Background Color:"
      Height          =   195
      Left            =   2280
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label lblbcolor 
      AutoSize        =   -1  'True
      Caption         =   "Border Color:"
      Height          =   195
      Left            =   2760
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblborder 
      AutoSize        =   -1  'True
      Caption         =   "Border:"
      Height          =   195
      Left            =   3120
      TabIndex        =   12
      Top             =   120
      Width           =   510
   End
   Begin VB.Label lblspacing 
      AutoSize        =   -1  'True
      Caption         =   "Cellspacing:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblpadding 
      AutoSize        =   -1  'True
      Caption         =   "Cellpadding:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   870
   End
   Begin VB.Label lblcolumns 
      AutoSize        =   -1  'True
      Caption         =   "Columns:"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   645
   End
   Begin VB.Label lblrows 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      Caption         =   "Rows:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   450
   End
   Begin VB.Label lblwidth 
      AutoSize        =   -1  'True
      Caption         =   "Table Width:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "frmtable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tw As Integer
Dim cp As Integer
Dim cs As Integer
Dim b As Integer
Dim bc As String
Dim bg As String
Dim t As String
Dim tbl As String
Public tr As Integer
Public td As Integer
Public rwcl As String
Sub rc()
tr = frmtable.txtrow.Text
td = frmtable.txtcol.Text

For i = 1 To tr
rwcl = rwcl + "<tr>"
For j = 1 To td
rwcl = rwcl + "<td> </td>"
Next j
rwcl = rwcl + "</tr>" & vbCrLf
Next i

End Sub
Sub bgc()
If cbbg.ListIndex = 0 Then
bg = "#00FFFF"
ElseIf cbbg.ListIndex = 1 Then
bg = "#000000"
ElseIf cbbg.ListIndex = 2 Then
bg = "#0000FF"
ElseIf cbbg.ListIndex = 3 Then
bg = "#FF00FF"
ElseIf cbbg.ListIndex = 4 Then
bg = "#808080"
ElseIf cbbg.ListIndex = 5 Then
bg = "#008000"
ElseIf cbbg.ListIndex = 6 Then
bg = "#00FF00"
ElseIf cbbg.ListIndex = 7 Then
bg = "#800000"
ElseIf cbbg.ListIndex = 8 Then
bg = "#000080"
ElseIf cbbg.ListIndex = 9 Then
bg = "#808000"
ElseIf cbbg.ListIndex = 10 Then
bg = "#FFC0CB"
ElseIf cbbg.ListIndex = 11 Then
bg = "#800080"
ElseIf cbbc.ListIndex = 12 Then
bg = "#FF0000"
ElseIf cbbg.ListIndex = 13 Then
bg = "#C0C0C0"
ElseIf cbbg.ListIndex = 14 Then
bg = "#008080"
ElseIf cbbg.ListIndex = 15 Then
bg = "#fffff"
ElseIf cbbg.ListIndex = 16 Then
bg = "#FFFF00"
End If
End Sub


Sub border()
If cbbc.ListIndex = 0 Then
bc = "#00FFFF"
ElseIf cbbc.ListIndex = 1 Then
bc = "#000000"
ElseIf cbbc.ListIndex = 2 Then
bc = "#0000FF"
ElseIf cbbc.ListIndex = 3 Then
bc = "#FF00FF"
ElseIf cbbc.ListIndex = 4 Then
bc = "#808080"
ElseIf cbbc.ListIndex = 5 Then
bc = "#008000"
ElseIf cbbc.ListIndex = 6 Then
bc = "#00FF00"
ElseIf cbbc.ListIndex = 7 Then
bc = "#800000"
ElseIf cbbc.ListIndex = 8 Then
bc = "#000080"
ElseIf cbbc.ListIndex = 9 Then
bc = "#808000"
ElseIf cbbc.ListIndex = 10 Then
bc = "#FFC0CB"
ElseIf cbbc.ListIndex = 11 Then
bc = "#800080"
ElseIf cbbc.ListIndex = 12 Then
bc = "#FF0000"
ElseIf cbbc.ListIndex = 13 Then
bc = "#C0C0C0"
ElseIf cbbc.ListIndex = 14 Then
bc = "#008080"
ElseIf cbbc.ListIndex = 15 Then
bc = "#fffff"
ElseIf cbbc.ListIndex = 16 Then
bc = "#FFFF00"
End If
End Sub


Sub brbg()
If cbbc.ListIndex = 0 Then
bc = "#00FFFF"
ElseIf cbbc.ListIndex = 1 Then
bc = "#000000"
ElseIf cbbc.ListIndex = 2 Then
bc = "#0000FF"
ElseIf cbbc.ListIndex = 3 Then
bc = "#FF00FF"
ElseIf cbbc.ListIndex = 4 Then
bc = "#808080"
ElseIf cbbc.ListIndex = 5 Then
bc = "#008000"
ElseIf cbbc.ListIndex = 6 Then
bc = "#00FF00"
ElseIf cbbc.ListIndex = 7 Then
bc = "#800000"
ElseIf cbbc.ListIndex = 8 Then
bc = "#000080"
ElseIf cbbc.ListIndex = 9 Then
bc = "#808000"
ElseIf cbbc.ListIndex = 10 Then
bc = "#FFC0CB"
ElseIf cbbc.ListIndex = 11 Then
bc = "#800080"
ElseIf cbbc.ListIndex = 12 Then
bc = "#FF0000"
ElseIf cbbc.ListIndex = 13 Then
bc = "#C0C0C0"
ElseIf cbbc.ListIndex = 14 Then
bc = "#008080"
ElseIf cbbc.ListIndex = 15 Then
bc = "#fffff"
ElseIf cbbc.ListIndex = 16 Then
bc = "#FFFF00"
ElseIf cbbg.ListIndex = 0 Then
bg = "#00FFFF"
ElseIf cbbg.ListIndex = 1 Then
bg = "#000000"
ElseIf cbbg.ListIndex = 2 Then
bg = "#0000FF"
ElseIf cbbg.ListIndex = 3 Then
bg = "#FF00FF"
ElseIf cbbg.ListIndex = 4 Then
bg = "#808080"
ElseIf cbbg.ListIndex = 5 Then
bg = "#008000"
ElseIf cbbg.ListIndex = 6 Then
bg = "#00FF00"
ElseIf cbbg.ListIndex = 7 Then
bg = "#800000"
ElseIf cbbg.ListIndex = 8 Then
bg = "#000080"
ElseIf cbbg.ListIndex = 9 Then
bg = "#808000"
ElseIf cbbg.ListIndex = 10 Then
bg = "#FFC0CB"
ElseIf cbbg.ListIndex = 11 Then
bg = "#800080"
ElseIf cbbc.ListIndex = 12 Then
bg = "#FF0000"
ElseIf cbbg.ListIndex = 13 Then
bg = "#C0C0C0"
ElseIf cbbg.ListIndex = 14 Then
bg = "#008080"
ElseIf cbbg.ListIndex = 15 Then
bg = "#fffff"
ElseIf cbbg.ListIndex = 16 Then
bg = "#FFFF00"
End If
End Sub

Private Sub chkbc_Click()
lblbcolor.Visible = True
cbbc.Visible = True

End Sub

Private Sub chkbg_Click()
lblbgcolor.Visible = True
cbbg.Visible = True

End Sub


Private Sub cmdcancel_Click()
Unload Me

End Sub

Private Sub cmdok_Click()
tw = txtwidth.Text
cp = txtcp.Text
cs = txtcs.Text
b = txtborder.Text

If chkbc.Value = Checked Then
Call border
t = "<table width=" & Chr(34) & tw & Chr(34) & " cellspacing=" & Chr(34) & cs & Chr(34) & " cellpadding=" & Chr(34) & cp & Chr(34) & " border=" & Chr(34) & b & Chr(34) & " bordercolor=" & Chr(34) & bc & Chr(34) & "> " & vbCrLf
Else
t = "<table width=" & Chr(34) & tw & Chr(34) & " cellspacing=" & Chr(34) & cs & Chr(34) & " cellpadding=" & Chr(34) & cp & Chr(34) & " border=" & Chr(34) & b & Chr(34) & "> " & vbCrLf
End If

If chkbg.Value = Checked Then
Call bgc
t = "<table width=" & Chr(34) & tw & Chr(34) & " cellspacing=" & Chr(34) & cs & Chr(34) & " cellpadding=" & Chr(34) & cp & Chr(34) & " border=" & Chr(34) & b & Chr(34) & " bgcolor=" & Chr(34) & bg & Chr(34) & "> " & vbCrLf
Else
t = "<table width=" & Chr(34) & tw & Chr(34) & " cellspacing=" & Chr(34) & cs & Chr(34) & " cellpadding=" & Chr(34) & cp & Chr(34) & " border=" & Chr(34) & b & Chr(34) & "> " & vbCrLf
End If

If chkbc.Value = Checked And chkbg.Value = Checked Then
Call border
Call bgc
t = "<table width=" & Chr(34) & tw & Chr(34) & " cellspacing=" & Chr(34) & cs & Chr(34) & " cellpadding=" & Chr(34) & cp & Chr(34) & " border=" & Chr(34) & b & Chr(34) & " bordercolor=" & Chr(34) & bc & Chr(34) & " bgcolor=" & Chr(34) & bg & Chr(34) & "> " & vbCrLf
End If

Call rc

tbl = t + rwcl
tbl = tbl + "</table>" & vbCrLf

frmMain.ActiveForm.CodeMax1.SelText = tbl
Unload Me

End Sub

