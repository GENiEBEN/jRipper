VERSION 5.00
Begin VB.Form frmmarq 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add a Marquee"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3780
   Icon            =   "frmmarq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3780
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkstop 
      Caption         =   "Stop on mouse over?"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdinsert 
      Caption         =   "Insert"
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtspeed 
      Height          =   285
      Left            =   3240
      TabIndex        =   11
      Text            =   "0"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtdelay 
      Height          =   285
      Left            =   3240
      TabIndex        =   10
      Text            =   "0"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtheight 
      Height          =   285
      Left            =   3240
      TabIndex        =   9
      Text            =   "0"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtwidth 
      Height          =   285
      Left            =   3240
      TabIndex        =   8
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.ComboBox cbbg 
      Height          =   315
      ItemData        =   "frmmarq.frx":030A
      Left            =   1320
      List            =   "frmmarq.frx":033B
      TabIndex        =   3
      Text            =   "None"
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox cbdir 
      Height          =   315
      ItemData        =   "frmmarq.frx":03AB
      Left            =   1320
      List            =   "frmmarq.frx":03BB
      TabIndex        =   2
      Text            =   "Select"
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblspeed 
      AutoSize        =   -1  'True
      Caption         =   "Speed:"
      Height          =   195
      Left            =   2640
      TabIndex        =   7
      Top             =   1200
      Width           =   510
   End
   Begin VB.Label lbldelay 
      AutoSize        =   -1  'True
      Caption         =   "Delay:"
      Height          =   195
      Left            =   2760
      TabIndex        =   6
      Top             =   840
      Width           =   450
   End
   Begin VB.Label lblheight 
      AutoSize        =   -1  'True
      Caption         =   "Height:"
      Height          =   195
      Left            =   2640
      TabIndex        =   5
      Top             =   480
      Width           =   510
   End
   Begin VB.Label lblwidth 
      AutoSize        =   -1  'True
      Caption         =   "Width:"
      Height          =   195
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   465
   End
   Begin VB.Label lblback 
      AutoSize        =   -1  'True
      Caption         =   "Background:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   915
   End
   Begin VB.Label lbldir 
      AutoSize        =   -1  'True
      Caption         =   "Scroll Direction:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1110
   End
End
Attribute VB_Name = "frmmarq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
Unload Me

End Sub

Private Sub cmdinsert_Click()
Dim dir As String
Dim d As Integer
Dim w As Integer
Dim h As Integer
Dim s As Integer
Dim marq As String
Dim bc As String
Dim sd As String

w = txtwidth.Text
h = txtheight.Text
d = txtdelay.Text
s = txtspeed.Text

If cbbg.ListIndex = 1 Then
bg = "#FF3300"
ElseIf cbbg.ListIndex = 2 Then
bg = "#006600"
ElseIf cbbg.ListIndex = 3 Then
bg = "#0000FF"
ElseIf cbbg.ListIndex = 4 Then
bg = "#FF6600"
ElseIf cbbg.ListIndex = 5 Then
bg = "#000000"
ElseIf cbbg.ListIndex = 6 Then
bg = "#FFFFFF"
ElseIf cbbg.ListIndex = 7 Then
bg = "#FFFF00"
ElseIf cbbg.ListIndex = 8 Then
bg = "#FF99CC"
ElseIf cbbg.ListIndex = 9 Then
bg = "#990000"
ElseIf cbbg.ListIndex = 10 Then
bg = "#660099"
ElseIf cbbg.ListIndex = 11 Then
bg = "#999999"
ElseIf cbbg.ListIndex = 12 Then
bg = "#009999"
ElseIf cbbg.ListIndex = 13 Then
bg = "#87BAE0"
ElseIf cbbg.ListIndex = 14 Then
bg = "#000099"
End If

If cbdir.ListIndex = 0 Then
dir = "up"
ElseIf cbdir.ListIndex = 1 Then
dir = "down"
ElseIf cbdir.ListIndex = 2 Then
dir = "left"
ElseIf cbdir.ListIndex = 3 Then
dir = "right"
End If

If cbbg.ListIndex = 0 And chkstop.Value = Checked Then
marq = "<marquee behivor=scroll direction=" & dir & " height=" & h & " width=" & w & " scrollamount=" & s & " scrolldelay=" & d & " onMouseOver=this.stop() onMouseOut=this.start()>" & vbCrLf & vbCrLf
marq = marq + "</marquee>"
frmMain.ActiveForm.CodeMax1.SelText = marq
Unload Me
ElseIf chkstop.Value = Unchecked Then
marq = "<marquee behivor=scroll direction=" & dir & " height=" & h & " width=" & w & " scrollamount=" & s & " scrolldelay=" & d & " bgcolor=" & bg & ">" & vbCrLf & vbCrLf
marq = marq + "</marquee>"
frmMain.ActiveForm.CodeMax1.SelText = marq
Unload Me
ElseIf chkstop.Value = Checked Then
marq = "<marquee behivor=scroll direction=" & dir & " height=" & h & " width=" & w & " scrollamount=" & s & " scrolldelay=" & d & " bgcolor=" & bg & " onMouseOver=this.stop() onMouseOut=this.start()>" & vbCrLf & vbCrLf
marq = marq + "</marquee>"
frmMain.ActiveForm.CodeMax1.SelText = marq
Unload Me
End If

End Sub

Private Sub lblborder_Click()

End Sub


