VERSION 5.00
Begin VB.Form frmframes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insert Frames"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9075
   Icon            =   "frmframes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtmh5 
      Height          =   285
      Left            =   8520
      TabIndex        =   55
      Text            =   "0"
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtmh4 
      Height          =   285
      Left            =   8520
      TabIndex        =   54
      Text            =   "0"
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtmh3 
      Height          =   285
      Left            =   8520
      TabIndex        =   53
      Text            =   "0"
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtmw5 
      Height          =   285
      Left            =   6840
      TabIndex        =   49
      Text            =   "0"
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtmw4 
      Height          =   285
      Left            =   6840
      TabIndex        =   48
      Text            =   "0"
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtmw3 
      Height          =   285
      Left            =   6840
      TabIndex        =   47
      Text            =   "0"
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtname5 
      Height          =   285
      Left            =   4680
      TabIndex        =   43
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtname4 
      Height          =   285
      Left            =   4680
      TabIndex        =   42
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtname3 
      Height          =   285
      Left            =   4680
      TabIndex        =   41
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtsource5 
      Height          =   285
      Left            =   1320
      TabIndex        =   37
      Top             =   3840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtsource4 
      Height          =   285
      Left            =   1320
      TabIndex        =   35
      Top             =   3480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtsource3 
      Height          =   285
      Left            =   1320
      TabIndex        =   33
      Top             =   3120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtmh2 
      Height          =   285
      Left            =   8520
      TabIndex        =   31
      Text            =   "0"
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtmw2 
      Height          =   285
      Left            =   6840
      TabIndex        =   29
      Text            =   "0"
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtname2 
      Height          =   285
      Left            =   4680
      TabIndex        =   27
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtsource2 
      Height          =   285
      Left            =   1320
      TabIndex        =   25
      Top             =   2760
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtmh1 
      Height          =   285
      Left            =   8520
      TabIndex        =   23
      Text            =   "0"
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtmw1 
      Height          =   285
      Left            =   6840
      TabIndex        =   21
      Text            =   "0"
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtname1 
      Height          =   285
      Left            =   4680
      TabIndex        =   19
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtsource1 
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Top             =   2400
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.OptionButton optbrlt 
      Caption         =   "Bottom-right-left-top"
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   1920
      Width           =   1695
   End
   Begin VB.OptionButton optbotrl 
      Caption         =   "Bottom-right-left"
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   1560
      Width           =   1455
   End
   Begin VB.OptionButton opttoprl 
      Caption         =   "Top-right-left"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdinsert 
      Caption         =   "Insert"
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   4560
      Width           =   975
   End
   Begin VB.OptionButton optbottop 
      Caption         =   "Bottom-top"
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton optbotright 
      Caption         =   "Bottom-right"
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton optbotleft 
      Caption         =   "Bottom-left"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton optbottom 
      Caption         =   "Bottom"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.OptionButton optright 
      Caption         =   "Right"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.OptionButton optleft 
      Caption         =   "Left"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.OptionButton opttopright 
      Caption         =   "Top-right"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton opttopleft 
      Caption         =   "Top-left"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.OptionButton opttop 
      Caption         =   "Top"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   1855
      Left            =   3360
      ScaleHeight     =   1800
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   360
      Width           =   1575
      Begin VB.Shape tops 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Shape sbottom 
         BackColor       =   &H0080FFFF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   0
         Top             =   1560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Shape sright 
         BackColor       =   &H00FFFF00&
         BackStyle       =   1  'Opaque
         Height          =   1815
         Left            =   1320
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape sleft 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         Height          =   1815
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Label lblmh5 
      AutoSize        =   -1  'True
      Caption         =   "Margin Height:"
      Height          =   195
      Left            =   7440
      TabIndex        =   52
      Top             =   3840
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblmh4 
      AutoSize        =   -1  'True
      Caption         =   "Margin Height:"
      Height          =   195
      Left            =   7440
      TabIndex        =   51
      Top             =   3480
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblmh3 
      AutoSize        =   -1  'True
      Caption         =   "Margin Height:"
      Height          =   195
      Left            =   7440
      TabIndex        =   50
      Top             =   3120
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblmw5 
      AutoSize        =   -1  'True
      Caption         =   "Margin width:"
      Height          =   195
      Left            =   5880
      TabIndex        =   46
      Top             =   3840
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblmw4 
      AutoSize        =   -1  'True
      Caption         =   "Margin width:"
      Height          =   195
      Left            =   5880
      TabIndex        =   45
      Top             =   3480
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblmw3 
      AutoSize        =   -1  'True
      Caption         =   "Margin width:"
      Height          =   195
      Left            =   5880
      TabIndex        =   44
      Top             =   3120
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblname5 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   4200
      TabIndex        =   40
      Top             =   3840
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblname4 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   4200
      TabIndex        =   39
      Top             =   3480
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblname3 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   4200
      TabIndex        =   38
      Top             =   3120
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblsource5 
      AutoSize        =   -1  'True
      Caption         =   "Frame soucre 5:"
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   3840
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label lblsource4 
      AutoSize        =   -1  'True
      Caption         =   "Frame source 4:"
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   3480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label lblsource3 
      AutoSize        =   -1  'True
      Caption         =   "Frame source 3:"
      Height          =   195
      Left            =   120
      TabIndex        =   32
      Top             =   3120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label lblmh2 
      AutoSize        =   -1  'True
      Caption         =   "Margin Height:"
      Height          =   195
      Left            =   7440
      TabIndex        =   30
      Top             =   2760
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblmw2 
      AutoSize        =   -1  'True
      Caption         =   "Margin width:"
      Height          =   195
      Left            =   5880
      TabIndex        =   28
      Top             =   2760
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblname2 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   4200
      TabIndex        =   26
      Top             =   2760
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblsource2 
      AutoSize        =   -1  'True
      Caption         =   "Frame source 2:"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   2760
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label lblmh1 
      AutoSize        =   -1  'True
      Caption         =   "Margin Height:"
      Height          =   195
      Left            =   7440
      TabIndex        =   22
      Top             =   2400
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblmw1 
      AutoSize        =   -1  'True
      Caption         =   "Margin width:"
      Height          =   195
      Left            =   5880
      TabIndex        =   20
      Top             =   2400
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblname1 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   4200
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblsource1 
      AutoSize        =   -1  'True
      Caption         =   "Frame source 1:"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label lblpreview 
      AutoSize        =   -1  'True
      Caption         =   "Preview"
      Height          =   195
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "frmframes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim source1 As String
Dim source2 As String
Dim source3 As String
Dim source4 As String
Dim source5 As String
Dim mw1 As Integer
Dim mw2 As Integer
Dim mw3 As Integer
Dim mw4 As Integer
Dim mw5 As Integer
Dim mh1 As Integer
Dim mh2 As Integer
Dim mh3 As Integer
Dim mh4 As Integer
Dim mh5 As Integer
Dim name1 As String
Dim name2 As String
Dim name3 As String
Dim name4 As String
Dim name5 As String
Dim frame As String
Dim r As String
Dim c As String
Dim c2 As String
Sub getinfo()
source1 = txtsource1.Text
source2 = txtsource2.Text
source3 = txtsource3.Text
source4 = txtsource4.Text
source5 = txtsource5.Text
name1 = txtname1.Text
name2 = txtname2.Text
name3 = txtname3.Text
name4 = txtname4.Text
name5 = txtname5.Text
mw1 = txtmw1.Text
mw2 = txtmw2.Text
mw3 = txtmw3.Text
mw4 = txtmw4.Text
mw5 = txtmw5.Text
mh1 = txtmh1.Text
mh2 = txtmh2.Text
mh3 = txtmh3.Text
mh4 = txtmh4.Text
mh5 = txtmh5.Text
End Sub


Sub makeframes()
If opttop.Value = True Then
r = "<frameset rows=""25%,75%"" border=0 noresize>" & vbCrLf
c = "<frameset cols=""100%"" border=0 frameborder=0 scrolling=""no"" noresize>" & vbCrLf
frame = r + c
frame = frame + "<frame src=" & source1 & " name=" & name1 & " marginwidth=" & mw1 & " marginheight=" & mh1 & " noresize>" & vbCrLf
frame = frame + "<frame src=" & source2 & " name=" & name2 & " marginwidht=" & mw2 & " marginheight=" & mh2 & " noresize>" & vbCrLf
frame = frame + "</frameset>" & vbCrLf & "</frameset>" & vbCrLf
fMainForm.ActiveForm.CodeMax1.SelText = frame
ElseIf opttopleft.Value = True Then
r = "<frameset rows=""25%,75%"" border=0 noresize>" & vbCrLf
c = "<frameset cols=""100%"" border=0 frameborder=0 scrolling=""no"" noresize>" & vbCrLf
frame = r + c
frame = frame + "<frame src=" & source1 & " name=" & name1 & " marginwidth=" & mw1 & " marginheight=" & mh1 & " noresize>" & vbCrLf
c2 = "<frameset cols=""25%,75%"" border=0 frameborder=0 noresize>" & vbCrLf
frame = frame + c2
frame = frame + "<frame src=" & source2 & " name=" & name2 & " marginwidht=" & mw2 & " marginheight=" & mh2 & " noresize>" & vbCrLf
frame = frame + "<frame src=" & source3 & " name=" & name3 & " marginwidht=" & mw3 & " marginheight=" & mh3 & " noresize>" & vbCrLf
frame = frame + "</frameset>" & vbCrLf & "</frameset>" & vbCrLf & "</frameset>" & vbCrLf
fMainForm.ActiveForm.CodeMax1.SelText = frame
ElseIf opttopright.Value = True Then
r = "<frameset rows=""25%,75%"" border=0 noresize>" & vbCrLf
c = "<frameset cols=""100%"" border=0 frameborder=0 scrolling=""no"" noresize>" & vbCrLf
frame = r + c
frame = frame + "<frame src=" & source1 & " name=" & name1 & " marginwidth=" & mw1 & " marginheight=" & mh1 & " noresize>" & vbCrLf
c2 = "<frameset cols=""75%,25%"" border=0 frameborder=0 noresize>" & vbCrLf
frame = frame + c2
frame = frame + "<frame src=" & source2 & " name=" & name2 & " marginwidht=" & mw2 & " marginheight=" & mh2 & " noresize>" & vbCrLf
frame = frame + "<frame src=" & source3 & " name=" & name3 & " marginwidht=" & mw3 & " marginheight=" & mh3 & " noresize>" & vbCrLf
frame = frame + "</frameset>" & vbCrLf & "</frameset>" & vbCrLf & "</frameset>" & vbCrLf
fMainForm.ActiveForm.CodeMax1.SelText = frame
ElseIf opttoprl.Value = True Then
r = "<frameset rows=""25%,75%"" border=0 noresize>" & vbCrLf
c = "<frameset cols=""100%"" border=0 frameborder=0 scrolling=""no"" noresize>" & vbCrLf
frame = r + c
frame = frame + "<frame src=" & source1 & " name=" & name1 & " marginwidth=" & mw1 & " marginheight=" & mh1 & " noresize>" & vbCrLf
c2 = "<frameset cols=""25%,50%,25%"" border=0 frameborder=0 noresize>" & vbCrLf
frame = frame + c2
frame = frame + "<frame src=" & source2 & " name=" & name2 & " marginwidht=" & mw2 & " marginheight=" & mh2 & " noresize>" & vbCrLf
frame = frame + "<frame src=" & source3 & " name=" & name3 & " marginwidht=" & mw3 & " marginheight=" & mh3 & " noresize>" & vbCrLf
frame = frame + "<frame src=" & source4 & " name=" & name4 & " marginwidht=" & mw4 & " marginheight=" & mh4 & " noresize>" & vbCrLf
frame = frame + "</frameset>" & vbCrLf & "</frameset>" & vbCrLf & "</frameset>" & vbCrLf
fMainForm.ActiveForm.CodeMax1.SelText = frame
ElseIf optleft.Value = True Then
c = "<frameset cols=""100%"" border=0 frameborder=0 scrolling=""no"" noresize>" & vbCrLf
c2 = "<frameset cols=""25%,75%"" border=0 frameborder=0 noresize>" & vbCrLf
frame = c + c2
frame = frame + "<frame src=" & source1 & " name=" & name1 & " marginwidth=" & mw1 & " marginheight=" & mh1 & " noresize>" & vbCrLf
frame = frame + "<frame src=" & source2 & " name=" & name2 & " marginwidht=" & mw2 & " marginheight=" & mh2 & " noresize>" & vbCrLf
frame = frame + "</frameset>" & vbCrLf & "</frameset>" & vbCrLf
fMainForm.ActiveForm.CodeMax1.SelText = frame
ElseIf optright.Value = True Then
c = "<frameset cols=""100%"" border=0 frameborder=0 scrolling=""no"" noresize>" & vbCrLf
c2 = "<frameset cols=""75%,25%"" border=0 frameborder=0 noresize>" & vbCrLf
frame = c + c2
frame = frame + "<frame src=" & source1 & " name=" & name1 & " marginwidth=" & mw1 & " marginheight=" & mh1 & " noresize>" & vbCrLf
frame = frame + "<frame src=" & source2 & " name=" & name2 & " marginwidht=" & mw2 & " marginheight=" & mh2 & " noresize>" & vbCrLf
frame = frame + "</frameset>" & vbCrLf & "</frameset>" & vbCrLf
fMainForm.ActiveForm.CodeMax1.SelText = frame
ElseIf optbottom.Value = True Then
r = "<frameset rows=""75%,25%"" border=0 noresize>" & vbCrLf
c = "<frameset cols=""100%"" border=0 frameborder=0 scrolling=""no"" noresize>" & vbCrLf
frame = r + c
frame = frame + "<frame src=" & source1 & " name=" & name1 & " marginwidth=" & mw1 & " marginheight=" & mh1 & " noresize>" & vbCrLf
frame = frame + "<frame src=" & source2 & " name=" & name2 & " marginwidht=" & mw2 & " marginheight=" & mh2 & " noresize>" & vbCrLf
frame = frame + "</frameset>" & vbCrLf & "</frameset>" & vbCrLf
fMainForm.ActiveForm.CodeMax1.SelText = frame
ElseIf optbotleft.Value = True Then
r = "<frameset rows=""75%,25%"" border=0 noresize>" & vbCrLf
c = "<frameset cols=""25%,75%"" border=0 frameborder=0 scrolling=""no"" noresize>" & vbCrLf
frame = r + c
frame = frame + "<frame src=" & source1 & " name=" & name1 & " marginwidth=" & mw1 & " marginheight=" & mh1 & " noresize>" & vbCrLf
frame = frame + "<frame src=" & source2 & " name=" & name2 & " marginwidht=" & mw2 & " marginheight=" & mh2 & " noresize>" & vbCrLf
frame = frame + "</frameset>" & vbCrLf
frame = frame + "<frame src=" & source3 & " name=" & name3 & " marginwidht=" & mw3 & " marginheight=" & mh3 & " noresize>" & vbCrLf
frame = frame + "</frameset>" & vbCrLf
fMainForm.ActiveForm.CodeMax1.SelText = frame
ElseIf optbotright.Value = True Then
r = "<frameset rows=""75%,25%"" border=0 noresize>" & vbCrLf
c = "<frameset cols=""75%,25%"" border=0 frameborder=0 scrolling=""no"" noresize>" & vbCrLf
frame = r + c
frame = frame + "<frame src=" & source1 & " name=" & name1 & " marginwidth=" & mw1 & " marginheight=" & mh1 & " noresize>" & vbCrLf
frame = frame + "<frame src=" & source2 & " name=" & name2 & " marginwidht=" & mw2 & " marginheight=" & mh2 & " noresize>" & vbCrLf
frame = frame + "</frameset>" & vbCrLf
frame = frame + "<frame src=" & source3 & " name=" & name3 & " marginwidht=" & mw3 & " marginheight=" & mh3 & " noresize>" & vbCrLf
frame = frame + "</frameset>" & vbCrLf
fMainForm.ActiveForm.CodeMax1.SelText = frame
ElseIf optbotrl.Value = True Then
r = "<frameset rows=""75%,25%"" border=0 noresize>" & vbCrLf
c = "<frameset cols=""25%,50%,25%"" border=0 frameborder=0 scrolling=""no"" noresize>" & vbCrLf
frame = r + c
frame = frame + "<frame src=" & source1 & " name=" & name1 & " marginwidth=" & mw1 & " marginheight=" & mh1 & " noresize>" & vbCrLf
frame = frame + "<frame src=" & source2 & " name=" & name2 & " marginwidht=" & mw2 & " marginheight=" & mh2 & " noresize>" & vbCrLf
frame = frame + "<frame src=" & source3 & " name=" & name3 & " marginwidht=" & mw3 & " marginheight=" & mh3 & " noresize>" & vbCrLf
frame = frame + "</frameset>" & vbCrLf
frame = frame + "<frame src=" & source4 & " name=" & name4 & " marginwidht=" & mw4 & " marginheight=" & mh4 & " noresize>" & vbCrLf
frame = frame + "</frameset>" & vbCrLf
fMainForm.ActiveForm.CodeMax1.SelText = frame
ElseIf optbottop.Value = True Then
r = "<frameset rows=""25%,50%,25%"" border=0 noresize>" & vbCrLf
c = "<frameset cols=""100%"" border=0 frameborder=0 scrolling=""no"" noresize>" & vbCrLf
frame = r + c
frame = frame + "<frame src=" & source1 & " name=" & name1 & " marginwidth=" & mw1 & " marginheight=" & mh1 & " noresize>" & vbCrLf
frame = frame + "<frame src=" & source2 & " name=" & name2 & " marginwidht=" & mw2 & " marginheight=" & mh2 & " noresize>" & vbCrLf
frame = frame + "<frame src=" & source3 & " name=" & name3 & " marginwidht=" & mw3 & " marginheight=" & mh3 & " noresize>" & vbCrLf
frame = frame + "</frameset>" & vbCrLf & "</frameset>" & vbCrLf
fMainForm.ActiveForm.CodeMax1.SelText = frame
ElseIf optbrlt.Value = True Then
r = "<frameset rows=""25%,50%,25%"" border=0 noresize>" & vbCrLf
c = "<frameset cols=""100%"" border=0 frameborder=0 scrolling=""no"" noresize>" & vbCrLf
frame = r + c
frame = frame + "<frame src=" & source1 & " name=" & name1 & " marginwidth=" & mw1 & " marginheight=" & mh1 & " noresize>" & vbCrLf
c2 = "<frameset cols=""25%,50%,25%"" border=0 frameborder=0 noresize>" & vbCrLf
frame = frame + c2
frame = frame + "<frame src=" & source2 & " name=" & name2 & " marginwidht=" & mw2 & " marginheight=" & mh2 & " noresize>" & vbCrLf
frame = frame + "<frame src=" & source3 & " name=" & name3 & " marginwidht=" & mw3 & " marginheight=" & mh3 & " noresize>" & vbCrLf
frame = frame + "<frame src=" & source4 & " name=" & name4 & " marginwidht=" & mw4 & " marginheight=" & mh4 & " noresize>" & vbCrLf
frame = frame + "</frameset>" & vbCrLf
frame = frame + "<frame src=" & source5 & " name=" & name5 & " marginwidht=" & mw5 & " marginheight=" & mh5 & " noresize>" & vbCrLf
frame = frame + "</frameset>" & vbCrLf & "</frameset>" & vbCrLf
fMainForm.ActiveForm.CodeMax1.SelText = frame

End If

End Sub

Sub shows()
If opttop.Value = True Then
tops.Visible = True
sleft.Visible = False
sright.Visible = False
sbottom.Visible = False
lblsource1.Visible = True
txtsource1.Visible = True
lblname1.Visible = True
txtname1.Visible = True
lblmw1.Visible = True
txtmw1.Visible = True
lblmh1.Visible = True
txtmh1.Visible = True
lblsource2.Visible = True
txtsource2.Visible = True
lblname2.Visible = True
txtname2.Visible = True
lblmw2.Visible = True
txtmw2.Visible = True
lblmh2.Visible = True
txtmh2.Visible = True
lblsource3.Visible = False
txtsource3.Visible = False
lblname3.Visible = False
txtname3.Visible = False
lblmw3.Visible = False
txtmw3.Visible = False
lblmh3.Visible = False
txtmh3.Visible = False
lblsource4.Visible = False
txtsource4.Visible = False
lblname4.Visible = False
txtname4.Visible = False
lblmw4.Visible = False
txtmw4.Visible = False
lblmh4.Visible = False
txtmh4.Visible = False
lblsource5.Visible = False
txtsource5.Visible = False
lblname5.Visible = False
txtname5.Visible = False
lblmw5.Visible = False
txtmw5.Visible = False
lblmh5.Visible = False
txtmh5.Visible = False
ElseIf opttopleft.Value = True Then
tops.Visible = True
sleft.Visible = True
sright.Visible = False
sbottom.Visible = False
lblsource1.Visible = True
txtsource1.Visible = True
lblname1.Visible = True
txtname1.Visible = True
lblmw1.Visible = True
txtmw1.Visible = True
lblmh1.Visible = True
txtmh1.Visible = True
lblsource2.Visible = True
txtsource2.Visible = True
lblname2.Visible = True
txtname2.Visible = True
lblmw2.Visible = True
txtmw2.Visible = True
lblmh2.Visible = True
txtmh2.Visible = True
lblsource3.Visible = True
txtsource3.Visible = True
lblname3.Visible = True
txtname3.Visible = True
lblmw3.Visible = True
txtmw3.Visible = True
lblmh3.Visible = True
txtmh3.Visible = True
lblsource4.Visible = False
txtsource4.Visible = False
lblname4.Visible = False
txtname4.Visible = False
lblmw4.Visible = False
txtmw4.Visible = False
lblmh4.Visible = False
txtmh4.Visible = False
lblsource5.Visible = False
txtsource5.Visible = False
lblname5.Visible = False
txtname5.Visible = False
lblmw5.Visible = False
txtmw5.Visible = False
lblmh5.Visible = False
txtmh5.Visible = False
ElseIf opttopright.Value = True Then
tops.Visible = True
sright.Visible = True
sbottom.Visible = False
sleft.Visible = False
lblsource1.Visible = True
txtsource1.Visible = True
lblname1.Visible = True
txtname1.Visible = True
lblmw1.Visible = True
txtmw1.Visible = True
lblmh1.Visible = True
txtmh1.Visible = True
lblsource2.Visible = True
txtsource2.Visible = True
lblname2.Visible = True
txtname2.Visible = True
lblmw2.Visible = True
txtmw2.Visible = True
lblmh2.Visible = True
txtmh2.Visible = True
lblsource3.Visible = True
txtsource3.Visible = True
lblname3.Visible = True
txtname3.Visible = True
lblmw3.Visible = True
txtmw3.Visible = True
lblmh3.Visible = True
txtmh3.Visible = True
lblsource4.Visible = False
txtsource4.Visible = False
lblname4.Visible = False
txtname4.Visible = False
lblmw4.Visible = False
txtmw4.Visible = False
lblmh4.Visible = False
txtmh4.Visible = False
lblsource5.Visible = False
txtsource5.Visible = False
lblname5.Visible = False
txtname5.Visible = False
lblmw5.Visible = False
txtmw5.Visible = False
lblmh5.Visible = False
txtmh5.Visible = False
ElseIf opttoprl.Value = True Then
tops.Visible = True
sright.Visible = True
sleft.Visible = True
sbottom.Visible = False
lblsource1.Visible = True
txtsource1.Visible = True
lblname1.Visible = True
txtname1.Visible = True
lblmw1.Visible = True
txtmw1.Visible = True
lblmh1.Visible = True
txtmh1.Visible = True
lblsource2.Visible = True
txtsource2.Visible = True
lblname2.Visible = True
txtname2.Visible = True
lblmw2.Visible = True
txtmw2.Visible = True
lblmh2.Visible = True
txtmh2.Visible = True
lblsource3.Visible = True
txtsource3.Visible = True
lblname3.Visible = True
txtname3.Visible = True
lblmw3.Visible = True
txtmw3.Visible = True
lblmh3.Visible = True
txtmh3.Visible = True
lblsource4.Visible = True
txtsource4.Visible = True
lblname4.Visible = True
txtname4.Visible = True
lblmw4.Visible = True
txtmw4.Visible = True
lblmh4.Visible = True
txtmh4.Visible = True
lblsource5.Visible = False
txtsource5.Visible = False
lblname5.Visible = False
txtname5.Visible = False
lblmw5.Visible = False
txtmw5.Visible = False
lblmh5.Visible = False
txtmh5.Visible = False
ElseIf optleft.Value = True Then
sleft.Visible = True
sbottom.Visible = False
sright.Visible = False
tops.Visible = False
lblsource1.Visible = True
txtsource1.Visible = True
lblname1.Visible = True
txtname1.Visible = True
lblmw1.Visible = True
txtmw1.Visible = True
lblmh1.Visible = True
txtmh1.Visible = True
lblsource2.Visible = True
txtsource2.Visible = True
lblname2.Visible = True
txtname2.Visible = True
lblmw2.Visible = True
txtmw2.Visible = True
lblmh2.Visible = True
txtmh2.Visible = True
lblsource3.Visible = False
txtsource3.Visible = False
lblname3.Visible = False
txtname3.Visible = False
lblmw3.Visible = False
txtmw3.Visible = False
lblmh3.Visible = False
txtmh3.Visible = False
lblsource4.Visible = False
txtsource4.Visible = False
lblname4.Visible = False
txtname4.Visible = False
lblmw4.Visible = False
txtmw4.Visible = False
lblmh4.Visible = False
txtmh4.Visible = False
lblsource5.Visible = False
txtsource5.Visible = False
lblname5.Visible = False
txtname5.Visible = False
lblmw5.Visible = False
txtmw5.Visible = False
lblmh5.Visible = False
txtmh5.Visible = False
ElseIf optright.Value = True Then
sright.Visible = True
sbottom.Visible = False
sleft.Visible = False
tops.Visible = False
lblsource1.Visible = True
txtsource1.Visible = True
lblname1.Visible = True
txtname1.Visible = True
lblmw1.Visible = True
txtmw1.Visible = True
lblmh1.Visible = True
txtmh1.Visible = True
lblsource2.Visible = True
txtsource2.Visible = True
lblname2.Visible = True
txtname2.Visible = True
lblmw2.Visible = True
txtmw2.Visible = True
lblmh2.Visible = True
txtmh2.Visible = True
lblsource3.Visible = False
txtsource3.Visible = False
lblname3.Visible = False
txtname3.Visible = False
lblmw3.Visible = False
txtmw3.Visible = False
lblmh3.Visible = False
txtmh3.Visible = False
lblsource4.Visible = False
txtsource4.Visible = False
lblname4.Visible = False
txtname4.Visible = False
lblmw4.Visible = False
txtmw4.Visible = False
lblmh4.Visible = False
txtmh4.Visible = False
lblsource5.Visible = False
txtsource5.Visible = False
lblname5.Visible = False
txtname5.Visible = False
lblmw5.Visible = False
txtmw5.Visible = False
lblmh5.Visible = False
txtmh5.Visible = False
ElseIf optbottom.Value = True Then
sbottom.Visible = True
sright.Visible = False
sleft.Visible = False
tops.Visible = False
lblsource1.Visible = True
txtsource1.Visible = True
lblname1.Visible = True
txtname1.Visible = True
lblmw1.Visible = True
txtmw1.Visible = True
lblmh1.Visible = True
txtmh1.Visible = True
lblsource2.Visible = True
txtsource2.Visible = True
lblname2.Visible = True
txtname2.Visible = True
lblmw2.Visible = True
txtmw2.Visible = True
lblmh2.Visible = True
txtmh2.Visible = True
lblsource3.Visible = False
txtsource3.Visible = False
lblname3.Visible = False
txtname3.Visible = False
lblmw3.Visible = False
txtmw3.Visible = False
lblmh3.Visible = False
txtmh3.Visible = False
lblsource4.Visible = False
txtsource4.Visible = False
lblname4.Visible = False
txtname4.Visible = False
lblmw4.Visible = False
txtmw4.Visible = False
lblmh4.Visible = False
txtmh4.Visible = False
lblsource5.Visible = False
txtsource5.Visible = False
lblname5.Visible = False
txtname5.Visible = False
lblmw5.Visible = False
txtmw5.Visible = False
lblmh5.Visible = False
txtmh5.Visible = False
ElseIf optbotleft.Value = True Then
sbottom.Visible = True
sleft.Visible = True
sright.Visible = False
tops.Visible = False
lblsource1.Visible = True
txtsource1.Visible = True
lblname1.Visible = True
txtname1.Visible = True
lblmw1.Visible = True
txtmw1.Visible = True
lblmh1.Visible = True
txtmh1.Visible = True
lblsource2.Visible = True
txtsource2.Visible = True
lblname2.Visible = True
txtname2.Visible = True
lblmw2.Visible = True
txtmw2.Visible = True
lblmh2.Visible = True
txtmh2.Visible = True
lblsource3.Visible = True
txtsource3.Visible = True
lblname3.Visible = True
txtname3.Visible = True
lblmw3.Visible = True
txtmw3.Visible = True
lblmh3.Visible = True
txtmh3.Visible = True
lblsource4.Visible = False
txtsource4.Visible = False
lblname4.Visible = False
txtname4.Visible = False
lblmw4.Visible = False
txtmw4.Visible = False
lblmh4.Visible = False
txtmh4.Visible = False
lblsource5.Visible = False
txtsource5.Visible = False
lblname5.Visible = False
txtname5.Visible = False
lblmw5.Visible = False
txtmw5.Visible = False
lblmh5.Visible = False
txtmh5.Visible = False
ElseIf optbotright.Value = True Then
sbottom.Visible = True
sright.Visible = True
sleft.Visible = False
tops.Visible = False
lblsource1.Visible = True
txtsource1.Visible = True
lblname1.Visible = True
txtname1.Visible = True
lblmw1.Visible = True
txtmw1.Visible = True
lblmh1.Visible = True
txtmh1.Visible = True
lblsource2.Visible = True
txtsource2.Visible = True
lblname2.Visible = True
txtname2.Visible = True
lblmw2.Visible = True
txtmw2.Visible = True
lblmh2.Visible = True
txtmh2.Visible = True
lblsource3.Visible = True
txtsource3.Visible = True
lblname3.Visible = True
txtname3.Visible = True
lblmw3.Visible = True
txtmw3.Visible = True
lblmh3.Visible = True
txtmh3.Visible = True
lblsource4.Visible = False
txtsource4.Visible = False
lblname4.Visible = False
txtname4.Visible = False
lblmw4.Visible = False
txtmw4.Visible = False
lblmh4.Visible = False
txtmh4.Visible = False
lblsource5.Visible = False
txtsource5.Visible = False
lblname5.Visible = False
txtname5.Visible = False
lblmw5.Visible = False
txtmw5.Visible = False
lblmh5.Visible = False
txtmh5.Visible = False
ElseIf optbottop.Value = True Then
sbottom.Visible = True
tops.Visible = True
sleft.Visible = False
sright.Visible = False
lblsource1.Visible = True
txtsource1.Visible = True
lblname1.Visible = True
txtname1.Visible = True
lblmw1.Visible = True
txtmw1.Visible = True
lblmh1.Visible = True
txtmh1.Visible = True
lblsource2.Visible = True
txtsource2.Visible = True
lblname2.Visible = True
txtname2.Visible = True
lblmw2.Visible = True
txtmw2.Visible = True
lblmh2.Visible = True
txtmh2.Visible = True
lblsource3.Visible = True
txtsource3.Visible = True
lblname3.Visible = True
txtname3.Visible = True
lblmw3.Visible = True
txtmw3.Visible = True
lblmh3.Visible = True
txtmh3.Visible = True
lblsource4.Visible = False
txtsource4.Visible = False
lblname4.Visible = False
txtname4.Visible = False
lblmw4.Visible = False
txtmw4.Visible = False
lblmh4.Visible = False
txtmh4.Visible = False
lblsource5.Visible = False
txtsource5.Visible = False
lblname5.Visible = False
txtname5.Visible = False
lblmw5.Visible = False
txtmw5.Visible = False
lblmh5.Visible = False
txtmh5.Visible = False
ElseIf optbotrl.Value = True Then
sbottom.Visible = True
sleft.Visible = True
sright.Visible = True
tops.Visible = False
lblsource1.Visible = True
txtsource1.Visible = True
lblname1.Visible = True
txtname1.Visible = True
lblmw1.Visible = True
txtmw1.Visible = True
lblmh1.Visible = True
txtmh1.Visible = True
lblsource2.Visible = True
txtsource2.Visible = True
lblname2.Visible = True
txtname2.Visible = True
lblmw2.Visible = True
txtmw2.Visible = True
lblmh2.Visible = True
txtmh2.Visible = True
lblsource3.Visible = True
txtsource3.Visible = True
lblname3.Visible = True
txtname3.Visible = True
lblmw3.Visible = True
txtmw3.Visible = True
lblmh3.Visible = True
txtmh3.Visible = True
lblsource4.Visible = True
txtsource4.Visible = True
lblname4.Visible = True
txtname4.Visible = True
lblmw4.Visible = True
txtmw4.Visible = True
lblmh4.Visible = True
txtmh4.Visible = True
lblsource5.Visible = False
txtsource5.Visible = False
lblname5.Visible = False
txtname5.Visible = False
lblmw5.Visible = False
txtmw5.Visible = False
lblmh5.Visible = False
txtmh5.Visible = False
ElseIf optbrlt.Value = True Then
sbottom.Visible = True
sleft.Visible = True
sright.Visible = True
tops.Visible = True
lblsource1.Visible = True
txtsource1.Visible = True
lblname1.Visible = True
txtname1.Visible = True
lblmw1.Visible = True
txtmw1.Visible = True
lblmh1.Visible = True
txtmh1.Visible = True
lblsource2.Visible = True
txtsource2.Visible = True
lblname2.Visible = True
txtname2.Visible = True
lblmw2.Visible = True
txtmw2.Visible = True
lblmh2.Visible = True
txtmh2.Visible = True
lblsource3.Visible = True
txtsource3.Visible = True
lblname3.Visible = True
txtname3.Visible = True
lblmw3.Visible = True
txtmw3.Visible = True
lblmh3.Visible = True
txtmh3.Visible = True
lblsource4.Visible = True
txtsource4.Visible = True
lblname4.Visible = True
txtname4.Visible = True
lblmw4.Visible = True
txtmw4.Visible = True
lblmh4.Visible = True
txtmh4.Visible = True
lblsource5.Visible = True
txtsource5.Visible = True
lblname5.Visible = True
txtname5.Visible = True
lblmw5.Visible = True
txtmw5.Visible = True
lblmh5.Visible = True
txtmh5.Visible = True
End If

End Sub


Private Sub cmdcancel_Click()
Unload Me

End Sub

Private Sub cmdinsert_Click()
getinfo
makeframes
Unload Me

End Sub

Private Sub optbotleft_Click()
shows

End Sub

Private Sub optbotright_Click()
shows

End Sub


Private Sub optbotrl_Click()
shows

End Sub

Private Sub optbottom_Click()
shows

End Sub

Private Sub optbottop_Click()
shows

End Sub

Private Sub optbrlt_Click()
shows

End Sub

Private Sub optleft_Click()
shows

End Sub

Private Sub optright_Click()
shows

End Sub


Private Sub opttop_Click()
shows

End Sub


Private Sub opttopleft_Click()
shows

End Sub


Private Sub opttopright_Click()
shows

End Sub


Private Sub opttoprl_Click()
shows

End Sub


