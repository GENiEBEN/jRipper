VERSION 5.00
Object = "{ECEDB943-AC41-11D2-AB20-000000000000}#2.0#0"; "CBOX.OCX"
Object = "{D1558013-91A7-11D4-AA5B-00A0CC334D72}#2.0#0"; "WWTabs.ocx"
Begin VB.Form frmscripts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Script Editor"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "frmscripts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin CodeMaxCtl.CodeMax cssmax 
      Height          =   495
      Left            =   1920
      OleObjectBlob   =   "frmscripts.frx":030A
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin CodeMaxCtl.CodeMax cgimax 
      Height          =   495
      Left            =   1680
      OleObjectBlob   =   "frmscripts.frx":046C
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin CodeMaxCtl.CodeMax jsmax 
      Height          =   495
      Left            =   120
      OleObjectBlob   =   "frmscripts.frx":05CE
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin CodeMaxCtl.CodeMax jspmax 
      Height          =   495
      Left            =   1440
      OleObjectBlob   =   "frmscripts.frx":0730
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin CodeMaxCtl.CodeMax phpmax 
      Height          =   615
      Left            =   720
      OleObjectBlob   =   "frmscripts.frx":0892
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin CodeMaxCtl.CodeMax aspmax 
      Height          =   735
      Left            =   240
      OleObjectBlob   =   "frmscripts.frx":09F4
      TabIndex        =   8
      Top             =   240
      Width           =   615
   End
   Begin WWTabs.WTabs WTabs1 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   4200
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionTips     =   "|||||"
      Captions        =   "ASP|PHP|JSP|Java Script| CGI/Perl| CSS"
   End
   Begin VB.CommandButton cmdcssinsert 
      Caption         =   "Insert CSS"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdcgiinsert 
      Caption         =   "Insert CGI"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdjsinsert 
      Caption         =   "Insert Java"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdjspinsert 
      Caption         =   "Insert JSP"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdphpinsert 
      Caption         =   "Insert PHP"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdaspinsert 
      Caption         =   "Insert ASP"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   4320
      Width           =   975
   End
End
Attribute VB_Name = "frmscripts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim scripts As String

Private Sub cmdaspinsert_Click()
asp = aspmax.Text
frmMain.ActiveForm.CodeMax1.SelText = asp
Unload Me

End Sub


Private Sub cmdcancel_Click()
Unload Me

End Sub

Private Sub cmdcgiinsert_Click()
cgi = cgimax.Text
frmMain.ActiveForm.CodeMax1.SelText = cgi
Unload Me

End Sub

Private Sub cmdcssinsert_Click()
css = cssmax.Text
frmMain.ActiveForm.CodeMax1.SelText = css
Unload Me

End Sub


Private Sub cmdjsinsert_Click()
js = jsmax.Text
frmMain.ActiveForm.CodeMax1.SelText = js
Unload Me

End Sub

Private Sub cmdjspinsert_Click()
jsp = jspmax.Text
frmMain.ActiveForm.CodeMax1.SelText = jsp
Unload Me

End Sub

Private Sub cmdphpinsert_Click()
php = phpmax.Text
frmMain.ActiveForm.CodeMax1.SelText = ph
Unload Me

End Sub

Private Sub Form_Load()
scripts = ReadINI("Options", "Scripts", inifile)
aspmax.Visible = True
cmdaspinsert.Visible = True
aspmax.Height = 4200
aspmax.Width = 7215
aspmax.Top = 0
aspmax.Left = 0
aspmax.Text = "<%@ Language=" & Chr(34) & "VBScript" & Chr(34) & "%>" & vbCrLf & "<%" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "%>"
aspmax.Language = "HTML"
aspmax.ColorSyntax = scripts
phpmax.Height = 4200
phpmax.Width = 7215
phpmax.Top = 0
phpmax.Left = 0
phpmax.Text = "<?php" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "?>"
phpmax.Language = "HTML"
phpmax.ColorSyntax = scripts
jspmax.Height = 4200
jspmax.Width = 7215
jspmax.Top = 0
jspmax.Left = 0
jspmax.Text = "<%@ page language=" & Chr(34) & "java" & Chr(34) & " %>" & vbCrLf & "<%" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "%>" & vbCrLf
jspmax.Language = "HTML"
jspmax.ColorSyntax = scripts
jsmax.Height = 4200
jsmax.Width = 7215
jsmax.Top = 0
jsmax.Left = 0
jsmax.Text = "<Script Language=" & Chr(34) & "JavaScript" & Chr(34) & ">" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "</script>" & vbCrLf
jsmax.Language = "HTML"
jsmax.ColorSyntax = scripts
cgimax.Height = 4200
cgimax.Width = 7215
cgimax.Top = 0
cgimax.Left = 0
cgimax.Text = "#!/usr/bin/perl" & vbCrLf
cgimax.Language = "HTML"
cgimax.ColorSyntax = scripts
cssmax.Height = 4200
cssmax.Width = 7215
cssmax.Top = 0
cssmax.Left = 0
cssmax.Text = "<style>" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "</style>" & vbCrLf
cssmax.Language = "HTML"
cssmax.ColorSyntax = scripts
cmdaspinsert.Top = 4320
cmdphpinsert.Top = 4320
cmdjspinsert.Top = 4320
cmdjsinsert.Top = 4320
cmdcgiinsert.Top = 4320
cmdcssinsert.Top = 4320
End Sub

Private Sub WTabs1_Click(ByVal ActualClick As Boolean)
If WTabs1.ActiveTab = 0 Then
aspmax.Visible = True
phpmax.Visible = False
jspmax.Visible = False
jsmax.Visible = False
cgimax.Visible = False
cssmax.Visible = False
cmdaspinsert.Visible = True
cmdphpinsert.Visible = False
cmdjspinsert.Visible = False
cmdjsinsert.Visible = False
cmdcgiinsert.Visible = False
cmdcssinsert.Visible = False
End If

If WTabs1.ActiveTab = 1 Then
aspmax.Visible = False
phpmax.Visible = True
jspmax.Visible = False
jsmax.Visible = False
cgimax.Visible = False
cssmax.Visible = False
cmdaspinsert.Visible = False
cmdphpinsert.Visible = True
cmdjspinsert.Visible = False
cmdjsinsert.Visible = False
cmdcgiinsert.Visible = False
cmdcssinsert.Visible = False
End If

If WTabs1.ActiveTab = 2 Then
aspmax.Visible = False
phpmax.Visible = False
jspmax.Visible = True
jsmax.Visible = False
cgimax.Visible = False
cssmax.Visible = False
cmdaspinsert.Visible = False
cmdphpinsert.Visible = False
cmdjspinsert.Visible = True
cmdjsinsert.Visible = False
cmdcgiinsert.Visible = False
cmdcssinsert.Visible = False
End If

If WTabs1.ActiveTab = 3 Then
aspmax.Visible = False
phpmax.Visible = False
jspmax.Visible = False
jsmax.Visible = True
cgimax.Visible = False
cssmax.Visible = False
cmdaspinsert.Visible = False
cmdphpinsert.Visible = False
cmdjspinsert.Visible = False
cmdjsinsert.Visible = True
cmdcgiinsert.Visible = False
cmdcssinsert.Visible = False
End If

If WTabs1.ActiveTab = 4 Then
aspmax.Visible = False
phpmax.Visible = False
jspmax.Visible = False
jsmax.Visible = False
cgimax.Visible = True
cssmax.Visible = False
cmdaspinsert.Visible = False
cmdphpinsert.Visible = False
cmdjspinsert.Visible = False
cmdjsinsert.Visible = False
cmdcgiinsert.Visible = True
cmdcssinsert.Visible = False
End If

If WTabs1.ActiveTab = 5 Then
aspmax.Visible = False
phpmax.Visible = False
jspmax.Visible = False
jsmax.Visible = False
cgimax.Visible = False
cssmax.Visible = True
cmdaspinsert.Visible = False
cmdphpinsert.Visible = False
cmdjspinsert.Visible = False
cmdjsinsert.Visible = False
cmdcgiinsert.Visible = False
cmdcssinsert.Visible = True
End If
End Sub


