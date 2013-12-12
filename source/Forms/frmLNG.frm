VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLNG 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "  TOCA Language Editor"
   ClientHeight    =   6405
   ClientLeft      =   6750
   ClientTop       =   5070
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   375
      Left            =   345
      TabIndex        =   2
      Top             =   4980
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lngData 
      Height          =   4305
      Left            =   180
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   495
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   7594
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList iconz 
      Left            =   5760
      Top             =   1470
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLNG.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLNG.frx":0257
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLNG.frx":04B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLNG.frx":089E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLNG.frx":0AD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLNG.frx":0D02
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLNG.frx":0F32
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLNG.frx":1190
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLNG.frx":157D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLNG.frx":17DF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Tools 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   1005
      ButtonWidth     =   529
      ButtonHeight    =   503
      Appearance      =   1
      Style           =   1
      ImageList       =   "iconz"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save Changes"
            Object.Tag             =   "Save"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "lng"
                  Text            =   "Export to LNG"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "ini"
                  Text            =   "Export to INI"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "txt"
                  Text            =   "Export to TXT"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label path 
      Caption         =   "path"
      Height          =   210
      Left            =   1950
      TabIndex        =   1
      Top             =   5385
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmLNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
lngData.Top = Tools.Height
lngData.Left = -(Screen.TwipsPerPixelX * 1)
lngData.Height = frmLNG.ScaleHeight - Tools.Height
lngData.Width = frmLNG.ScaleWidth + (Screen.TwipsPerPixelX * 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lngData_Click()
lngData_KeyPress 13
End Sub

Private Sub lngData_ItemClick(ByVal Item As MSComctlLib.ListItem)
'lngData.ToolTipText = lngData.SelectedItem.text
End Sub

Private Sub lngData_KeyPress(KeyAscii As Integer)
On Error GoTo quit
If KeyAscii = 13 Then
If Not frmLNG.lngData.SelectedItem.Text = "<unused>" Then
frmLNG.lngData.StartLabelEdit
End If
End If
Exit Sub
quit:
End Sub

Private Sub lngData_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
If lngData.HitTest(x, Y) Is Nothing Then
lngData.ToolTipText = ""
Else
lngData.ToolTipText = lngData.HitTest(x, Y).Text & " ID: " & lngData.HitTest(x, Y).Index
End If
End Sub

Private Sub Tools_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Tag = "Save" Then
LNG.LNG_saveasLNG Me.Path.Caption, Me.lngData
End If
End Sub

Private Sub Tools_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If ButtonMenu.Tag = "ini" Then
LNG.LNG_saveasINI Replace(frmLNG.Path.Caption, ".lng", ".ini"), "jRipper", frmLNG.lngData, frmLNG.pb1
ElseIf ButtonMenu.Tag = "txt" Then
LNG.LNG_saveasTXT Replace(frmLNG.Path.Caption, ".lng", ".txt"), frmLNG.lngData, frmLNG.pb1
End If
End Sub
