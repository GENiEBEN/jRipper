VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H004D483F&
   Caption         =   "HTML Edit"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageKey        =   "Undo"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Object.ToolTipText     =   "Redo"
            ImageKey        =   "Redo"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "About"
            ImageKey        =   "Help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   8040
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12726
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "11.09.2006"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "16:28"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1050
      Top             =   630
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   225
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":041C
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":052E
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0640
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0752
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0864
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0976
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A88
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B9A
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CAC
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufilenew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnufileclose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditredo 
         Caption         =   "&Redo"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnumnuinsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuinsertfont 
         Caption         =   "Font"
      End
      Begin VB.Menu mnuinsertheading 
         Caption         =   "Heading"
         Begin VB.Menu mnuinsertheadingh1 
            Caption         =   "H1"
         End
         Begin VB.Menu mnuinsertheadingh2 
            Caption         =   "H2"
         End
         Begin VB.Menu mnuinsertheadingh3 
            Caption         =   "H3"
         End
         Begin VB.Menu mnuinsertheadingh4 
            Caption         =   "H4"
         End
         Begin VB.Menu mnuinsertheadingh5 
            Caption         =   "H5"
         End
         Begin VB.Menu mnuinsertheadingh6 
            Caption         =   "H6"
         End
      End
      Begin VB.Menu mnutags 
         Caption         =   "Tags"
         Begin VB.Menu mnutagsunderline 
            Caption         =   "Underline"
         End
         Begin VB.Menu mnutagsbold 
            Caption         =   "Bold"
         End
         Begin VB.Menu mnutagsitalic 
            Caption         =   "Italicized"
         End
         Begin VB.Menu mnutagsdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnutagslinebreak 
            Caption         =   "Line Break"
         End
         Begin VB.Menu mnutagshr 
            Caption         =   "Horizontal Rule"
         End
         Begin VB.Menu mnutagsparagraph 
            Caption         =   "Paragraph"
         End
         Begin VB.Menu mnutagsimage 
            Caption         =   "Image"
         End
         Begin VB.Menu mnutagsdivision 
            Caption         =   "Division"
         End
         Begin VB.Menu mnutagsspan 
            Caption         =   "Span"
         End
         Begin VB.Menu mnutagscomment 
            Caption         =   "Comment"
         End
         Begin VB.Menu mnutagsdash1 
            Caption         =   "-"
         End
         Begin VB.Menu mnutagslink 
            Caption         =   "Link"
         End
         Begin VB.Menu mnutagsimagelink 
            Caption         =   "Image Link"
         End
         Begin VB.Menu mnutagsanchor 
            Caption         =   "Anchor"
         End
         Begin VB.Menu mnutagsdash2 
            Caption         =   "-"
         End
         Begin VB.Menu mnutagssuper 
            Caption         =   "Superscript"
         End
         Begin VB.Menu mnutagssub 
            Caption         =   "Subscript"
         End
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "T&ools"
      Begin VB.Menu mnuscriptsedit 
         Caption         =   "Script Editor"
      End
      Begin VB.Menu mnutoolsdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuviewoptions 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu mnuwizards 
      Caption         =   "W&izards"
      Begin VB.Menu mnuwizardstable 
         Caption         =   "Table"
      End
      Begin VB.Menu mnuwizardsframes 
         Caption         =   "Frames"
      End
      Begin VB.Menu mnuwizardsmarquee 
         Caption         =   "Marquee"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnunewWindow 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuhelpcontents 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnuhelpdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Public lDocumentCount As Long

Private Sub MDIForm_Load()
    Me.Left = GetSetting(App.title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.title, "Settings", "MainHeight", 6500)
    LoadNewDoc
End Sub

Private Sub LoadNewDoc()
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Caption = "Untitled " & lDocumentCount
Dim nd As String
nd = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.01 Transitional//EN" & Chr(34) & ">" & vbCrLf
nd = nd + "<html>" & vbCrLf
nd = nd + "<head>" & vbCrLf
nd = nd + "<title>Untitled Document</title>" & vbCrLf
nd = nd + "<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=UTF-8" & Chr(34) & ">" & vbCrLf
nd = nd + "</head>" & vbCrLf
nd = nd + "<body>" & vbCrLf & vbCrLf & vbCrLf & vbCrLf
nd = nd + "</body>" & vbCrLf
nd = nd + "</html>"
frmD.CodeMax1.Text = nd
frmD.Show

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.title, "Settings", "MainLeft", Me.Left
        SaveSetting App.title, "Settings", "MainTop", Me.Top
        SaveSetting App.title, "Settings", "MainWidth", Me.Width
        SaveSetting App.title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub mnuEditredo_Click()
ActiveForm.CodeMax1.Redo

End Sub

Private Sub mnuhelpcontents_Click()
HHShowContents App.Path & "\Help\html edit.chm", Me.hwnd
End Sub

Private Sub mnuinsertfont_Click()
frmfont.Show
End Sub

Private Sub mnuinsertheadingh1_Click()
ActiveForm.CodeMax1.SelText = "<h1> </h1>"

End Sub

Private Sub mnuinsertheadingh2_Click()
ActiveForm.CodeMax1.SelText = "<h2> </h2>"

End Sub


Private Sub mnuinsertheadingh3_Click()
ActiveForm.CodeMax1.SelText = "<h3> </h3>"

End Sub


Private Sub mnuinsertheadingh4_Click()
ActiveForm.CodeMax1.SelText = "<h4> </h4>"

End Sub

Private Sub mnuinsertheadingh5_Click()
ActiveForm.CodeMax1.SelText = "<h5> </h5>"

End Sub

Private Sub mnuinsertheadingh6_Click()
ActiveForm.CodeMax1.SelText = "<h6> </h6>"

End Sub

Private Sub mnunewWindow_Click()
LoadNewDoc

End Sub

Private Sub mnuscriptsedit_Click()
frmscripts.Show

End Sub

Private Sub mnutagsanchor_Click()
ActiveForm.CodeMax1.SelText = "<a href=" & Chr(34) & "#anchorname" & Chr(34) & ">text here</a>" & vbCrLf & vbCrLf & "<a name=" & Chr(34) & "anchorname" & Chr(34) & ">"

End Sub

Private Sub mnutagsbold_Click()
ActiveForm.CodeMax1.SelText = "<b> </b>"

End Sub

Private Sub mnutagscomment_Click()
ActiveForm.CodeMax1.SelText = "<!--comment here-->"

End Sub

Private Sub mnutagsdivision_Click()
ActiveForm.CodeMax1.SelText = "<div> </div>"

End Sub

Private Sub mnutagshr_Click()
ActiveForm.CodeMax1.SelText = "<hr>"

End Sub

Private Sub mnutagsimage_Click()
ActiveForm.CodeMax1.SelText = "<img src=" & Chr(34) & "image.gif" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & ">"

End Sub

Private Sub mnutagsimagelink_Click()
ActiveForm.CodeMax1.SelText = "<a href=" & Chr(34) & "page.html" & Chr(34) & "><img src=" & Chr(34) & "image.gif" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></a>"

End Sub

Private Sub mnutagsitalic_Click()
ActiveForm.CodeMax1.SelText = "<i> </i>"

End Sub


Private Sub mnutagslinebreak_Click()
ActiveForm.CodeMax1.SelText = "<br>"

End Sub

Private Sub mnutagslink_Click()
ActiveForm.CodeMax1.SelText = "<a href=" & Chr(34) & "page.html" & Chr(34) & ">click here</a>"

End Sub

Private Sub mnutagsparagraph_Click()
ActiveForm.CodeMax1.SelText = "<p> </p>"

End Sub

Private Sub mnutagsspan_Click()
ActiveForm.CodeMax1.SelText = "<span> </span>"
End Sub

Private Sub mnutagssub_Click()
ActiveForm.CodeMax1.SelText = "<sub> </sub>"

End Sub

Private Sub mnutagssuper_Click()
ActiveForm.CodeMax1.SelText = "<sup> </sup>"

End Sub

Private Sub mnutagsunderline_Click()
ActiveForm.CodeMax1.SelText = "<u> </u>"

End Sub









Private Sub mnuviewoptions_Click()
frmoptions.Show

End Sub

Private Sub mnuwizardsframes_Click()
frmframes.Show

End Sub

Private Sub mnuwizardsmarquee_Click()
frmmarq.Show

End Sub


Private Sub mnuwizardstable_Click()
frmtable.Show

End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            LoadNewDoc
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Undo"
            mnuEditUndo_Click
        Case "Redo"
            mnuEditredo_Click
        Case "Help"
            mnuhelpcontents_Click
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show

End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub


Private Sub mnuEditPaste_Click()
    On Error Resume Next
    ActiveForm.CodeMax1.SelText = Clipboard.GetText

End Sub

Private Sub mnuEditCopy_Click()
Clipboard.SetText ActiveForm.CodeMax1.SelText

End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.CodeMax1.SelRTF
    ActiveForm.CodeMax1.SelText = vbNullString

End Sub


Private Sub mnuEditUndo_Click()
ActiveForm.CodeMax1.Undo

End Sub


Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFilePrint_Click()
On Error Resume Next
ActiveForm.CodeMax1.PrintContents 0, cmPrnDefaultPrn + cmPrnRichFonts

End Sub

Private Sub mnuFilePageSetup_Click()
  On Error Resume Next
  ActiveForm.CodeMax1.PrintContents 0, cmPrnColor + cmPrnRichFonts

End Sub



Private Sub mnuFileSaveAs_Click()
Dim filenumber
On Error Resume Next
filenumber = FreeFile
dlgCommonDialog.InitDir = App.Path + "\Projects\"
dlgCommonDialog.Filter = "HTML Files (*.html;*.htm)|*.html;*.htm|ASP Files(*.asp;*.aspx;*.asa;*.asax)|*.asp;*.aspx;*.asa;*.asax|PHP Files(*.php;*.php3;*.php4)|*.php;*.php3;*.php4|Style Sheets(*.css)|*.css|Java Script Pages(*.jsp;*.js)|*.jsp;*.js|CGI/Perl(*.cgi;*.pl;*.pm)|*.cgi;*.pl;*.pm|Xml files(*.xml;*.xsl;*.xsd)|*.xml;*.xsl;*.xsd|SHTML files(*.shtml;*.shtm)|*.shtml;*.shtm|DHTML files(*.dhtml)|*.dhtml|Template files(*.tpl)|*.tpl|Configuration file(*.cfg)|*.cfg|All files(*.*)|*.*"
dlgCommonDialog.ShowSave
Open dlgCommonDialog.Filename For Output As #filenumber
Print #filenumber, ActiveForm.CodeMax1.Text
Close #filenumber

ActiveForm.Caption = dlgCommonDialog.FileTitle
ActiveForm.CodeMax1.ToolTipText = dlgCommonDialog.Filename


End Sub

Private Sub mnuFileSave_Click()
dlgCommonDialog.Filename = ActiveForm.CodeMax1.ToolTipText

If ActiveForm.CodeMax1.ToolTipText = "" Then
mnuFileSaveAs_Click
Else
Open dlgCommonDialog.Filename For Output As #1
Print #1, ActiveForm.CodeMax1.Text
Close #1
End If


End Sub

Private Sub mnuFileClose_Click()
Unload ActiveForm
End Sub

Private Sub mnuFileOpen_Click()
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    
Dim filenumber
On Error Resume Next
filenumber = FreeFile
dlgCommonDialog.InitDir = App.Path + "\Projects\"
dlgCommonDialog.Filter = "HTML Files (*.html;*.htm)|*.html;*.htm|ASP Files(*.asp;*.aspx;*.asa;*.asax)|*.asp;*.aspx;*.asa;*.asax|PHP Files(*.php;*.php3;*.php4)|*.php;*.php3;*.php4|Style Sheets(*.css)|*.css|Java Script Pages(*.jsp;*.js)|*.jsp;*.js|CGI/Perl(*.cgi;*.pl;*.pm)|*.cgi;*.pl;*.pm|Xml files(*.xml;*.xsl;*.xsd)|*.xml;*.xsl;*.xsd|SHTML files(*.shtml;*.shtm)|*.shtml;*.shtm|DHTML files(*.dhtml)|*.dhtml|Template files(*.tpl)|*.tpl|Configuration file(*.cfg)|*.cfg|All files(*.*)|*.*"
dlgCommonDialog.ShowOpen
          
   Open dlgCommonDialog.Filename For Input As #filenumber
   frmD.CodeMax1.Text = Input(LOF(filenumber), #filenumber)
   frmD.Caption = dlgCommonDialog.FileTitle
   Close

frmD.CodeMax1.ToolTipText = dlgCommonDialog.Filename
frmD.Show

End Sub

Private Sub mnuFileNew_Click()
    LoadNewDoc
End Sub

