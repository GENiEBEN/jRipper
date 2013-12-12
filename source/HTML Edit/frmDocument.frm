VERSION 5.00
Object = "{ECEDB943-AC41-11D2-AB20-000000000000}#2.0#0"; "CBOX.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{D1558013-91A7-11D4-AA5B-00A0CC334D72}#2.0#0"; "WWTabs.ocx"
Begin VB.Form frmDocument 
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11970
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cd 
      Left            =   4560
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin CodeMaxCtl.CodeMax CodeMax1 
      Height          =   6975
      Left            =   0
      OleObjectBlob   =   "frmDocument.frx":030A
      TabIndex        =   2
      Top             =   0
      Width           =   11955
   End
   Begin WWTabs.WTabs WTabs1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8325
      Width           =   1740
      _ExtentX        =   3069
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
      CaptionTips     =   "|"
      Captions        =   "Normal|Preview"
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   495
      Left            =   2985
      TabIndex        =   0
      Top             =   7080
      Visible         =   0   'False
      Width           =   735
      ExtentX         =   1296
      ExtentY         =   873
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public UndoStack As New Collection
Public RedoStack As New Collection
Public trapUndo As Boolean
Const EM_UNDO = &HC7
Public SelStart As Long
Public TextLen As Long
Public Text As String
Public bChanged As Boolean
Public bUpdateFlag As Boolean
Public Positions As Collection


Private Function Change(ByVal lParam1 As String, ByVal lParam2 As String, startSearch As Long) As String
On Error Resume Next
Dim tempParam$
Dim d&
    If Len(lParam1) > Len(lParam2) Then 'swap
        tempParam$ = lParam1
        lParam1 = lParam2
        lParam2 = tempParam$
    End If
    d& = Len(lParam2) - Len(lParam1)
    Change = Mid(lParam2, startSearch - d&, d&)
End Function
Sub exitsave()
Dim filenumber
On Error Resume Next
filenumber = FreeFile
cd.InitDir = App.Path + "\Projects\"
cd.Filter = "HTML Files (*.html;*.htm)|*.html;*.htm|ASP Files(*.asp;*.aspx;*.asa;*.asax)|*.asp;*.aspx;*.asa;*.asax|PHP Files(*.php;*.php3;*.php4)|*.php;*.php3;*.php4|Style Sheets(*.css)|*.css|Java Script Pages(*.jsp;*.js)|*.jsp;*.js|CGI/Perl(*.cgi;*.pl;*.pm)|*.cgi;*.pl;*.pm|Xml files(*.xml;*.xsl;*.xsd)|*.xml;*.xsl;*.xsd|SHTML files(*.shtml;*.shtm)|*.shtml;*.shtm|DHTML files(*.dhtml)|*.dhtml|Template files(*.tpl)|*.tpl|Configuration file(*.cfg)|*.cfg|All files(*.*)|*.*"
cd.ShowSave
Open cd.Filename For Output As #filenumber
Print #filenumber, CodeMax1.Text
Close

End Sub

Sub ResetUndo()
Dim i As Long
For i = 1 To UndoStack.Count
UndoStack.Remove i
Next i
For i = 1 To RedoStack.Count
RedoStack.Remove i
Next i
End Sub
Public Sub Undo()
On Error Resume Next
Dim chg$, x&
Dim DeleteFlag As Boolean 'flag as to whether or not to delete text or append text
Dim objElement As Object, objElement2 As Object

    If UndoStack.Count > 1 And trapUndo Then 'we can proceed
        trapUndo = False
        DeleteFlag = UndoStack(UndoStack.Count - 1).TextLen < UndoStack(UndoStack.Count).TextLen
        If DeleteFlag Then  'delete some text
            x& = SendMessage(rtftext.hwnd, EM_UNDO, 1&, 1&)
            Set objElement = UndoStack(UndoStack.Count)
            Set objElement2 = UndoStack(UndoStack.Count - 1)
            rtftext.SelStart = objElement.SelStart - (objElement.TextLen - objElement2.TextLen)
            rtftext.SelLength = objElement.TextLen - objElement2.TextLen
            rtftext.SelText = ""
            x& = SendMessage(rtftext.hwnd, EM_UNDO, 0&, 0&)
        Else 'append something
            Set objElement = UndoStack(UndoStack.Count - 1)
            Set objElement2 = UndoStack(UndoStack.Count)
            chg$ = Change(objElement.Text, objElement2.Text, objElement2.SelStart + 1 + Abs(Len(objElement.Text) - Len(objElement2.Text)))
            rtftext.SelStart = objElement2.SelStart
            rtftext.SelLength = 0
            rtftext.SelText = chg$
            rtftext.SelStart = objElement2.SelStart
            If Len(chg$) > 1 And chg$ <> vbCrLf Then
                rtftext.SelLength = Len(chg$)
            Else
                rtftext.SelStart = rtftext.SelStart + Len(chg$)
            End If
        End If
        RedoStack.Add UndoStack(UndoStack.Count)
        UndoStack.Remove UndoStack.Count
    End If
    trapUndo = True
    rtftext.SetFocus
End Sub
Public Sub Redo()
Dim chg$
Dim DeleteFlag As Boolean 'flag as to whether or not to delete text or append text
Dim objElement As Object
If RedoStack.Count > 0 And trapUndo Then
    trapUndo = False
    DeleteFlag = RedoStack(RedoStack.Count).TextLen < Len(rtftext.Text)
    If DeleteFlag Then  'delete last item
        Set objElement = RedoStack(RedoStack.Count)
        rtftext.SelStart = objElement.SelStart
        rtftext.SelLength = Len(rtftext.Text) - objElement.TextLen
        rtftext.SelText = ""
    Else 'append something
        Set objElement = RedoStack(RedoStack.Count)
        chg$ = Change(rtftext.Text, objElement.Text, objElement.SelStart + 1)
        rtftext.SelStart = objElement.SelStart - Len(chg$)
        rtftext.SelLength = 0
        rtftext.SelText = chg$
        rtftext.SelStart = objElement.SelStart - Len(chg$)
        If Len(chg$) > 1 And chg$ <> vbCrLf Then
            rtftext.SelLength = Len(chg$)
        Else
            rtftext.SelStart = rtftext.SelStart + Len(chg$)
        End If
    End If
    UndoStack.Add objElement
    RedoStack.Remove RedoStack.Count
    trapUndo = True
End If

rtftext.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
color = ReadINI("Options", "Colored", inifile)
bUpdateFlag = False
bChanged = False
trapUndo = True
wb.Top = 0
wb.Left = 0
wb.Width = 11975
wb.Height = 7035
CodeMax1.Language = "HTML"
CodeMax1.ColorSyntax = color


End Sub



Private Sub Form_Resize()
CodeMax1.Width = Me.Width - 100
CodeMax1.Height = Me.Height - 800
wb.Width = Me.Width - 100
wb.Height = Me.Height - 800
WTabs1.Top = Me.Height - 800

End Sub

Private Sub WTabs1_Click(ByVal ActualClick As Boolean)
If WTabs1.ActiveTab = 0 Then
CodeMax1.Visible = True
wb.Visible = False
End If

If WTabs1.ActiveTab = 1 Then
CodeMax1.Visible = False
wb.Visible = True
On Error Resume Next
Open App.Path & "\projects\temp.html" For Output As #1
Print #1, CodeMax1.Text
Close #1
wb.Navigate App.Path & "\projects\temp.html"
End If

End Sub


