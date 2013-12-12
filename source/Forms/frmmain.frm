VERSION 5.00
Object = "{F924C9A7-D9B7-4808-8A32-108A70944450}#1.0#0"; "HookMenu.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmmain 
   Caption         =   "Text Editor 1.00"
   ClientHeight    =   6180
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   8730
   WindowState     =   2  'Maximized
   Begin VB.CommandButton bbar 
      BackColor       =   &H00291E16&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   -15
      TabIndex        =   2
      Top             =   5865
      Width           =   8800
   End
   Begin RichTextLib.RichTextBox txtmain 
      Height          =   5835
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   10292
      _Version        =   393217
      BackColor       =   2694678
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmmain.frx":0DE6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   8115
      Top             =   3975
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox filelist 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4485
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog diag 
      Left            =   8115
      Top             =   2595
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lines 
      Caption         =   "Label1"
      Height          =   735
      Left            =   2265
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufnew 
         Caption         =   "New"
      End
      Begin VB.Menu mnufopen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnufsave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnufsaveas 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnufsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnufclose 
         Caption         =   "Close File"
      End
      Begin VB.Menu mnufsep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnufexit 
         Caption         =   "Exit"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edit"
      Begin VB.Menu mnuecut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuecopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuepaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuesep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuefind 
         Caption         =   "Find"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuffr 
         Caption         =   "Find && Replace"
      End
      Begin VB.Menu mnuesort 
         Caption         =   "Sort Lines"
         Begin VB.Menu mnusa 
            Caption         =   "Ascending"
         End
         Begin VB.Menu mnusd 
            Caption         =   "Descending"
         End
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim saved As Boolean, edited As Boolean, loadedfile As Boolean
Dim fso As FileSystemObject
Dim txtstr As TextStream
Dim FileName As String, answer As String
Const AppName As String = "Text Editor"

Private Sub Form_Load()
'    Dim pname As String
'    If Command <> "" Then
'        pname = Command
'        pname = Replace(pname, Chr(34), "")
'        If fso.FileExists(pname) = True Then
'            Set fso = New FileSystemObject
'            Set txtstr = fso.OpenTextFile(pname, ForReading, False)
'            With txtstr
'                txtmain.Text = .ReadAll
'                .Close
'            End With
            Set txtstr = Nothing
            Set fso = Nothing
            FileName = jrMain.Path.Caption
            edited = False
            saved = True
            Me.Caption = "Text Editor [" & Get_FNAME(FileName) & "]"
'        Else
'            Me.Caption = "Text Editor [New File]"
'        End If
'    Else
'        Me.Caption = "Text Editor [New File]"
'    End If
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        txtmain.Left = -30
        txtmain.Top = -30
        txtmain.Width = Me.Width - 70
        txtmain.Height = Me.Height - bbar.Height - 360
        bbar.Top = Me.Height - bbar.Height - 400
        bbar.Width = Me.Width - 90
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If edited = True Then
        answer = MsgBox("This file has been altered, do you wish to save?", vbQuestion + vbYesNoCancel, AppName)
        If answer = vbYes Then
            SaveFile FileName, loadedfile
            Unload Me
        ElseIf answer = vbNo Then
            Unload Me
        ElseIf answer = vbCancel Then
            Cancel = -1
        End If
    End If
    
End Sub

Private Sub mnuecopy_Click()
    Clipboard.SetText txtmain.SelText
End Sub

Private Sub mnuecut_Click()
    Clipboard.SetText txtmain.SelText
    txtmain.SelText = ""
End Sub

Private Sub mnuedit_Click()
    If Clipboard.GetText <> "" Then
        mnuepaste.Enabled = True
    Else
        mnuepaste.Caption = False
    End If
    If txtmain.SelLength <> 0 Then
        mnuecut.Enabled = True
        mnuecopy.Enabled = True
    Else
        mnuecut.Enabled = False
        mnuecopy.Enabled = False
    End If
    If txtmain.Text <> "" Then
        mnuesort.Enabled = True
    Else
        mnuesort.Enabled = False
    End If
End Sub

Private Sub mnuepaste_Click()
    txtmain.SelText = Clipboard.GetText
End Sub

Private Sub mnufclose_Click()
    New_file
End Sub

Private Sub mnufexit_Click()
    Unload Me
End Sub

Private Sub mnuffr_Click()
    frmfind.Show 1
    'Me.Enabled = False
End Sub

Private Sub mnufnew_Click()
    New_file
End Sub

Private Sub mnufopen_Click()
    OpenFile
        frmmain.bbar.Caption = "Lines: " & TXT.lineCount(FileName) & " File: " & FileName
frmmain.txtmain.SelStart = 0
frmmain.txtmain.SelLength = Len(frmmain.txtmain.Text)
frmmain.txtmain.SelColor = vbWhite
frmmain.txtmain.SelFontName = "Terminal"
frmmain.txtmain.SelFontSize = 9
frmmain.txtmain.SelStart = 0
End Sub

Private Sub mnufsave_Click()
    SaveFile FileName, loadedfile
End Sub

Private Sub mnufsaveas_Click()
    SaveFile FileName, False
End Sub

Private Sub mnusa_Click()
    sort True
    frmmain.txtmain.SelStart = 0
    frmmain.txtmain.SelLength = Len(frmmain.txtmain.Text)
    frmmain.txtmain.SelColor = vbWhite
    frmmain.txtmain.SelStart = 0
End Sub

Private Sub mnusd_Click()
    sort False
    frmmain.txtmain.SelStart = 0
    frmmain.txtmain.SelLength = Len(frmmain.txtmain.Text)
    frmmain.txtmain.SelColor = vbWhite
    frmmain.txtmain.SelStart = 0
End Sub

Private Sub txtmain_Change()
    If edited <> True Then
        edited = True
        Me.Caption = Me.Caption & "*"
    End If
End Sub

Private Function Get_FNAME(ByVal filesname As String) As String
On Error GoTo errorhand
    Dim loggednumber, strlen As Integer, i As Integer
    Dim letter As String
    filesname = Replace(filesname, Chr(34), "")
    strlen = Len(FileName)
    For i = 1 To strlen
        letter = Mid(FileName, Len(filesname) - i, 1)
        If Mid(filesname, Len(filesname) - i, 1) = "\" Then
            loggednumber = i
            i = Len(filesname)
        End If
    Next i
    Get_FNAME = CStr(Mid(filesname, Len(filesname) - (loggednumber - 1), loggednumber))
    Exit Function
errorhand:
    Exit Function
End Function

Private Sub New_file()
    If edited = True Then
        answer = MsgBox("This file has been altered, do you wish to save?", vbQuestion + vbYesNoCancel, AppName)
        If answer = vbYes Then
            If SaveFile(FileName, loadedfile) = True Then
                txtmain.Text = ""
                loadedfile = False
                edited = False
                saved = False
                Me.Caption = "Text Editor [New File]"
            End If
        ElseIf answer = vbNo Then
            txtmain.Text = ""
            loadedfile = False
            edited = False
            saved = False
            Me.Caption = "Text Editor [New File]"
        End If
    Else
        txtmain.Text = ""
        loadedfile = False
        edited = False
        saved = False
        Me.Caption = "Text Editor [New File]"
    End If
End Sub

Private Function SaveFile(ByVal fName As String, ByVal exist As Boolean) As Boolean
On Error GoTo errx
    If exist = True Then
        If fName <> "" Then
            Set fso = New FileSystemObject
            Set txtstr = fso.CreateTextFile(fName, True)
            With txtstr
                .Write txtmain.Text
                .Close
            End With
            Set txtstr = Nothing
            Set fso = Nothing
            SaveFile = True
            Me.Caption = "Text Editor [" & Get_FNAME(FileName) & "]"
            edited = False
        End If
    Else
        diag.DialogTitle = "Save File"
        diag.FileName = fName
        diag.Filter = "Text File (*.txt)|*.txt|All Formats (*.*)|*.*|"
        diag.ShowSave
        FileName = diag.FileName
        If FileName <> "" Then
            If fso.FileExists(FileName) = True Then
                answer = MsgBox("This file already exists, do you wish to overwrite?", vbQuestion + vbYesNo, AppName)
                If answer = vbYes Then
                    Set fso = New FileSystemObject
                    Set txtstr = fso.CreateTextFile(FileName, True)
                    With txtstr
                        .Write txtmain.Text
                        .Close
                    End With
                    Set txtstr = Nothing
                    Set fso = Nothing
                    SaveFile = True
                    Me.Caption = "Text Editor [" & Get_FNAME(FileName) & "]"
                    edited = False
                Else
                    SaveFile = False
                End If
            Else
                Set fso = New FileSystemObject
                Set txtstr = fso.CreateTextFile(FileName, True)
                With txtstr
                    .Write txtmain.Text
                    .Close
                End With
                Set txtstr = Nothing
                Set fso = Nothing
                SaveFile = True
                Me.Caption = "Text Editor [" & Get_FNAME(FileName) & "]"
                edited = False
            End If
        End If
    End If
'ERROR MANAGER
errx:
If Err.Number = 91 Then ' cancel pressed
Exit Function
Else
MsgBox "err #" & Err.Number & " : " & Err.Description
Exit Function
End If
    
End Function

Public Sub OpenFile()
Dim myFilters As String
Dim f
Dim NOFilters As String: NOFilters = ReadINI(App.Path & "\bin\jr.ini", "TextEditor", "NumberOfFilters")
Dim filterformat As String: filterformat = ReadINI(App.Path & "\bin\jr.ini", "TextEditor", "ShowExtension")
filterformat = LCase(filterformat)
Dim fStore As String
Dim fName As String
Dim fExt As String

' Load Filters
For f = 1 To NOFilters
fStore = ReadINI(App.Path & "\bin\jr.ini", "TextEditor", f)
fName = Split(fStore, "$")(0)
fExt = Split(fStore, "$")(1)
If filterformat = "true" Then
myFilters = myFilters & fName & "(" & fExt & ")" & vbNullChar & fExt & vbNullChar
Else
myFilters = myFilters & fName & vbNullChar & fExt & vbNullChar
End If
Next f
myFilters = myFilters & vbNullChar & vbNullChar

    If edited = True Then
        answer = MsgBox("This file has been altered, do you wish to save?", vbQuestion + vbYesNoCancel, AppName)
        If answer = vbYes Then
            If SaveFile(FileName, loadedfile) = True Then
                diag.DialogTitle = "Open File"
                diag.FileName = ""
                diag.Filter = myFilters '"All Files (*.*)|*.*|"
                diag.ShowOpen
                FileName = diag.FileName
                If FileName <> "" Then
                    txtmain.Text = ""
                    Set fso = New FileSystemObject
                    Set txtstr = fso.OpenTextFile(FileName, ForReading, False)
                    With txtstr
                        txtmain.Text = .ReadAll
                        .Close
                    End With
                    Set txtstr = Nothing
                    Set fso = Nothing
                    edited = False
                    loadedfile = True
                    saved = True
                    Me.Caption = "Text Editor [" & Get_FNAME(FileName) & "]"
                End If
            End If
        Else
            diag.DialogTitle = "Open File"
            diag.FileName = ""
            diag.Filter = myFilters '"All Files (*.*)|*.*|"
            diag.ShowOpen
            FileName = diag.FileName
            If FileName <> "" Then
                txtmain.Text = ""
                Set fso = New FileSystemObject
                Set txtstr = fso.OpenTextFile(FileName, ForReading, False)
                With txtstr
                    txtmain.Text = .ReadAll
                    .Close
                End With
                Set txtstr = Nothing
                Set fso = Nothing
                edited = False
                loadedfile = True
                saved = True
                Me.Caption = "Text Editor [" & Get_FNAME(FileName) & "]"
            End If
        End If
    Else
        diag.DialogTitle = "Open File"
        diag.FileName = ""
        diag.Filter = myFilters '"All Files (*.*)|*.*|"
        diag.ShowOpen
        FileName = diag.FileName
        If FileName <> "" Then
            txtmain.Text = ""
            Set fso = New FileSystemObject
            Set txtstr = fso.OpenTextFile(FileName, ForReading, False)
            With txtstr
                txtmain.Text = .ReadAll
                .Close
            End With
            Set txtstr = Nothing
            Set fso = Nothing
            edited = False
            loadedfile = True
            saved = True
            Me.Caption = "Text Editor [" & Get_FNAME(FileName) & "]"
        End If
    End If
End Sub

 Private Function Count_Lines(TextBox As TextBox) As Long
    Dim lnc As Long, lns As String
    lnc = 1
    lns = TextBox.Text
    Do While InStr(lns, Chr(13))
        lnc = lnc + 1
        lns = Mid(lns, InStr(lns, Chr(13)) + 1)
    Loop
    Count_Lines = lnc
End Function

Public Sub sort(ByVal ascending As Boolean)
On Error Resume Next
    Dim i As Integer, lineCount As Long, perc As Integer
    Dim linedata As String
    If txtmain.Text <> "" Then
        lineCount = 0
        filelist.Clear
        txtmain.SelStart = 0
        txtmain.SelLength = Len(txtmain.Text)
        Open App.Path & "\tmp.txt" For Output As #1
            Print #1, txtmain.SelText
        Close #1
        'count amount of lines to be sorted
        lineCount = Me.lines.Caption * 2
        frmworking.Show
        frmworking.progbar1.Max = lineCount
        frmworking.Caption = "Working, Please Wait..." & "   0% Complete"
        Me.Enabled = False
        Open App.Path & "\tmp.txt" For Input As #1
            Do Until EOF(1)
                Line Input #1, linedata
                If linedata <> "" Then
                    filelist.AddItem linedata
                End If
                frmworking.progbar1.Value = frmworking.progbar1.Value + 1
                perc = (frmworking.progbar1.Value / frmworking.progbar1.Max) * 100
                frmworking.Caption = "Working, Please Wait..." & "  " & perc & "% Complete"
                DoEvents
            Loop
        Close #1
        Kill App.Path & "\tmp.txt"
        txtmain.Text = ""
        If ascending = True Then
            For i = 0 To filelist.ListCount - 1
                If i = filelist.ListCount - 1 Then
                    txtmain.Text = txtmain.Text & filelist.list(i)
                Else
                    txtmain.Text = txtmain.Text & filelist.list(i) & vbCrLf
                End If
                frmworking.progbar1.Value = frmworking.progbar1.Value + 1
                perc = (frmworking.progbar1.Value / frmworking.progbar1.Max) * 100
                frmworking.Caption = "Working, Please Wait..." & "  " & perc & "% Complete"
            Next i
        Else
            For i = 0 To filelist.ListCount - 1
                If (filelist.ListCount - 1) - i = 0 Then
                    txtmain.Text = txtmain.Text & filelist.list((filelist.ListCount - 1) - i)
                Else
                    txtmain.Text = txtmain.Text & filelist.list((filelist.ListCount - 1) - i) & vbCrLf
                End If
                frmworking.progbar1.Value = frmworking.progbar1.Value + 1
                perc = (frmworking.progbar1.Value / frmworking.progbar1.Max) * 100
                frmworking.Caption = "Working, Please Wait..." & "  " & perc & "% Complete"
            Next i
        End If
        filelist.Clear
        Me.Enabled = True
        Unload frmworking
    End If
End Sub
