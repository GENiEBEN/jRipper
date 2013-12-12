VERSION 5.00
Object = "{F924C9A7-D9B7-4808-8A32-108A70944450}#1.0#0"; "HookMenu.ocx"
Begin VB.Form fMAIN 
   BackColor       =   &H00291E16&
   Caption         =   "PAMCEV"
   ClientHeight    =   8220
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin jR_RC2.ScrollingPic pic 
      Align           =   1  'Align Top
      Height          =   7545
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   13309
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   10815
      Top             =   7140
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      AutoSize        =   -1  'True
      BackColor       =   &H00291E16&
      Height          =   450
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   11820
      TabIndex        =   0
      Top             =   7770
      Width           =   11880
      Begin VB.ListBox files 
         Height          =   255
         Left            =   990
         TabIndex        =   4
         Top             =   60
         Width           =   2475
      End
      Begin HookMenu.ctxHookMenu ctxHookMenu1 
         Left            =   1380
         Top             =   675
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
      Begin VB.TextBox filecount 
         Height          =   255
         Left            =   45
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "FileNumber/TotalFiles"
         Top             =   60
         Width           =   915
      End
      Begin VB.Image topbar2 
         Height          =   675
         Left            =   -15
         Picture         =   "fMAIN.frx":0000
         Stretch         =   -1  'True
         Top             =   -210
         Width           =   13440
      End
   End
   Begin VB.Label fileno 
      Caption         =   "0"
      Height          =   345
      Left            =   9780
      TabIndex        =   2
      Top             =   6510
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_File_Open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_File_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Props 
         Caption         =   "&Properties"
      End
      Begin VB.Menu mnu_File_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Exit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnu_image 
      Caption         =   "&Image"
      Begin VB.Menu mnu_Image_properties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnu_image_sep1 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "fMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim ShowExtension As Boolean
Private Resp As Long


Private Sub Form_Load()
Dim fn
fn = FolderPath
loadfiles fn, "bmp"
loadfiles fn, "jpg"
loadfiles fn, "jpeg"
loadfiles fn, "ico"
loadfiles fn, "cur"

displayfile 1
End Sub

Private Sub Form_Resize()
On Error Resume Next
topbar2.Width = Me.Width
pic.Height = Me.Height - Picture1.Height - 500

End Sub

Function loadfiles(Folder, Extension)
On Error Resume Next
Dim mp3FileName As String
Dim tmpStrg As String
ShowExtension = True

tmpStrg = FileSystem.dir(Folder & "\*." & Extension)
tmpStrg = Replace(tmpStrg, "\\", "\")
If tmpStrg <> "" Then
    If ShowExtension = False Then
        mp3FileName = functions.get_FileNameWithoutExtension(tmpStrg)
    Else
        mp3FileName = tmpStrg
    End If
    files.AddItem mp3FileName
    tmpStrg = FileSystem.dir
    While Len(tmpStrg) > 0
        If ShowExtension = False Then
            mp3FileName = Left$(tmpStrg, Len(tmpStrg) - 4)
        Else
            mp3FileName = tmpStrg
        End If
        files.AddItem mp3FileName
        tmpStrg = FileSystem.dir
    Wend
Else
    'MsgBox "No " & Extension & " files found in this root folder", vbExclamation
End If
End Function

Private Sub mnu_File_Exit_Click()
Unload Me
End Sub


Private Sub Timer1_Timer()
Dim x
' Space
If GetAsyncKeyState(vbKeySpace) Then
    displayfile 1
End If
' Backspace
If GetAsyncKeyState(vbKeyBack) Then
    displayfile2
End If
' Enter
If GetAsyncKeyState(vbKeyReturn) Then
'    If Me.Picture1.Visible = True Then
'    Me.WindowState = vbMaximized
'    Picture1.Visible = False
'    pic.Height = Me.Height - 680
'    Else
'    Me.WindowState = vbNormal
'    Picture1.Visible = True
'    End If

End If
End Sub

Function displayfile(IncreaseNumber)
    ' Dead ends reached?
    If fileno.Caption = files.ListCount Then
    fileno.Caption = 0
    End If
    ' Check file type and load it
    Select Case get_ExtensionFromFileName(files.list(fileno.Caption))
    Case "bmp"
    Me.pic.Picture = FolderPath & files.list(fileno.Caption)
    fileno.Caption = Val(fileno.Caption) + IncreaseNumber
    filecount = fileno.Caption & "/" & files.ListCount
    
    Case "jpg"
    Me.pic.Picture = FolderPath & files.list(fileno.Caption)
    fileno.Caption = Val(fileno.Caption) + IncreaseNumber
    filecount = fileno.Caption & "/" & files.ListCount
    
    End Select
End Function

Function displayfile2()
    ' Dead ends reached?
    If fileno.Caption = 1 Then
    fileno.Caption = files.ListCount + 1
    End If
    fileno.Caption = Val(fileno.Caption) - 1

    ' Check file type and load it
    Dim i
    i = 1
    Select Case get_ExtensionFromFileName(files.list(Val(fileno.Caption) - i))
    Case "bmp"
    Me.pic.Picture = FolderPath & files.list(Val(fileno.Caption) - i)
    filecount = fileno.Caption & "/" & files.ListCount
    
    Case "jpg"
    Me.pic.Picture = FolderPath & files.list(Val(fileno.Caption) - i)
    filecount = fileno.Caption & "/" & files.ListCount
    
    End Select
End Function


