VERSION 5.00
Begin VB.Form NFSMW_MT 
   BackColor       =   &H004D483F&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NFS Most Wanted --MENUS TWEAK 3.01--"
   ClientHeight    =   2910
   ClientLeft      =   345
   ClientTop       =   330
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "nfsmw_mt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo6 
      BackColor       =   &H00131313&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1290
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   1785
      Width           =   3945
   End
   Begin VB.ComboBox Combo5 
      BackColor       =   &H00131313&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1290
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   1455
      Width           =   3945
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00131313&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1290
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1125
      Width           =   3945
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00131313&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1290
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   795
      Width           =   3945
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00131313&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1290
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   465
      Width           =   3945
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00131313&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1290
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   135
      Width           =   3945
   End
   Begin jR_RC2.Butt Butt1 
      Height          =   420
      Left            =   4170
      TabIndex        =   13
      Top             =   2355
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   741
      Caption         =   "APPLY"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin jR_RC2.Butt Butt2 
      Height          =   420
      Left            =   3015
      TabIndex        =   14
      Top             =   2355
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   741
      Caption         =   "CANCEL"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin jR_RC2.Butt Butt3 
      Height          =   420
      Left            =   1890
      TabIndex        =   15
      Top             =   2355
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   741
      Caption         =   "ABOUT"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Image Image1 
      Height          =   45
      Left            =   -15
      Picture         =   "nfsmw_mt.frx":199A
      Stretch         =   -1  'True
      Top             =   2220
      Width           =   5520
   End
   Begin VB.Label ret 
      Caption         =   "return"
      Height          =   225
      Left            =   225
      TabIndex        =   12
      Top             =   3540
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "NextGenSky"
      ForeColor       =   &H00BDB8AF&
      Height          =   225
      Left            =   120
      TabIndex        =   10
      Top             =   1815
      Width           =   1035
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop Backroom"
      ForeColor       =   &H00BDB8AF&
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   1485
      Width           =   1170
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
      ForeColor       =   &H00BDB8AF&
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   1155
      Width           =   825
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SafeHouse"
      ForeColor       =   &H00BDB8AF&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   825
      Width           =   825
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Car Lot"
      ForeColor       =   &H00BDB8AF&
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   495
      Width           =   825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
      ForeColor       =   &H00BDB8AF&
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   825
   End
End
Attribute VB_Name = "NFSMW_MT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim fso As New FileSystemObject

Private Sub Butt1_Click()
Dim file1 As String: file1 = "PlatformCrib.BIN"
Dim file2 As String: file2 = "CAR_LOT.BIN"
Dim file3 As String: file3 = "CAREER_SAFEHOUSE.BIN"
Dim file4 As String: file4 = "CUSTOMIZATION_SHOP.BIN"
Dim file5 As String: file5 = "CUSTOMIZATION_SHOP_BACKROOM.BIN"
Dim file6 As String: file6 = "NextGenSky.BIN"
' bACKUP
BF file1
BF file2
BF file3
BF file4
BF file5
BF file6
' rEPLACE
SC Combo1, file1
SC Combo2, file2
SC Combo3, file3
SC Combo4, file4
SC Combo5, file5
SC Combo6, file6
' eXIT
MsgBox "Done!"
Unload Me
End Sub

Function BF(File)
If fso.FileExists(Me.ret.Caption & File & ".BAK") = False Then
fso.CopyFile Me.ret.Caption & File, Me.ret.Caption & File & ".BAK"
End If
End Function
Function SC(ComboBox As ComboBox, Dest)
Dim file1 As String: file1 = "PlatformCrib.BIN"
Dim file2 As String: file2 = "CAR_LOT.BIN"
Dim file3 As String: file3 = "CAREER_SAFEHOUSE.BIN"
Dim file4 As String: file4 = "CUSTOMIZATION_SHOP.BIN"
Dim file5 As String: file5 = "CUSTOMIZATION_SHOP_BACKROOM.BIN"
Dim file6 As String: file6 = "NextGenSky.BIN"

Select Case ComboBox.Text
    Case "Crib"
    RF file1, Dest
    Case "Car Lot"
    RF file2, Dest
    Case "Safehouse"
    RF file3, Dest
    Case "Shop"
    RF file4, Dest
    Case "Shop Backroom"
    RF file5, Dest
    Case "Empty Black Screen"
    RF file6, Dest
End Select

End Function
Function RF(Source, Dest)
' NOW REPLACE DEST WITH SOURCE
fso.DeleteFile Me.ret.Caption & Dest
fso.CopyFile Me.ret.Caption & Source & ".BAK", Me.ret.Caption & Dest
End Function

Private Sub Form_Load()
' Load registry path
Dim regKey As String: regKey = "HKEY_LOCAL_MACHINE\SOFTWARE\EA GAMES\Need for Speed Most Wanted"
Dim regEnt As String: regEnt = "Install Dir"
Me.ret.Caption = jrRegistry.GetStringValue(regKey, regEnt)
Me.ret.Caption = ret.Caption & "FRONTEND\PLATFORMS\"
' Fill comboboxes
With Combo1
.Text = "Crib"
.AddItem "Crib"
.AddItem "Car Lot"
.AddItem "Safehouse"
.AddItem "Shop"
.AddItem "Shop Backroom"
.AddItem "Empty Black Screen"

End With
With Combo2
.Text = "Car Lot"
.AddItem "Crib"
.AddItem "Car Lot"
.AddItem "Safehouse"
.AddItem "Shop"
.AddItem "Shop Backroom"
.AddItem "Empty Black Screen"
End With
With Combo3
.Text = "Safehouse"
.AddItem "Crib"
.AddItem "Car Lot"
.AddItem "Safehouse"
.AddItem "Shop"
.AddItem "Shop Backroom"
.AddItem "Empty Black Screen"
End With
With Combo4
.Text = "Shop"
.AddItem "Crib"
.AddItem "Car Lot"
.AddItem "Safehouse"
.AddItem "Shop"
.AddItem "Shop Backroom"
.AddItem "Empty Black Screen"
End With
With Combo5
.Text = "Shop Backroom"
.AddItem "Crib"
.AddItem "Car Lot"
.AddItem "Safehouse"
.AddItem "Shop"
.AddItem "Shop Backroom"
.AddItem "Empty Black Screen"
End With
With Combo6
.Text = "Empty Black Screen"
.AddItem "Crib"
.AddItem "Car Lot"
.AddItem "Safehouse"
.AddItem "Shop"
.AddItem "Shop Backroom"
.AddItem "Empty Black Screen"
End With
' WndAnimation
FadeIn Me, , 6
End Sub

'=================================================================================
Private Sub Butt2_Click()
FadeOut Me, , 6
Unload Me
End Sub

Private Sub Butt3_Click()
MsgBox "Tool Code    : codin/GENiEBEN" & vbCrLf & _
              "Tool GFX      : codin/GENiEBEN" & vbCrLf & vbCrLf & _
              "Tool site       : http://genieben.t35.com", vbInformation, _
              "NFSMW MT3.01"
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
'F1
If GetAsyncKeyState(vbKeyF1) Then
    If fso.FileExists(App.path & "\GENiEBEN.nfo") = True Then
    Shell "C:\Windows\Notepad.exe " & App.path & "\GENiEBEN.nfo", vbNormalFocus
    End If
End If
'ESC
If GetAsyncKeyState(vbKeyEscape) Then
    FadeOut Me, , 6
    End
End If
'F2
If GetAsyncKeyState(vbKeyF2) Then
    
End If
'F3
If GetAsyncKeyState(vbKeyF3) Then
    
End If
End Sub

Private Function SaveResItemToDisk(ByVal iResourceNum As Integer, ByVal sResourceType As String, ByVal sDestFileName As String) As Long
    Dim bytResourceData()   As Byte
    Dim iFileNumOut         As Integer
    On Error GoTo SaveResItemToDisk_err
    bytResourceData = LoadResData(iResourceNum, sResourceType)
    iFileNumOut = FreeFile
    Open sDestFileName For Binary Access Write As #iFileNumOut
        Put #iFileNumOut, , bytResourceData
    Close #iFileNumOut
    SaveResItemToDisk = 0
    Exit Function
SaveResItemToDisk_err:
    SaveResItemToDisk = Err.Number
End Function

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
