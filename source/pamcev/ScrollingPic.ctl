VERSION 5.00
Begin VB.UserControl ScrollingPic 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   ScaleHeight     =   191
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   165
   ToolboxBitmap   =   "ScrollingPic.ctx":0000
   Begin VB.PictureBox picSource 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   2640
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.VScrollBar yScroll 
      Height          =   2535
      LargeChange     =   25
      Left            =   2160
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar xScroll 
      Height          =   255
      LargeChange     =   25
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Width           =   2415
   End
   Begin VB.PictureBox picShow 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H004D483F&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   0
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "ScrollingPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim picLoaded As Boolean
Dim xPos As Long, yPos As Long



Public Property Get Picture() As Variant
Picture = picSource.hdc
End Property

Public Property Let Picture(ByVal vNewValue As Variant)
picShow.Cls
picSource.Cls
xPos = 0: yPos = 0
xScroll.Value = 0
yScroll.Value = 0
picShow.Width = picSource.Width
picSource.Picture = LoadPicture(vNewValue)
UserControl_Resize
MapImage
End Property



Private Sub UserControl_Resize()
On Local Error Resume Next
picShow.Width = UserControl.ScaleWidth - 17
picShow.Height = UserControl.ScaleHeight - 17
yScroll.Left = UserControl.ScaleWidth - 17
xScroll.Top = UserControl.ScaleHeight - 17
yScroll.Height = UserControl.ScaleHeight - 17
xScroll.Width = UserControl.ScaleWidth '- 17
If picShow.Width < picSource.Width Then
  xScroll.Enabled = True
  xScroll.Max = picSource.Width - picShow.Width
  If Err Then
    xScroll.Value = picSource.Width - picShow.Width
    xScroll.Max = picSource.Width - picShow.Width
  End If
  xPos = xScroll.Value
Else
  xScroll.Enabled = False
  xPos = 0
End If
If picShow.Height < picSource.Height Then
  yScroll.Enabled = True
  yScroll.Max = picSource.Height - picShow.Height
  If Err Then
    yScroll.Value = picSource.Height - picShow.Height
    yScroll.Max = picSource.Height - picShow.Height
  End If
  yPos = yScroll.Value
Else
  yScroll.Enabled = False
  yPos = 0
End If
MapImage
End Sub

Private Sub MapImage()
On Local Error Resume Next
BitBlt picShow.hdc, 0, 0, picShow.Width, picShow.Height, picSource.hdc, xPos, yPos, SRCCOPY
picShow.Refresh
End Sub

Private Sub xScroll_Change()
xPos = xScroll.Value
MapImage
End Sub

Private Sub yScroll_Change()
yPos = yScroll.Value
MapImage
End Sub
