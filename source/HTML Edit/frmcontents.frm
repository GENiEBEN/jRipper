VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcontents 
   Caption         =   "HTML Edit Help Contents"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8820
   Icon            =   "frmcontents.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView tv1 
      Height          =   6255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   11033
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.TextBox txthelp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufileclose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmcontents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub display(ByVal h As Integer)
Dim help As String

If h = 0 Then
help = "Welcome to HTML Edit"
ElseIf h = 1 Then
help = "When opening a file click on the open file icon (the one that looks like an open folder) or go to the menu bar click on File and then Open."
ElseIf h = 2 Then
help = "When saving a file you can either click on the save file icon (the one that looks like a floppy disk) or you can go to the menu bar click on File and then on either Save or Save As."
ElseIf h = 3 Then
help = "To close a file you can click on the close button on the menu bar or you can go to the menu bar click on File and then click on Close"
ElseIf h = 4 Then
help = "When printing your file click on the printer icon or go to the menu bar and click on File and then on Print"
ElseIf h = 5 Then
help = "When inserting a table go to Insert on the menu bar and click on Table." & vbCrLf & vbCrLf
help = help + "Once the dialog box appears you will be asked to enter the table width, the number of rows, the number of columns, the cellpadding, the cellspacing, the border, weather or not you want a border color, and weather or not you want a background color."
ElseIf h = 6 Then
help = "The table width determines the width of the table. It can be entered as a number (ex. 150) or as a percentage (ex. 50%)."
ElseIf h = 7 Then
help = "The number of rows you have in a table determines the height of the table. A table row is represented by the <tr></tr> tags."
ElseIf h = 8 Then
help = "The number of columns in a table determines the number of cells in the table. For example you have 2 rows and 5 columns you would then have 10 cells in the table. A column and a cell is represented by the <td></td> tags."
ElseIf h = 9 Then
help = "The cellpadding of a table determines the amount of space between the border and the cell contents."
ElseIf h = 10 Then
help = "The cellspacing of a table determines the amount of space between each cell in the table."
ElseIf h = 11 Then
help = "By setting the border of a table determines how thin or how thick you want the border to be. The smaller the number the smaller the border and the larger the number the larger the border. When the border is set to 0 that means that there will be no border in the table."
ElseIf h = 12 Then
help = "The border color is the color of the border. However if you have a border of 0 then it is pointless setting the border color because the border is not going to appear."
ElseIf h = 13 Then
help = "The background color is the background color of the table and you can even set the background color of each cell in the table all the same color or all different colors if you want."
End If

txthelp.Text = help

End Sub






Sub distree()
Dim help As String

If tv1.SelectedItem.Key = "Welcome" Then
help = "Welcome to HTML Edit"
ElseIf tv1.SelectedItem.Key = "opening" Then
help = "When opening a file click on the open file icon (the one that looks like an open folder) or go to the menu bar click on File and then Open."
ElseIf tv1.SelectedItem.Key = "saving" Then
help = "When saving a file you can either click on the save file icon (the one that looks like a floppy disk) or you can go to the menu bar click on File and then on either Save or Save As."
ElseIf tv1.SelectedItem.Key = "closing" Then
help = "To close a file you can click on the close button on the menu bar or you can go to the menu bar click on File and then click on Close"
ElseIf tv1.SelectedItem.Key = "printing" Then
help = "When printing your file click on the printer icon or go to the menu bar and click on File and then on Print"
End If

txthelp.Text = help
End Sub

Sub loadtree()
'welcome
Set tempnode = tv1.Nodes.Add(, , "Welcome", "Welcome", , 0)
Set tempnode = tv1.Nodes.Add("Welcome", tvwChild, "opening", "Opening a file", , 0)
Set tempnode = tv1.Nodes.Add("Welcome", tvwChild, "saving", "Saving a file", , 0)
Set tempnode = tv1.Nodes.Add("Welcome", tvwChild, "closing", "Closing a file", , 0)
Set tempnode = tv1.Nodes.Add("Welcome", tvwChild, "printing", "Printing", , 0)

'tables
Set tempnode = tv1.Nodes.Add(, , "tables", "Tables", , 0)
Set tempnode = tv1.Nodes.Add("tables", tvwChild, "twidth", "Table Width", , 0)
Set tempnode = tv1.Nodes.Add("tables", tvwChild, "trows", "Rows", , 0)
Set tempnode = tv1.Nodes.Add("tables", tvwChild, "tcols", "Columns", , 0)
Set tempnode = tv1.Nodes.Add("tables", tvwChild, "cellpad", "Cellpadding", , 0)
Set tempnode = tv1.Nodes.Add("tables", tvwChild, "cellspac", "Cellspacing", , 0)
Set tempnode = tv1.Nodes.Add("tables", tvwChild, "tborder", "Border", , 0)
Set tempnode = tv1.Nodes.Add("tables", tvwChild, "tborcol", "Border Color", , 0)
Set tempnode = tv1.Nodes.Add("tables", tvwChild, "tbgcol", "Background Color", , 0)


End Sub

Private Sub Form_Load()
Dim hehelp As String
hehelp = "Welcome to HTML Edit"
txthelp.Text = hehelp

loadtree

End Sub






Private Sub mnufileclose_Click()
Unload Me

End Sub


Private Sub tv1_Click()
distree

End Sub


