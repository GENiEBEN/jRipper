Attribute VB_Name = "ListVideos"
'TODO :
' Check_FEAR nu merge ... nu returneaza valoarea care vreau sa o citesc in mod hex

Option Explicit
Dim fso As New FileSystemObject
Global outLong(1 To 2) As Byte

'==========================================================================================
Public Function GetIcon(fileEXT)
fileEXT = StrConv(fileEXT, vbLowerCase)
Select Case fileEXT
    Case "bik": GetIcon = "bik"
    Case "xmv": GetIcon = "bik"
    Case "rmv": GetIcon = "bik"
    Case "arc": GetIcon = "bik"
    Case "bmp": GetIcon = "bmp"
    Case "ogg": GetIcon = "ogg"
    Case "wmv": GetIcon = "wmv"
    Case "exe": GetIcon = "exe"
    Case "ini": GetIcon = "ini_txt"
    Case "txt": GetIcon = "ini_txt"
    Case Else: GetIcon = "unknown"
End Select
End Function

Public Function add(VideoName, FileName)
On Error Resume Next
Dim strEXT As String * 3
Dim l As ListItem
    Set l = NIMP.list.ListItems.add(, , VideoName)
        l.SubItems(1) = FileName
strEXT = Strings.Split(l.SubItems(1), ".")(1)
        l.SmallIcon = GetIcon(strEXT)
End Function

Public Function dir(regPATH As String, regKEY As String, VideoFolder As String) ' Get video folder path
Dim tmp As String
With NIMP.ret
          tmp = GetStringValue(regPATH, regKEY)
         .Caption = tmp
         .Caption = .Caption & VideoFolder
dir = .Caption
End With
End Function
Public Function dir2(regPATH As String, regKEY As String, VideoFolder As String) ' Get video folder path (2)
Dim tmp As String
With NIMP.ret2
          tmp = GetStringValue(regPATH, regKEY)
         .Caption = tmp
         .Caption = .Caption & VideoFolder
dir2 = .Caption
End With
End Function
Public Function del(ListViewINDEX) ' delete a file
On Error Resume Next
Dim vidName As String
Dim l As ListItem
If ListViewINDEX > NIMP.list.ListItems.Count Then
MsgBox "Index not existent"
Exit Function
End If
vidName = NIMP.list.ListItems(ListViewINDEX).SubItems(1) ' Get video file name from ListView
vidName = NIMP.ret.Caption & vidName ' Return full path of video file
SetAttr vidName, vbNormal ' Remove Read-Only attribute
fso.CopyFile vidName, vidName & ".NIMP" ' Backup file
fso.DeleteFile vidName, True ' Delete original file
End Function
Public Function del2(ListViewINDEX) ' delete a file(2)
On Error Resume Next
Dim vidName As String
Dim l As ListItem
If ListViewINDEX > NIMP.list.ListItems.Count Then
MsgBox "Index not existent"
Exit Function
End If
vidName = NIMP.list.ListItems(ListViewINDEX).SubItems(1) ' Get video file name from ListView
vidName = NIMP.ret2.Caption & vidName ' Return full path of video file
SetAttr vidName, vbNormal ' Remove Read-Only attribute
fso.CopyFile vidName, vidName & ".NIMP" ' Backup file
fso.DeleteFile vidName, True ' Delete original file
End Function
Public Function del_CondemnedCO(ListViewINDEX, Offset) ' noop video entry in a archive
On Error Resume Next
Dim vidName As String
Dim l As ListItem
If ListViewINDEX > NIMP.list.ListItems.Count Then
MsgBox "Index not existent"
Exit Function
End If
vidName = NIMP.list.ListItems(ListViewINDEX).SubItems(1) ' Get video file name from ListView
vidName = NIMP.ret.Caption & vidName ' Return full path of video file
SetAttr vidName, vbNormal ' Remove Read-Only attribute
' Open file in bynary mode and fill Offset with 0 (this way u can noop file entry instead wasting time deleting it from archive)
Open vidName For Binary As #1
'NumberToByte 0
Put #1, Offset, "X" ' outLong
Close #1
End Function
Public Function del_FEAR(ListViewINDEX, Offset) ' noop video entry in a archive
On Error Resume Next
Dim vidName As String
Dim l As ListItem
If ListViewINDEX > NIMP.list.ListItems.Count Then
MsgBox "Index not existent"
Exit Function
End If
vidName = NIMP.list.ListItems(ListViewINDEX).SubItems(1) ' Get video file name from ListView
vidName = NIMP.ret.Caption & vidName ' Return full path of video file
SetAttr vidName, vbNormal ' Remove Read-Only attribute
' Open file in bynary mode and fill Offset with 0 (this way u can noop file entry instead wasting time deleting it from archive)
Open vidName For Binary As #1
'NumberToByte 0
Put #1, Offset, "X" ' outLong
Close #1
End Function
Public Function restore(ListViewINDEX) ' restore a backup
On Error Resume Next
Dim vidName As String
Dim l As ListItem
If ListViewINDEX > NIMP.list.ListItems.Count Then
MsgBox "Index not existent"
Exit Function
End If
vidName = NIMP.list.ListItems(ListViewINDEX).SubItems(1) ' Get video file name from ListView
vidName = NIMP.ret.Caption & vidName ' Return full path of video file
SetAttr vidName, vbNormal ' Remove Read-Only attribute
fso.CopyFile vidName & ".NIMP", vidName ' Restore File
End Function
Public Function restore2(ListViewINDEX) ' restore a backup(2)
On Error Resume Next
Dim vidName As String
Dim l As ListItem
If ListViewINDEX > NIMP.list.ListItems.Count Then
MsgBox "Index not existent"
Exit Function
End If
vidName = NIMP.list.ListItems(ListViewINDEX).SubItems(1) ' Get video file name from ListView
vidName = NIMP.ret2.Caption & vidName ' Return full path of video file
SetAttr vidName, vbNormal ' Remove Read-Only attribute
fso.CopyFile vidName & ".NIMP", vidName ' Restore File
End Function
Public Function restore_CondemnedCO(ListViewINDEX, Offset, OriginalLetter As String) ' noop video entry in a archive
On Error Resume Next
Dim vidName As String
Dim l As ListItem
If ListViewINDEX > NIMP.list.ListItems.Count Then
MsgBox "Index not existent"
Exit Function
End If
vidName = NIMP.list.ListItems(ListViewINDEX).SubItems(1) ' Get video file name from ListView
vidName = NIMP.ret.Caption & vidName ' Return full path of video file
SetAttr vidName, vbNormal ' Remove Read-Only attribute
' Open file in bynary mode and fill Offset with 0 (this way u can noop file entry instead wasting time deleting it from archive)
Open vidName For Binary As #1
'NumberToByte 0
Put #1, Offset, OriginalLetter ' outLong
Close #1
End Function
Public Function restore_FEAR(ListViewINDEX, Offset, OriginalLetter As String) ' noop video entry in a archive
On Error Resume Next
Dim vidName As String
Dim l As ListItem
If ListViewINDEX > NIMP.list.ListItems.Count Then
MsgBox "Index not existent"
Exit Function
End If
vidName = NIMP.list.ListItems(ListViewINDEX).SubItems(1) ' Get video file name from ListView
vidName = NIMP.ret.Caption & vidName ' Return full path of video file
SetAttr vidName, vbNormal ' Remove Read-Only attribute
' Open file in bynary mode and fill Offset with 0 (this way u can noop file entry instead wasting time deleting it from archive)
Open vidName For Binary As #1
'NumberToByte 0
Put #1, Offset, OriginalLetter ' outLong
Close #1
End Function
Public Function check(ListViewINDEX) ' file or backup exists?
On Error Resume Next
Dim vidName As String
Dim l As ListItem
If ListViewINDEX > NIMP.list.ListItems.Count Then
MsgBox "Index not existent"
Exit Function
End If
vidName = NIMP.list.ListItems(ListViewINDEX).SubItems(1) ' Get video file name from ListView
vidName = NIMP.ret.Caption & vidName ' Return full path of video file
Dim errorAhh As String: errorAhh = Left(NIMP.ret.Caption, 5)
If errorAhh = "Error" Then
NIMP.selall.Enabled = False
Else
NIMP.selall.Enabled = True
If fso.FileExists(vidName) = True Then ' if original file exists then ...
    If FileLen(vidName) <> 3984 Then 'is empty BIK?
        If FileLen(vidName) <> 13572 Then 'is empty WMV?
            If FileLen(vidName) <> 16388 Then ' is empty MPG?
                If FileLen(vidName) <> 925696 Then   'is this a ChampionSheep Rally file?
                    NIMP.list.ListItems(ListViewINDEX).Checked = True ' enable checkbox so user knows he should remove them
                End If
            End If
        End If
    End If
Else
    If fso.FileExists(vidName & ".NIMP") = True Then ' if backup exists then ..
    NIMP.list.ListItems(ListViewINDEX).Checked = False ' .. disable checkbox in listview
    End If
End If
End If
End Function
Public Function check2(ListViewINDEX) ' file or backup exists? (2)
On Error Resume Next
Dim vidName As String
Dim l As ListItem
If ListViewINDEX > NIMP.list.ListItems.Count Then
MsgBox "Index not existent"
Exit Function
End If
vidName = NIMP.list.ListItems(ListViewINDEX).SubItems(1) ' Get video file name from ListView
vidName = NIMP.ret2.Caption & vidName ' Return full path of video file
Dim errorAhh As String: errorAhh = Left(NIMP.ret2.Caption, 5)
If errorAhh = "Error" Then
NIMP.selall.Enabled = False
Else
NIMP.selall.Enabled = True
If fso.FileExists(vidName) = True Then ' if original file exists then ...
    If FileLen(vidName) <> 3984 Then 'is empty BIK?
        If FileLen(vidName) <> 13572 Then 'is empty WMV?
            If FileLen(vidName) <> 16388 Then ' is empty MPG?
                If FileLen(vidName) <> 1849344 Then  'is this a ChampionSheep Rally file?
                    NIMP.list.ListItems(ListViewINDEX).Checked = True ' enable checkbox so user knows he should remove them
                End If
            End If
        End If
    End If
Else
    If fso.FileExists(vidName & ".NIMP") = True Then ' if backup exists then ..
    NIMP.list.ListItems(ListViewINDEX).Checked = False ' .. disable checkbox in listview
    End If
End If
End If
End Function

Public Function check_CondemnedCO(ListViewINDEX, Offset) ' file or backup exists? (for Condemned Criminal Origins)
On Error Resume Next
Dim vidName As String
Dim l As ListItem
If ListViewINDEX > NIMP.list.ListItems.Count Then
MsgBox "Index not existent"
Exit Function
End If
vidName = NIMP.list.ListItems(ListViewINDEX).SubItems(1) ' Get video file name from ListView
vidName = NIMP.ret.Caption & vidName ' Return full path of video file
' Open file in bynary mode and check if we made changes 2 it
Dim tmp As String
Open vidName For Binary As #1
Get #1, Offset, tmp
Close #1
If tmp = "X" Then
NIMP.list.ListItems(1).Checked = False
Else
NIMP.list.ListItems(1).Checked = True
End If
End Function
Public Function check_FEAR(ListViewINDEX, Offset) ' file or backup exists? (for F.E.A.R.)
On Error Resume Next
Dim vidName As String
Dim l As ListItem
If ListViewINDEX > NIMP.list.ListItems.Count Then
MsgBox "Index not existent"
Exit Function
End If
vidName = NIMP.list.ListItems(ListViewINDEX).SubItems(1) ' Get video file name from ListView
vidName = NIMP.ret.Caption & vidName ' Return full path of video file
' Open file in bynary mode and check if we made changes 2 it
Dim tmp As String
Open vidName For Binary As #1
Get #1, Offset, tmp
Close #1
If tmp = "X" Then
NIMP.list.ListItems(1).Checked = False
Else
NIMP.list.ListItems(1).Checked = True
End If
End Function
Public Function SaveResItemToDisk(ByVal iResourceNum As Integer, ByVal sResourceType As String, ByVal sDestFileName As String) As Long
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

Public Function Ovr(ListViewINDEX, resindex, resname) ' replace a file
On Error Resume Next
Dim vidName As String
Dim l As ListItem
If ListViewINDEX > NIMP.list.ListItems.Count Then
MsgBox "Index not existent"
Exit Function
End If
vidName = NIMP.list.ListItems(ListViewINDEX).SubItems(1) ' Get video file name from ListView
vidName = NIMP.ret.Caption & vidName ' Return full path of video file
SaveResItemToDisk resindex, resname, vidName
End Function

Public Function Ovr2(ListViewINDEX)  ' replace a file
On Error Resume Next
Dim vidName As String
Dim vidName2 As String
Dim l As ListItem
If ListViewINDEX > NIMP.list.ListItems.Count Then
MsgBox "Index not existent"
Exit Function
End If
vidName = NIMP.list.ListItems(ListViewINDEX).SubItems(1) ' Get video file name from ListView
vidName = NIMP.ret.Caption & vidName ' Return full path of video file
vidName2 = NIMP.ret.Caption & "TITLE.DBC" ' Return full path of video file

fso.CopyFile vidName2, vidName, True
End Function
'================================================================================
'================================================================================

Function Apply() ' APPLY FILE CHANGES
Dim X
For X = 1 To NIMP.list.ListItems.Count
    If NIMP.list.ListItems(X).Checked = True Then
    del (X)
    Else
    restore (X)
    End If
Next X
End Function

Function Apply2(resindex, resname) ' APPLY FILE CHANGES (SPECIAL)
Dim X
For X = 1 To NIMP.list.ListItems.Count
    If NIMP.list.ListItems(X).Checked = True Then
    del (X)
    Ovr X, resindex, resname
    Else
    restore (X)
    End If
Next X
End Function

Function Apply3() ' APPLY FILE CHANGES
Dim X
For X = 1 To NIMP.list.ListItems.Count
    If NIMP.list.ListItems(X).Checked = True Then
    del2 (X)
    Else
    restore2 (X)
    End If
Next X
End Function

Function Apply_CSR() ' APPLY FILE CHANGES (ChampionSheep Rally)
Dim X
For X = 1 To NIMP.list.ListItems.Count
    If NIMP.list.ListItems(X).Checked = True Then
    del (X)
    Ovr2 (X)
    Else
    restore (X)
    End If
Next X
End Function

Function Apply_CondemnedCO() ' APPLY FILE CHANGES (for CONDEMNED CRIMINAL ORIGINS)
If NIMP.list.ListItems(1).Checked = True Then
    del_CondemnedCO 1, 247421
    del_CondemnedCO 1, 247373
Else
    restore_CondemnedCO 1, 247421, "s"
    restore_CondemnedCO 1, 247373, "M"
End If
End Function
Function Apply_FEAR() ' APPLY FILE CHANGES (for F.E.A.R.)
If NIMP.list.ListItems(1).Checked = True Then
    del_FEAR 1, 213597
    del_FEAR 1, 214473
    del_FEAR 1, 214717
Else
    restore_FEAR 1, 213597, "M"
    restore_FEAR 1, 214473, "s"
    restore_FEAR 1, 214717, "W"
End If
End Function
Function Apply_LoTRWoR() ' APPLY FILE CHANGES (for Lord of the Rings War of the Ring)
If NIMP.list.ListItems(1).Checked = True Then
    del_FEAR 1, 39605175
Else
    restore_FEAR 1, 213597, "B"
End If
End Function
Function CheckFiles() ' CHECK IF THERE IS A BACKUP OR WE DIDN'T TOUCHED THE ORIGINAL FILES
Dim X
For X = 1 To NIMP.list.ListItems.Count
    check X
    check2 X
    If NIMP.Combo1.Text = "Condemned - Criminal Origins" Then
    check_CondemnedCO 1, 247421
    End If
    If NIMP.Combo1.Text = "F.E.A.R. First Encounter Assault and Recon" Then
    check_FEAR X, 214473
    End If
    If NIMP.Combo1.Text = "Lord of the Rings: War of the Ring" Then
    check_FEAR X, 39605175
    End If
    Next X
End Function

''''''
Public Function NumberToByte(newnum)
Dim byloop As Long
Dim newstr As String
Dim newbyte(1 To 4) As Byte
Dim myNum As String
myNum = Hex(newnum)
If Len(myNum) = 1 Then
myNum = "0" & myNum
End If
If Len(myNum) = 3 Then
myNum = "0" & myNum
End If
If Len(myNum) = 5 Then
myNum = "0" & myNum
End If
If Len(myNum) = 7 Then
myNum = "0" & myNum
End If
myNum = ConvertIntelToMotorola(myNum)
outLong(1) = Val("&H" & Mid(myNum, 1, 2) & "&")
End Function
Public Function ConvertIntelToMotorola(IntelHex As String)
ConvertIntelToMotorola = Mid(IntelHex, 7, 2) & Mid(IntelHex, 5, 2) & Mid(IntelHex, 3, 2) & Mid(IntelHex, 1, 2)
End Function

