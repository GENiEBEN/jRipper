Attribute VB_Name = "SChl"
Option Explicit

' version 1.0
'
' Various games (including FIFA/NFS/MVP/NHL) (*.ast/*.dat)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Function CVL(ValToConvert As String) As Long
Dim WorkLong As Long
CopyMemory WorkLong, ByVal ValToConvert, 4
CVL = WorkLong
End Function

Private Function ByteToNumber(mybyte() As Byte)
Dim byloop As Long
Dim newstr As String
For byloop = 1 To UBound(mybyte)
newstr = newstr & Chr(mybyte(byloop))
Next byloop
ByteToNumber = CVL(newstr)
End Function

'=> Get Header
Function SCHl_get_Header(SCHlFilePath As String)
SCHl_get_Header = jrBytes.get_Header_Len4(SCHlFilePath)
End Function
'=> Valid Header?
Function SCHl_checkIfValid(SCHlFilePath As String) As Boolean
If SCHl_get_Header(SCHlFilePath) = "SCHl" Then
SCHl_checkIfValid = True
Else
SCHl_checkIfValid = False
End If
End Function

'=> Split Header offset
Function SCHl_get_SplitHeader(SCHlFilePath As String)
SCHl_get_SplitHeader = jrBytes.get_1Byte(SCHlFilePath, 5)
End Function

'=> Count Files (not fastest way) TODO!
Function SCHl_get_Headers(SCHlFilePath As String, ProgressBar As ProgressBar, ListV As ListView)
Dim X As Long
Dim c As Long
Dim l As ListItem
Dim speed As Long
'
ListV.Visible = True
ProgressBar.Visible = True
DoEvents
ProgressBar.Max = FileLen(SCHlFilePath)

speed = 4

'
For X = 1 To FileLen(SCHlFilePath) Step speed
SCHl_get_Headers = get_4chars(SCHlFilePath, X)
    If SCHl_get_Headers = "SCHl" Then
        c = c + 1
        Set l = ListV.ListItems.add(, , "File " & c) ' Add filenames
        l.SubItems(2) = X ' Add Start Offset
        l.SubItems(3) = "ASF" ' Add Extension (there is no filename.ext , so i use the extension Game Extractor uses :-P )
        l.SmallIcon = "wav"
        DoEvents
    ElseIf SCHl_get_Headers = "SCEl" Or Left(SCHl_get_Headers, 3) = "CEl" Then
    l.ListSubItems(2).Text = l.ListSubItems(2).Text & "-" & X + 4
    End If
ProgressBar.Value = X
Next X
'
SCHl_get_Headers = c
ProgressBar.Visible = False
End Function

Function SCHl_extractOne(SCHlFilePath As String, DestFolder As String, Filenumber, ListV As ListView)
' Dims
Dim l As ListItem
Set l = ListV.ListItems(Filenumber)
Dim strFFO As String: strFFO = (Split(l.SubItems(2), "-")(0)) - 1
Dim strLFO As String: strLFO = (Split(l.SubItems(2), "-")(1)) - 1
Dim strFOP As String: strFOP = DestFolder & "\" & l.Text & ".ASF"
Dim byteSTORE() As Byte
Dim filesize As Long
filesize = strLFO - strFFO
ReDim byteSTORE(filesize - 1)
' Open SCHl
    Open SCHlFilePath For Binary As #1
    Open strFOP For Binary As #2
    Get #1, strFFO + 1, byteSTORE
    DoEvents
    Put #2, 1, byteSTORE
    Close #1
    Close #2
    'DoEvents
End Function

Function get_4chars(FilePath As String, Pos)
' Dims
Dim strHEADER As String * 4
' Open file and get 4chars
Open FilePath For Binary As #1
Get #1, Pos, strHEADER
Close #1
' Print result
get_4chars = strHEADER
End Function

