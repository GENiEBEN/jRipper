Attribute VB_Name = "VP6"
'=>  version 2.0
'===============================
' Need For Speed Most Wanted
' Need For Speed Undergroud 2
' FIFA Manager 2006
' EA Cricket 2005
'........+other..................

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

'=> Check if a file is a valid VP6 file
Function MVhd_get_Header(MVhdFilePath As String)
' Print result
MVhd_get_Header = get_Header_Len4(MVhdFilePath)
End Function

'=> Check if a file is a valid MVhd file
Public Function MVhd_CheckIfValid(MVhdFilePath As String) As Boolean
' Dims
Dim strTMP As String: strTMP = get_Header_Len4(MVhdFilePath)
' Check
If strTMP = "MVhd" Then
MVhd_CheckIfValid = True
End If
End Function

'=> Split Header offset
Function MVhd_get_SplitHeader(MVhdFilePath As String)
MVhd_get_SplitHeader = MyFunctions.get_1Byte(MVhdFilePath, 5)
End Function

'=> Count Files (not fastest way, but is good enough)
Function MVhd_get_Headers(MVhdFilePath As String, ProgressBar As ProgressBar, ListV As ListView)
Dim x As Long
Dim c As Long
Dim l As ListItem
Dim speed As Long
'
ListV.Visible = True
ProgressBar.Visible = True
DoEvents
ProgressBar.Max = FileLen(MVhdFilePath)

speed = 4

'
For x = 1 To FileLen(MVhdFilePath) Step speed
MVhd_get_Headers = get_4chars(MVhdFilePath, x)
    If MVhd_get_Headers = "SCHl" Then
        c = c + 1
        Set l = ListV.ListItems.add(, , "File " & c) ' Add filenames
        l.SubItems(2) = x ' Add Start Offset
        l.SubItems(3) = "MVHD" ' Add Extension
        l.SmallIcon = "unknown"
        DoEvents
    ElseIf MVhd_get_Headers = "SCEl" Or Left(MVhd_get_Headers, 3) = "CEl" Then
    l.ListSubItems(2).Text = l.ListSubItems(2).Text & "-" & x + 4
    End If
ProgressBar.Value = x
Next x
'
MVhd_get_Headers = c
ProgressBar.Visible = False
End Function

Function MVhd_extractOne(MVhdFilePath As String, DestFolder As String, Filenumber, ListV As ListView)
' Dims
Dim l As ListItem
Set l = ListV.ListItems(Filenumber)
Dim strFFO As String: strFFO = (Split(l.SubItems(2), "-")(0)) - 1
Dim strLFO As String: strLFO = (Split(l.SubItems(2), "-")(1)) - 1
Dim strFOP As String: strFOP = DestFolder & "\" & l.Text & ".MVhd"
Dim byteSTORE() As Byte
Dim filesize As Long
filesize = strLFO - strFFO
ReDim byteSTORE(filesize - 1)
' Open MVhd
    Open MVhdFilePath For Binary As #1
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



