Attribute VB_Name = "WAD"
' version 2.0

' Tomb Raider 3 - The Lost Artifact/ Adventures of Lara Croft

' NEW in v2.0
' *Added code for Extracting .wad files

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

Private Function GetIcon(fileEXT)
' Hmm, what icon to use?
If StrConv(fileEXT, vbLowerCase) = "bctg" And Len(fileEXT) = 3 Then GetIcon = "3ds": Exit Function 'bctg
If StrConv(fileEXT, vbLowerCase) = "btag" And Len(fileEXT) = 3 Then GetIcon = "3ds": Exit Function 'btag
If StrConv(fileEXT, vbLowerCase) = "ctag" And Len(fileEXT) = 3 Then GetIcon = "3ds": Exit Function 'ctag
If StrConv(fileEXT, vbLowerCase) = "itag" And Len(fileEXT) = 3 Then GetIcon = "3ds": Exit Function 'itag
If Right(fileEXT, 1) = "5" Then GetIcon = "big" '5
If fileEXT = "345" Then GetIcon = "unknown" '345
fileEXT = StrConv(fileEXT, vbLowerCase)
Select Case fileEXT
    ' Archives
    Case "rar": GetIcon = "rar"
    Case "lzs": GetIcon = "big"
    Case "big": GetIcon = "big"
    Case "bgx": GetIcon = "big"
    Case "tbf": GetIcon = "big"
    ' Language Files
    Case "lng":       GetIcon = "lng"
    ' Text Files
    Case "txt":       GetIcon = "txt"
    Case "ref":       GetIcon = "txt"
    Case "play":      GetIcon = "txt"
    Case "xmod":      GetIcon = "txt"
    Case "gt":        GetIcon = "txt"
    Case "cc":        GetIcon = "txt"
    Case "em":        GetIcon = "txt"
    Case "plist":     GetIcon = "txt"
    Case "lvl":       GetIcon = "txt"
    Case "pdef":      GetIcon = "txt"
    Case "hood":      GetIcon = "txt"
    Case "bnd":       GetIcon = "txt"
    Case "td":        GetIcon = "txt"
    Case "prop":      GetIcon = "txt"
    Case "sndref":    GetIcon = "txt"
    Case "parfileio": GetIcon = "txt"
    Case "ini":       GetIcon = "ini"
    Case "inf":       GetIcon = "ini"
    Case "icz":       GetIcon = "ini"
    ' gorky17/odium
    Case "dlg":      GetIcon = "txt"
    Case "ar":        GetIcon = "txt" ' no use because i can't return "AR" from filename (small bug)
    Case "cm":        GetIcon = "txt" ' no use because i can't return "CM" from filename (small bug)
    Case "lev":        GetIcon = "txt"
    Case "lts":        GetIcon = "txt"
    Case "ftr":        GetIcon = "txt"
    Case "wpn":        GetIcon = "txt"
    Case "hro":        GetIcon = "txt"
    Case "cfg":        GetIcon = "txt"
    Case "ar":        GetIcon = "txt"
    Case "itm":        GetIcon = "txt"
    Case "are":        GetIcon = "txt"
    Case "aba":        GetIcon = "txt"
    Case "pth":        GetIcon = "txt"
    Case "ba":        GetIcon = "txt"
    Case "tab":        GetIcon = "txt"
    Case "dsc":        GetIcon = "txt"
    Case "var":        GetIcon = "txt"
    Case "spr": GetIcon = "3ds"
    ' 3DS files
    Case "dff": GetIcon = "3ds"
    Case "col": GetIcon = "3ds"
    Case "ipl": GetIcon = "3ds"
    Case "p3d": GetIcon = "3ds"
    Case "wtr": GetIcon = "3ds"
    Case "tas": GetIcon = "3ds"
    Case "tmf": GetIcon = "3ds"
    Case "tan": GetIcon = "3ds"
    Case "hz1": GetIcon = "3ds"
    Case "rsg": GetIcon = "3ds"
    Case "sh2": GetIcon = "3ds"
    ' Image Files
    Case "dds": GetIcon = "image"
    Case "tga": GetIcon = "image"
    Case "bmp": GetIcon = "image"
    Case "jpg": GetIcon = "image"
    Case "pct": GetIcon = "image"
    Case "gif": GetIcon = "image"
    Case "tif": GetIcon = "image"
    Case "tex": GetIcon = "image"
    Case "txd": GetIcon = "image"
    ' Sound Files
    Case "wav": GetIcon = "wav"
    Case "mp3": GetIcon = "wav"
    ' Movie Files
    Case "bik": GetIcon = "bik"
    ' ?! extension not mapped
    Case Else: GetIcon = "unknown"
End Select
End Function


'=> Check were first file is stored
Function WAD_get_FirstFileOffset(WADFilePath As String)
' Dims
Dim byteNOF(1 To 4) As Byte
' Open and read
    Open WADFilePath For Binary As #1
    Get #1, 801, byteNOF
    Close #1
' Print results
WAD_get_FirstFileOffset = ByteToNumber(byteNOF)
End Function

'=> Get FileName
Function WAD_getOne_FileName(WADFilePath As String, FileNumber)
' Dims
Dim strFileName As String * 260 ' those guys that made the .WAD format were stupid...windowz can't handle more than 255
Dim strFILEEXT As String * 3
' Get Files
If FileNumber = 1 Then
    Open WADFilePath For Binary As #2 ' => open file in binary mode
    Get #2, 537, strFileName ' => jump to offset where info is stored
    WAD_getOne_FileName = strFileName
    Close #2 ' => close file
Else
    Open WADFilePath For Binary As #9 ' => open file in binary mode
    Get #9, 537 + (268 * (FileNumber - 1)), strFileName ' => jump to offset where info is stored
    WAD_getOne_FileName = strFileName
    Close #9 ' => close file
End If
End Function

'=> Get FileLength
Function WAD_getOne_FileLength(WADFilePath As String, FileNumber)
' Dims
Dim byteNOF(1 To 4) As Byte
' Get Files
If FileNumber = 1 Then
    Open WADFilePath For Binary As #2 ' => open file in binary mode
    Get #2, 797, byteNOF ' => jump to offset where info is stored
    WAD_getOne_FileLength = ByteToNumber(byteNOF)
    Close #2 ' => close file
Else
    Open WADFilePath For Binary As #9 ' => open file in binary mode
    Get #9, 797 + (268 * (FileNumber - 1)), byteNOF ' => jump to offset where info is stored
    WAD_getOne_FileLength = ByteToNumber(byteNOF)
    Close #9 ' => close file
End If
End Function

'=> Get FileOffset Start
Function WAD_getOne_FileStartOffset(WADFilePath As String, FileNumber)
' Dims
Dim byteNOF(1 To 4) As Byte
' Get Files
If FileNumber = 1 Then
    Open WADFilePath For Binary As #2 ' => open file in binary mode
    Get #2, 801, byteNOF ' => jump to offset where info is stored
    WAD_getOne_FileStartOffset = ByteToNumber(byteNOF)
    Close #2 ' => close file
Else
    Open WADFilePath For Binary As #9 ' => open file in binary mode
    Get #9, 801 + (268 * (FileNumber - 1)), byteNOF ' => jump to offset where info is stored
    WAD_getOne_FileStartOffset = ByteToNumber(byteNOF)
    Close #9 ' => close file
End If
End Function

Function WAD_extractOne(WADFilePath As String, DestFolder As String, FileNumber)
' Dims
Dim strFFO As String: strFFO = WAD_getOne_FileStartOffset(WADFilePath, FileNumber)
Dim strLFO As String: strLFO = WAD_getOne_FileStartOffset(WADFilePath, FileNumber) + WAD_getOne_FileLength(WADFilePath, FileNumber)
Dim strFOP As String: strFOP = DestFolder & "\" & WAD_getOne_FileName(WADFilePath, FileNumber)
Dim byteSTORE() As Byte
Dim filesize As Long: filesize = strLFO - strFFO
ReDim byteSTORE(filesize - 1)
' Open WAD
    Open WADFilePath For Binary As #1
    Open strFOP For Binary As #2
    Get #1, strFFO + 1, byteSTORE
    DoEvents
    Put #2, 1, byteSTORE
    Close #1
    Close #2
End Function

Function WAD_Decode(WADFilePath As String, ListV As ListView)
On Error Resume Next
' Dims
Dim strFileName As String * 56
Dim strFILEEXT As String * 3
Dim byteFL(1 To 4) As Byte
Dim l As ListItem
Dim str As String * 4
' Get Files
ListV.ListItems.Clear
For x = 1 To 122 '(we know that there are 122 files stored in the cdAudio.wad....there is now way to count files)
    If x > 77 And x < 81 Then 'Skip 79 80 81 ..
    x = 81                    '.. to 81
    End If
strFileName = WAD_getOne_FileName(WADFilePath, x)
strFILEEXT = Split(strFileName, ".")(1)
        Set l = ListV.ListItems.Add(, , strFileName) ' Add filenames
        l.SmallIcon = GetIcon(strFILEEXT) ' Set icon
        l.SubItems(1) = WAD_getOne_FileLength(WADFilePath, x) ' Put File Size
        l.SubItems(2) = WAD_getOne_FileStartOffset(WADFilePath, x) & "-" & WAD_getOne_FileStartOffset(WADFilePath, x) + WAD_getOne_FileLength(WADFilePath, x)
        l.SubItems(3) = StrConv(strFILEEXT, vbUpperCase) ' Put Extension
Next x
End Function



