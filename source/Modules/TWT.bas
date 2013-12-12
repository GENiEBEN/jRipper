Attribute VB_Name = "TWT"
' Carmageddon II

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Function CVL(ByVal ValToConvert As String) As Long
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
If fileEXT = "345" Then GetIcon = "unknown"
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


'=> Check how many files are stored in the TWT archive
Function TWT_get_TotalFiles(ByVal TWTFilePath As String)
' Dims
Dim byteNOF(1 To 4) As Byte
' Open file and get bytes 5-9
Open TWTFilePath For Binary As #5
Get #5, 5, byteNOF
Close #5
' Print result
TWT_get_TotalFiles = ByteToNumber(byteNOF)
End Function

'=> Check archive size
Function TWT_get_ArchiveSize(ByVal TWTFilePath As String)
' Dims
Dim byteNOF(1 To 4) As Byte
' Open file and get bytes 1-5
Open TWTFilePath For Binary As #5
Get #5, 1, byteNOF
Close #5
' Print result
TWT_get_ArchiveSize = ByteToNumber(byteNOF)
End Function

'=> Check were first file is stored
Function TWT_get_FirstFileOffset(ByVal TWTFilePath As String)
' Dims
Dim byteNOF(1 To 4) As Byte
Dim FileNumber As Long: FileNumber = TWT_get_TotalFiles(TWTFilePath)
Dim pos As Long: pos = 8 + (56 * (FileNumber - 1)) + 1 + 56
' Print result
TWT_get_FirstFileOffset = pos
End Function

'=> Check were files are stored (I don't no if there is any byte used as space between files, so i hope it's good code)..on test2.twt 6 BYTES are missing...maybe there are spaces between files
Function TWT_get_FileOffset(ByVal TWTFilePath As String, ByVal FileNumber As Long)
' Dims
Dim byteNOF(1 To 4) As Byte
Dim FN As Long: FN = TWT_get_TotalFiles(TWTFilePath)
Dim pos As Long
If FileNumber = 1 Then
 pos = 8 + (56 * (FN - 1)) + 1 + 56
 TWT_get_FileOffset = pos
 Exit Function
Else
    Dim x As Long
    Dim s As Long
    For x = 1 To (FileNumber)
    s = TWT_getOne_FileSize(TWTFilePath, x) + Val(s)
    Next x
    TWT_get_FileOffset = TWT_get_FirstFileOffset(TWTFilePath) + s
End If
End Function

'=> Check filesize
Function TWT_getOne_FileSize(ByVal TWTFilePath As String, ByVal FileNumber As Long)
' Dims
Dim byteNOF(1 To 4) As Byte
' Get first file
If FileNumber = 1 Then
    Open TWTFilePath For Binary As #5
    Get #5, 9, byteNOF
    Close #5
TWT_getOne_FileSize = ByteToNumber(byteNOF)
Exit Function
Else
' Get rest of files
    Open TWTFilePath For Binary As #5
    Get #5, 8 + (56 * (FileNumber - 1)) + 1, byteNOF
    Close #5
TWT_getOne_FileSize = ByteToNumber(byteNOF)
Exit Function
End If
' Print result
End Function

'=> Get filename for a stored file
Function TWT_getOne_FileName(ByVal TWTFilePath As String, ByVal FileNumber As Long)
' Dims
Dim FN As String * 56
Dim FE As String * 3
' Get Files
If FileNumber > TWT_get_TotalFiles(TWTFilePath) Then
TWT_getOne_FileName = "File Number not existent"
Exit Function
End If
If FileNumber = 1 Then
    Open TWTFilePath For Binary As #1
    Get #1, 12 + 1, FN
    Close #1
    TWT_getOne_FileName = FN
    Exit Function
Else
    Open TWTFilePath For Binary As #1
    Get #1, 12 + (56 * (FileNumber - 1)) + 1, FN
    Close #1
    TWT_getOne_FileName = FN
    Exit Function
End If
End Function

Function TWT_Decode(ByVal TWTFilePath As String, ByVal ListV As ListView)
' Dims
Dim strFileName As String * 56
Dim strFILEEXT As String * 3
Dim byteFL(1 To 4) As Byte
Dim l As ListItem
Dim str As String * 4
' Get Files
ListV.ListItems.Clear
For x = 1 To TWT_get_TotalFiles(TWTFilePath)
strFileName = TWT_getOne_FileName(TWTFilePath, x)
strFILEEXT = Split(strFileName, ".")(1)
        Set l = ListV.ListItems.Add(, , strFileName) ' Add filenames
        l.SmallIcon = GetIcon(strFILEEXT) ' Set icon
        l.SubItems(1) = TWT_getOne_FileSize(TWTFilePath, x) ' Put File Size
        If x = 1 Then
        l.SubItems(2) = TWT_get_FileOffset(TWTFilePath, x) & "-" & TWT_getOne_FileSize(TWTFilePath, 2)
        ElseIf x = TWT_get_TotalFiles(TWTFilePath) Then
        l.SubItems(2) = TWT_get_FileOffset(TWTFilePath, x) & "-" & TWT_get_ArchiveSize(TWTFilePath)
        Else
            If (TWT_getOne_FileSize(TWTFilePath, x + 1)) < (TWT_get_FileOffset(TWTFilePath, x)) Then
        l.SubItems(2) = TWT_get_FileOffset(TWTFilePath, x) & "-" & (TWT_get_FileOffset(TWTFilePath, x)) + (TWT_getOne_FileSize(TWTFilePath, x + 1))
            Else
        l.SubItems(2) = TWT_get_FileOffset(TWTFilePath, x) & "-" & (TWT_getOne_FileSize(TWTFilePath, x + 1))
            End If
        End If
        l.SubItems(3) = StrConv(strFILEEXT, vbUpperCase) ' Put Extension
Next x
End Function

