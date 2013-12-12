Attribute VB_Name = "SLF"
' version 1.1
'
' Jagged Alliance 2
'
' NEW IN 1.1
'* loading is faster (on a 221MB file is 2 sec faster)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Dim fso As New FileSystemObject

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
    Case "sti": GetIcon = "image"
    ' Sound Files
    Case "wav": GetIcon = "wav"
    Case "mp3": GetIcon = "wav"
    Case "ogg": GetIcon = "wav"
    ' Movie Files
    Case "bik": GetIcon = "bik"
    ' ?! extension not mapped
    Case Else: GetIcon = "unknown"
End Select
End Function
'======================================================================

'=> Get Archive Name
Function SLF_get_ArchiveName(SLF_FilePath As String)
' Dims
Dim strHEADER As String * 256
' Open file and get first 256 bytes
Open SLF_FilePath For Binary As #1
Get #1, 1, strHEADER
Close #1
' Print result
SLF_get_ArchiveName = strHEADER
End Function

'=> Set ArchiveName
Function SLF_set_ArchiveName(SLF_FilePath As String, NewHeaderName As String)
' Dims
Dim strHEADER As String * 256: strHEADER = NewHeaderName
' Open file and set first 256 offsetsbytes
Open SLF_FilePath For Binary As #1
Put #1, 1, strHEADER
Close #1
End Function

'=> Get Directory Name
Function SLF_get_DirName(SLF_FilePath As String)
' Dims
Dim strHEADER As String * 256
' Open file and get next bytes 256-512
Open SLF_FilePath For Binary As #9
Get #9, 257, strHEADER
Close #9
' Print result
SLF_get_DirName = strHEADER
End Function

'=> Set dir name
Function SLF_set_Header(SLF_FilePath As String, NewHeaderName As String)
' Dims
Dim strHEADER As String * 256: strHEADER = NewHeaderName
' Open file and set bytes 256-512
Open SLF_FilePath For Binary As #1
Put #1, 1, strHEADER
Close #1
End Function

'=> Check how many files are stored in the SLF archive ( -1 file, which is added later)
Function SLF_get_TotalFiles(SLFFilePath As String)
' Dims
Dim byteNOF(1 To 4) As Byte
' Open file and get first byte
Open SLFFilePath For Binary As #3
Get #3, 513, byteNOF
Close #3
' Print result
SLF_get_TotalFiles = (ByteToNumber(byteNOF)) - 1
End Function

'=> Check FileTable Entry Offset
Function SLF_get_FileTableEntryOffset(SLF_FilePath As String)
' Dims
Dim nof As String: nof = SLF_get_TotalFiles(SLF_FilePath)
Dim fLen As String: fLen = FileLen(SLF_FilePath)
' Print result
SLF_get_FileTableEntryOffset = fLen - (nof * 280) - 280
End Function

'=> Get filename
Function SLF_getOne_FileName(SLFFilePath As String, FileNumber)
' Dims
Dim strFileName As String * 256
' Get offset start
SLF_getOne_FileName = 280 * (FileNumber - 1) + SLF_get_FileTableEntryOffset(SLFFilePath)
' Get FileName
    Open SLFFilePath For Binary As #1            ' => open file in binary mode
    Get #1, SLF_getOne_FileName + 1, strFileName ' => jump to offset where info is stored
    Close #1                                     ' => close file
' return result
SLF_getOne_FileName = strFileName
End Function

'=> Check data offset
Function SLF_getOne_FileDataOffset(SLFFilePath As String, FileNumber)
' Dims
Dim byteNOF(1 To 4) As Byte
' Open file and get first byte
Open SLFFilePath For Binary As #1
Get #1, 280 * (FileNumber - 1) + SLF_get_FileTableEntryOffset(SLFFilePath) + 1 + 256, byteNOF
Close #1
' Print result
SLF_getOne_FileDataOffset = ByteToNumber(byteNOF)
End Function

'=> Check length of file
Function SLF_getOne_FileLength(SLFFilePath As String, FileNumber)
' Dims
Dim byteNOF(1 To 4) As Byte
' Open file and get bytes
Open SLFFilePath For Binary As #2
Get #2, 280 * (FileNumber - 1) + SLF_get_FileTableEntryOffset(SLFFilePath) + 1 + 260, byteNOF
Close #2
' Print result
SLF_getOne_FileLength = ByteToNumber(byteNOF)
End Function

'=> count how many times "\" occurs in a FilePath
Private Function count_Char(filepath)
Dim tmp1
Dim tmp2
Dim C
For tmp1 = 1 To Len(filepath)
    tmp2 = Strings.Mid(filepath, tmp1, 1)
    If tmp2 = "\" Then
    C = Val(C) + 1
    End If
Next tmp1
count_Char = C
End Function

'=> Create dirs and subdirs (used when extracting)
Private Function create_Dirs(DestFolder, DirName, DirName2, DirName3, HowMany)
If HowMany > 3 Then
    MsgBox "Max. 3 Dirs can be created"
    Exit Function
Else
    If HowMany = 1 Then
        If fso.FolderExists(DestFolder & "\" & DirName) = False Then
        fso.CreateFolder DestFolder & "\" & DirName
        End If
        Exit Function
    Else
        If HowMany = 2 Then
            If fso.FolderExists(DestFolder & "\" & DirName) = False Then
            fso.CreateFolder DestFolder & "\" & DirName
            End If
            If fso.FolderExists(DestFolder & "\" & DirName & "\" & DirName2) = False Then
            fso.CreateFolder DestFolder & "\" & DirName & "\" & DirName2
            End If
            Exit Function
        ElseIf HowMany = 3 Then
            If fso.FolderExists(DestFolder & "\" & DirName) = False Then
            fso.CreateFolder DestFolder & "\" & DirName
            End If
            If fso.FolderExists(DestFolder & "\" & DirName & "\" & DirName2) = False Then
            fso.CreateFolder DestFolder & "\" & DirName & "\" & DirName2
            End If
            If fso.FolderExists(DestFolder & "\" & DirName & "\" & DirName2 & "\" & DirName3) = False Then
            fso.CreateFolder DestFolder & "\" & DirName & "\" & DirName2 & "\" & DirName3
            End If
            Exit Function
        End If
    End If
End If
End Function

''=> Extract a file v2.0 (fastest way - 19MB/s)
Function SLF_extractOne(SLFFilePath As String, DestFolder As String, FileNumber, ListV As ListView)
On Error Resume Next
'===DIMS========================================================================================
Dim FN: FN = SLF_getOne_FileName(SLFFilePath, FileNumber)         ' Get FileName
Dim xN                                                            ' like FN ... but avoiding duplicates
Dim SN: SN = get_FileNameWithoutExtension(FN)                     ' filename without extension
Dim FE: FE = get_ExtensionFromFileName(FN)                        ' File Extension
Dim FO: FO = SLF_getOne_FileDataOffset(SLFFilePath, FileNumber) ' Get Starting Offset
Dim EO                                                            ' store Ending Offset
Dim byteSTORE() As Byte                                             ' store bytes
Dim DN                                                            ' DestFolder 1 (maindir)
Dim DN2                                                           ' DestFolder 2 (subdir1)
Dim DN3                                                           ' DestFolder 3 (subdir2)
Dim l As ListItem                                                 '
Set l = ListV.ListItems(FileNumber)
Dim filesize As Long
'===CREATE-DIRS-AND-SUBDIRS======================================================================
EO = Split(l.SubItems(2), "-")(1)                                 ' Get Ending offset of a file
filesize = EO - FO
ReDim byteSTORE(filesize - 1)
C = count_Char(FN)                                                ' check how many "\" contains filename
If C = 1 Then                                                     ' if there's one "\" then create one dir
DN = Split(FN, "\")(0)                                            ' dirname1
FN = Split(FN, "\")(1)                                            ' filename (dir removed)
create_Dirs DestFolder, DN, "", "", 1                             ' create dir in Windows
ElseIf C = 2 Then                                                 ' if there are two "\" then create one dir and a subdir
DN = Split(FN, "\")(0)                                            ' dirname1
DN2 = Split(FN, "\")(1)                                           ' dirname2
FN = Split(FN, "\")(2)                                            ' filename (dir's removed)
create_Dirs DestFolder, DN, DN2, "", 2                            ' create dir\subdir
ElseIf C = 3 Then                                                 ' if there are three "\" then create one dir a subdir and sub-subdir
DN = Split(FN, "\")(0)                                            ' dirname1
DN2 = Split(FN, "\")(1)                                           ' dirname2
DN3 = Split(FN, "\")(2)                                           ' dirname3
FN = Split(FN, "\")(3)                                            ' filename (dir's removed)
create_Dirs DestFolder, DN, DN2, DN3, 3                           ' create dir\subdir\subdir
End If                                                            '
'===EXTRACT======================================================================================
Open SLFFilePath For Binary As #1                                 ' open source archive
    If C = 1 Then                                                     ' check if we're gonna extract in main_dir\
    DestFolder = DestFolder & "\" & DN                                ' OK, extract in main_dir\
    ElseIf C = 2 Then                                                 ' check if we're gonna extract in main_dir\sub_dir1\
    DestFolder = DestFolder & "\" & DN & "\" & DN2                    ' OK, extract in main_dir\sub_dir1\
    ElseIf C = 3 Then                                                 ' check if we're gonna extract in main_dir\sub_dir1\sub_dir2\
    DestFolder = DestFolder & "\" & DN & "\" & DN2 & "\" & DN3        ' OK, extract in main_dir\sub_dir1\sub_dir2\
    End If                                                            '

If fso.FileExists(DestFolder & "\" & FN) = False Then
    Open DestFolder & "\" & FN For Binary As #2                       ' open destination file (empty)
    Get #1, FO + 1, byteSTORE                                    ' get from source archive
    Put #2, 1, byteSTORE                                      ' put in new file
    If C = 1 Then                                                     ' check dest folder
    DestFolder = Replace(DestFolder, DN, "")                          ' go one level up
    ElseIf C = 2 Then                                                 ' check dest folder
    DestFolder = Replace(DestFolder, DN & "\" & DN2, "")              ' go two levels up
    ElseIf C = 3 Then                                                 ' check dest folder
    DestFolder = Replace(DestFolder, DN & "\" & DN2 & "\" & DN3, "")  ' go three levels up
    End If                                                            '
Else
 xN = SN & " ~2nd~." & FE                                         '
    Open DestFolder & "\" & xN For Binary As #2                       ' open destination file (empty)
    Get #1, FO + 1, byteSTORE                                    ' get from source archive
    Put #2, 1, byteSTORE                                      ' put in new file
    If C = 1 Then                                                     ' check dest folder
    DestFolder = Replace(DestFolder, DN, "")                          ' go one level up
    ElseIf C = 2 Then                                                 ' check dest folder
    DestFolder = Replace(DestFolder, DN & "\" & DN2, "")              ' go two levels up
    ElseIf C = 3 Then                                                 ' check dest folder
    DestFolder = Replace(DestFolder, DN & "\" & DN2 & "\" & DN3, "")  ' go three levels up
    End If                                                            '
End If
Close #1                                                          ' Close source archive
Close #2                                                          ' close dest file
Exit Function                                                     ' Exit this shit :)
End Function

'=> Extract all files
Function SLF_extractAll(SLFFilePath As String, DestFolder As String, ListV As ListView)
' Dims
Dim TF: TF = SLF_get_TotalFiles(SLFFilePath)
Dim X
' Extract
For X = 1 To TF
SLF_extractOne SLFFilePath, DestFolder, X, ListV
Next X
End Function
'============
' Fill a ListView with all the FileNames (and starting-ending offset + size)
Function SLF_decode(SLFFilePath As String, ListV As ListView, ProgressBar As ProgressBar)
On Error Resume Next
' Dims
Dim X
Dim Y: Y = SLF_get_TotalFiles(SLFFilePath) + 1
Dim Z
Dim l As ListItem
Dim strFE As String * 3
Dim strFE2 As String
Dim GY
Dim GX
Dim FX
Dim FY
Dim GZ
'
ProgressBar.Visible = True
' Fill listbox
For X = 1 To Y
    Z = SLF_getOne_FileName(SLFFilePath, X)
    GY = SLF_getOne_FileDataOffset(SLFFilePath, Y)
    GX = SLF_getOne_FileDataOffset(SLFFilePath, X)
    GZ = SLF_getOne_FileDataOffset(SLFFilePath, 1)
    FX = SLF_getOne_FileLength(SLFFilePath, X)
    FY = SLF_getOne_FileLength(SLFFilePath, Y)
    ' Get file extension
    strFE2 = Z
    strFE = get_ExtensionFromFileName(strFE2)
    ' Add FileNames
    Set l = ListV.ListItems.Add(, , Z)
    ' Set smallicon for each file (cos whe know the file extension)
    l.SmallIcon = GetIcon(strFE)
    ' Get first and last offset of each stored file
    Dim byteNOF(1 To 4) As Byte
    Dim X1: X1 = GX
    Open SLFFilePath For Binary As #1
    Get #1, X1 + 1, byteNOF
    Close #1
    If X = 1 Then
    l.SubItems(2) = GZ & "-" & GZ + FX
    ElseIf X = Y Then
        If FileLen(SLFFilePath) > GY Then
    l.SubItems(2) = GY & "-" & GY + FY
        Else
    l.SubItems(2) = GY & "-" & GY + FY
        End If
    Else
    l.SubItems(2) = GX & "-" & GX + FX
    End If
    ' Get size of each file
    Dim LO
    Dim FO
    l.SubItems(1) = FX
    ' Add extension in a ListView Column
    l.SubItems(3) = StrConv(strFE, vbUpperCase)
    ' Progress bar
    ProgressBar.Max = Y \ 200
    ProgressBar.Value = X \ 200
Next X
'End
ProgressBar.Visible = False
End Function
