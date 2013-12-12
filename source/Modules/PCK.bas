Attribute VB_Name = "PCK"
Option Explicit

' version 1.0
'
' Commandos II - Men of Courage (*.pck)

Dim fso As New FileSystemObject
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
    ' Files with no extension
    Case "dir": GetIcon = "big"
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

'=> Get Header
Function PCK_get_Header(PCKFilePath As String)
PCK_get_Header = get_Header_Len4(PCKFilePath)
End Function
'=> Valid Header?
Function PCK_checkIfValid(PCKFilePath As String) As Boolean
If PCK_get_Header(PCKFilePath) = "DATA" Then
PCK_checkIfValid = True
Else
PCK_checkIfValid = False
End If
End Function

'=> Get FileName
Function PCK_getOne_FileName(PCKFilePath As String, FileNumber)
' Dims
Dim strFileName As String * 36
' Get Files
    Open PCKFilePath For Binary As #9 ' => open file in binary mode
    Get #9, 49 + (48 * (FileNumber - 1)), strFileName ' => jump to offset where info is stored
    PCK_getOne_FileName = strFileName
    Close #9 ' => close file
End Function

'=> Get Type
Function PCK_getOne_FileType(PCKFilePath As String, FileNumber)
' Dims
Dim byteNOF(1 To 4) As Byte
' Open file and get byte
Open PCKFilePath For Binary As #1
Get #1, 85 + (48 * (FileNumber - 1)), byteNOF
Close #1
' Print result
PCK_getOne_FileType = ByteToNumber(byteNOF)
If PCK_getOne_FileType = 1 Then
PCK_getOne_FileType = "DIR"
ElseIf PCK_getOne_FileType = 0 Then
PCK_getOne_FileType = "FILE"
ElseIf PCK_getOne_FileType = 255 Then
PCK_getOne_FileType = "DIR_END"
End If
End Function

'=> Get Size
Function PCK_getOne_FileSize(PCKFilePath As String, FileNumber)
' Dims
Dim byteNOF(1 To 4) As Byte
' Open file and get byte
Open PCKFilePath For Binary As #1
Get #1, 89 + (48 * (FileNumber - 1)), byteNOF
Close #1
' Print result
PCK_getOne_FileSize = ByteToNumber(byteNOF)
If PCK_getOne_FileSize = -1 Then
PCK_getOne_FileSize = "Folder"
Else
PCK_getOne_FileSize = PCK_getOne_FileSize
End If
End Function

'=> Get Size
Function PCK_getOne_FileOffset(PCKFilePath As String, FileNumber)
' Dims
Dim byteNOF(1 To 4) As Byte
' Open file and get byte
Open PCKFilePath For Binary As #1
Get #1, 93 + (48 * (FileNumber - 1)), byteNOF
Close #1
' Print result
PCK_getOne_FileOffset = ByteToNumber(byteNOF)
End Function

Function PCK_Decode(PCKFilePath As String, ListV As ListView, HowManyFiles As Long, ProgressBar As ProgressBar) '-> progressbar slows down things with 0,2 seconds on DATA.PCK
On Error Resume Next
' Dims
Dim strFileName As String * 36
Dim strFILEEXT As String * 3
Dim byteFL(1 To 4) As Byte
Dim l As ListItem
Dim str As String * 4
Dim x As Long
' Get Files
ListV.ListItems.Clear
ProgressBar.Visible = True
PCKFilePath = get_FileNameWithoutExtension(PCKFilePath)
If PCKFilePath = "DATA2.PCK" Then
ProgressBar.Max = 304 \ 100
ElseIf PCKFilePath = "DATA.PCK" Then
ProgressBar.Max = 4916 \ 100
End If
For x = 1 To HowManyFiles
strFileName = PCK_getOne_FileName(PCKFilePath, x)
strFILEEXT = get_ExtensionFromFileName_prv(strFileName)
        Set l = ListV.ListItems.add(, , strFileName) ' Add filenames
        l.SmallIcon = GetIcon(strFILEEXT) ' Set icon
        l.SubItems(1) = PCK_getOne_FileSize(PCKFilePath, x) ' Put File Size
        l.SubItems(2) = PCK_getOne_FileOffset(PCKFilePath, x) & "-" & PCK_getOne_FileOffset(PCKFilePath, x) + PCK_getOne_FileSize(PCKFilePath, x)
        l.SubItems(3) = StrConv(strFILEEXT, vbUpperCase) ' Put Extension
        ProgressBar.Value = x \ 100
Next x
' End
ProgressBar.Visible = False
End Function

Function ex(PCKFilePath As String, DestFolder As String, FileNumber) 'extract, shared by PCK_extractOne_2
' Dims
Dim strFFO As String: strFFO = PCK.PCK_getOne_FileOffset(PCKFilePath, FileNumber)
Dim strLFO As String: strLFO = (PCK.PCK_getOne_FileOffset(PCKFilePath, FileNumber) + PCK.PCK_getOne_FileSize(PCKFilePath, FileNumber))
Dim strFOP As String: strFOP = DestFolder & "\" & PCK_getOne_FileName(PCKFilePath, FileNumber)
Dim byteSTORE() As Byte
Dim filesize As Long
filesize = strLFO - strFFO
ReDim byteSTORE(filesize - 1)
' Open PCK
    Open PCKFilePath For Binary As #1
    Open strFOP For Binary As #2
    Get #1, strFFO + 1, byteSTORE
    DoEvents
    Put #2, 1, byteSTORE
    Close #1
    Close #2
End Function

Function PCK_extractOne_1(PCKFilePath As String, DestFolder As String, FileNumber) 'extract, just for DATA2.PCK
dE DestFolder, 1, "SONIDOS", "", ""
ex PCKFilePath, DestFolder & "\SONIDOS", FileNumber
End Function

Function PCK_extractOne_2(PCKFilePath As String, DestFolder As String, ProgressBar As ProgressBar) 'extract, just for DATA.PCK
Dim x
'
ex PCKFilePath, DestFolder, 1
ex PCKFilePath, DestFolder, 2
'
dE DestFolder, 2, "ANIMS", "ABI", ""
For x = 4 To 536
ex PCKFilePath, DestFolder & "\ANIMS", x
ProgressBar.Value = x
Next x
'
For x = 538 To 1643
ex PCKFilePath, DestFolder & "\ANIMS\ABI", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "ANIMS", "GRL", ""
For x = 1646 To 1708
ex PCKFilePath, DestFolder & "\ANIMS\GRL", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 1, "BRIEFING", "", ""
For x = 1712 To 1799
ex PCKFilePath, DestFolder & "\BRIEFING", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 1, "CARANIM", "", ""
ex PCKFilePath, DestFolder & "\CARANIM", 1802
'
dE DestFolder, 1, "FONTS", "", ""
For x = 1805 To 1812
ex PCKFilePath, DestFolder & "\FONTS", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 1, "FX", "", ""
ex PCKFilePath, DestFolder & "\FX", 1815
'
dE DestFolder, 1, "INTENDENCIA", "", ""
ex PCKFilePath, DestFolder & "\INTENDENCIA", 1818
ex PCKFilePath, DestFolder & "\INTENDENCIA", 1819
'
dE DestFolder, 1, "INTERFAZ", "", ""
For x = 1822 To 2230
ex PCKFilePath, DestFolder & "\INTERFAZ", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "INTERLUDIOS", "ABI", ""
For x = 2233 To 2236
ex PCKFilePath, DestFolder & "\INTERLUDIOS", x
ProgressBar.Value = x
Next x
'
ex PCKFilePath, DestFolder & "\INTERLUDIOS\ABI", 2238
'
dE DestFolder, 1, "MACROS", "", ""
For x = 2242 To 2573
ex PCKFilePath, DestFolder & "\MACROS", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 1, "MENUS", "", ""
For x = 2576 To 2599
ex PCKFilePath, DestFolder & "\MENUS", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "MISIONES", "BU", ""
For x = 2602 To 2605
ex PCKFilePath, DestFolder & "\MISIONES", x
ProgressBar.Value = x
Next x
'
For x = 2607 To 2685
ex PCKFilePath, DestFolder & "\MISIONES\BU", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "MISIONES", "CZ", ""
For x = 2688 To 2899
ex PCKFilePath, DestFolder & "\MISIONES\CZ", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "MISIONES", "ECL", ""
For x = 2902 To 2910
ex PCKFilePath, DestFolder & "\MISIONES\ECL", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "MISIONES", "HL", ""
For x = 2913 To 3003
ex PCKFilePath, DestFolder & "\MISIONES\HL", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "MISIONES", "IS", ""
For x = 3006 To 3101
ex PCKFilePath, DestFolder & "\MISIONES\IS", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "MISIONES", "KW", ""
For x = 3104 To 3143
ex PCKFilePath, DestFolder & "\MISIONES\KW", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "MISIONES", "PA", ""
For x = 3146 To 3206
ex PCKFilePath, DestFolder & "\MISIONES\PA", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "MISIONES", "PT", ""
For x = 3209 To 3297
ex PCKFilePath, DestFolder & "\MISIONES\PT", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "MISIONES", "RY", ""
For x = 3300 To 3342
ex PCKFilePath, DestFolder & "\MISIONES\RY", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "MISIONES", "SB", ""
For x = 3345 To 3463
ex PCKFilePath, DestFolder & "\MISIONES\SB", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "MISIONES", "SH", ""
For x = 3466 To 3483
ex PCKFilePath, DestFolder & "\MISIONES\SH", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "MISIONES", "TK", ""
For x = 3486 To 3496
ex PCKFilePath, DestFolder & "\MISIONES\TK", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "MISIONES", "TU01", ""
For x = 3499 To 3514
ex PCKFilePath, DestFolder & "\MISIONES\TU01", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "MISIONES", "TU02", ""
For x = 3517 To 3527
ex PCKFilePath, DestFolder & "\MISIONES\TU02", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "MISIONES", "TU03", ""
For x = 3530 To 3541
ex PCKFilePath, DestFolder & "\MISIONES\TU03", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "MISIONES", "TU04", ""
For x = 3544 To 3552
ex PCKFilePath, DestFolder & "\MISIONES\TU04", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 2, "MISIONES", "TU05", ""
For x = 3555 To 3562
ex PCKFilePath, DestFolder & "\MISIONES\TU05", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 1, "OBJETOSESPECIALES", "", ""
ex PCKFilePath, DestFolder & "\OBJETOSESPECIALES", 3566
'
dE DestFolder, 1, "PARAMETRICA", "", ""
ex PCKFilePath, DestFolder & "\PARAMETRICA", 3569
ex PCKFilePath, DestFolder & "\PARAMETRICA", 3570
'
dE DestFolder, 1, "RED", "", ""
ex PCKFilePath, DestFolder & "\RED", 3573
'
dE DestFolder, 3, "SONIDOS", "ESA", "ARTIFICIERO"
For x = 3576 To 3999
ex PCKFilePath, DestFolder & "\SONIDOS", x
ProgressBar.Value = x
Next x
'
ex PCKFilePath, DestFolder & "\SONIDOS\ESA", 4001
ex PCKFilePath, DestFolder & "\SONIDOS\ESA", 4002
'
For x = 4004 To 4058
ex PCKFilePath, DestFolder & "\SONIDOS\ESA\ARTIFICIERO", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 3, "SONIDOS", "ESA", "CAPITAN"
For x = 4061 To 4107
ex PCKFilePath, DestFolder & "\SONIDOS\ESA\CAPITAN", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 3, "SONIDOS", "ESA", "CHUSMA"
For x = 4110 To 4156
ex PCKFilePath, DestFolder & "\SONIDOS\ESA\CHUSMA", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 3, "SONIDOS", "ESA", "COMANDO"
For x = 4159 To 4213
ex PCKFilePath, DestFolder & "\SONIDOS\ESA\COMANDO", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 3, "SONIDOS", "ESA", "CONDUCTOR"
For x = 4216 To 4270
ex PCKFilePath, DestFolder & "\SONIDOS\ESA\CONDUCTOR", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 3, "SONIDOS", "ESA", "CORONEL"
For x = 4273 To 4319
ex PCKFilePath, DestFolder & "\SONIDOS\ESA\CORONEL", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 3, "SONIDOS", "ESA", "ESPIA"
For x = 4322 To 4376
ex PCKFilePath, DestFolder & "\SONIDOS\ESA\ESPIA", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 3, "SONIDOS", "ESA", "FRANCO"
For x = 4379 To 4433
ex PCKFilePath, DestFolder & "\SONIDOS\ESA\FRANCO", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 3, "SONIDOS", "ESA", "GHURKA"
For x = 4436 To 4482
ex PCKFilePath, DestFolder & "\SONIDOS\ESA\GHURKA", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 3, "SONIDOS", "ESA", "LADRON"
For x = 4485 To 4539
ex PCKFilePath, DestFolder & "\SONIDOS\ESA\LADRON", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 3, "SONIDOS", "ESA", "LAMA"
For x = 4542 To 4588
ex PCKFilePath, DestFolder & "\SONIDOS\ESA\LAMA", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 3, "SONIDOS", "ESA", "LANCHERO"
For x = 4591 To 4645
ex PCKFilePath, DestFolder & "\SONIDOS\ESA\LANCHERO", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 3, "SONIDOS", "ESA", "NATACHA"
For x = 4648 To 4702
ex PCKFilePath, DestFolder & "\SONIDOS\ESA\NATACHA", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 3, "SONIDOS", "ESA", "NAUFRAGO"
For x = 4705 To 4751
ex PCKFilePath, DestFolder & "\SONIDOS\ESA\NAUFRAGO", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 3, "SONIDOS", "ESA", "PILOTO"
For x = 4754 To 4800
ex PCKFilePath, DestFolder & "\SONIDOS\ESA\PILOTO", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 3, "SONIDOS", "ESA", "SMITH"
For x = 4803 To 4849
ex PCKFilePath, DestFolder & "\SONIDOS\ESA\SMITH", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 3, "SONIDOS", "ESA", "TRIPULACION"
For x = 4852 To 4898
ex PCKFilePath, DestFolder & "\SONIDOS\ESA\TRIPULACION", x
ProgressBar.Value = x
Next x
'
dE DestFolder, 1, "STR", "", ""
ex PCKFilePath, DestFolder & "\STR", 4903
'
dE DestFolder, 1, "CREDITOS", "", ""
For x = 4906 To 4914
ex PCKFilePath, DestFolder & "\CREDITOS", x
ProgressBar.Value = x
Next x
'
End Function

'=> Create dirs and subdirs (used when extracting)
Private Function dE(DestFolder, HowMany, DirName, DirName2, DirName3)
If HowMany > 3 Then
    MsgBox "Max. 3 folders can be created"
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

