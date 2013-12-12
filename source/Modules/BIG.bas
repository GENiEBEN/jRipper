Attribute VB_Name = "BIG"
' version 1.1
'
' ToCA Race Driver 1
' ToCA Race Driver 2
' ToCA Race Driver 3

' NEW in v1.1
' * Added progressbar when decoding BIGF archive

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
    Case "icz":       GetIcon = "ini"
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

'======================================================================
'=> Check if a file is a valid BIGF file
Public Function BIGF_CheckIfValid(BIGFilePath As String) As Boolean
' Dims
Dim strTMP As String: strTMP = BIGF_get_Header(BIGFilePath)
' Check
If strTMP = "BIGF" Then
BIGF_CheckIfValid = True
End If
End Function
'=> Check if a file is a valid BIGF file (BIGC - compressed BIGF)
Public Function BIGF_CheckIfValidBIGC(BIGFilePath As String) As Boolean
' Dims
Dim strTMP As String: strTMP = BIGF_get_Header(BIGFilePath)
' Check
If strTMP = "BIGC" Then
BIGF_CheckIfValidBIGC = True
End If
End Function
'=> Get file header (so we can check later if its a valid BIG_BIGF file)
Function BIGF_get_Header(BIGFilePath As String)
' Dims
Dim strHEADER As String * 4
' Open file and get first 4 offsets
Open BIGFilePath For Binary As #1
Get #1, 1, strHEADER
Close #1
' Print result
BIGF_get_Header = strHEADER
End Function

'=> Set file header (so we can check later if its a valid BIG_BIGF file)
Function BIGF_set_Header(BIGFilePath As String, NewHeaderName As String)
' Dims
Dim strHEADER As String * 4: strHEADER = NewHeaderName
' Open file and set first 4 offsets
Open BIGFilePath For Binary As #1
Put #1, 1, strHEADER
Close #1
End Function
'=> Check how many files are stored in the BIG_BIGF archive
Function BIGF_get_TotalFiles(BIGFilePath As String)
' Dims
Dim byteNOF(1 To 4) As Byte
' Open file and get bytes 5-9
Open BIGFilePath For Binary As #5
Get #5, 5, byteNOF
Close #5
' Print result
BIGF_get_TotalFiles = ByteToNumber(byteNOF)
End Function

'=> Check were first file offset is (thats where the files starts being stored)
Function BIGF_get_FirstFileOffset(BIGFilePath As String)
' Dims
Dim byteFFO(1 To 4) As Byte
' Open file and get bytes 9-12
Open BIGFilePath For Binary As #6
Get #6, 9, byteFFO
Close #6
' Print result
BIGF_get_FirstFileOffset = ByteToNumber(byteFFO)
End Function

'=> Not sure where this is used :-??
Function BIGF_get_PaddingMultiple(BIGFilePath As String)
' Dims
Dim bytePM(1 To 4) As Byte
' Open file and get bytes 13-16
Open BIGFilePath For Binary As #1
Get #1, 13, bytePM
Close #1
' Print result
BIGF_get_PaddingMultiple = ByteToNumber(bytePM)
End Function

'=> Get Additional info (who created the BIG_BIGF file)
Function BIGF_get_ArchiveCreator(BIGFilePath As String)
' Dims
Dim strAUTHOR As String * 20
' Open file and get bytes 17-36
Open BIGFilePath For Binary As #1
Get #1, 17, strAUTHOR
Close #1
' Print result
BIGF_get_ArchiveCreator = strAUTHOR
End Function

'=> Set Additional info (who created the BIG_BIGF file)
Function BIGF_set_ArchiveCreator(BIGFilePath As String, ArchiveCreator)
' Dims
Dim author As String * 20: author = ArchiveCreator 'Max 20 chars even if u type more
Dim free: free = FreeFile
' Open file
Open BIGFilePath For Binary As #free
Put #free, 17, author
Close #free
End Function
'=> Loop thru BIG_BIGF archive and fill a ListView object with all stored files
Function BIGF_getAll_FileName(BIGFilePath As String, ListV As ListView)
' Dims
Dim strFileName As String * 16
Dim strFILEEXT As String
' Get Files
ListV.ListItems.Clear
For x = 1 To BIGF_get_TotalFiles(BIGFilePath) ' => check how many files are stored in archive and start loop
    Open BIGFilePath For Binary As #1 ' => open file in binary mode
    Get #1, 13 + (24 * x), strFileName ' => loop thru filenames (each file is stored on 24bit : 16bit for fileName + 4bit for fileLen + 4bit for fileOffset)
    strFILEEXT = get_ExtensionFromFileName(strFileName) ' => get file extensions
    ListV.ListItems.add(, , strFileName).SmallIcon = GetIcon(strFILEEXT) ' => add files in a ListView object and set icon according fileExtension
    Close #1 ' => close file
Next x ' => loop until last file
End Function

'=> Get filename for a stored file (filname is in 12.3 filename format)
Function BIGF_getOne_FileName(BIGFilePath As String, Filenumber)
' Dims
Dim strFileName As String * 16
' Get Files
If Filenumber > BIGF_get_TotalFiles(BIGFilePath) Then
BIGF_getOne_FileName = "File Number not existent"
Exit Function
End If
    Open BIGFilePath For Binary As #9 ' => open file in binary mode
    Get #9, 13 + (24 * Filenumber), strFileName ' => jump to offset where info is stored
    BIGF_getOne_FileName = strFileName
    Close #9 ' => close file
End Function

'=> Loop thru BIG_BIGF archive and fill a ListView object with size in bytes of all stored files
Function BIGF_getAll_FileLength(BIGFilePath As String, ListV As ListView)
' Dims
Dim byteFL(1 To 4) As Byte
' Get Files
ListV.ListItems.Clear
For x = 1 To BIGF_get_TotalFiles(BIGFilePath) ' => check how many files are stored in archive and start loop
    Open BIGFilePath For Binary As #3 ' => open file in binary mode
    Get #3, 29 + (24 * x), byteFL ' => loop thru fileLens (each file is stored on 24bit : 16bit for fileName + 4bit for fileLen + 4bit for fileOffset)
    ListV.ListItems.add , , Format(ByteToNumber(byteFL), "###,###") ' => fill listview with all colected info
    Close #3
Next x
End Function

'=> Get filesize for a stored file and format it for better view
Function BIGF_getOne_FileLength(BIGFilePath As String, Filenumber)
' Dims
Dim byteFL(1 To 4) As Byte
' Get Files
If Filenumber > BIGF_get_TotalFiles(BIGFilePath) Then
BIGF_getOne_FileLength = "File Number not existent"
Exit Function
End If
For x = 1 To BIGF_get_TotalFiles(BIGFilePath) ' => check how many files are stored in archive and start loop
    Open BIGFilePath For Binary As #3 ' => open file in binary mode
    Get #3, (29 + (24 * Filenumber)), byteFL ' => jump to offset where info is stored
    BIGF_getOne_FileLength = Format(ByteToNumber(byteFL), "###,###") ' => store extracted info
    Close #3
Next x
End Function

'=> Get filesize for a stored file and dont format
Function BIGF_getOne_FileLength_UnFormated(BIGFilePath As String, Filenumber)
' Dims
Dim byteFL(1 To 4) As Byte
' Get Files
If Filenumber > BIGF_get_TotalFiles(BIGFilePath) Then
BIGF_getOne_FileLength_UnFormated = "File Number not existent"
Exit Function
End If
For x = 1 To BIGF_get_TotalFiles(BIGFilePath) ' => check how many files are stored in archive and start loop
    Open BIGFilePath For Binary As #3 ' => open file in binary mode
    Get #3, (29 + (24 * Filenumber)), byteFL ' => jump to offset where info is stored
    BIGF_getOne_FileLength_UnFormated = ByteToNumber(byteFL) ' => store extracted info
    Close #3
Next x
End Function

'=> Loop thru BIG_BIGF archive and fill a ListView object with the first offset where file starts being stored
Function BIGF_getAll_FileOffset_First(BIGFilePath As String, ListV As ListView)
' Dims
Dim byteFO As Byte
Dim byteFL(1 To 4) As Byte
' Get Files
ListV.ListItems.Clear
For x = 1 To BIGF_get_TotalFiles(BIGFilePath) ' => check how many files are stored in archive and start loop
    Open BIGFilePath For Binary As #3 ' => open file in binary mode
    Get #3, (33 + (24 * x)), byteFL  ' => loop thru fileLens (each file is stored on 24bit : 16bit for fileName + 4bit for fileLen + 4bit for fileOffset)
    Close #3
    ListV.ListItems.add , , ByteToNumber(byteFL) + BIGF_get_FirstFileOffset(BIGFilePath)
Next x
End Function

'=> Get first offset of a stored file (we need this for extracting a file)
Function BIGF_getOne_FileOffset_First(BIGFilePath As String, Filenumber)
' Dims
Dim byteFO As Byte
Dim byteFL(1 To 4) As Byte
If Filenumber > BIGF_get_TotalFiles(BIGFilePath) Then
BIGF_getOne_FileOffset_First = "File Number not existent"
Exit Function
End If
' Get Files
    Open BIGFilePath For Binary As #3 ' => open file in binary mode
    Get #3, (33 + (24 * Filenumber)), byteFL  ' => jump to offset where info is stored
    Close #3
    BIGF_getOne_FileOffset_First = ByteToNumber(byteFL) + BIGF_get_FirstFileOffset(BIGFilePath)
End Function

'=> Get last offset of a stored file (we need this for extracting a file)
Function BIGF_getOne_FileOffset_Last(BIGFilePath As String, Filenumber)
' Dims
Dim byteFO As Byte
Dim byteFL(1 To 4) As Byte
If Filenumber > BIGF_get_TotalFiles(BIGFilePath) Then
BIGF_getOne_FileOffset_Last = "File Number not existent"
Exit Function
End If
' Get Files
    Open BIGFilePath For Binary As #3 ' => open file in binary mode
    Get #3, (33 + (24 * Filenumber)), byteFL  ' => jump to offset where info is stored
    Close #3
    BIGF_getOne_FileOffset_Last = ByteToNumber(byteFL) + BIGF_get_FirstFileOffset(BIGFilePath) + BIGF_getOne_FileLength_UnFormated(BIGFilePath, Filenumber)
End Function

Function BIGF_extractOne(BIGFilePath As String, DestFolder As String, Filenumber)
' Dims
Dim strFFO As String: strFFO = BIGF_getOne_FileOffset_First(BIGFilePath, Filenumber)
Dim strLFO As String: strLFO = BIGF_getOne_FileOffset_Last(BIGFilePath, Filenumber) - 1
Dim strFOP As String: strFOP = DestFolder & "\" & BIGF_getOne_FileName(BIGFilePath, Filenumber)
Dim byteSTORE() As Byte
Dim filesize As Long
filesize = strLFO - strFFO
ReDim byteSTORE(filesize - 1)
' Open BIGF
    Open BIGFilePath For Binary As #1
    Open strFOP For Binary As #2
    Get #1, strFFO + 1, byteSTORE
    DoEvents
    Put #2, 1, byteSTORE
    Close #1
    Close #2
    'DoEvents
End Function

'=> Extract all files inside of BIGF archive
Function BIGF_extractAll(BIGFilePath As String, DestFolder As String)
Dim strNOF: strNOF = BIGF_get_TotalFiles(BIGFilePath)
For x = 1 To strNOF
BIGF_extractOne BIGFilePath, DestFolder, x
Next x
End Function

'=> Create an empty BIGF file : padding 2048
Public Function BIGF_makeEmptyBIGF(DestFile, ArchiveCreator, Optional Buffing As String)
On Error Resume Next
' Dims
Dim newdata As String * 2048
Dim author As String * 20: author = ArchiveCreator
Dim free: free = FreeFile
Dim ConvertIntelToMotorola
' Delete file if already exists
Kill DestFile
' Set buffer
If Buffing = "" Then Buffing = 2048
newdata = "BIGF" & Space(4) & Space(4) & Space(4) & author & String(20 - Len(author), 0) & String(2012, 0)
' Open File and write
Open DestFile For Binary As #free
Put #free, 1, newdata
    Dim newnum: newnum = 0
    Dim byloop As Long: Dim newstr As String: Dim newbyte(1 To 4) As Byte: Dim myNum As String: Dim outLong(1 To 4) As Byte
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
    ConvertIntelToMotorola = Mid(myNum, 7, 2) & Mid(myNum, 5, 2) & Mid(myNum, 3, 2) & Mid(myNum, 1, 2)
    myNum = ConvertIntelToMotorola
    outLong(1) = Val("&H" & Mid(myNum, 1, 2) & "&")
    outLong(2) = Val("&H" & Mid(myNum, 3, 2) & "&")
    outLong(3) = Val("&H" & Mid(myNum, 5, 2) & "&")
    outLong(4) = Val("&H" & Mid(myNum, 7, 2) & "&")
Put #free, 5, outLong
    newnum = 2048
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
    ConvertIntelToMotorola = Mid(myNum, 7, 2) & Mid(myNum, 5, 2) & Mid(myNum, 3, 2) & Mid(myNum, 1, 2)
    myNum = ConvertIntelToMotorola
    outLong(1) = Val("&H" & Mid(myNum, 1, 2) & "&")
    outLong(2) = Val("&H" & Mid(myNum, 3, 2) & "&")
    outLong(3) = Val("&H" & Mid(myNum, 5, 2) & "&")
    outLong(4) = Val("&H" & Mid(myNum, 7, 2) & "&")
Put #free, 9, outLong
    newnum = CLng(Buffing)
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
    ConvertIntelToMotorola = Mid(myNum, 7, 2) & Mid(myNum, 5, 2) & Mid(myNum, 3, 2) & Mid(myNum, 1, 2)
    myNum = ConvertIntelToMotorola
    outLong(1) = Val("&H" & Mid(myNum, 1, 2) & "&")
    outLong(2) = Val("&H" & Mid(myNum, 3, 2) & "&")
    outLong(3) = Val("&H" & Mid(myNum, 5, 2) & "&")
    outLong(4) = Val("&H" & Mid(myNum, 7, 2) & "&")
Put #free, 13, outLong
' Close file
Close #free
End Function

'=> This function reads all necesary info and starts filling ListView object with filename/filesize/offset/type etc....
Function BIGF_Decode(BIGFilePath As String, ListV As ListView, ProgressBar As ProgressBar)
' Dims
Dim strFileName As String * 16
Dim strFILEEXT As String * 3
Dim strFILEEXT2 As String
Dim byteFL(1 To 4) As Byte
Dim l As ListItem
Dim str As String * 4
' Get Files
ListV.ListItems.Clear
ProgressBar.Visible = True
ProgressBar.Max = BIGF_get_TotalFiles(BIGFilePath)
For x = 1 To BIGF_get_TotalFiles(BIGFilePath)
For Y = x To x
    Open BIGFilePath For Binary As #1
    Get #1, 13 + (24 * x), strFileName
    strFILEEXT2 = get_ExtensionFromFileName(strFileName)
    strFILEEXT = Split(strFileName, ".")(1)
        Open BIGFilePath For Binary As #2
        Get #2, (29 + (24 * Y)), byteFL
        Set l = ListV.ListItems.add(, , strFileName) ' Add filenames
        l.SmallIcon = GetIcon(strFILEEXT) ' Set icon
        l.SubItems(1) = Format(ByteToNumber(byteFL), "###,###") ' Put File Size
        l.SubItems(2) = BIGF_getOne_FileOffset_First(BIGFilePath, x) & "-" & BIGF_getOne_FileOffset_Last(BIGFilePath, x)
        l.SubItems(3) = StrConv(strFILEEXT2, vbUpperCase) ' Put Extension
        Close #2
    Close #1
Next Y
ProgressBar.Value = x
Next x
' End
ProgressBar.Visible = False
End Function




