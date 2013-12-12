Attribute VB_Name = "BSP"
' version 1.0.0
'
' Return to Castle Wolfenstein SBWL (russian)

' KNOWN BUGS
' * some archives returns bad getTotalFiles

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
    Case "zzz": GetIcon = "big"
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

'=> Get Header
Function BSP_get_Header(BSPFilePath As String)
BSP_get_Header = jrBytes.get_Header_Len4(BSPFilePath)
End Function
'=> Valid Header?
Function BSP_checkIfValid(BSPFilePath As String) As Boolean
If BSP_get_Header(BSPFilePath) = "IBSP" Then
BSP_checkIfValid = True
Else
BSP_checkIfValid = False
End If
End Function
'=> Get Total Files
Function BSP_get_TotalFiles(BSPFilePath As String)
BSP_get_TotalFiles = jrBytes.get_1Byte(BSPFilePath, 9)
End Function
'=> Get First FileName List Entry
Function BSP_get_FirstFileNameEntry(BSPFilePath As String)
BSP_get_FirstFileNameEntry = jrBytes.get_4Bytes(BSPFilePath, 17)
End Function
'=> Get First File Starting Offset
Function BSP_get_FirstFileOffset(BSPFilePath As String)
BSP_get_FirstFileOffset = jrBytes.get_4Bytes(BSPFilePath, 25)
End Function
'=> Get filename
Function BSP_getOne_FileName(BSPFilePath As String, FileNumber)
' Dims
Dim strFileName As String * 72
' Get offset start
If FileNumber = 1 Then
BSP_getOne_FileName = BSP_get_FirstFileNameEntry(BSPFilePath)
Else
BSP_getOne_FileName = 72 * (FileNumber - 1) + BSP_get_FirstFileNameEntry(BSPFilePath)
End If
' Get FileName
    Open BSPFilePath For Binary As #1            ' => open file in binary mode
    Get #1, BSP_getOne_FileName + 1, strFileName ' => jump to offset where info is stored
    BSP_getOne_FileName = strFileName            ' => Store filename so we can return it as result
    Close #1                                     ' => close file
' return
BSP_getOne_FileName = strFileName
End Function

'=======================================================================================
' Fill a ListView with all the FileNames
Function BSP_decode(BSPFilePath As String, ListV As ListView, ProgressBar As ProgressBar)
On Error Resume Next
' Dims
Dim x
Dim y
Dim strFN
Dim l As ListItem
Dim strFE
Dim strFE2 As String * 3
Dim strFE3
' Set progressbar
ProgressBar.Visible = True
ProgressBar.Max = BSP_get_TotalFiles(BSPFilePath) / 10
' Fill listbox
For x = 1 To BSP_get_TotalFiles(BSPFilePath)
    ' Get file extension
    strFE = jRFct.get_ExtensionFromFileName(BSP_getOne_FileName(BSPFilePath, x))
    strFE2 = StrConv(strFE, vbUpperCase)
    ' Add FileNames
    FN = BSP_getOne_FileName(BSPFilePath, x)
    Set l = ListV.ListItems.add(, , FN)
    ' Set smallicon for each file (cos whe know the file extension)
    l.SmallIcon = GetIcon(strFE2)
    ' Add extension in a ListView Column
    l.SubItems(3) = strFE2
    ' Progress bar
    ProgressBar.Value = x
Next x
'End
ProgressBar.Visible = False
End Function



