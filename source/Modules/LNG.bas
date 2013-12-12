Attribute VB_Name = "LNG"
' ToCA Race Driver 3

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Dim lDex As Long
Dim outLong(1 To 4) As Byte
Dim outByte() As Byte

'================================================================================================================================================================================================
'=LNG=OPEN=======================================================================================================================================================================================
'================================================================================================================================================================================================

Public Function LNG_open(FilePath, ListV As ListView)
' Dims
Dim lngLoop As Long
Dim lngData() As Byte
Dim lngdat
Dim flnum As Long
Dim LNG(1 To 4) As Byte
Dim lngTmp As Long
Dim lngTmp2 As Long
Dim flen As Long
Dim free
free = FreeFile
flen = FileLen(FilePath)
' Open File and get items
Open FilePath For Binary As #free
Get #free, 9, LNG
flnum = ByteToNumber(LNG)

For lngLoop = 0 To flnum - 1
    Get #free, (lngLoop * 4) + 13, LNG
    lngTmp = ByteToNumber(LNG)
    
    If (lngLoop + 1) < flnum Then
        Get #free, ((lngLoop + 1) * 4) + 13, LNG
        lngTmp2 = ByteToNumber(LNG)
        ReDim lngData(1 To (lngTmp2 - lngTmp))
    Else
        ReDim lngData(1 To (flen - (((flnum * 4) + 13) + lngTmp)) + 1)
    End If
    '
    Get #free, ((flnum * 4) + 13) + lngTmp, lngData
    
    If ByteToText(lngData) = Chr(0) & Chr(0) & Chr(0) & Chr(0) Then
        Set lngdat = ListV.ListItems.add(, "a" & lngLoop, "<unused>")
    Else
        Set lngdat = ListV.ListItems.add(, "a" & lngLoop, ByteToText(lngData))
    End If
Next lngLoop
' Close file and skip errors
Close #free
On Error Resume Next
ListV.ListItems(1).Selected = False
ListV.SelectedItem = Nothing
End Function

'================================================================================================================================================================================================
'=LNG=SAVE=AS=LNG================================================================================================================================================================================
'================================================================================================================================================================================================

Public Function LNG_saveasLNG(fpath, ListV As ListView)
' Dims
Dim lngOff As Long
Dim lngDataOff As Long
Dim sLoop As Long
Dim numString As Long
Dim nLng As String
Dim iOff As Long
Dim iLen As Long
Dim tmpStr As String
Dim trueLen As Long
Dim free: free = FreeFile
lngOff = 1
numString = ListV.ListItems.Count
' Open file and write it
Open fpath For Binary As #free
nLng = "LANG" & Space(4) & Space(4)
Put #free, lngOff, nLng
lngOff = 5
NumberToByte 1
Put #free, lngOff, outLong
lngOff = 9
NumberToByte numString
Put #free, lngOff, outLong

iOff = 0
lngDataOff = (numString * 4) + 13
lngOff = 13
' Make offsets and strings
For sLoop = 1 To numString
' Header
  NumberToByte iOff
  Put #free, lngOff, outLong
  lngOff = lngOff + 4
   ' Get trueLen
  If ListV.ListItems(sLoop).Text = "<unused>" Then
   trueLen = 4
  Else
   iLen = 1
   ' Get true size for current string
  Do Until iLen * 4 > Len(ListV.ListItems(sLoop).Text)
   iLen = iLen + 1
  Loop
   trueLen = iLen * 4
  End If
   ' Add up offset to print iOff
  iOff = iOff + trueLen
   ' Print strings
   If ListV.ListItems(sLoop).Text = "<unused>" Then
    tmpStr = Chr(0) & Chr(0) & Chr(0) & Chr(0)
   Else
    tmpStr = ListV.ListItems(sLoop).Text & String(trueLen - Len(ListV.ListItems(sLoop).Text), 0)
   End If
   Put #free, lngDataOff, tmpStr
    ' Add up offset to print tmpstr
   lngDataOff = lngDataOff + trueLen
Next sLoop
' Close file
Close #free
End Function

'================================================================================================================================================================================================
'=LNG=SAVE=AS=TXT================================================================================================================================================================================
'================================================================================================================================================================================================

Public Function LNG_saveasTXT(DestFile, ListV As ListView, ProgressBar As ProgressBar)
' Dims
Dim counter As String: counter = ListV.ListItems.Count
Dim lbl As String
Dim X As Long
Dim Y As Long
' Fill textbox with all strings
ProgressBar.Max = counter
lbl = ListV.ListItems(1)
For X = 2 To counter
lbl = lbl & vbNewLine & ListV.ListItems(X).Text
DoEvents
ProgressBar.Value = X
Next X
' Save output file
Open DestFile For Binary As #1
Put #1, 1, lbl
Close #1
End Function

'================================================================================================================================================================================================
'=LNG=SAVE=AS=INI================================================================================================================================================================================
'================================================================================================================================================================================================

Public Function LNG_saveasINI(DestFile, SectionName, ListV As ListView, ProgressBar As ProgressBar)
' Dims
Dim counter As String: counter = ListV.ListItems.Count
Dim lbl As String
Dim X As Long
Dim Y As Long
' Fill textbox with all strings
ProgressBar.Max = counter
lbl = "[" & SectionName & "]" & vbNewLine
lbl = lbl & "TotalStrings=" & counter & vbNewLine & vbNewLine
lbl = lbl & "1=" & ListV.ListItems(1)
For X = 2 To counter
lbl = lbl & vbNewLine & X & "=" & ListV.ListItems(X).Text
DoEvents
ProgressBar.Value = X
Next X
' Save output file
Open DestFile For Binary As #1
Put #1, 1, lbl
Close #1
End Function

'================================================================================================================================================================================================
'=PRIVATE=FUNCTIONS==============================================================================================================================================================================
'================================================================================================================================================================================================

Private Function ByteToNumber(mybyte() As Byte)
Dim byloop As Long: Dim newstr As String
For byloop = 1 To UBound(mybyte)
newstr = newstr & Chr(mybyte(byloop))
Next byloop
Dim WorkLong As Long
CopyMemory WorkLong, ByVal newstr, 4
ByteToNumber = WorkLong
End Function
Private Function ByteToText(mybyte() As Byte)
Dim byloop As Long: Dim newstr
For byloop = 1 To UBound(mybyte)
newstr = newstr & ChrW(mybyte(byloop))
Next byloop
ByteToText = newstr
End Function
Private Function ConvertIntelToMotorola(IntelHex As String)
ConvertIntelToMotorola = Mid(IntelHex, 7, 2) & Mid(IntelHex, 5, 2) & Mid(IntelHex, 3, 2) & Mid(IntelHex, 1, 2)
End Function
Private Function NumberToByte(newnum)
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
outLong(2) = Val("&H" & Mid(myNum, 3, 2) & "&")
outLong(3) = Val("&H" & Mid(myNum, 5, 2) & "&")
outLong(4) = Val("&H" & Mid(myNum, 7, 2) & "&")
End Function



