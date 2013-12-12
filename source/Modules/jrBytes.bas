Attribute VB_Name = "jrBytes"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

Public Function CVL(ValToConvert As String) As Long
Dim WorkLong As Long
CopyMemory WorkLong, ByVal ValToConvert, 4
CVL = WorkLong
End Function

Public Function ByteToNumber(mybyte() As Byte)
Dim newstr As String
Dim byloop As Long
For byloop = 1 To UBound(mybyte)
newstr = newstr & Chr(mybyte(byloop))
Next byloop
ByteToNumber = CVL(newstr)
End Function

Public Function ByteToNumber_1byte(mybyte As Byte)
Dim newstr As String: newstr = Chr(mybyte)
ByteToNumber_1byte = CVL(newstr)
End Function

Function get_1Byte(FilePath As String, StartOffset)
Dim byteNOF As Byte
Open FilePath For Binary As #5
Get #5, StartOffset, byteNOF
Close #5
get_1Byte = ByteToNumber_1byte(byteNOF)
End Function

Function get_2Bytes(FilePath As String, StartOffset)
Dim byteNOF(1 To 2) As Byte
Open FilePath For Binary As #5
Get #5, StartOffset, byteNOF
Close #5
get_2Bytes = ByteToNumber(byteNOF)
End Function

Function get_4Bytes(FilePath As String, StartOffset)
Dim byteNOF(1 To 4) As Byte
Open FilePath For Binary As #5
Get #5, StartOffset, byteNOF
Close #5
get_4Bytes = ByteToNumber(byteNOF)
End Function

Function get_Header_Len2(FilePath As String)
Dim strHEADER As String * 2
Open FilePath For Binary As #1
Get #1, 1, strHEADER
Close #1
get_Header_Len2 = strHEADER
End Function

Function get_Header_Len3(FilePath As String)
Dim strHEADER As String * 3
Open FilePath For Binary As #1
Get #1, 1, strHEADER
Close #1
get_Header_Len3 = strHEADER
End Function

Function get_Header_Len4(FilePath As String)
Dim strHEADER As String * 4
Open FilePath For Binary As #1
Get #1, 1, strHEADER
Close #1
get_Header_Len4 = strHEADER
End Function

Function get_Header_Len5(FilePath As String)
Dim strHEADER As String * 5
Open FilePath For Binary As #1
Get #1, 1, strHEADER
Close #1
get_Header_Len5 = strHEADER
End Function

Function get_Header_Len6(FilePath As String)
Dim strHEADER As String * 6
Open FilePath For Binary As #1
Get #1, 1, strHEADER
Close #1
get_Header_Len6 = strHEADER
End Function

Function get_Header_Len7(FilePath As String)
Dim strHEADER As String * 7
Open FilePath For Binary As #1
Get #1, 1, strHEADER
Close #1
get_Header_Len7 = strHEADER
End Function

Function get_Header_Len600(FilePath As String)
Dim strHEADER As String * 600
Open FilePath For Binary As #1
Get #1, 1, strHEADER
Close #1
get_Header_Len600 = strHEADER
End Function

