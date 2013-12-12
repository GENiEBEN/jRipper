Attribute VB_Name = "TXT"
Dim fso As New FileSystemObject
Function lineCount(myInFile As String) As Long
On Error Resume Next
    Dim lFileSize As Long, lChunk As Long
    Dim bFile() As Byte
    Dim lSize As Long
    Dim strText As String
    
    'the size of the chunk to read in. You c
    '     an experiment
    'with this to see what works fastest.
    lSize = CLng(1024) * 10
    
    'size the array to the chunk size
    ReDim bFile(lSize - 1) As Byte
    
    Open myInFile For Binary As #1
    'get the file size
    lFileSize = LOF(1)
    
    'set the chunk number to 1
    lChunk = 1


    Do While (lSize * lChunk) < lFileSize
        'get the data from the in file
        Get #1, , bFile
        strText = StrConv(bFile, vbUnicode)
        'get the line count for this chunk
        lineCount = lineCount + searchText(strText)
        'increment the chunk count
        lChunk = lChunk + 1
    Loop
    
    'redim the array to the remaining size
    ReDim bFile((lFileSize - (lSize * (lChunk - 1))) - 1) As Byte
    'get the remaining data
    Get #1, , bFile
    strText = StrConv(bFile, vbUnicode)
    'get line count for this chunk
    lineCount = lineCount + searchText(strText)
    
    'close the file
    Close #1
    
    lineCount = lineCount + 1
End Function


Private Function searchText(strText As String) As Long
    Static blPossible As Boolean
    Dim lp1 As Long
    
    'if we have a possible line count


    If blPossible = True Then
        'if the fist charcter is chr(10) then we
        '     have a new line


        If Left$(strText, 1) = Chr(10) Then
            searchText = searchText + 1
        End If
    End If
    
    blPossible = False
    
    'loop through counting vbCrLf's
    lp1 = 1


    Do
        lp1 = InStr(lp1, strText, vbCrLf)


        If lp1 <> 0 Then
            searchText = searchText + 1
            lp1 = lp1 + 2
        End If
    Loop Until lp1 = 0
    
    'if the last character is a chr(13) then
    '     we may have a
    'new line, so we mark it as possible


    If Right$(strText, 1) = Chr(13) Then
        blPossible = True
    End If
    
End Function

Function TXT_load(Filetxt, TextBox As TextBox)
        Open Filetxt For Input As #1
        TextBox.Text = Input$(LOF(1), #1)
        Close #1
End Function

Function TXT2_load(FileTXT_Special, RTFTextBox As RichTextBox)
RTFTextBox.LOADFILE FileTXT_Special
End Function

Function TXT_Save(Filetxt As String, TextBox As TextBox)
    Open Filetxt For Binary As #1
    Put #1, 1, TextBox.Text
    Close #1
End Function

Function WriteTextFile(fName As String, _
  sText As String) As Boolean

  Dim fso As New FileSystemObject
  Dim FSTR As Scripting.TextStream
  On Error Resume Next
  Set FSTR = fso.OpenTextFile(fName, ForWriting, _
    Not fso.FileExists(fName))
  FSTR.Write sText
  WriteTextFile = True
  FSTR.Close
  If Err.Number Then WriteTextFile = False
  On Error GoTo 0
  Set FSTR = Nothing
  Set fso = Nothing
End Function



