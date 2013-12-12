'Game Name: Septerra Core: Legacy of the Creator
'File Type: Encrypted Text File
'Extension: TX0
'Support  : Decrypt (partial)
'Author/yr: GENiEBEN (2013)
'==============================================================================================================
Namespace Extensions.SepterraCore.TX0
    Module TX0
        Public Function Decrypt(InFile As String, OutFile As String) As Boolean
            'To decrypt TX0 files you must first XOR 96 and then increment uppercase letters +2 (i.e A = C, B = D)
            Dim bytes() As Byte = IO.File.ReadAllBytes(InFile)
            Dim m_XOR_MASK As Byte = 96
            Dim buf As Byte = bytes(4) Xor m_XOR_MASK
            Dim i As Integer = 0
            Do
                bytes(i) = bytes(i) Xor m_XOR_MASK
                i = i + 1
            Loop While i < CInt(bytes.Length)
            IO.File.WriteAllBytes(OutFile, bytes)
            Return True
        End Function
    End Module
End Namespace