'Game Name: Septerra Core: Legacy of the Creator
'File Type: Archive (no compression)
'Extension: DB
'Support  : Read, Extract
'Author/yr: GENiEBEN (2013)
'==============================================================================================================

Namespace Extensions.SepterraCore.DB
    Module DB

        'Is this a Septerra Core file?
        Public Function CheckMagic(InFile As String) As Boolean
            On Error Resume Next
            Dim bytes() As Byte = IO.File.ReadAllBytes(InFile)
            Dim i As Integer = 0
            Dim foundheader As Boolean = False
            Do
                If bytes(i) = &H43 And bytes(i + 1) = &H48 And bytes(i + 2) = &H31 And bytes(i + 3) = &H34 Then
                    foundheader = True
                    i = i + 4
                ElseIf bytes(i) = &H54 And bytes(i + 1) = &H58 And bytes(i + 2) = &H30 And bytes(i + 3) = &H30 Then
                    foundheader = True
                    i = i + 4
                ElseIf bytes(i) = &H56 And bytes(i + 1) = &H53 And bytes(i + 2) = &H53 And bytes(i + 3) = &H46 Then
                    foundheader = True
                    i = i + 4
                ElseIf bytes(i) = &H41 And bytes(i + 1) = &H4D And bytes(i + 2) = &H30 And bytes(i + 3) = &H34 Then
                    foundheader = True
                    i = i + 4
                ElseIf bytes(i) = &H6D And bytes(i + 1) = &H6F And bytes(i + 2) = &H6F And bytes(i + 3) = &H76 Then
                    foundheader = True
                    i = i + 4
                ElseIf bytes(i) = &H4C And bytes(i + 1) = &H56 And bytes(i + 2) = &H32 And bytes(i + 3) = &H35 Then
                    foundheader = True
                    i = i + 4
                ElseIf bytes(i) = &H47 And bytes(i + 1) = &H56 And bytes(i + 2) = &H30 And bytes(i + 3) = &H30 Then
                    foundheader = True
                    i = i + 4
                ElseIf bytes(i) = &H49 And bytes(i + 1) = &H4C And bytes(i + 2) = &H30 And bytes(i + 3) = &H30 Then
                    foundheader = True
                    i = i + 4
                End If
                i = i + 1
            Loop While i < CInt(bytes.Length)

            If foundheader = True Then
                Return True
            Else
                Return False
            End If
        End Function

        'Scan the file for known headers and populate a list with results.
        Public Function OpenFile(InFile As String, DGV As DataGridView) As Boolean
            On Error Resume Next
            Dim bytes() As Byte = IO.File.ReadAllBytes(InFile)
            Dim i As Integer = 0
            Do
                If bytes(i) = &H43 And bytes(i + 1) = &H48 And bytes(i + 2) = &H31 And bytes(i + 3) = &H34 Then
                    DGV.Rows.Add(1, DGV.Rows.Count + 1, "CH14", i)
                    i = i + 4
                ElseIf bytes(i) = &H54 And bytes(i + 1) = &H58 And bytes(i + 2) = &H30 And bytes(i + 3) = &H30 Then
                    DGV.Rows.Add(1, DGV.Rows.Count + 1, "TX00", i)
                    i = i + 4
                ElseIf bytes(i) = &H56 And bytes(i + 1) = &H53 And bytes(i + 2) = &H53 And bytes(i + 3) = &H46 Then
                    DGV.Rows.Add(1, DGV.Rows.Count + 1, "VSSF", i)
                    i = i + 4
                ElseIf bytes(i) = &H41 And bytes(i + 1) = &H4D And bytes(i + 2) = &H30 And bytes(i + 3) = &H34 Then
                    DGV.Rows.Add(1, DGV.Rows.Count + 1, "AM04", i)
                    i = i + 4
                ElseIf bytes(i) = &H4C And bytes(i + 1) = &H56 And bytes(i + 2) = &H32 And bytes(i + 3) = &H35 Then
                    DGV.Rows.Add(1, DGV.Rows.Count + 1, "LV25", i)
                    i = i + 4
                ElseIf bytes(i) = &H47 And bytes(i + 1) = &H56 And bytes(i + 2) = &H30 And bytes(i + 3) = &H30 Then
                    DGV.Rows.Add(1, DGV.Rows.Count + 1, "GV00", i)
                    i = i + 4
                ElseIf bytes(i) = &H49 And bytes(i + 1) = &H4C And bytes(i + 2) = &H30 And bytes(i + 3) = &H30 Then
                    DGV.Rows.Add(1, DGV.Rows.Count + 1, "IL00", i)
                    i = i + 4
                ElseIf bytes(i) = &H6D And bytes(i + 1) = &H6F And bytes(i + 2) = &H6F And bytes(i + 3) = &H76 Then
                    DGV.Rows.Add(1, DGV.Rows.Count + 1, "MOOV", i - 4)
                    i = i + 4
                End If
                i = i + 1
            Loop While i < CInt(bytes.Length)
            DGV.Rows.RemoveAt(DGV.Rows.Count - 1) 'because of a flaw in the code there's always an extra item at the end of the list that is not real. Too lazy to fix the code so I rather write a hugeass comment.
            For i = 0 To DGV.Rows.Count
                DGV.Item(4, i - 1).Value = CInt(DGV.Item(3, i).Value) - CInt(DGV.Item(3, i - 1).Value)
            Next
            DGV.Item(4, DGV.Rows.Count - 1).Value = CInt(bytes.Length) - CInt(DGV.Item(3, DGV.Rows.Count - 1).Value)
            Return True
        End Function

        'Extract a single file from a DB archive
        'TODO: Not working
        Public Function ExtractFile(Index As Integer, InFile As String, OutFile As String, DGV As DataGridView) As Boolean

        End Function
    End Module
End Namespace