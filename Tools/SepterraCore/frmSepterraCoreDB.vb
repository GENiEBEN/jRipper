Public Class frmSepterraCoreDB

    Private Sub cbSelectAll_CheckStateChanged(sender As Object, e As EventArgs) Handles cbSelectAll.CheckStateChanged
        For i As Integer = 0 To DGV1.Rows.Count - 1
            DGV1.Item(0, i).Value = cbSelectAll.Checked
        Next
    End Sub

    Private Sub tsbTX00Decrypt_Click(sender As Object, e As EventArgs) Handles tsbTX00Decrypt.Click

    End Sub

    Private Sub tsbExtract_Click(sender As Object, e As EventArgs) Handles tsbExtract.Click
        Extensions.SepterraCore.DB.ExtractFile(1, "C:\Program Files (x86)\GOG.com\Septerra Core\text.db", "C:\out.txt", DGV1)
    End Sub
End Class