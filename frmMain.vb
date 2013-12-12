Imports System.Windows.Forms
Imports System.IO

Public Class frmMain

    Private Sub OpenFile(ByVal sender As Object, ByVal e As EventArgs) Handles OpenToolStripMenuItem.Click
        Dim OpenFileDialog As New OpenFileDialog
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFileDialog.Filter = _
            "All Files|*.*|" & _
            "Text Files|*.txt|" & _
            "Septerra Core|*.db;*.mft"


        If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = OpenFileDialog.FileName
            Select Case LCase(Path.GetExtension(FileName))



                Case ".db"
                    'Septerra Core
                    If Extensions.SepterraCore.DB.CheckMagic(FileName) = True Then
                        frmSepterraCoreDB.MdiParent = Me
                        frmSepterraCoreDB.Show()
                        If Extensions.SepterraCore.DB.OpenFile(FileName, frmSepterraCoreDB.DGV1) = True Then
                            frmSepterraCoreDB.statusbar_label.Text = "Found " & frmSepterraCoreDB.DGV1.Rows.Count & " files in " & FileName
                        End If
                    End If







            End Select
        End If
    End Sub


    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CascadeToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileVerticalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileHorizontalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ArrangeIconsToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CloseAllToolStripMenuItem.Click
        ' Close all child forms of the parent.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private m_ChildFormNumber As Integer

    Private Sub DBAnalyzerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DBAnalyzerToolStripMenuItem.Click
        frmSepterraCoreDB.MdiParent = Me
        frmSepterraCoreDB.Show()
    End Sub
End Class
