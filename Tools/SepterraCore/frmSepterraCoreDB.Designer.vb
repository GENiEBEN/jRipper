<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSepterraCoreDB
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSepterraCoreDB))
        Me.DGV1 = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Column14 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.statusbar_label = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.tsbFile = New System.Windows.Forms.ToolStripDropDownButton()
        Me.tsbNew = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsbOpen = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsbExtract = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbTX00 = New System.Windows.Forms.ToolStripDropDownButton()
        Me.tsbTX00Decrypt = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsbTX00OpenView = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsbMOOV = New System.Windows.Forms.ToolStripDropDownButton()
        Me.tsbMOOVPlay = New System.Windows.Forms.ToolStripMenuItem()
        Me.cbSelectAll = New System.Windows.Forms.CheckBox()
        CType(Me.DGV1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DGV1
        '
        Me.DGV1.AllowUserToAddRows = False
        Me.DGV1.AllowUserToDeleteRows = False
        Me.DGV1.AllowUserToResizeRows = False
        Me.DGV1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DGV1.BackgroundColor = System.Drawing.SystemColors.Control
        Me.DGV1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.DGV1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column14, Me.Column2, Me.DataGridViewTextBoxColumn1, Me.Column3})
        Me.DGV1.Location = New System.Drawing.Point(0, 39)
        Me.DGV1.Margin = New System.Windows.Forms.Padding(0)
        Me.DGV1.MultiSelect = False
        Me.DGV1.Name = "DGV1"
        Me.DGV1.RowHeadersVisible = False
        Me.DGV1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DGV1.Size = New System.Drawing.Size(492, 555)
        Me.DGV1.TabIndex = 36
        '
        'Column1
        '
        Me.Column1.HeaderText = ""
        Me.Column1.Name = "Column1"
        Me.Column1.Width = 30
        '
        'Column14
        '
        Me.Column14.HeaderText = "#"
        Me.Column14.Name = "Column14"
        Me.Column14.ReadOnly = True
        Me.Column14.Width = 40
        '
        'Column2
        '
        Me.Column2.HeaderText = "Type"
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridViewTextBoxColumn1.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridViewTextBoxColumn1.HeaderText = "Offset"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        '
        'Column3
        '
        Me.Column3.HeaderText = "Size"
        Me.Column3.Name = "Column3"
        Me.Column3.ReadOnly = True
        Me.Column3.Width = 150
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.statusbar_label})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 594)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(492, 22)
        Me.StatusStrip1.TabIndex = 38
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'statusbar_label
        '
        Me.statusbar_label.Name = "statusbar_label"
        Me.statusbar_label.Size = New System.Drawing.Size(12, 17)
        Me.statusbar_label.Text = "-"
        '
        'ToolStrip1
        '
        Me.ToolStrip1.AllowMerge = False
        Me.ToolStrip1.GripMargin = New System.Windows.Forms.Padding(0)
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbFile, Me.tsbExtract, Me.ToolStripSeparator1, Me.tsbTX00, Me.tsbMOOV})
        Me.ToolStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Padding = New System.Windows.Forms.Padding(0)
        Me.ToolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.ToolStrip1.Size = New System.Drawing.Size(492, 39)
        Me.ToolStrip1.Stretch = True
        Me.ToolStrip1.TabIndex = 39
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'tsbFile
        '
        Me.tsbFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbNew, Me.tsbOpen})
        Me.tsbFile.Image = CType(resources.GetObject("tsbFile.Image"), System.Drawing.Image)
        Me.tsbFile.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbFile.Name = "tsbFile"
        Me.tsbFile.Size = New System.Drawing.Size(70, 36)
        Me.tsbFile.Text = "File"
        '
        'tsbNew
        '
        Me.tsbNew.Name = "tsbNew"
        Me.tsbNew.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
        Me.tsbNew.Size = New System.Drawing.Size(146, 22)
        Me.tsbNew.Text = "New"
        '
        'tsbOpen
        '
        Me.tsbOpen.Name = "tsbOpen"
        Me.tsbOpen.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
        Me.tsbOpen.Size = New System.Drawing.Size(146, 22)
        Me.tsbOpen.Text = "Open"
        '
        'tsbExtract
        '
        Me.tsbExtract.Image = CType(resources.GetObject("tsbExtract.Image"), System.Drawing.Image)
        Me.tsbExtract.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbExtract.Name = "tsbExtract"
        Me.tsbExtract.Size = New System.Drawing.Size(125, 36)
        Me.tsbExtract.Text = "Extract Selected"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 39)
        '
        'tsbTX00
        '
        Me.tsbTX00.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbTX00Decrypt, Me.tsbTX00OpenView})
        Me.tsbTX00.Image = CType(resources.GetObject("tsbTX00.Image"), System.Drawing.Image)
        Me.tsbTX00.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbTX00.Name = "tsbTX00"
        Me.tsbTX00.Size = New System.Drawing.Size(78, 36)
        Me.tsbTX00.Text = "TX00"
        '
        'tsbTX00Decrypt
        '
        Me.tsbTX00Decrypt.Name = "tsbTX00Decrypt"
        Me.tsbTX00Decrypt.Size = New System.Drawing.Size(175, 22)
        Me.tsbTX00Decrypt.Text = "Decrypt..."
        '
        'tsbTX00OpenView
        '
        Me.tsbTX00OpenView.Name = "tsbTX00OpenView"
        Me.tsbTX00OpenView.Size = New System.Drawing.Size(175, 22)
        Me.tsbTX00OpenView.Text = "Open in text viewer"
        '
        'tsbMOOV
        '
        Me.tsbMOOV.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbMOOVPlay})
        Me.tsbMOOV.Image = CType(resources.GetObject("tsbMOOV.Image"), System.Drawing.Image)
        Me.tsbMOOV.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbMOOV.Name = "tsbMOOV"
        Me.tsbMOOV.Size = New System.Drawing.Size(88, 36)
        Me.tsbMOOV.Text = "MOOV"
        '
        'tsbMOOVPlay
        '
        Me.tsbMOOVPlay.Name = "tsbMOOVPlay"
        Me.tsbMOOVPlay.Size = New System.Drawing.Size(128, 22)
        Me.tsbMOOVPlay.Text = "Play video"
        '
        'cbSelectAll
        '
        Me.cbSelectAll.AutoSize = True
        Me.cbSelectAll.Checked = True
        Me.cbSelectAll.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbSelectAll.Location = New System.Drawing.Point(9, 42)
        Me.cbSelectAll.Name = "cbSelectAll"
        Me.cbSelectAll.Size = New System.Drawing.Size(15, 14)
        Me.cbSelectAll.TabIndex = 40
        Me.cbSelectAll.UseVisualStyleBackColor = True
        '
        'frmSepterraCoreDB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(492, 616)
        Me.Controls.Add(Me.cbSelectAll)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.DGV1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSepterraCoreDB"
        Me.Text = "Septerra Core | DB Analyzer v0.1"
        CType(Me.DGV1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DGV1 As System.Windows.Forms.DataGridView
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents statusbar_label As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsbTX00 As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents tsbTX00Decrypt As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsbMOOV As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents tsbMOOVPlay As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsbTX00OpenView As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cbSelectAll As System.Windows.Forms.CheckBox
    Friend WithEvents tsbExtract As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbFile As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents tsbNew As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsbOpen As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Column14 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
