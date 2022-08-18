<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmResults_3
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.lblBeginTime = New System.Windows.Forms.Label()
        Me.lblEndtime = New System.Windows.Forms.Label()
        Me.lblTotalTime = New System.Windows.Forms.Label()
        Me.lblDirection = New System.Windows.Forms.Label()
        Me.lblOrder = New System.Windows.Forms.Label()
        Me.lblNumBarriers = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdExportXLS = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToOrderColumns = True
        Me.DataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DataGridView1.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.RaisedVertical
        Me.DataGridView1.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Me.DataGridView1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.ColumnHeadersHeight = 46
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView1.Location = New System.Drawing.Point(38, 182)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DataGridView1.MinimumSize = New System.Drawing.Size(46, 0)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.DataGridView1.Size = New System.Drawing.Size(1231, 511)
        Me.DataGridView1.TabIndex = 0
        '
        'lblBeginTime
        '
        Me.lblBeginTime.AutoSize = True
        Me.lblBeginTime.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBeginTime.Location = New System.Drawing.Point(292, 34)
        Me.lblBeginTime.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblBeginTime.Name = "lblBeginTime"
        Me.lblBeginTime.Size = New System.Drawing.Size(117, 22)
        Me.lblBeginTime.TabIndex = 2
        Me.lblBeginTime.Text = "Begin Time:"
        '
        'lblEndtime
        '
        Me.lblEndtime.AutoSize = True
        Me.lblEndtime.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEndtime.Location = New System.Drawing.Point(310, 75)
        Me.lblEndtime.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblEndtime.Name = "lblEndtime"
        Me.lblEndtime.Size = New System.Drawing.Size(101, 22)
        Me.lblEndtime.TabIndex = 3
        Me.lblEndtime.Text = "End Time:"
        '
        'lblTotalTime
        '
        Me.lblTotalTime.AutoSize = True
        Me.lblTotalTime.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalTime.Location = New System.Drawing.Point(300, 117)
        Me.lblTotalTime.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblTotalTime.Name = "lblTotalTime"
        Me.lblTotalTime.Size = New System.Drawing.Size(112, 22)
        Me.lblTotalTime.TabIndex = 4
        Me.lblTotalTime.Text = "Total Time:"
        '
        'lblDirection
        '
        Me.lblDirection.AutoSize = True
        Me.lblDirection.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDirection.Location = New System.Drawing.Point(742, 75)
        Me.lblDirection.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblDirection.Name = "lblDirection"
        Me.lblDirection.Size = New System.Drawing.Size(177, 22)
        Me.lblDirection.TabIndex = 5
        Me.lblDirection.Text = "Analysis Direction:"
        '
        'lblOrder
        '
        Me.lblOrder.AutoSize = True
        Me.lblOrder.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOrder.Location = New System.Drawing.Point(752, 34)
        Me.lblOrder.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblOrder.Name = "lblOrder"
        Me.lblOrder.Size = New System.Drawing.Size(171, 22)
        Me.lblOrder.TabIndex = 6
        Me.lblOrder.Text = "Order of Analysis:"
        '
        'lblNumBarriers
        '
        Me.lblNumBarriers.AutoSize = True
        Me.lblNumBarriers.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNumBarriers.Location = New System.Drawing.Point(659, 117)
        Me.lblNumBarriers.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblNumBarriers.Name = "lblNumBarriers"
        Me.lblNumBarriers.Size = New System.Drawing.Size(260, 22)
        Me.lblNumBarriers.TabIndex = 7
        Me.lblNumBarriers.Text = "Number of Barriers / Nodes:"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.FIPEX_DCI_Logo_2020_90x90a
        Me.PictureBox1.Location = New System.Drawing.Point(38, 14)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(136, 140)
        Me.PictureBox1.TabIndex = 1
        Me.PictureBox1.TabStop = False
        '
        'cmdExportXLS
        '
        Me.cmdExportXLS.Location = New System.Drawing.Point(568, 716)
        Me.cmdExportXLS.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdExportXLS.Name = "cmdExportXLS"
        Me.cmdExportXLS.Size = New System.Drawing.Size(154, 35)
        Me.cmdExportXLS.TabIndex = 8
        Me.cmdExportXLS.Text = "Export to XLS"
        Me.cmdExportXLS.UseVisualStyleBackColor = True
        '
        'frmResults_3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1295, 765)
        Me.Controls.Add(Me.cmdExportXLS)
        Me.Controls.Add(Me.lblNumBarriers)
        Me.Controls.Add(Me.lblOrder)
        Me.Controls.Add(Me.lblDirection)
        Me.Controls.Add(Me.lblTotalTime)
        Me.Controls.Add(Me.lblEndtime)
        Me.Controls.Add(Me.lblBeginTime)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "frmResults_3"
        Me.Text = "FiPEx Results Summary"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents lblBeginTime As System.Windows.Forms.Label
    Friend WithEvents lblEndtime As System.Windows.Forms.Label
    Friend WithEvents lblTotalTime As System.Windows.Forms.Label
    Friend WithEvents lblDirection As System.Windows.Forms.Label
    Friend WithEvents lblOrder As System.Windows.Forms.Label
    Friend WithEvents lblNumBarriers As System.Windows.Forms.Label
    Friend WithEvents cmdExportXLS As System.Windows.Forms.Button
End Class
