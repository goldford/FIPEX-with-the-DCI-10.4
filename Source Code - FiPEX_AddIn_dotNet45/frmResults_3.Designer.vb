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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Step1 = New System.Windows.Forms.Label()
        Me.Step1b = New System.Windows.Forms.Label()
        Me.Step1a = New System.Windows.Forms.Label()
        Me.Step2 = New System.Windows.Forms.Label()
        Me.Step2a = New System.Windows.Forms.Label()
        Me.Step2b = New System.Windows.Forms.Label()
        Me.Step2c = New System.Windows.Forms.Label()
        Me.Step2d = New System.Windows.Forms.Label()
        Me.Step2e = New System.Windows.Forms.Label()
        Me.Step3 = New System.Windows.Forms.Label()
        Me.Step3a = New System.Windows.Forms.Label()
        Me.Step3b = New System.Windows.Forms.Label()
        Me.Step4 = New System.Windows.Forms.Label()
        Me.Step4a = New System.Windows.Forms.Label()
        Me.Step5 = New System.Windows.Forms.Label()
        Me.Step6 = New System.Windows.Forms.Label()
        Me.cmdCancel = New System.Windows.Forms.Button()
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
        Me.DataGridView1.Location = New System.Drawing.Point(304, 182)
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
        Me.PictureBox1.Image = Global.FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.My.Resources.Resources.FIPEX_DCI_Logo_2020_90x90a
        Me.PictureBox1.Location = New System.Drawing.Point(38, 14)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(136, 140)
        Me.PictureBox1.TabIndex = 1
        Me.PictureBox1.TabStop = False
        '
        'cmdExportXLS
        '
        Me.cmdExportXLS.Location = New System.Drawing.Point(673, 716)
        Me.cmdExportXLS.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdExportXLS.Name = "cmdExportXLS"
        Me.cmdExportXLS.Size = New System.Drawing.Size(154, 35)
        Me.cmdExportXLS.TabIndex = 8
        Me.cmdExportXLS.Text = "Export to XLS"
        Me.cmdExportXLS.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(48, 172)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 20)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Progress:"
        '
        'Step1
        '
        Me.Step1.AutoSize = True
        Me.Step1.Location = New System.Drawing.Point(53, 207)
        Me.Step1.Name = "Step1"
        Me.Step1.Size = New System.Drawing.Size(100, 20)
        Me.Step1.TabIndex = 10
        Me.Step1.Text = "1) Setting up"
        '
        'Step1b
        '
        Me.Step1b.AutoSize = True
        Me.Step1b.Location = New System.Drawing.Point(81, 271)
        Me.Step1b.Name = "Step1b"
        Me.Step1b.Size = New System.Drawing.Size(117, 20)
        Me.Step1b.TabIndex = 11
        Me.Step1b.Text = "Saving network"
        '
        'Step1a
        '
        Me.Step1a.AutoSize = True
        Me.Step1a.Location = New System.Drawing.Point(80, 237)
        Me.Step1a.Name = "Step1a"
        Me.Step1a.Size = New System.Drawing.Size(169, 20)
        Me.Step1a.TabIndex = 12
        Me.Step1a.Text = "Getting FIPEX options"
        '
        'Step2
        '
        Me.Step2.AutoSize = True
        Me.Step2.Location = New System.Drawing.Point(48, 305)
        Me.Step2.Name = "Step2"
        Me.Step2.Size = New System.Drawing.Size(137, 20)
        Me.Step2.TabIndex = 13
        Me.Step2.Text = "2) Network Traces"
        '
        'Step2a
        '
        Me.Step2a.AutoSize = True
        Me.Step2a.Location = New System.Drawing.Point(81, 334)
        Me.Step2a.Name = "Step2a"
        Me.Step2a.Size = New System.Drawing.Size(136, 20)
        Me.Step2a.TabIndex = 14
        Me.Step2a.Text = "Current flag / sink:"
        '
        'Step2b
        '
        Me.Step2b.AutoSize = True
        Me.Step2b.Location = New System.Drawing.Point(81, 364)
        Me.Step2b.Name = "Step2b"
        Me.Step2b.Size = New System.Drawing.Size(115, 20)
        Me.Step2b.TabIndex = 15
        Me.Step2b.Text = "Current barrier:"
        '
        'Step2c
        '
        Me.Step2c.AutoSize = True
        Me.Step2c.Location = New System.Drawing.Point(81, 395)
        Me.Step2c.Name = "Step2c"
        Me.Step2c.Size = New System.Drawing.Size(107, 20)
        Me.Step2c.TabIndex = 16
        Me.Step2c.Text = "Current order:"
        '
        'Step2d
        '
        Me.Step2d.AutoSize = True
        Me.Step2d.Location = New System.Drawing.Point(80, 424)
        Me.Step2d.Name = "Step2d"
        Me.Step2d.Size = New System.Drawing.Size(87, 20)
        Me.Step2d.TabIndex = 17
        Me.Step2d.Text = "Trace type:"
        '
        'Step2e
        '
        Me.Step2e.AutoSize = True
        Me.Step2e.Location = New System.Drawing.Point(80, 455)
        Me.Step2e.Name = "Step2e"
        Me.Step2e.Size = New System.Drawing.Size(156, 20)
        Me.Step2e.TabIndex = 18
        Me.Step2e.Text = "Intersecting features"
        '
        'Step3
        '
        Me.Step3.AutoSize = True
        Me.Step3.Location = New System.Drawing.Point(53, 489)
        Me.Step3.Name = "Step3"
        Me.Step3.Size = New System.Drawing.Size(125, 20)
        Me.Step3.TabIndex = 19
        Me.Step3.Text = "3) DCI (optional)"
        '
        'Step3a
        '
        Me.Step3a.AutoEllipsis = True
        Me.Step3a.AutoSize = True
        Me.Step3a.Location = New System.Drawing.Point(80, 519)
        Me.Step3a.Name = "Step3a"
        Me.Step3a.Size = New System.Drawing.Size(113, 20)
        Me.Step3a.TabIndex = 20
        Me.Step3a.Text = "DCI_d / DCI_p"
        '
        'Step3b
        '
        Me.Step3b.AutoSize = True
        Me.Step3b.Location = New System.Drawing.Point(81, 550)
        Me.Step3b.Name = "Step3b"
        Me.Step3b.Size = New System.Drawing.Size(104, 20)
        Me.Step3b.TabIndex = 21
        Me.Step3b.Text = "DCI sectional"
        '
        'Step4
        '
        Me.Step4.AutoSize = True
        Me.Step4.Location = New System.Drawing.Point(50, 582)
        Me.Step4.Name = "Step4"
        Me.Step4.Size = New System.Drawing.Size(123, 20)
        Me.Step4.TabIndex = 22
        Me.Step4.Text = "4) Output tables"
        '
        'Step4a
        '
        Me.Step4a.AutoSize = True
        Me.Step4a.Location = New System.Drawing.Point(81, 612)
        Me.Step4a.Name = "Step4a"
        Me.Step4a.Size = New System.Drawing.Size(115, 20)
        Me.Step4a.TabIndex = 23
        Me.Step4a.Text = "Writing to table"
        '
        'Step5
        '
        Me.Step5.AutoSize = True
        Me.Step5.Location = New System.Drawing.Point(48, 642)
        Me.Step5.Name = "Step5"
        Me.Step5.Size = New System.Drawing.Size(130, 20)
        Me.Step5.TabIndex = 24
        Me.Step5.Text = "5) Reset network"
        '
        'Step6
        '
        Me.Step6.AutoSize = True
        Me.Step6.Location = New System.Drawing.Point(48, 673)
        Me.Step6.Name = "Step6"
        Me.Step6.Size = New System.Drawing.Size(164, 20)
        Me.Step6.TabIndex = 25
        Me.Step6.Text = "6) Print results to form"
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(88, 716)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(86, 35)
        Me.cmdCancel.TabIndex = 26
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'frmResults_3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1588, 765)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.Step6)
        Me.Controls.Add(Me.Step5)
        Me.Controls.Add(Me.Step4a)
        Me.Controls.Add(Me.Step4)
        Me.Controls.Add(Me.Step3b)
        Me.Controls.Add(Me.Step3a)
        Me.Controls.Add(Me.Step3)
        Me.Controls.Add(Me.Step2e)
        Me.Controls.Add(Me.Step2d)
        Me.Controls.Add(Me.Step2c)
        Me.Controls.Add(Me.Step2b)
        Me.Controls.Add(Me.Step2a)
        Me.Controls.Add(Me.Step2)
        Me.Controls.Add(Me.Step1a)
        Me.Controls.Add(Me.Step1b)
        Me.Controls.Add(Me.Step1)
        Me.Controls.Add(Me.Label1)
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Step1 As System.Windows.Forms.Label
    Friend WithEvents Step1b As System.Windows.Forms.Label
    Friend WithEvents Step1a As System.Windows.Forms.Label
    Friend WithEvents Step2 As System.Windows.Forms.Label
    Friend WithEvents Step2a As System.Windows.Forms.Label
    Friend WithEvents Step2b As System.Windows.Forms.Label
    Friend WithEvents Step2c As System.Windows.Forms.Label
    Friend WithEvents Step2d As System.Windows.Forms.Label
    Friend WithEvents Step2e As System.Windows.Forms.Label
    Friend WithEvents Step3 As System.Windows.Forms.Label
    Friend WithEvents Step3a As System.Windows.Forms.Label
    Friend WithEvents Step3b As System.Windows.Forms.Label
    Friend WithEvents Step4 As System.Windows.Forms.Label
    Friend WithEvents Step4a As System.Windows.Forms.Label
    Friend WithEvents Step5 As System.Windows.Forms.Label
    Friend WithEvents Step6 As System.Windows.Forms.Label
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
End Class
