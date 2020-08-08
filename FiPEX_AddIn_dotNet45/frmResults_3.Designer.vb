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
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DataGridView1.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(12, 118)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(1035, 332)
        Me.DataGridView1.TabIndex = 0
        '
        'lblBeginTime
        '
        Me.lblBeginTime.AutoSize = True
        Me.lblBeginTime.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBeginTime.Location = New System.Drawing.Point(195, 22)
        Me.lblBeginTime.Name = "lblBeginTime"
        Me.lblBeginTime.Size = New System.Drawing.Size(84, 15)
        Me.lblBeginTime.TabIndex = 2
        Me.lblBeginTime.Text = "Begin Time:"
        '
        'lblEndtime
        '
        Me.lblEndtime.AutoSize = True
        Me.lblEndtime.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEndtime.Location = New System.Drawing.Point(207, 49)
        Me.lblEndtime.Name = "lblEndtime"
        Me.lblEndtime.Size = New System.Drawing.Size(72, 15)
        Me.lblEndtime.TabIndex = 3
        Me.lblEndtime.Text = "End Time:"
        '
        'lblTotalTime
        '
        Me.lblTotalTime.AutoSize = True
        Me.lblTotalTime.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalTime.Location = New System.Drawing.Point(200, 76)
        Me.lblTotalTime.Name = "lblTotalTime"
        Me.lblTotalTime.Size = New System.Drawing.Size(79, 15)
        Me.lblTotalTime.TabIndex = 4
        Me.lblTotalTime.Text = "Total Time:"
        '
        'lblDirection
        '
        Me.lblDirection.AutoSize = True
        Me.lblDirection.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDirection.Location = New System.Drawing.Point(495, 49)
        Me.lblDirection.Name = "lblDirection"
        Me.lblDirection.Size = New System.Drawing.Size(125, 15)
        Me.lblDirection.TabIndex = 5
        Me.lblDirection.Text = "Analysis Direction:"
        '
        'lblOrder
        '
        Me.lblOrder.AutoSize = True
        Me.lblOrder.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOrder.Location = New System.Drawing.Point(501, 22)
        Me.lblOrder.Name = "lblOrder"
        Me.lblOrder.Size = New System.Drawing.Size(119, 15)
        Me.lblOrder.TabIndex = 6
        Me.lblOrder.Text = "Order of Analysis:"
        '
        'lblNumBarriers
        '
        Me.lblNumBarriers.AutoSize = True
        Me.lblNumBarriers.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNumBarriers.Location = New System.Drawing.Point(487, 76)
        Me.lblNumBarriers.Name = "lblNumBarriers"
        Me.lblNumBarriers.Size = New System.Drawing.Size(133, 15)
        Me.lblNumBarriers.TabIndex = 7
        Me.lblNumBarriers.Text = "Number of Barriers:"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.My.Resources.Resources.FiPEx_LOGOv3b_90x90
        Me.PictureBox1.Location = New System.Drawing.Point(25, 9)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(91, 91)
        Me.PictureBox1.TabIndex = 1
        Me.PictureBox1.TabStop = False
        '
        'cmdExportXLS
        '
        Me.cmdExportXLS.Location = New System.Drawing.Point(393, 462)
        Me.cmdExportXLS.Name = "cmdExportXLS"
        Me.cmdExportXLS.Size = New System.Drawing.Size(103, 23)
        Me.cmdExportXLS.TabIndex = 8
        Me.cmdExportXLS.Text = "Export to XLS"
        Me.cmdExportXLS.UseVisualStyleBackColor = True
        '
        'frmResults_3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1059, 497)
        Me.Controls.Add(Me.cmdExportXLS)
        Me.Controls.Add(Me.lblNumBarriers)
        Me.Controls.Add(Me.lblOrder)
        Me.Controls.Add(Me.lblDirection)
        Me.Controls.Add(Me.lblTotalTime)
        Me.Controls.Add(Me.lblEndtime)
        Me.Controls.Add(Me.lblBeginTime)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.DataGridView1)
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
