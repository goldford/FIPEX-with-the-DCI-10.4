<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCalculateOverlap
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCalculateOverlap))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.cmdCalculate = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdBrowseTab1 = New System.Windows.Forms.Button()
        Me.cmdBrowseTab2 = New System.Windows.Forms.Button()
        Me.txtTable1 = New System.Windows.Forms.TextBox()
        Me.txtTable2 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmdGDBOut = New System.Windows.Forms.Button()
        Me.txtGDBOut = New System.Windows.Forms.TextBox()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.txtTableName = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(158, 75)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(291, 91)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = resources.GetString("Label1.Text")
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 406)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(452, 23)
        Me.ProgressBar1.TabIndex = 1
        '
        'cmdCalculate
        '
        Me.cmdCalculate.Location = New System.Drawing.Point(67, 370)
        Me.cmdCalculate.Name = "cmdCalculate"
        Me.cmdCalculate.Size = New System.Drawing.Size(75, 23)
        Me.cmdCalculate.TabIndex = 2
        Me.cmdCalculate.Text = "Calculate"
        Me.cmdCalculate.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(329, 370)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "Close"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdBrowseTab1
        '
        Me.cmdBrowseTab1.Location = New System.Drawing.Point(12, 197)
        Me.cmdBrowseTab1.Name = "cmdBrowseTab1"
        Me.cmdBrowseTab1.Size = New System.Drawing.Size(130, 23)
        Me.cmdBrowseTab1.TabIndex = 4
        Me.cmdBrowseTab1.Text = "Browse For Table 1:"
        Me.cmdBrowseTab1.UseVisualStyleBackColor = True
        '
        'cmdBrowseTab2
        '
        Me.cmdBrowseTab2.Location = New System.Drawing.Point(12, 234)
        Me.cmdBrowseTab2.Name = "cmdBrowseTab2"
        Me.cmdBrowseTab2.Size = New System.Drawing.Size(130, 23)
        Me.cmdBrowseTab2.TabIndex = 5
        Me.cmdBrowseTab2.Text = "Browse For Table 2:"
        Me.cmdBrowseTab2.UseVisualStyleBackColor = True
        '
        'txtTable1
        '
        Me.txtTable1.Location = New System.Drawing.Point(161, 199)
        Me.txtTable1.Name = "txtTable1"
        Me.txtTable1.ReadOnly = True
        Me.txtTable1.Size = New System.Drawing.Size(303, 20)
        Me.txtTable1.TabIndex = 6
        '
        'txtTable2
        '
        Me.txtTable2.Location = New System.Drawing.Point(161, 236)
        Me.txtTable2.Name = "txtTable2"
        Me.txtTable2.ReadOnly = True
        Me.txtTable2.Size = New System.Drawing.Size(303, 20)
        Me.txtTable2.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(173, 31)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(252, 18)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Calculate 'Decision Overlap' Table"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdGDBOut
        '
        Me.cmdGDBOut.Location = New System.Drawing.Point(12, 275)
        Me.cmdGDBOut.Name = "cmdGDBOut"
        Me.cmdGDBOut.Size = New System.Drawing.Size(130, 23)
        Me.cmdGDBOut.TabIndex = 11
        Me.cmdGDBOut.Text = "Output Geodatabase:"
        Me.cmdGDBOut.UseVisualStyleBackColor = True
        '
        'txtGDBOut
        '
        Me.txtGDBOut.Location = New System.Drawing.Point(161, 278)
        Me.txtGDBOut.Name = "txtGDBOut"
        Me.txtGDBOut.ReadOnly = True
        Me.txtGDBOut.Size = New System.Drawing.Size(303, 20)
        Me.txtGDBOut.TabIndex = 12
        '
        'lblProgress
        '
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProgress.Location = New System.Drawing.Point(226, 380)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(0, 13)
        Me.lblProgress.TabIndex = 13
        '
        'BackgroundWorker3
        '
        Me.BackgroundWorker3.WorkerReportsProgress = True
        Me.BackgroundWorker3.WorkerSupportsCancellation = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.My.Resources.Resources.FiPEx_LOGOv3b_90x90
        Me.PictureBox1.Location = New System.Drawing.Point(32, 31)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(90, 90)
        Me.PictureBox1.TabIndex = 14
        Me.PictureBox1.TabStop = False
        '
        'txtTableName
        '
        Me.txtTableName.Location = New System.Drawing.Point(161, 319)
        Me.txtTableName.Name = "txtTableName"
        Me.txtTableName.Size = New System.Drawing.Size(303, 20)
        Me.txtTableName.TabIndex = 15
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(39, 322)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(103, 13)
        Me.Label2.TabIndex = 16
        Me.Label2.Text = "Output Table Name:"
        '
        'frmCalculateOverlap
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(479, 450)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtTableName)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.txtGDBOut)
        Me.Controls.Add(Me.cmdGDBOut)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtTable2)
        Me.Controls.Add(Me.txtTable1)
        Me.Controls.Add(Me.cmdBrowseTab2)
        Me.Controls.Add(Me.cmdBrowseTab1)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdCalculate)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmCalculateOverlap"
        Me.Text = "FiPEX - Calculate Decision Overlap Table"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents cmdCalculate As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseTab1 As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseTab2 As System.Windows.Forms.Button
    Friend WithEvents txtTable1 As System.Windows.Forms.TextBox
    Friend WithEvents txtTable2 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdGDBOut As System.Windows.Forms.Button
    Friend WithEvents txtGDBOut As System.Windows.Forms.TextBox
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents txtTableName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
