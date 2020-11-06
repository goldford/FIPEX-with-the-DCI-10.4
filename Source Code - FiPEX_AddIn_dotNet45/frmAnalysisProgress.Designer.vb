<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAnalysisProgress
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
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(25, 113)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(417, 23)
        Me.ProgressBar1.TabIndex = 0
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(164, 154)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(117, 23)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "Cancel / Close"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'lblProgress
        '
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProgress.Location = New System.Drawing.Point(32, 18)
        Me.lblProgress.MinimumSize = New System.Drawing.Size(400, 80)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(400, 80)
        Me.lblProgress.TabIndex = 2
        Me.lblProgress.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'frmAnalysisProgress
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(458, 189)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Name = "frmAnalysisProgress"
        Me.Text = "FIPEX Analysis Results"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents lblProgress As System.Windows.Forms.Label
End Class
