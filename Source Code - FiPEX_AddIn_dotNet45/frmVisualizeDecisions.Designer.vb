<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmVisualizeDecisions
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
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdHighlight = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.txtTable = New System.Windows.Forms.TextBox()
        Me.lstBudgets = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblHighlighted = New System.Windows.Forms.Label()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.FiPEx_LOGOv3b_90x90
        Me.PictureBox1.Location = New System.Drawing.Point(12, 12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(90, 90)
        Me.PictureBox1.TabIndex = 15
        Me.PictureBox1.TabStop = False
        '
        'lblProgress
        '
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Location = New System.Drawing.Point(54, 145)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(0, 13)
        Me.lblProgress.TabIndex = 17
        Me.lblProgress.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(299, 140)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 18
        Me.cmdCancel.Text = "Close"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdHighlight
        '
        Me.cmdHighlight.Location = New System.Drawing.Point(165, 140)
        Me.cmdHighlight.Name = "cmdHighlight"
        Me.cmdHighlight.Size = New System.Drawing.Size(75, 23)
        Me.cmdHighlight.TabIndex = 19
        Me.cmdHighlight.Text = "Highlight"
        Me.cmdHighlight.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(165, 17)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(117, 23)
        Me.Button1.TabIndex = 20
        Me.Button1.Text = "Browse for Table"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'txtTable
        '
        Me.txtTable.Location = New System.Drawing.Point(165, 46)
        Me.txtTable.Name = "txtTable"
        Me.txtTable.Size = New System.Drawing.Size(236, 20)
        Me.txtTable.TabIndex = 21
        '
        'lstBudgets
        '
        Me.lstBudgets.FormattingEnabled = True
        Me.lstBudgets.Location = New System.Drawing.Point(254, 87)
        Me.lstBudgets.Name = "lstBudgets"
        Me.lstBudgets.Size = New System.Drawing.Size(120, 43)
        Me.lstBudgets.TabIndex = 22
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(136, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(23, 15)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "1. "
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(136, 87)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 15)
        Me.Label2.TabIndex = 24
        Me.Label2.Text = "2. Select Budget"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(136, 143)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(23, 15)
        Me.Label3.TabIndex = 25
        Me.Label3.Text = "3. "
        '
        'lblHighlighted
        '
        Me.lblHighlighted.AutoSize = True
        Me.lblHighlighted.Location = New System.Drawing.Point(12, 150)
        Me.lblHighlighted.Name = "lblHighlighted"
        Me.lblHighlighted.Size = New System.Drawing.Size(63, 13)
        Me.lblHighlighted.TabIndex = 26
        Me.lblHighlighted.Text = "Highlighted:"
        '
        'frmVisualizeDecisions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(458, 178)
        Me.Controls.Add(Me.lblHighlighted)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lstBudgets)
        Me.Controls.Add(Me.txtTable)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.cmdHighlight)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.PictureBox1)
        Me.Name = "frmVisualizeDecisions"
        Me.Text = "Highlight Decision Sets"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdHighlight As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents txtTable As System.Windows.Forms.TextBox
    Friend WithEvents lstBudgets As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblHighlighted As System.Windows.Forms.Label
End Class
