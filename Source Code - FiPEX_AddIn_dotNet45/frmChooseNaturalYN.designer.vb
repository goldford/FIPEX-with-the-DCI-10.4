<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChooseNaturalYN
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.cboBarrierNaturalYNFields = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cboBarrierNaturalYNFields
        '
        Me.cboBarrierNaturalYNFields.FormattingEnabled = True
        Me.cboBarrierNaturalYNFields.Location = New System.Drawing.Point(54, 35)
        Me.cboBarrierNaturalYNFields.Name = "cboBarrierNaturalYNFields"
        Me.cboBarrierNaturalYNFields.Size = New System.Drawing.Size(121, 21)
        Me.cboBarrierNaturalYNFields.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(38, 19)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(157, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Barrier Layer Fields (select one):"
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(25, 84)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 8
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(130, 84)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 9
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(38, 59)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(167, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "(Must be of length = 1; type 'Text')"
        '
        'frmChooseNaturalYN
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(253, 119)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cboBarrierNaturalYNFields)
        Me.Name = "frmChooseNaturalYN"
        Me.Text = "Choose Natural T/F field"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboBarrierNaturalYNFields As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
