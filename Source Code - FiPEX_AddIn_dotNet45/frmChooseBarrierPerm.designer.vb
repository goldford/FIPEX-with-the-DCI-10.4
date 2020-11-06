<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChooseBarrierPerm
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
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboBarrierPermFields = New System.Windows.Forms.ComboBox()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(231, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Barrier Layer Fields for Permeability (select one):"
        '
        'cboBarrierPermFields
        '
        Me.cboBarrierPermFields.FormattingEnabled = True
        Me.cboBarrierPermFields.Location = New System.Drawing.Point(59, 38)
        Me.cboBarrierPermFields.Name = "cboBarrierPermFields"
        Me.cboBarrierPermFields.Size = New System.Drawing.Size(121, 21)
        Me.cboBarrierPermFields.TabIndex = 6
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(37, 90)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 7
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(131, 90)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 8
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(34, 62)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(181, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "(must be of format 0.00; type Double)"
        '
        'frmChooseBarrierPerm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(247, 130)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cboBarrierPermFields)
        Me.Controls.Add(Me.Label3)
        Me.Name = "frmChooseBarrierPerm"
        Me.Text = "Choose Barrier Permeability Field"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboBarrierPermFields As System.Windows.Forms.ComboBox
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
