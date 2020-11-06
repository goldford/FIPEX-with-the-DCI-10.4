<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChooseBarrID
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
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmboBarrierIDFields = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(12, 87)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 0
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(129, 87)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmboBarrierIDFields
        '
        Me.cmboBarrierIDFields.FormattingEnabled = True
        Me.cmboBarrierIDFields.Location = New System.Drawing.Point(48, 33)
        Me.cmboBarrierIDFields.Name = "cmboBarrierIDFields"
        Me.cmboBarrierIDFields.Size = New System.Drawing.Size(121, 21)
        Me.cmboBarrierIDFields.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(45, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Barrier ID Field:"
        '
        'frmChooseBarrID
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(218, 122)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmboBarrierIDFields)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdSave)
        Me.Name = "frmChooseBarrID"
        Me.Text = "Choose Barrier ID Field"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmboBarrierIDFields As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
