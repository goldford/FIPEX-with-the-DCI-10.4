<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ChoosePolyHabParam
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
        Me.cmboHabClsField = New System.Windows.Forms.ComboBox()
        Me.cmboHabQuanField = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmdSaveHabSettings = New System.Windows.Forms.Button()
        Me.cboUnits = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cmboHabClsField
        '
        Me.cmboHabClsField.FormattingEnabled = True
        Me.cmboHabClsField.Location = New System.Drawing.Point(54, 114)
        Me.cmboHabClsField.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmboHabClsField.Name = "cmboHabClsField"
        Me.cmboHabClsField.Size = New System.Drawing.Size(180, 28)
        Me.cmboHabClsField.TabIndex = 0
        '
        'cmboHabQuanField
        '
        Me.cmboHabQuanField.FormattingEnabled = True
        Me.cmboHabQuanField.Location = New System.Drawing.Point(54, 38)
        Me.cmboHabQuanField.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmboHabQuanField.Name = "cmboHabQuanField"
        Me.cmboHabQuanField.Size = New System.Drawing.Size(180, 28)
        Me.cmboHabQuanField.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(27, 89)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(346, 20)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Habitat Class Field  (select <none> if undesired)"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(27, 14)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(234, 20)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Habitat Quantity Field (required)"
        '
        'cmdSaveHabSettings
        '
        Me.cmdSaveHabSettings.Location = New System.Drawing.Point(84, 169)
        Me.cmdSaveHabSettings.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdSaveHabSettings.Name = "cmdSaveHabSettings"
        Me.cmdSaveHabSettings.Size = New System.Drawing.Size(99, 35)
        Me.cmdSaveHabSettings.TabIndex = 4
        Me.cmdSaveHabSettings.Text = "Save"
        Me.cmdSaveHabSettings.UseVisualStyleBackColor = True
        '
        'cboUnits
        '
        Me.cboUnits.FormattingEnabled = True
        Me.cboUnits.Location = New System.Drawing.Point(291, 38)
        Me.cboUnits.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cboUnits.Name = "cboUnits"
        Me.cboUnits.Size = New System.Drawing.Size(133, 28)
        Me.cboUnits.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(336, 14)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(50, 20)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Units:"
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(274, 169)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(112, 35)
        Me.cmdCancel.TabIndex = 7
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'ChoosePolyHabParam
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(459, 228)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cboUnits)
        Me.Controls.Add(Me.cmdSaveHabSettings)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmboHabQuanField)
        Me.Controls.Add(Me.cmboHabClsField)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "ChoosePolyHabParam"
        Me.Text = "Set Habitat Layer Options"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmboHabClsField As System.Windows.Forms.ComboBox
    Friend WithEvents cmboHabQuanField As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmdSaveHabSettings As System.Windows.Forms.Button
    Friend WithEvents cboUnits As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
End Class
