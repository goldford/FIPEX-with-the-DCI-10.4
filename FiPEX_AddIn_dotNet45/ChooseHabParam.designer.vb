<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ChooseHabParam
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
        Me.cmboHabClsField.Location = New System.Drawing.Point(36, 74)
        Me.cmboHabClsField.Name = "cmboHabClsField"
        Me.cmboHabClsField.Size = New System.Drawing.Size(121, 21)
        Me.cmboHabClsField.TabIndex = 0
        '
        'cmboHabQuanField
        '
        Me.cmboHabQuanField.FormattingEnabled = True
        Me.cmboHabQuanField.Location = New System.Drawing.Point(36, 25)
        Me.cmboHabQuanField.Name = "cmboHabQuanField"
        Me.cmboHabQuanField.Size = New System.Drawing.Size(121, 21)
        Me.cmboHabQuanField.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(18, 58)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(230, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Habitat Class Field  (select <none> if undesired)"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(18, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(155, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Habitat Quantity Field (required)"
        '
        'cmdSaveHabSettings
        '
        Me.cmdSaveHabSettings.Location = New System.Drawing.Point(56, 110)
        Me.cmdSaveHabSettings.Name = "cmdSaveHabSettings"
        Me.cmdSaveHabSettings.Size = New System.Drawing.Size(66, 23)
        Me.cmdSaveHabSettings.TabIndex = 4
        Me.cmdSaveHabSettings.Text = "Save"
        Me.cmdSaveHabSettings.UseVisualStyleBackColor = True
        '
        'cboUnits
        '
        Me.cboUnits.FormattingEnabled = True
        Me.cboUnits.Location = New System.Drawing.Point(194, 25)
        Me.cboUnits.Name = "cboUnits"
        Me.cboUnits.Size = New System.Drawing.Size(90, 21)
        Me.cboUnits.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(160, 33)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Units:"
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(183, 110)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 7
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'ChooseHabParam
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(306, 148)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cboUnits)
        Me.Controls.Add(Me.cmdSaveHabSettings)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmboHabQuanField)
        Me.Controls.Add(Me.cmboHabClsField)
        Me.Name = "ChooseHabParam"
        Me.Text = "Habitat Layer Options"
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
