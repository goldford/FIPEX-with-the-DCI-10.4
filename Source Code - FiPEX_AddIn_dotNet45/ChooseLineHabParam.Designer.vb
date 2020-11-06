<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ChooseLineHabParam
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmdSaveLineHabSettings = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cboLengthField = New System.Windows.Forms.ComboBox()
        Me.cboLengthUnits = New System.Windows.Forms.ComboBox()
        Me.cmboHabQuanField = New System.Windows.Forms.ComboBox()
        Me.cboHabUnits = New System.Windows.Forms.ComboBox()
        Me.cboHabClassField = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(70, 128)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(162, 20)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Habitat Quantity Field"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(268, 128)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(165, 20)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Habitat Quantity Units"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(70, 197)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(142, 20)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Habitat Class Field"
        '
        'cmdSaveLineHabSettings
        '
        Me.cmdSaveLineHabSettings.Location = New System.Drawing.Point(119, 275)
        Me.cmdSaveLineHabSettings.Name = "cmdSaveLineHabSettings"
        Me.cmdSaveLineHabSettings.Size = New System.Drawing.Size(93, 35)
        Me.cmdSaveLineHabSettings.TabIndex = 3
        Me.cmdSaveLineHabSettings.Text = "Save"
        Me.cmdSaveLineHabSettings.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(272, 275)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(91, 35)
        Me.cmdCancel.TabIndex = 4
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(268, 44)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 20)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Length Units"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(70, 44)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(97, 20)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "Length Field"
        '
        'cboLengthField
        '
        Me.cboLengthField.FormattingEnabled = True
        Me.cboLengthField.Location = New System.Drawing.Point(74, 67)
        Me.cboLengthField.Name = "cboLengthField"
        Me.cboLengthField.Size = New System.Drawing.Size(164, 28)
        Me.cboLengthField.TabIndex = 7
        '
        'cboLengthUnits
        '
        Me.cboLengthUnits.FormattingEnabled = True
        Me.cboLengthUnits.Location = New System.Drawing.Point(272, 67)
        Me.cboLengthUnits.Name = "cboLengthUnits"
        Me.cboLengthUnits.Size = New System.Drawing.Size(161, 28)
        Me.cboLengthUnits.TabIndex = 8
        '
        'cmboHabQuanField
        '
        Me.cmboHabQuanField.FormattingEnabled = True
        Me.cmboHabQuanField.Location = New System.Drawing.Point(74, 151)
        Me.cmboHabQuanField.Name = "cmboHabQuanField"
        Me.cmboHabQuanField.Size = New System.Drawing.Size(164, 28)
        Me.cmboHabQuanField.TabIndex = 9
        '
        'cboHabUnits
        '
        Me.cboHabUnits.FormattingEnabled = True
        Me.cboHabUnits.Location = New System.Drawing.Point(272, 151)
        Me.cboHabUnits.Name = "cboHabUnits"
        Me.cboHabUnits.Size = New System.Drawing.Size(161, 28)
        Me.cboHabUnits.TabIndex = 10
        '
        'cboHabClassField
        '
        Me.cboHabClassField.FormattingEnabled = True
        Me.cboHabClassField.Location = New System.Drawing.Point(74, 220)
        Me.cboHabClassField.Name = "cboHabClassField"
        Me.cboHabClassField.Size = New System.Drawing.Size(164, 28)
        Me.cboHabClassField.TabIndex = 11
        '
        'ChooseLineHabParam
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(516, 334)
        Me.Controls.Add(Me.cboHabClassField)
        Me.Controls.Add(Me.cboHabUnits)
        Me.Controls.Add(Me.cmboHabQuanField)
        Me.Controls.Add(Me.cboLengthUnits)
        Me.Controls.Add(Me.cboLengthField)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdSaveLineHabSettings)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "ChooseLineHabParam"
        Me.Text = "Set Line Layer Options"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmdSaveLineHabSettings As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboLengthField As System.Windows.Forms.ComboBox
    Friend WithEvents cboLengthUnits As System.Windows.Forms.ComboBox
    Friend WithEvents cmboHabQuanField As System.Windows.Forms.ComboBox
    Friend WithEvents cboHabUnits As System.Windows.Forms.ComboBox
    Friend WithEvents cboHabClassField As System.Windows.Forms.ComboBox
End Class
