<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDistanceToMouth
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDistanceToMouth))
        Me.chkUseFiPExQuantity = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.chklistSimpleFCs = New System.Windows.Forms.CheckedListBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmdDoIt = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.chkMouth = New System.Windows.Forms.CheckBox()
        Me.chkSource = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkUseFiPExQuantity
        '
        Me.chkUseFiPExQuantity.AutoSize = True
        Me.chkUseFiPExQuantity.Location = New System.Drawing.Point(25, 48)
        Me.chkUseFiPExQuantity.Name = "chkUseFiPExQuantity"
        Me.chkUseFiPExQuantity.Size = New System.Drawing.Size(44, 17)
        Me.chkUseFiPExQuantity.TabIndex = 0
        Me.chkUseFiPExQuantity.Text = "Yes"
        Me.chkUseFiPExQuantity.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(19, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(280, 26)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Use FiPEx 'Options' quantity setting for distance measure?" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(otherwise tool will " & _
            "use shape_length field)"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(22, 11)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(300, 39)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Select feature classes to calculate network distance-to-mouth:" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(only SIMPLE netw" & _
            "ork edges and junctions with currently " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "selected features will appear here)"
        '
        'chklistSimpleFCs
        '
        Me.chklistSimpleFCs.FormattingEnabled = True
        Me.chklistSimpleFCs.Location = New System.Drawing.Point(22, 65)
        Me.chklistSimpleFCs.Name = "chklistSimpleFCs"
        Me.chklistSimpleFCs.Size = New System.Drawing.Size(300, 109)
        Me.chklistSimpleFCs.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(93, 75)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(252, 65)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = resources.GetString("Label3.Text")
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdDoIt
        '
        Me.cmdDoIt.Location = New System.Drawing.Point(45, 528)
        Me.cmdDoIt.Name = "cmdDoIt"
        Me.cmdDoIt.Size = New System.Drawing.Size(75, 23)
        Me.cmdDoIt.TabIndex = 5
        Me.cmdDoIt.Text = "Calculate"
        Me.cmdDoIt.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(292, 528)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 6
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chklistSimpleFCs)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(45, 211)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(343, 187)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.chkUseFiPExQuantity)
        Me.GroupBox2.Location = New System.Drawing.Point(45, 413)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(342, 83)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(97, 27)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(237, 36)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Calculate 'Distance to Mouth' Or" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "'Distance to Source'"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'chkMouth
        '
        Me.chkMouth.AutoSize = True
        Me.chkMouth.Checked = True
        Me.chkMouth.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMouth.Location = New System.Drawing.Point(70, 170)
        Me.chkMouth.Name = "chkMouth"
        Me.chkMouth.Size = New System.Drawing.Size(113, 17)
        Me.chkMouth.TabIndex = 10
        Me.chkMouth.Text = "Distance to Mouth"
        Me.chkMouth.UseVisualStyleBackColor = True
        '
        'chkSource
        '
        Me.chkSource.AutoSize = True
        Me.chkSource.Location = New System.Drawing.Point(227, 170)
        Me.chkSource.Name = "chkSource"
        Me.chkSource.Size = New System.Drawing.Size(117, 17)
        Me.chkSource.TabIndex = 11
        Me.chkSource.Text = "Distance to Source"
        Me.chkSource.UseVisualStyleBackColor = True
        '
        'frmDistanceToMouth
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(429, 595)
        Me.Controls.Add(Me.chkSource)
        Me.Controls.Add(Me.chkMouth)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdDoIt)
        Me.Controls.Add(Me.Label3)
        Me.Name = "frmDistanceToMouth"
        Me.Text = "FiPEx: Distance To Mouth"
        Me.TopMost = True
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkUseFiPExQuantity As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chklistSimpleFCs As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmdDoIt As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents chkMouth As System.Windows.Forms.CheckBox
    Friend WithEvents chkSource As System.Windows.Forms.CheckBox
End Class
