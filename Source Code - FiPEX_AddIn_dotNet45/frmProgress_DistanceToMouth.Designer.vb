<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProgress_DistanceToMouth
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
        Me.lblCurrentFeature = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblTotalFeatures = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblGetNetworkLines = New System.Windows.Forms.Label()
        Me.lblGeonetSettings = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(44, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Feature"
        '
        'lblCurrentFeature
        '
        Me.lblCurrentFeature.AutoSize = True
        Me.lblCurrentFeature.Location = New System.Drawing.Point(62, 46)
        Me.lblCurrentFeature.Name = "lblCurrentFeature"
        Me.lblCurrentFeature.Size = New System.Drawing.Size(14, 13)
        Me.lblCurrentFeature.TabIndex = 1
        Me.lblCurrentFeature.Text = "X"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(44, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Out of"
        '
        'lblTotalFeatures
        '
        Me.lblTotalFeatures.AutoSize = True
        Me.lblTotalFeatures.Location = New System.Drawing.Point(62, 104)
        Me.lblTotalFeatures.Name = "lblTotalFeatures"
        Me.lblTotalFeatures.Size = New System.Drawing.Size(14, 13)
        Me.lblTotalFeatures.TabIndex = 3
        Me.lblTotalFeatures.Text = "X"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(60, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(163, 14)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Distance-to-Mouth Progress"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblTotalFeatures)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.lblCurrentFeature)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(53, 198)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(183, 146)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        '
        'lblGetNetworkLines
        '
        Me.lblGetNetworkLines.AutoSize = True
        Me.lblGetNetworkLines.Location = New System.Drawing.Point(63, 83)
        Me.lblGetNetworkLines.Name = "lblGetNetworkLines"
        Me.lblGetNetworkLines.Size = New System.Drawing.Size(95, 13)
        Me.lblGetNetworkLines.TabIndex = 6
        Me.lblGetNetworkLines.Text = "Get Network Lines"
        '
        'lblGeonetSettings
        '
        Me.lblGeonetSettings.AutoSize = True
        Me.lblGeonetSettings.Location = New System.Drawing.Point(63, 110)
        Me.lblGeonetSettings.Name = "lblGeonetSettings"
        Me.lblGeonetSettings.Size = New System.Drawing.Size(204, 13)
        Me.lblGeonetSettings.TabIndex = 7
        Me.lblGeonetSettings.Text = "Save Current Geometric Network Settings"
        '
        'frmProgress_DistanceToMouth
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(363, 380)
        Me.Controls.Add(Me.lblGeonetSettings)
        Me.Controls.Add(Me.lblGetNetworkLines)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label3)
        Me.Name = "frmProgress_DistanceToMouth"
        Me.Text = "Distance to Mouth Calculation"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblCurrentFeature As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblTotalFeatures As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblGetNetworkLines As System.Windows.Forms.Label
    Friend WithEvents lblGeonetSettings As System.Windows.Forms.Label
End Class
