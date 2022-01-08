<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLabelHypergraph
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLabelHypergraph))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdRunLabel = New System.Windows.Forms.Button()
        Me.lblCheck1 = New System.Windows.Forms.Label()
        Me.lblCheck2 = New System.Windows.Forms.Label()
        Me.lblCheck3 = New System.Windows.Forms.Label()
        Me.lblCheck4 = New System.Windows.Forms.Label()
        Me.lblPrep1 = New System.Windows.Forms.Label()
        Me.lblPrep2 = New System.Windows.Forms.Label()
        Me.lblPrep3 = New System.Windows.Forms.Label()
        Me.lblAnalysis1 = New System.Windows.Forms.Label()
        Me.lblAnalysis2 = New System.Windows.Forms.Label()
        Me.lblAnalysis3 = New System.Windows.Forms.Label()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(446, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(448, 47)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Hypergraph Labelling:" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Label River ""Segments"" and ""Sub-segments""" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(437, 71)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(457, 595)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = resources.GetString("Label2.Text")
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Image = Global.FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.My.Resources.Resources.FIPEX_DCI_Logo_2020_90x90a
        Me.PictureBox1.InitialImage = Global.FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.My.Resources.Resources.FIPEX_DCI_Logo_2020_90x90a
        Me.PictureBox1.Location = New System.Drawing.Point(113, 9)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(170, 194)
        Me.PictureBox1.TabIndex = 2
        Me.PictureBox1.TabStop = False
        '
        'cmdRunLabel
        '
        Me.cmdRunLabel.Location = New System.Drawing.Point(113, 198)
        Me.cmdRunLabel.Name = "cmdRunLabel"
        Me.cmdRunLabel.Size = New System.Drawing.Size(151, 56)
        Me.cmdRunLabel.TabIndex = 3
        Me.cmdRunLabel.Text = "Run Hypergraph Labeling"
        Me.cmdRunLabel.UseVisualStyleBackColor = True
        '
        'lblCheck1
        '
        Me.lblCheck1.AutoSize = True
        Me.lblCheck1.Location = New System.Drawing.Point(22, 271)
        Me.lblCheck1.Name = "lblCheck1"
        Me.lblCheck1.Size = New System.Drawing.Size(67, 20)
        Me.lblCheck1.TabIndex = 4
        Me.lblCheck1.Text = "Check 1"
        '
        'lblCheck2
        '
        Me.lblCheck2.AutoSize = True
        Me.lblCheck2.Location = New System.Drawing.Point(22, 305)
        Me.lblCheck2.Name = "lblCheck2"
        Me.lblCheck2.Size = New System.Drawing.Size(67, 20)
        Me.lblCheck2.TabIndex = 5
        Me.lblCheck2.Text = "Check 2"
        '
        'lblCheck3
        '
        Me.lblCheck3.AutoSize = True
        Me.lblCheck3.Location = New System.Drawing.Point(22, 337)
        Me.lblCheck3.Name = "lblCheck3"
        Me.lblCheck3.Size = New System.Drawing.Size(67, 20)
        Me.lblCheck3.TabIndex = 6
        Me.lblCheck3.Text = "Check 3"
        '
        'lblCheck4
        '
        Me.lblCheck4.AutoSize = True
        Me.lblCheck4.Location = New System.Drawing.Point(22, 368)
        Me.lblCheck4.Name = "lblCheck4"
        Me.lblCheck4.Size = New System.Drawing.Size(67, 20)
        Me.lblCheck4.TabIndex = 7
        Me.lblCheck4.Text = "Check 4"
        '
        'lblPrep1
        '
        Me.lblPrep1.AutoSize = True
        Me.lblPrep1.Location = New System.Drawing.Point(22, 403)
        Me.lblPrep1.Name = "lblPrep1"
        Me.lblPrep1.Size = New System.Drawing.Size(55, 20)
        Me.lblPrep1.TabIndex = 8
        Me.lblPrep1.Text = "Prep 1"
        '
        'lblPrep2
        '
        Me.lblPrep2.AutoSize = True
        Me.lblPrep2.Location = New System.Drawing.Point(22, 434)
        Me.lblPrep2.Name = "lblPrep2"
        Me.lblPrep2.Size = New System.Drawing.Size(55, 20)
        Me.lblPrep2.TabIndex = 9
        Me.lblPrep2.Text = "Prep 2"
        '
        'lblPrep3
        '
        Me.lblPrep3.AutoSize = True
        Me.lblPrep3.Location = New System.Drawing.Point(22, 465)
        Me.lblPrep3.Name = "lblPrep3"
        Me.lblPrep3.Size = New System.Drawing.Size(55, 20)
        Me.lblPrep3.TabIndex = 10
        Me.lblPrep3.Text = "Prep 3"
        '
        'lblAnalysis1
        '
        Me.lblAnalysis1.AutoSize = True
        Me.lblAnalysis1.Location = New System.Drawing.Point(22, 498)
        Me.lblAnalysis1.Name = "lblAnalysis1"
        Me.lblAnalysis1.Size = New System.Drawing.Size(61, 20)
        Me.lblAnalysis1.TabIndex = 12
        Me.lblAnalysis1.Text = "Label 1"
        '
        'lblAnalysis2
        '
        Me.lblAnalysis2.AutoSize = True
        Me.lblAnalysis2.Location = New System.Drawing.Point(22, 533)
        Me.lblAnalysis2.Name = "lblAnalysis2"
        Me.lblAnalysis2.Size = New System.Drawing.Size(61, 20)
        Me.lblAnalysis2.TabIndex = 15
        Me.lblAnalysis2.Text = "Label 2"
        '
        'lblAnalysis3
        '
        Me.lblAnalysis3.AutoSize = True
        Me.lblAnalysis3.Location = New System.Drawing.Point(22, 572)
        Me.lblAnalysis3.Name = "lblAnalysis3"
        Me.lblAnalysis3.Size = New System.Drawing.Size(61, 20)
        Me.lblAnalysis3.TabIndex = 16
        Me.lblAnalysis3.Text = "Label 3"
        '
        'frmLabelHypergraph
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(906, 716)
        Me.Controls.Add(Me.lblAnalysis3)
        Me.Controls.Add(Me.lblAnalysis2)
        Me.Controls.Add(Me.lblAnalysis1)
        Me.Controls.Add(Me.lblPrep3)
        Me.Controls.Add(Me.lblPrep2)
        Me.Controls.Add(Me.lblPrep1)
        Me.Controls.Add(Me.lblCheck4)
        Me.Controls.Add(Me.lblCheck3)
        Me.Controls.Add(Me.lblCheck2)
        Me.Controls.Add(Me.lblCheck1)
        Me.Controls.Add(Me.cmdRunLabel)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmLabelHypergraph"
        Me.Text = "FIPEX Label Hypergraph Components"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdRunLabel As System.Windows.Forms.Button
    Friend WithEvents lblCheck1 As System.Windows.Forms.Label
    Friend WithEvents lblCheck2 As System.Windows.Forms.Label
    Friend WithEvents lblCheck3 As System.Windows.Forms.Label
    Friend WithEvents lblCheck4 As System.Windows.Forms.Label
    Friend WithEvents lblPrep1 As System.Windows.Forms.Label
    Friend WithEvents lblPrep2 As System.Windows.Forms.Label
    Friend WithEvents lblPrep3 As System.Windows.Forms.Label
    Friend WithEvents lblAnalysis1 As System.Windows.Forms.Label
    Friend WithEvents lblAnalysis2 As System.Windows.Forms.Label
    Friend WithEvents lblAnalysis3 As System.Windows.Forms.Label
End Class
