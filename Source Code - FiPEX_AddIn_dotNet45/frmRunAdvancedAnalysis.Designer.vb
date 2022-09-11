<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRunAdvancedAnalysis
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRunAdvancedAnalysis))
        Me.cmdRun = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.TabBarriers = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.lblGeneralDesc = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.PictureBox4 = New System.Windows.Forms.PictureBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.chkTotalPathDownHab = New System.Windows.Forms.CheckBox()
        Me.chkTotalDownHab = New System.Windows.Forms.CheckBox()
        Me.chkTotalUpHab = New System.Windows.Forms.CheckBox()
        Me.chkPathDownHab = New System.Windows.Forms.CheckBox()
        Me.chkDownHab = New System.Windows.Forms.CheckBox()
        Me.chkUpHab = New System.Windows.Forms.CheckBox()
        Me.frmLayersInclude = New System.Windows.Forms.GroupBox()
        Me.GroupBox8 = New System.Windows.Forms.GroupBox()
        Me.lstLineLengthUnits = New System.Windows.Forms.ListBox()
        Me.lstLineLength = New System.Windows.Forms.ListBox()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.lstLineLayers = New System.Windows.Forms.CheckedListBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lstLineHabCls = New System.Windows.Forms.ListBox()
        Me.lstLineHabUnits = New System.Windows.Forms.ListBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.lstLineHabQuan = New System.Windows.Forms.ListBox()
        Me.cmdChngLineCls = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lstPolyLayers = New System.Windows.Forms.CheckedListBox()
        Me.lstPolyUnit = New System.Windows.Forms.ListBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.cmdChngPolyCls = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lstPolyHabCls = New System.Windows.Forms.ListBox()
        Me.lstPolyHabQuan = New System.Windows.Forms.ListBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.cmdRemove2 = New System.Windows.Forms.Button()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.lstVlsExcld = New System.Windows.Forms.ListBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.lstFtrsExcld = New System.Windows.Forms.ListBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.lstLyrsExcld = New System.Windows.Forms.ListBox()
        Me.Farme3 = New System.Windows.Forms.GroupBox()
        Me.cmdAddExcld = New System.Windows.Forms.Button()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.lstFields = New System.Windows.Forms.ListBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.lstValues = New System.Windows.Forms.ListBox()
        Me.lstLayers = New System.Windows.Forms.ListBox()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.GroupBox9 = New System.Windows.Forms.GroupBox()
        Me.lstBarrierField = New System.Windows.Forms.ListBox()
        Me.lstNaturalTFField = New System.Windows.Forms.ListBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.cmdSelectNaturalTF = New System.Windows.Forms.Button()
        Me.lstPermField = New System.Windows.Forms.ListBox()
        Me.cmdBarrierID = New System.Windows.Forms.Button()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.cmdSelectBarrierPerm = New System.Windows.Forms.Button()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.chkLstBarriersLayers = New System.Windows.Forms.CheckedListBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.GroupBox10 = New System.Windows.Forms.GroupBox()
        Me.rdoAdvancedNet = New System.Windows.Forms.RadioButton()
        Me.rdoBasicConnect = New System.Windows.Forms.RadioButton()
        Me.chkAdvConnect = New System.Windows.Forms.CheckBox()
        Me.chkNaturalTF = New System.Windows.Forms.CheckBox()
        Me.chkBarrierPerm = New System.Windows.Forms.CheckBox()
        Me.chkConnect = New System.Windows.Forms.CheckBox()
        Me.cmdAddGDB = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.ChkDBFOutput = New System.Windows.Forms.CheckBox()
        Me.txtTablesPrefix = New System.Windows.Forms.TextBox()
        Me.txtGDB = New System.Windows.Forms.TextBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.TabPage6 = New System.Windows.Forms.TabPage()
        Me.PictureBox8 = New System.Windows.Forms.PictureBox()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.frmDirection = New System.Windows.Forms.GroupBox()
        Me.OptDown = New System.Windows.Forms.RadioButton()
        Me.OptUp = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TxtOrder = New System.Windows.Forms.TextBox()
        Me.ChkMaxOrd = New System.Windows.Forms.CheckBox()
        Me.PictureBox5 = New System.Windows.Forms.PictureBox()
        Me.TabPage5 = New System.Windows.Forms.TabPage()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.GroupBox13 = New System.Windows.Forms.GroupBox()
        Me.rdoHabArea = New System.Windows.Forms.RadioButton()
        Me.rdoHabLength = New System.Windows.Forms.RadioButton()
        Me.GroupBox11 = New System.Windows.Forms.GroupBox()
        Me.PictureBox7 = New System.Windows.Forms.PictureBox()
        Me.GroupBox12 = New System.Windows.Forms.GroupBox()
        Me.rdoCircle = New System.Windows.Forms.RadioButton()
        Me.chkDistanceDecay = New System.Windows.Forms.CheckBox()
        Me.rdoLinear = New System.Windows.Forms.RadioButton()
        Me.rdoNatExp1 = New System.Windows.Forms.RadioButton()
        Me.rdoSigmoid = New System.Windows.Forms.RadioButton()
        Me.chkDistanceLimit = New System.Windows.Forms.CheckBox()
        Me.txtMaxDistance = New System.Windows.Forms.TextBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.chkDCISectional = New System.Windows.Forms.CheckBox()
        Me.cmdDCIModelDir = New System.Windows.Forms.Button()
        Me.txtDCIModelDir = New System.Windows.Forms.TextBox()
        Me.txtRInstallDir = New System.Windows.Forms.TextBox()
        Me.cmdRInstallDir = New System.Windows.Forms.Button()
        Me.chkDCI = New System.Windows.Forms.CheckBox()
        Me.PictureBox6 = New System.Windows.Forms.PictureBox()
        Me.TabBarriers.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.frmLayersInclude.SuspendLayout()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.Farme3.SuspendLayout()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox9.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox10.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage6.SuspendLayout()
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.frmDirection.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage5.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox13.SuspendLayout()
        Me.GroupBox11.SuspendLayout()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox12.SuspendLayout()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdRun
        '
        Me.cmdRun.BackColor = System.Drawing.Color.PaleGreen
        Me.cmdRun.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRun.Location = New System.Drawing.Point(388, 832)
        Me.cmdRun.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdRun.Name = "cmdRun"
        Me.cmdRun.Size = New System.Drawing.Size(181, 35)
        Me.cmdRun.TabIndex = 9
        Me.cmdRun.Text = "Run Analysis"
        Me.cmdRun.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.Location = New System.Drawing.Point(679, 832)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(173, 35)
        Me.cmdCancel.TabIndex = 10
        Me.cmdCancel.Text = "Cancel and Exit"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(97, 832)
        Me.cmdSave.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(201, 35)
        Me.cmdSave.TabIndex = 11
        Me.cmdSave.Text = "Save Settings"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'TabBarriers
        '
        Me.TabBarriers.AllowDrop = True
        Me.TabBarriers.Controls.Add(Me.TabPage1)
        Me.TabBarriers.Controls.Add(Me.TabPage2)
        Me.TabBarriers.Controls.Add(Me.TabPage3)
        Me.TabBarriers.Controls.Add(Me.TabPage4)
        Me.TabBarriers.Controls.Add(Me.TabPage6)
        Me.TabBarriers.Controls.Add(Me.TabPage5)
        Me.TabBarriers.Location = New System.Drawing.Point(34, 35)
        Me.TabBarriers.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabBarriers.Name = "TabBarriers"
        Me.TabBarriers.SelectedIndex = 0
        Me.TabBarriers.Size = New System.Drawing.Size(942, 769)
        Me.TabBarriers.TabIndex = 12
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.lblGeneralDesc)
        Me.TabPage1.Controls.Add(Me.Label23)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.PictureBox4)
        Me.TabPage1.Controls.Add(Me.GroupBox3)
        Me.TabPage1.Controls.Add(Me.frmLayersInclude)
        Me.TabPage1.Location = New System.Drawing.Point(4, 29)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabPage1.Size = New System.Drawing.Size(934, 736)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "1) General Setup"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'lblGeneralDesc
        '
        Me.lblGeneralDesc.AutoSize = True
        Me.lblGeneralDesc.Location = New System.Drawing.Point(434, 39)
        Me.lblGeneralDesc.Name = "lblGeneralDesc"
        Me.lblGeneralDesc.Size = New System.Drawing.Size(411, 60)
        Me.lblGeneralDesc.TabIndex = 31
        Me.lblGeneralDesc.Text = "General options set the basic behaviour of the network" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "tracing algorithm when th" & _
    "e 'one click' analysis button and " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "the 'advanced analysis' button are clicked."
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Font = New System.Drawing.Font("Calibri", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.Location = New System.Drawing.Point(192, 35)
        Me.Label23.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(203, 35)
        Me.Label23.TabIndex = 30
        Me.Label23.Text = "General Options"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.MenuText
        Me.Label3.Location = New System.Drawing.Point(804, 711)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(114, 20)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "(* = Required)"
        '
        'PictureBox4
        '
        Me.PictureBox4.Image = Global.FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.FIPEX_DCI_Logo_2020_90x90a
        Me.PictureBox4.Location = New System.Drawing.Point(33, 35)
        Me.PictureBox4.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(135, 138)
        Me.PictureBox4.TabIndex = 11
        Me.PictureBox4.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.chkTotalPathDownHab)
        Me.GroupBox3.Controls.Add(Me.chkTotalDownHab)
        Me.GroupBox3.Controls.Add(Me.chkTotalUpHab)
        Me.GroupBox3.Controls.Add(Me.chkPathDownHab)
        Me.GroupBox3.Controls.Add(Me.chkDownHab)
        Me.GroupBox3.Controls.Add(Me.chkUpHab)
        Me.GroupBox3.Location = New System.Drawing.Point(185, 104)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox3.Size = New System.Drawing.Size(697, 118)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Trace Types"
        '
        'chkTotalPathDownHab
        '
        Me.chkTotalPathDownHab.AutoSize = True
        Me.chkTotalPathDownHab.Location = New System.Drawing.Point(456, 66)
        Me.chkTotalPathDownHab.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkTotalPathDownHab.Name = "chkTotalPathDownHab"
        Me.chkTotalPathDownHab.Size = New System.Drawing.Size(201, 24)
        Me.chkTotalPathDownHab.TabIndex = 5
        Me.chkTotalPathDownHab.Text = "Downstream Path Total"
        Me.chkTotalPathDownHab.UseVisualStyleBackColor = True
        '
        'chkTotalDownHab
        '
        Me.chkTotalDownHab.AutoSize = True
        Me.chkTotalDownHab.Location = New System.Drawing.Point(230, 66)
        Me.chkTotalDownHab.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkTotalDownHab.Name = "chkTotalDownHab"
        Me.chkTotalDownHab.Size = New System.Drawing.Size(164, 24)
        Me.chkTotalDownHab.TabIndex = 4
        Me.chkTotalDownHab.Text = "Downstream Total"
        Me.chkTotalDownHab.UseVisualStyleBackColor = True
        '
        'chkTotalUpHab
        '
        Me.chkTotalUpHab.AutoSize = True
        Me.chkTotalUpHab.Location = New System.Drawing.Point(22, 65)
        Me.chkTotalUpHab.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkTotalUpHab.Name = "chkTotalUpHab"
        Me.chkTotalUpHab.Size = New System.Drawing.Size(144, 24)
        Me.chkTotalUpHab.TabIndex = 3
        Me.chkTotalUpHab.Text = "Upstream Total"
        Me.chkTotalUpHab.UseVisualStyleBackColor = True
        '
        'chkPathDownHab
        '
        Me.chkPathDownHab.AutoSize = True
        Me.chkPathDownHab.Location = New System.Drawing.Point(456, 32)
        Me.chkPathDownHab.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkPathDownHab.Name = "chkPathDownHab"
        Me.chkPathDownHab.Size = New System.Drawing.Size(201, 24)
        Me.chkPathDownHab.TabIndex = 2
        Me.chkPathDownHab.Text = "Downstream Path Imm."
        Me.chkPathDownHab.UseVisualStyleBackColor = True
        '
        'chkDownHab
        '
        Me.chkDownHab.AutoSize = True
        Me.chkDownHab.Location = New System.Drawing.Point(230, 29)
        Me.chkDownHab.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkDownHab.Name = "chkDownHab"
        Me.chkDownHab.Size = New System.Drawing.Size(204, 24)
        Me.chkDownHab.TabIndex = 1
        Me.chkDownHab.Text = "Downstream Immediate"
        Me.chkDownHab.UseVisualStyleBackColor = True
        '
        'chkUpHab
        '
        Me.chkUpHab.AutoSize = True
        Me.chkUpHab.Location = New System.Drawing.Point(22, 29)
        Me.chkUpHab.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkUpHab.Name = "chkUpHab"
        Me.chkUpHab.Size = New System.Drawing.Size(184, 24)
        Me.chkUpHab.TabIndex = 0
        Me.chkUpHab.Text = "Upstream Immediate"
        Me.chkUpHab.UseVisualStyleBackColor = True
        '
        'frmLayersInclude
        '
        Me.frmLayersInclude.Controls.Add(Me.GroupBox8)
        Me.frmLayersInclude.Controls.Add(Me.GroupBox2)
        Me.frmLayersInclude.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmLayersInclude.Location = New System.Drawing.Point(29, 232)
        Me.frmLayersInclude.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.frmLayersInclude.Name = "frmLayersInclude"
        Me.frmLayersInclude.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.frmLayersInclude.Size = New System.Drawing.Size(897, 474)
        Me.frmLayersInclude.TabIndex = 1
        Me.frmLayersInclude.TabStop = False
        Me.frmLayersInclude.Text = "Network and Habitat Setup"
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.lstLineLengthUnits)
        Me.GroupBox8.Controls.Add(Me.lstLineLength)
        Me.GroupBox8.Controls.Add(Me.Label34)
        Me.GroupBox8.Controls.Add(Me.Label33)
        Me.GroupBox8.Controls.Add(Me.lstLineLayers)
        Me.GroupBox8.Controls.Add(Me.Label18)
        Me.GroupBox8.Controls.Add(Me.Label1)
        Me.GroupBox8.Controls.Add(Me.lstLineHabCls)
        Me.GroupBox8.Controls.Add(Me.lstLineHabUnits)
        Me.GroupBox8.Controls.Add(Me.Label19)
        Me.GroupBox8.Controls.Add(Me.lstLineHabQuan)
        Me.GroupBox8.Controls.Add(Me.cmdChngLineCls)
        Me.GroupBox8.Location = New System.Drawing.Point(30, 29)
        Me.GroupBox8.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox8.Size = New System.Drawing.Size(844, 217)
        Me.GroupBox8.TabIndex = 26
        Me.GroupBox8.TabStop = False
        Me.GroupBox8.Text = "Line Layers"
        '
        'lstLineLengthUnits
        '
        Me.lstLineLengthUnits.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.lstLineLengthUnits.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstLineLengthUnits.FormattingEnabled = True
        Me.lstLineLengthUnits.ItemHeight = 20
        Me.lstLineLengthUnits.Location = New System.Drawing.Point(639, 53)
        Me.lstLineLengthUnits.Name = "lstLineLengthUnits"
        Me.lstLineLengthUnits.Size = New System.Drawing.Size(165, 20)
        Me.lstLineLengthUnits.TabIndex = 27
        '
        'lstLineLength
        '
        Me.lstLineLength.BackColor = System.Drawing.SystemColors.Info
        Me.lstLineLength.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstLineLength.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstLineLength.FormattingEnabled = True
        Me.lstLineLength.ItemHeight = 20
        Me.lstLineLength.Location = New System.Drawing.Point(639, 28)
        Me.lstLineLength.Name = "lstLineLength"
        Me.lstLineLength.Size = New System.Drawing.Size(169, 20)
        Me.lstLineLength.TabIndex = 26
        '
        'Label34
        '
        Me.Label34.AutoSize = True
        Me.Label34.Location = New System.Drawing.Point(447, 53)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(179, 20)
        Me.Label34.TabIndex = 25
        Me.Label34.Text = "Length / Distance Units:"
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(450, 24)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(182, 20)
        Me.Label33.TabIndex = 24
        Me.Label33.Text = "Length / Distance Field*:"
        '
        'lstLineLayers
        '
        Me.lstLineLayers.FormattingEnabled = True
        Me.lstLineLayers.Location = New System.Drawing.Point(15, 34)
        Me.lstLineLayers.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstLineLayers.Name = "lstLineLayers"
        Me.lstLineLayers.Size = New System.Drawing.Size(419, 130)
        Me.lstLineLayers.TabIndex = 17
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.ForeColor = System.Drawing.SystemColors.InfoText
        Me.Label18.Location = New System.Drawing.Point(454, 83)
        Me.Label18.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(172, 20)
        Me.Label18.TabIndex = 8
        Me.Label18.Text = "Habitat Quantity Field*:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.SystemColors.InfoText
        Me.Label1.Location = New System.Drawing.Point(457, 112)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(169, 20)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "Habitat Quantity Units:"
        '
        'lstLineHabCls
        '
        Me.lstLineHabCls.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.lstLineHabCls.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstLineHabCls.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstLineHabCls.ForeColor = System.Drawing.Color.Black
        Me.lstLineHabCls.FormattingEnabled = True
        Me.lstLineHabCls.ItemHeight = 20
        Me.lstLineHabCls.Location = New System.Drawing.Point(639, 141)
        Me.lstLineHabCls.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstLineHabCls.Name = "lstLineHabCls"
        Me.lstLineHabCls.SelectionMode = System.Windows.Forms.SelectionMode.None
        Me.lstLineHabCls.Size = New System.Drawing.Size(170, 20)
        Me.lstLineHabCls.TabIndex = 11
        '
        'lstLineHabUnits
        '
        Me.lstLineHabUnits.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.lstLineHabUnits.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstLineHabUnits.ForeColor = System.Drawing.SystemColors.MenuText
        Me.lstLineHabUnits.FormattingEnabled = True
        Me.lstLineHabUnits.ItemHeight = 20
        Me.lstLineHabUnits.Location = New System.Drawing.Point(639, 112)
        Me.lstLineHabUnits.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstLineHabUnits.Name = "lstLineHabUnits"
        Me.lstLineHabUnits.Size = New System.Drawing.Size(169, 20)
        Me.lstLineHabUnits.TabIndex = 23
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.ForeColor = System.Drawing.Color.Black
        Me.Label19.Location = New System.Drawing.Point(480, 141)
        Me.Label19.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(146, 20)
        Me.Label19.TabIndex = 7
        Me.Label19.Text = "Habitat Class Field:"
        '
        'lstLineHabQuan
        '
        Me.lstLineHabQuan.BackColor = System.Drawing.SystemColors.Info
        Me.lstLineHabQuan.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstLineHabQuan.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstLineHabQuan.ForeColor = System.Drawing.SystemColors.InfoText
        Me.lstLineHabQuan.FormattingEnabled = True
        Me.lstLineHabQuan.ItemHeight = 20
        Me.lstLineHabQuan.Location = New System.Drawing.Point(639, 83)
        Me.lstLineHabQuan.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstLineHabQuan.Name = "lstLineHabQuan"
        Me.lstLineHabQuan.SelectionMode = System.Windows.Forms.SelectionMode.None
        Me.lstLineHabQuan.Size = New System.Drawing.Size(170, 20)
        Me.lstLineHabQuan.TabIndex = 12
        '
        'cmdChngLineCls
        '
        Me.cmdChngLineCls.Location = New System.Drawing.Point(582, 172)
        Me.cmdChngLineCls.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdChngLineCls.Name = "cmdChngLineCls"
        Me.cmdChngLineCls.Size = New System.Drawing.Size(92, 35)
        Me.cmdChngLineCls.TabIndex = 15
        Me.cmdChngLineCls.Text = "Change..."
        Me.cmdChngLineCls.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lstPolyLayers)
        Me.GroupBox2.Controls.Add(Me.lstPolyUnit)
        Me.GroupBox2.Controls.Add(Me.Label21)
        Me.GroupBox2.Controls.Add(Me.cmdChngPolyCls)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.lstPolyHabCls)
        Me.GroupBox2.Controls.Add(Me.lstPolyHabQuan)
        Me.GroupBox2.Controls.Add(Me.Label20)
        Me.GroupBox2.Location = New System.Drawing.Point(30, 272)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox2.Size = New System.Drawing.Size(844, 181)
        Me.GroupBox2.TabIndex = 25
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Polygon Layers"
        '
        'lstPolyLayers
        '
        Me.lstPolyLayers.FormattingEnabled = True
        Me.lstPolyLayers.Location = New System.Drawing.Point(15, 29)
        Me.lstPolyLayers.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstPolyLayers.Name = "lstPolyLayers"
        Me.lstPolyLayers.Size = New System.Drawing.Size(419, 130)
        Me.lstPolyLayers.TabIndex = 18
        '
        'lstPolyUnit
        '
        Me.lstPolyUnit.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.lstPolyUnit.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstPolyUnit.ForeColor = System.Drawing.SystemColors.InfoText
        Me.lstPolyUnit.FormattingEnabled = True
        Me.lstPolyUnit.ItemHeight = 20
        Me.lstPolyUnit.Location = New System.Drawing.Point(639, 72)
        Me.lstPolyUnit.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstPolyUnit.Name = "lstPolyUnit"
        Me.lstPolyUnit.Size = New System.Drawing.Size(170, 20)
        Me.lstPolyUnit.TabIndex = 22
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.ForeColor = System.Drawing.Color.Black
        Me.Label21.Location = New System.Drawing.Point(480, 105)
        Me.Label21.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(146, 20)
        Me.Label21.TabIndex = 9
        Me.Label21.Text = "Habitat Class Field:"
        '
        'cmdChngPolyCls
        '
        Me.cmdChngPolyCls.Location = New System.Drawing.Point(582, 135)
        Me.cmdChngPolyCls.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdChngPolyCls.Name = "cmdChngPolyCls"
        Me.cmdChngPolyCls.Size = New System.Drawing.Size(92, 35)
        Me.cmdChngPolyCls.TabIndex = 16
        Me.cmdChngPolyCls.Text = "Change..."
        Me.cmdChngPolyCls.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.SystemColors.InfoText
        Me.Label2.Location = New System.Drawing.Point(457, 72)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(169, 20)
        Me.Label2.TabIndex = 20
        Me.Label2.Text = "Habitat Quantity Units:"
        '
        'lstPolyHabCls
        '
        Me.lstPolyHabCls.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.lstPolyHabCls.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstPolyHabCls.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstPolyHabCls.ForeColor = System.Drawing.Color.Black
        Me.lstPolyHabCls.FormattingEnabled = True
        Me.lstPolyHabCls.ItemHeight = 20
        Me.lstPolyHabCls.Location = New System.Drawing.Point(639, 105)
        Me.lstPolyHabCls.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstPolyHabCls.Name = "lstPolyHabCls"
        Me.lstPolyHabCls.SelectionMode = System.Windows.Forms.SelectionMode.None
        Me.lstPolyHabCls.Size = New System.Drawing.Size(170, 20)
        Me.lstPolyHabCls.TabIndex = 13
        '
        'lstPolyHabQuan
        '
        Me.lstPolyHabQuan.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.lstPolyHabQuan.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstPolyHabQuan.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstPolyHabQuan.ForeColor = System.Drawing.SystemColors.InfoText
        Me.lstPolyHabQuan.FormattingEnabled = True
        Me.lstPolyHabQuan.ItemHeight = 20
        Me.lstPolyHabQuan.Location = New System.Drawing.Point(639, 36)
        Me.lstPolyHabQuan.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstPolyHabQuan.Name = "lstPolyHabQuan"
        Me.lstPolyHabQuan.SelectionMode = System.Windows.Forms.SelectionMode.None
        Me.lstPolyHabQuan.Size = New System.Drawing.Size(170, 20)
        Me.lstPolyHabQuan.TabIndex = 14
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.ForeColor = System.Drawing.SystemColors.InfoText
        Me.Label20.Location = New System.Drawing.Point(460, 36)
        Me.Label20.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(166, 20)
        Me.Label20.TabIndex = 10
        Me.Label20.Text = "Habitat Quantity Field:"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Label28)
        Me.TabPage2.Controls.Add(Me.Label4)
        Me.TabPage2.Controls.Add(Me.GroupBox5)
        Me.TabPage2.Controls.Add(Me.Farme3)
        Me.TabPage2.Controls.Add(Me.PictureBox3)
        Me.TabPage2.Location = New System.Drawing.Point(4, 29)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabPage2.Size = New System.Drawing.Size(934, 736)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "2) Exclusions"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(322, 35)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(454, 120)
        Me.Label28.TabIndex = 30
        Me.Label28.Text = resources.GetString("Label28.Text")
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Calibri", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(176, 35)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(129, 70)
        Me.Label4.TabIndex = 29
        Me.Label4.Text = "Exclusion " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Options"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.cmdRemove2)
        Me.GroupBox5.Controls.Add(Me.Label17)
        Me.GroupBox5.Controls.Add(Me.lstVlsExcld)
        Me.GroupBox5.Controls.Add(Me.Label16)
        Me.GroupBox5.Controls.Add(Me.lstFtrsExcld)
        Me.GroupBox5.Controls.Add(Me.Label15)
        Me.GroupBox5.Controls.Add(Me.lstLyrsExcld)
        Me.GroupBox5.Location = New System.Drawing.Point(9, 431)
        Me.GroupBox5.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox5.Size = New System.Drawing.Size(886, 229)
        Me.GroupBox5.TabIndex = 7
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "2. Current Exclusions"
        '
        'cmdRemove2
        '
        Me.cmdRemove2.Location = New System.Drawing.Point(327, 172)
        Me.cmdRemove2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdRemove2.Name = "cmdRemove2"
        Me.cmdRemove2.Size = New System.Drawing.Size(207, 35)
        Me.cmdRemove2.TabIndex = 6
        Me.cmdRemove2.Text = "Remove From Exclusions"
        Me.cmdRemove2.UseVisualStyleBackColor = True
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(645, 38)
        Me.Label17.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(54, 20)
        Me.Label17.TabIndex = 5
        Me.Label17.Text = "Value:"
        '
        'lstVlsExcld
        '
        Me.lstVlsExcld.BackColor = System.Drawing.SystemColors.MenuBar
        Me.lstVlsExcld.FormattingEnabled = True
        Me.lstVlsExcld.ItemHeight = 20
        Me.lstVlsExcld.Location = New System.Drawing.Point(650, 63)
        Me.lstVlsExcld.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstVlsExcld.Name = "lstVlsExcld"
        Me.lstVlsExcld.Size = New System.Drawing.Size(208, 84)
        Me.lstVlsExcld.TabIndex = 4
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(414, 38)
        Me.Label16.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(47, 20)
        Me.Label16.TabIndex = 3
        Me.Label16.Text = "Field:"
        '
        'lstFtrsExcld
        '
        Me.lstFtrsExcld.BackColor = System.Drawing.SystemColors.MenuBar
        Me.lstFtrsExcld.FormattingEnabled = True
        Me.lstFtrsExcld.ItemHeight = 20
        Me.lstFtrsExcld.Location = New System.Drawing.Point(418, 63)
        Me.lstFtrsExcld.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstFtrsExcld.Name = "lstFtrsExcld"
        Me.lstFtrsExcld.Size = New System.Drawing.Size(220, 84)
        Me.lstFtrsExcld.TabIndex = 2
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(22, 38)
        Me.Label15.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(182, 20)
        Me.Label15.TabIndex = 1
        Me.Label15.Text = "Layer (select to remove):"
        '
        'lstLyrsExcld
        '
        Me.lstLyrsExcld.FormattingEnabled = True
        Me.lstLyrsExcld.ItemHeight = 20
        Me.lstLyrsExcld.Location = New System.Drawing.Point(27, 63)
        Me.lstLyrsExcld.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstLyrsExcld.Name = "lstLyrsExcld"
        Me.lstLyrsExcld.Size = New System.Drawing.Size(364, 84)
        Me.lstLyrsExcld.TabIndex = 0
        '
        'Farme3
        '
        Me.Farme3.Controls.Add(Me.cmdAddExcld)
        Me.Farme3.Controls.Add(Me.Label14)
        Me.Farme3.Controls.Add(Me.Label13)
        Me.Farme3.Controls.Add(Me.lstFields)
        Me.Farme3.Controls.Add(Me.Label12)
        Me.Farme3.Controls.Add(Me.lstValues)
        Me.Farme3.Controls.Add(Me.lstLayers)
        Me.Farme3.Location = New System.Drawing.Point(9, 189)
        Me.Farme3.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Farme3.Name = "Farme3"
        Me.Farme3.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Farme3.Size = New System.Drawing.Size(886, 232)
        Me.Farme3.TabIndex = 1
        Me.Farme3.TabStop = False
        Me.Farme3.Text = "1. Select Exclusion "
        '
        'cmdAddExcld
        '
        Me.cmdAddExcld.Location = New System.Drawing.Point(354, 186)
        Me.cmdAddExcld.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdAddExcld.Name = "cmdAddExcld"
        Me.cmdAddExcld.Size = New System.Drawing.Size(153, 37)
        Me.cmdAddExcld.TabIndex = 6
        Me.cmdAddExcld.Text = "Add to Exclusions"
        Me.cmdAddExcld.UseVisualStyleBackColor = True
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(645, 40)
        Me.Label14.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(165, 20)
        Me.Label14.TabIndex = 5
        Me.Label14.Text = "Value (unique values):"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(414, 40)
        Me.Label13.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(199, 20)
        Me.Label13.TabIndex = 4
        Me.Label13.Text = "Field (populate 'Value' box)"
        '
        'lstFields
        '
        Me.lstFields.FormattingEnabled = True
        Me.lstFields.ItemHeight = 20
        Me.lstFields.Location = New System.Drawing.Point(418, 65)
        Me.lstFields.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstFields.Name = "lstFields"
        Me.lstFields.Size = New System.Drawing.Size(220, 104)
        Me.lstFields.TabIndex = 2
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(22, 40)
        Me.Label12.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(260, 20)
        Me.Label12.TabIndex = 1
        Me.Label12.Text = "Layer (select to populate 'Field' box:"
        '
        'lstValues
        '
        Me.lstValues.FormattingEnabled = True
        Me.lstValues.ItemHeight = 20
        Me.lstValues.Location = New System.Drawing.Point(650, 65)
        Me.lstValues.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstValues.Name = "lstValues"
        Me.lstValues.Size = New System.Drawing.Size(208, 104)
        Me.lstValues.TabIndex = 3
        '
        'lstLayers
        '
        Me.lstLayers.FormattingEnabled = True
        Me.lstLayers.ItemHeight = 20
        Me.lstLayers.Location = New System.Drawing.Point(27, 65)
        Me.lstLayers.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstLayers.Name = "lstLayers"
        Me.lstLayers.Size = New System.Drawing.Size(364, 104)
        Me.lstLayers.TabIndex = 0
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = Global.FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.FIPEX_DCI_Logo_2020_90x90a
        Me.PictureBox3.Location = New System.Drawing.Point(33, 35)
        Me.PictureBox3.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(135, 138)
        Me.PictureBox3.TabIndex = 11
        Me.PictureBox3.TabStop = False
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.Label29)
        Me.TabPage3.Controls.Add(Me.Label5)
        Me.TabPage3.Controls.Add(Me.GroupBox6)
        Me.TabPage3.Controls.Add(Me.PictureBox2)
        Me.TabPage3.Location = New System.Drawing.Point(4, 29)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabPage3.Size = New System.Drawing.Size(934, 736)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "3) Barriers"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(328, 38)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(530, 160)
        Me.Label29.TabIndex = 29
        Me.Label29.Text = resources.GetString("Label29.Text")
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Calibri", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(192, 35)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(106, 70)
        Me.Label5.TabIndex = 28
        Me.Label5.Text = "Barrier " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Options"
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.GroupBox9)
        Me.GroupBox6.Controls.Add(Me.Label8)
        Me.GroupBox6.Controls.Add(Me.chkLstBarriersLayers)
        Me.GroupBox6.Location = New System.Drawing.Point(24, 203)
        Me.GroupBox6.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox6.Size = New System.Drawing.Size(834, 437)
        Me.GroupBox6.TabIndex = 0
        Me.GroupBox6.TabStop = False
        '
        'GroupBox9
        '
        Me.GroupBox9.Controls.Add(Me.lstBarrierField)
        Me.GroupBox9.Controls.Add(Me.lstNaturalTFField)
        Me.GroupBox9.Controls.Add(Me.Label9)
        Me.GroupBox9.Controls.Add(Me.Label11)
        Me.GroupBox9.Controls.Add(Me.cmdSelectNaturalTF)
        Me.GroupBox9.Controls.Add(Me.lstPermField)
        Me.GroupBox9.Controls.Add(Me.cmdBarrierID)
        Me.GroupBox9.Controls.Add(Me.Label10)
        Me.GroupBox9.Controls.Add(Me.cmdSelectBarrierPerm)
        Me.GroupBox9.Location = New System.Drawing.Point(496, 28)
        Me.GroupBox9.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox9.Name = "GroupBox9"
        Me.GroupBox9.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox9.Size = New System.Drawing.Size(300, 394)
        Me.GroupBox9.TabIndex = 22
        Me.GroupBox9.TabStop = False
        Me.GroupBox9.Text = "2. Choose Barrier Layer Settings"
        '
        'lstBarrierField
        '
        Me.lstBarrierField.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lstBarrierField.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstBarrierField.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstBarrierField.ForeColor = System.Drawing.SystemColors.InfoText
        Me.lstBarrierField.FormattingEnabled = True
        Me.lstBarrierField.ItemHeight = 20
        Me.lstBarrierField.Location = New System.Drawing.Point(58, 69)
        Me.lstBarrierField.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstBarrierField.Name = "lstBarrierField"
        Me.lstBarrierField.SelectionMode = System.Windows.Forms.SelectionMode.None
        Me.lstBarrierField.Size = New System.Drawing.Size(170, 20)
        Me.lstBarrierField.TabIndex = 13
        '
        'lstNaturalTFField
        '
        Me.lstNaturalTFField.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lstNaturalTFField.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstNaturalTFField.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstNaturalTFField.ForeColor = System.Drawing.SystemColors.InfoText
        Me.lstNaturalTFField.FormattingEnabled = True
        Me.lstNaturalTFField.ItemHeight = 20
        Me.lstNaturalTFField.Location = New System.Drawing.Point(57, 317)
        Me.lstNaturalTFField.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstNaturalTFField.Name = "lstNaturalTFField"
        Me.lstNaturalTFField.SelectionMode = System.Windows.Forms.SelectionMode.None
        Me.lstNaturalTFField.Size = New System.Drawing.Size(170, 20)
        Me.lstNaturalTFField.TabIndex = 21
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(54, 45)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(170, 20)
        Me.Label9.TabIndex = 14
        Me.Label9.Text = "Barrier ID / Label Field:"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(52, 292)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(129, 20)
        Me.Label11.TabIndex = 19
        Me.Label11.Text = "Natural T/F Field:"
        '
        'cmdSelectNaturalTF
        '
        Me.cmdSelectNaturalTF.Location = New System.Drawing.Point(57, 346)
        Me.cmdSelectNaturalTF.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdSelectNaturalTF.Name = "cmdSelectNaturalTF"
        Me.cmdSelectNaturalTF.Size = New System.Drawing.Size(228, 35)
        Me.cmdSelectNaturalTF.TabIndex = 17
        Me.cmdSelectNaturalTF.Text = "Change Natural T/F Field..."
        Me.cmdSelectNaturalTF.UseVisualStyleBackColor = True
        '
        'lstPermField
        '
        Me.lstPermField.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lstPermField.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstPermField.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstPermField.ForeColor = System.Drawing.SystemColors.InfoText
        Me.lstPermField.FormattingEnabled = True
        Me.lstPermField.ItemHeight = 20
        Me.lstPermField.Location = New System.Drawing.Point(58, 200)
        Me.lstPermField.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstPermField.Name = "lstPermField"
        Me.lstPermField.SelectionMode = System.Windows.Forms.SelectionMode.None
        Me.lstPermField.Size = New System.Drawing.Size(170, 20)
        Me.lstPermField.TabIndex = 20
        '
        'cmdBarrierID
        '
        Me.cmdBarrierID.Location = New System.Drawing.Point(58, 98)
        Me.cmdBarrierID.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdBarrierID.Name = "cmdBarrierID"
        Me.cmdBarrierID.Size = New System.Drawing.Size(168, 35)
        Me.cmdBarrierID.TabIndex = 15
        Me.cmdBarrierID.Text = "Change ID Field..."
        Me.cmdBarrierID.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(54, 175)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(142, 20)
        Me.Label10.TabIndex = 18
        Me.Label10.Text = "Permeability Field*:"
        '
        'cmdSelectBarrierPerm
        '
        Me.cmdSelectBarrierPerm.Location = New System.Drawing.Point(57, 229)
        Me.cmdSelectBarrierPerm.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdSelectBarrierPerm.Name = "cmdSelectBarrierPerm"
        Me.cmdSelectBarrierPerm.Size = New System.Drawing.Size(224, 35)
        Me.cmdSelectBarrierPerm.TabIndex = 16
        Me.cmdSelectBarrierPerm.Text = "Change Permeability Field..."
        Me.cmdSelectBarrierPerm.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(20, 28)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(267, 20)
        Me.Label8.TabIndex = 3
        Me.Label8.Text = "1. Select Barrier Layers For Analysis:"
        '
        'chkLstBarriersLayers
        '
        Me.chkLstBarriersLayers.FormattingEnabled = True
        Me.chkLstBarriersLayers.Location = New System.Drawing.Point(24, 57)
        Me.chkLstBarriersLayers.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkLstBarriersLayers.Name = "chkLstBarriersLayers"
        Me.chkLstBarriersLayers.Size = New System.Drawing.Size(408, 340)
        Me.chkLstBarriersLayers.TabIndex = 2
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = Global.FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.FIPEX_DCI_Logo_2020_90x90a
        Me.PictureBox2.Location = New System.Drawing.Point(33, 35)
        Me.PictureBox2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(135, 138)
        Me.PictureBox2.TabIndex = 23
        Me.PictureBox2.TabStop = False
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.Label30)
        Me.TabPage4.Controls.Add(Me.Label22)
        Me.TabPage4.Controls.Add(Me.GroupBox4)
        Me.TabPage4.Controls.Add(Me.PictureBox1)
        Me.TabPage4.Location = New System.Drawing.Point(4, 29)
        Me.TabPage4.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabPage4.Size = New System.Drawing.Size(934, 736)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "4) Output"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Location = New System.Drawing.Point(336, 35)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(535, 140)
        Me.Label30.TabIndex = 28
        Me.Label30.Text = resources.GetString("Label30.Text")
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("Calibri", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.Location = New System.Drawing.Point(195, 35)
        Me.Label22.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(106, 70)
        Me.Label22.TabIndex = 27
        Me.Label22.Text = "Output " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Options"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.GroupBox10)
        Me.GroupBox4.Controls.Add(Me.chkAdvConnect)
        Me.GroupBox4.Controls.Add(Me.chkNaturalTF)
        Me.GroupBox4.Controls.Add(Me.chkBarrierPerm)
        Me.GroupBox4.Controls.Add(Me.chkConnect)
        Me.GroupBox4.Controls.Add(Me.cmdAddGDB)
        Me.GroupBox4.Controls.Add(Me.Label7)
        Me.GroupBox4.Controls.Add(Me.ChkDBFOutput)
        Me.GroupBox4.Controls.Add(Me.txtTablesPrefix)
        Me.GroupBox4.Controls.Add(Me.txtGDB)
        Me.GroupBox4.Location = New System.Drawing.Point(33, 206)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox4.Size = New System.Drawing.Size(869, 496)
        Me.GroupBox4.TabIndex = 8
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Table Output (requires at least one barrier layer to be selected in 'Barriers' ta" & _
    "b)"
        '
        'GroupBox10
        '
        Me.GroupBox10.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox10.Controls.Add(Me.rdoAdvancedNet)
        Me.GroupBox10.Controls.Add(Me.rdoBasicConnect)
        Me.GroupBox10.Location = New System.Drawing.Point(37, 365)
        Me.GroupBox10.Name = "GroupBox10"
        Me.GroupBox10.Size = New System.Drawing.Size(247, 97)
        Me.GroupBox10.TabIndex = 13
        Me.GroupBox10.TabStop = False
        Me.GroupBox10.Text = "Network Connectivity Output"
        Me.GroupBox10.Visible = False
        '
        'rdoAdvancedNet
        '
        Me.rdoAdvancedNet.AutoSize = True
        Me.rdoAdvancedNet.Location = New System.Drawing.Point(20, 58)
        Me.rdoAdvancedNet.Name = "rdoAdvancedNet"
        Me.rdoAdvancedNet.Size = New System.Drawing.Size(169, 24)
        Me.rdoAdvancedNet.TabIndex = 12
        Me.rdoAdvancedNet.TabStop = True
        Me.rdoAdvancedNet.Text = "Advanced Connect"
        Me.rdoAdvancedNet.UseVisualStyleBackColor = True
        '
        'rdoBasicConnect
        '
        Me.rdoBasicConnect.AutoSize = True
        Me.rdoBasicConnect.Location = New System.Drawing.Point(20, 28)
        Me.rdoBasicConnect.Name = "rdoBasicConnect"
        Me.rdoBasicConnect.Size = New System.Drawing.Size(137, 24)
        Me.rdoBasicConnect.TabIndex = 11
        Me.rdoBasicConnect.TabStop = True
        Me.rdoBasicConnect.Text = "Basic Connect"
        Me.rdoBasicConnect.UseVisualStyleBackColor = True
        '
        'chkAdvConnect
        '
        Me.chkAdvConnect.AutoSize = True
        Me.chkAdvConnect.BackColor = System.Drawing.Color.Transparent
        Me.chkAdvConnect.Location = New System.Drawing.Point(37, 308)
        Me.chkAdvConnect.Name = "chkAdvConnect"
        Me.chkAdvConnect.Size = New System.Drawing.Size(354, 24)
        Me.chkAdvConnect.TabIndex = 10
        Me.chkAdvConnect.Text = "Generate Advanced Network Summary Table"
        Me.chkAdvConnect.UseVisualStyleBackColor = False
        Me.chkAdvConnect.Visible = False
        '
        'chkNaturalTF
        '
        Me.chkNaturalTF.AutoSize = True
        Me.chkNaturalTF.Location = New System.Drawing.Point(37, 242)
        Me.chkNaturalTF.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkNaturalTF.Name = "chkNaturalTF"
        Me.chkNaturalTF.Size = New System.Drawing.Size(202, 24)
        Me.chkNaturalTF.TabIndex = 8
        Me.chkNaturalTF.Text = "Include Natural T/F field"
        Me.chkNaturalTF.UseVisualStyleBackColor = True
        '
        'chkBarrierPerm
        '
        Me.chkBarrierPerm.AutoSize = True
        Me.chkBarrierPerm.Location = New System.Drawing.Point(37, 208)
        Me.chkBarrierPerm.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkBarrierPerm.Name = "chkBarrierPerm"
        Me.chkBarrierPerm.Size = New System.Drawing.Size(265, 24)
        Me.chkBarrierPerm.TabIndex = 7
        Me.chkBarrierPerm.Text = "Include Barrier Permeability Field"
        Me.chkBarrierPerm.UseVisualStyleBackColor = True
        '
        'chkConnect
        '
        Me.chkConnect.AutoSize = True
        Me.chkConnect.Location = New System.Drawing.Point(37, 276)
        Me.chkConnect.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkConnect.Name = "chkConnect"
        Me.chkConnect.Size = New System.Drawing.Size(213, 24)
        Me.chkConnect.TabIndex = 6
        Me.chkConnect.Text = "Export Connectivity Table"
        Me.chkConnect.UseVisualStyleBackColor = True
        '
        'cmdAddGDB
        '
        Me.cmdAddGDB.Location = New System.Drawing.Point(30, 102)
        Me.cmdAddGDB.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdAddGDB.Name = "cmdAddGDB"
        Me.cmdAddGDB.Size = New System.Drawing.Size(141, 35)
        Me.cmdAddGDB.TabIndex = 2
        Me.cmdAddGDB.Text = "Browse for GDB:"
        Me.cmdAddGDB.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(65, 146)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(107, 20)
        Me.Label7.TabIndex = 4
        Me.Label7.Text = "Tables Prefix: "
        '
        'ChkDBFOutput
        '
        Me.ChkDBFOutput.AutoSize = True
        Me.ChkDBFOutput.Location = New System.Drawing.Point(31, 42)
        Me.ChkDBFOutput.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.ChkDBFOutput.Name = "ChkDBFOutput"
        Me.ChkDBFOutput.Size = New System.Drawing.Size(295, 24)
        Me.ChkDBFOutput.TabIndex = 0
        Me.ChkDBFOutput.Text = "Output to GeoDatabase DBF Tables"
        Me.ChkDBFOutput.UseVisualStyleBackColor = True
        '
        'txtTablesPrefix
        '
        Me.txtTablesPrefix.Location = New System.Drawing.Point(180, 141)
        Me.txtTablesPrefix.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtTablesPrefix.Name = "txtTablesPrefix"
        Me.txtTablesPrefix.Size = New System.Drawing.Size(178, 26)
        Me.txtTablesPrefix.TabIndex = 3
        '
        'txtGDB
        '
        Me.txtGDB.Location = New System.Drawing.Point(180, 105)
        Me.txtGDB.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtGDB.Name = "txtGDB"
        Me.txtGDB.ReadOnly = True
        Me.txtGDB.Size = New System.Drawing.Size(373, 26)
        Me.txtGDB.TabIndex = 1
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.FIPEX_DCI_Logo_2020_90x90a
        Me.PictureBox1.Location = New System.Drawing.Point(33, 35)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(135, 138)
        Me.PictureBox1.TabIndex = 10
        Me.PictureBox1.TabStop = False
        '
        'TabPage6
        '
        Me.TabPage6.Controls.Add(Me.PictureBox8)
        Me.TabPage6.Controls.Add(Me.Label31)
        Me.TabPage6.Controls.Add(Me.Label24)
        Me.TabPage6.Controls.Add(Me.frmDirection)
        Me.TabPage6.Controls.Add(Me.GroupBox1)
        Me.TabPage6.Controls.Add(Me.PictureBox5)
        Me.TabPage6.Location = New System.Drawing.Point(4, 29)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage6.Size = New System.Drawing.Size(934, 736)
        Me.TabPage6.TabIndex = 5
        Me.TabPage6.Text = "5) Advanced"
        Me.TabPage6.UseVisualStyleBackColor = True
        '
        'PictureBox8
        '
        Me.PictureBox8.Image = Global.FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.orderdiagram
        Me.PictureBox8.Location = New System.Drawing.Point(59, 343)
        Me.PictureBox8.Name = "PictureBox8"
        Me.PictureBox8.Size = New System.Drawing.Size(677, 347)
        Me.PictureBox8.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox8.TabIndex = 28
        Me.PictureBox8.TabStop = False
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(331, 35)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(511, 100)
        Me.Label31.TabIndex = 27
        Me.Label31.Text = resources.GetString("Label31.Text")
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Font = New System.Drawing.Font("Calibri", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.Location = New System.Drawing.Point(176, 35)
        Me.Label24.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(130, 70)
        Me.Label24.TabIndex = 26
        Me.Label24.Text = "Advanced" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Options"
        '
        'frmDirection
        '
        Me.frmDirection.Controls.Add(Me.OptDown)
        Me.frmDirection.Controls.Add(Me.OptUp)
        Me.frmDirection.Location = New System.Drawing.Point(394, 208)
        Me.frmDirection.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.frmDirection.Name = "frmDirection"
        Me.frmDirection.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.frmDirection.Size = New System.Drawing.Size(194, 103)
        Me.frmDirection.TabIndex = 11
        Me.frmDirection.TabStop = False
        Me.frmDirection.Text = "Analysis Direction"
        '
        'OptDown
        '
        Me.OptDown.AutoSize = True
        Me.OptDown.Location = New System.Drawing.Point(33, 66)
        Me.OptDown.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.OptDown.Name = "OptDown"
        Me.OptDown.Size = New System.Drawing.Size(124, 24)
        Me.OptDown.TabIndex = 1
        Me.OptDown.TabStop = True
        Me.OptDown.Text = "Downstream"
        Me.OptDown.UseVisualStyleBackColor = True
        '
        'OptUp
        '
        Me.OptUp.AutoSize = True
        Me.OptUp.Location = New System.Drawing.Point(33, 31)
        Me.OptUp.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.OptUp.Name = "OptUp"
        Me.OptUp.Size = New System.Drawing.Size(104, 24)
        Me.OptUp.TabIndex = 0
        Me.OptUp.TabStop = True
        Me.OptUp.Text = "Upstream"
        Me.OptUp.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.TxtOrder)
        Me.GroupBox1.Controls.Add(Me.ChkMaxOrd)
        Me.GroupBox1.Location = New System.Drawing.Point(67, 208)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox1.Size = New System.Drawing.Size(319, 103)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Max 'Depth' of BFS Network Algorithm"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(8, 66)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(57, 20)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "Depth:"
        '
        'TxtOrder
        '
        Me.TxtOrder.Location = New System.Drawing.Point(69, 63)
        Me.TxtOrder.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TxtOrder.Name = "TxtOrder"
        Me.TxtOrder.Size = New System.Drawing.Size(86, 26)
        Me.TxtOrder.TabIndex = 1
        '
        'ChkMaxOrd
        '
        Me.ChkMaxOrd.AutoSize = True
        Me.ChkMaxOrd.Location = New System.Drawing.Point(8, 29)
        Me.ChkMaxOrd.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.ChkMaxOrd.Name = "ChkMaxOrd"
        Me.ChkMaxOrd.Size = New System.Drawing.Size(165, 24)
        Me.ChkMaxOrd.TabIndex = 0
        Me.ChkMaxOrd.Text = "Complete Network"
        Me.ChkMaxOrd.UseVisualStyleBackColor = True
        '
        'PictureBox5
        '
        Me.PictureBox5.Image = Global.FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.FIPEX_DCI_Logo_2020_90x90a
        Me.PictureBox5.Location = New System.Drawing.Point(33, 35)
        Me.PictureBox5.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(135, 138)
        Me.PictureBox5.TabIndex = 12
        Me.PictureBox5.TabStop = False
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.Label32)
        Me.TabPage5.Controls.Add(Me.Label25)
        Me.TabPage5.Controls.Add(Me.GroupBox7)
        Me.TabPage5.Controls.Add(Me.PictureBox6)
        Me.TabPage5.Location = New System.Drawing.Point(4, 29)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage5.Size = New System.Drawing.Size(934, 736)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "6) DCI"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Location = New System.Drawing.Point(327, 35)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(559, 140)
        Me.Label32.TabIndex = 28
        Me.Label32.Text = resources.GetString("Label32.Text")
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Font = New System.Drawing.Font("Calibri", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.Location = New System.Drawing.Point(176, 35)
        Me.Label25.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(152, 35)
        Me.Label25.TabIndex = 27
        Me.Label25.Text = "DCI Options"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.GroupBox13)
        Me.GroupBox7.Controls.Add(Me.GroupBox11)
        Me.GroupBox7.Controls.Add(Me.chkDCISectional)
        Me.GroupBox7.Controls.Add(Me.cmdDCIModelDir)
        Me.GroupBox7.Controls.Add(Me.txtDCIModelDir)
        Me.GroupBox7.Controls.Add(Me.txtRInstallDir)
        Me.GroupBox7.Controls.Add(Me.cmdRInstallDir)
        Me.GroupBox7.Controls.Add(Me.chkDCI)
        Me.GroupBox7.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.GroupBox7.Location = New System.Drawing.Point(33, 199)
        Me.GroupBox7.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox7.Size = New System.Drawing.Size(853, 511)
        Me.GroupBox7.TabIndex = 12
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "DCI Calculation in R 3.6.1"
        '
        'GroupBox13
        '
        Me.GroupBox13.Controls.Add(Me.rdoHabArea)
        Me.GroupBox13.Controls.Add(Me.rdoHabLength)
        Me.GroupBox13.Location = New System.Drawing.Point(549, 29)
        Me.GroupBox13.Name = "GroupBox13"
        Me.GroupBox13.Size = New System.Drawing.Size(269, 123)
        Me.GroupBox13.TabIndex = 17
        Me.GroupBox13.TabStop = False
        Me.GroupBox13.Text = "DCI Habitat"
        '
        'rdoHabArea
        '
        Me.rdoHabArea.AutoSize = True
        Me.rdoHabArea.Location = New System.Drawing.Point(32, 74)
        Me.rdoHabArea.Name = "rdoHabArea"
        Me.rdoHabArea.Size = New System.Drawing.Size(179, 24)
        Me.rdoHabArea.TabIndex = 16
        Me.rdoHabArea.TabStop = True
        Me.rdoHabArea.Text = "Use Area (Polygons)"
        Me.rdoHabArea.UseVisualStyleBackColor = True
        '
        'rdoHabLength
        '
        Me.rdoHabLength.AutoSize = True
        Me.rdoHabLength.Location = New System.Drawing.Point(32, 34)
        Me.rdoHabLength.Name = "rdoHabLength"
        Me.rdoHabLength.Size = New System.Drawing.Size(169, 24)
        Me.rdoHabLength.TabIndex = 15
        Me.rdoHabLength.TabStop = True
        Me.rdoHabLength.Text = "Use Length (Lines)"
        Me.rdoHabLength.UseVisualStyleBackColor = True
        '
        'GroupBox11
        '
        Me.GroupBox11.Controls.Add(Me.PictureBox7)
        Me.GroupBox11.Controls.Add(Me.GroupBox12)
        Me.GroupBox11.Controls.Add(Me.chkDistanceLimit)
        Me.GroupBox11.Controls.Add(Me.txtMaxDistance)
        Me.GroupBox11.Controls.Add(Me.Label26)
        Me.GroupBox11.Controls.Add(Me.Label27)
        Me.GroupBox11.Location = New System.Drawing.Point(30, 174)
        Me.GroupBox11.Name = "GroupBox11"
        Me.GroupBox11.Size = New System.Drawing.Size(788, 318)
        Me.GroupBox11.TabIndex = 14
        Me.GroupBox11.TabStop = False
        Me.GroupBox11.Text = "Distance Limits"
        '
        'PictureBox7
        '
        Me.PictureBox7.Location = New System.Drawing.Point(426, 18)
        Me.PictureBox7.Name = "PictureBox7"
        Me.PictureBox7.Size = New System.Drawing.Size(300, 300)
        Me.PictureBox7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox7.TabIndex = 15
        Me.PictureBox7.TabStop = False
        '
        'GroupBox12
        '
        Me.GroupBox12.Controls.Add(Me.rdoCircle)
        Me.GroupBox12.Controls.Add(Me.chkDistanceDecay)
        Me.GroupBox12.Controls.Add(Me.rdoLinear)
        Me.GroupBox12.Controls.Add(Me.rdoNatExp1)
        Me.GroupBox12.Controls.Add(Me.rdoSigmoid)
        Me.GroupBox12.Location = New System.Drawing.Point(52, 136)
        Me.GroupBox12.Name = "GroupBox12"
        Me.GroupBox12.Size = New System.Drawing.Size(325, 176)
        Me.GroupBox12.TabIndex = 14
        Me.GroupBox12.TabStop = False
        '
        'rdoCircle
        '
        Me.rdoCircle.AutoSize = True
        Me.rdoCircle.Location = New System.Drawing.Point(91, 116)
        Me.rdoCircle.Name = "rdoCircle"
        Me.rdoCircle.Size = New System.Drawing.Size(167, 24)
        Me.rdoCircle.TabIndex = 11
        Me.rdoCircle.TabStop = True
        Me.rdoCircle.Text = "Circular (1-x^2)^0.5"
        Me.rdoCircle.UseVisualStyleBackColor = True
        '
        'chkDistanceDecay
        '
        Me.chkDistanceDecay.AutoSize = True
        Me.chkDistanceDecay.Location = New System.Drawing.Point(19, 25)
        Me.chkDistanceDecay.Name = "chkDistanceDecay"
        Me.chkDistanceDecay.Size = New System.Drawing.Size(190, 24)
        Me.chkDistanceDecay.TabIndex = 10
        Me.chkDistanceDecay.Text = "Apply Distance Decay"
        Me.chkDistanceDecay.UseVisualStyleBackColor = True
        '
        'rdoLinear
        '
        Me.rdoLinear.AutoSize = True
        Me.rdoLinear.Location = New System.Drawing.Point(91, 56)
        Me.rdoLinear.Name = "rdoLinear"
        Me.rdoLinear.Size = New System.Drawing.Size(78, 24)
        Me.rdoLinear.TabIndex = 7
        Me.rdoLinear.TabStop = True
        Me.rdoLinear.Text = "Linear"
        Me.rdoLinear.UseVisualStyleBackColor = True
        '
        'rdoNatExp1
        '
        Me.rdoNatExp1.AutoSize = True
        Me.rdoNatExp1.Location = New System.Drawing.Point(91, 86)
        Me.rdoNatExp1.Name = "rdoNatExp1"
        Me.rdoNatExp1.Size = New System.Drawing.Size(209, 24)
        Me.rdoNatExp1.TabIndex = 9
        Me.rdoNatExp1.TabStop = True
        Me.rdoNatExp1.Text = "Natural Exponential (e^x)"
        Me.rdoNatExp1.UseVisualStyleBackColor = True
        '
        'rdoSigmoid
        '
        Me.rdoSigmoid.AutoSize = True
        Me.rdoSigmoid.Location = New System.Drawing.Point(91, 146)
        Me.rdoSigmoid.Name = "rdoSigmoid"
        Me.rdoSigmoid.Size = New System.Drawing.Size(91, 24)
        Me.rdoSigmoid.TabIndex = 8
        Me.rdoSigmoid.TabStop = True
        Me.rdoSigmoid.Text = "Sigmoid"
        Me.rdoSigmoid.UseVisualStyleBackColor = True
        '
        'chkDistanceLimit
        '
        Me.chkDistanceLimit.AutoSize = True
        Me.chkDistanceLimit.Location = New System.Drawing.Point(25, 39)
        Me.chkDistanceLimit.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkDistanceLimit.Name = "chkDistanceLimit"
        Me.chkDistanceLimit.Size = New System.Drawing.Size(178, 24)
        Me.chkDistanceLimit.TabIndex = 6
        Me.chkDistanceLimit.Text = "Apply Distance Limit"
        Me.chkDistanceLimit.UseVisualStyleBackColor = True
        '
        'txtMaxDistance
        '
        Me.txtMaxDistance.Location = New System.Drawing.Point(67, 104)
        Me.txtMaxDistance.Name = "txtMaxDistance"
        Me.txtMaxDistance.Size = New System.Drawing.Size(123, 26)
        Me.txtMaxDistance.TabIndex = 11
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.Location = New System.Drawing.Point(199, 104)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(139, 20)
        Me.Label26.TabIndex = 12
        Me.Label26.Text = "(units set in Tab 1)"
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(63, 77)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(193, 20)
        Me.Label27.TabIndex = 13
        Me.Label27.Text = "Maximum distance / cutoff"
        '
        'chkDCISectional
        '
        Me.chkDCISectional.AutoSize = True
        Me.chkDCISectional.Location = New System.Drawing.Point(252, 29)
        Me.chkDCISectional.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkDCISectional.Name = "chkDCISectional"
        Me.chkDCISectional.Size = New System.Drawing.Size(214, 24)
        Me.chkDCISectional.TabIndex = 5
        Me.chkDCISectional.Text = "Calculate Segmental DCI"
        Me.chkDCISectional.UseVisualStyleBackColor = True
        '
        'cmdDCIModelDir
        '
        Me.cmdDCIModelDir.Location = New System.Drawing.Point(30, 119)
        Me.cmdDCIModelDir.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdDCIModelDir.Name = "cmdDCIModelDir"
        Me.cmdDCIModelDir.Size = New System.Drawing.Size(194, 31)
        Me.cmdDCIModelDir.TabIndex = 4
        Me.cmdDCIModelDir.Text = "DCI Model Dir:"
        Me.cmdDCIModelDir.UseVisualStyleBackColor = True
        '
        'txtDCIModelDir
        '
        Me.txtDCIModelDir.Location = New System.Drawing.Point(233, 124)
        Me.txtDCIModelDir.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtDCIModelDir.Name = "txtDCIModelDir"
        Me.txtDCIModelDir.Size = New System.Drawing.Size(299, 26)
        Me.txtDCIModelDir.TabIndex = 3
        '
        'txtRInstallDir
        '
        Me.txtRInstallDir.Location = New System.Drawing.Point(233, 79)
        Me.txtRInstallDir.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtRInstallDir.Name = "txtRInstallDir"
        Me.txtRInstallDir.Size = New System.Drawing.Size(299, 26)
        Me.txtRInstallDir.TabIndex = 2
        '
        'cmdRInstallDir
        '
        Me.cmdRInstallDir.Location = New System.Drawing.Point(30, 70)
        Me.cmdRInstallDir.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdRInstallDir.Name = "cmdRInstallDir"
        Me.cmdRInstallDir.Size = New System.Drawing.Size(194, 35)
        Me.cmdRInstallDir.TabIndex = 1
        Me.cmdRInstallDir.Text = "R Installation Directory:"
        Me.cmdRInstallDir.UseVisualStyleBackColor = True
        '
        'chkDCI
        '
        Me.chkDCI.AutoSize = True
        Me.chkDCI.Location = New System.Drawing.Point(30, 29)
        Me.chkDCI.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkDCI.Name = "chkDCI"
        Me.chkDCI.Size = New System.Drawing.Size(214, 24)
        Me.chkDCI.TabIndex = 0
        Me.chkDCI.Text = "Calculate DCId and DCIp"
        Me.chkDCI.UseVisualStyleBackColor = True
        '
        'PictureBox6
        '
        Me.PictureBox6.Image = Global.FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.FIPEX_DCI_Logo_2020_90x90a
        Me.PictureBox6.Location = New System.Drawing.Point(33, 35)
        Me.PictureBox6.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.PictureBox6.Name = "PictureBox6"
        Me.PictureBox6.Size = New System.Drawing.Size(135, 138)
        Me.PictureBox6.TabIndex = 11
        Me.PictureBox6.TabStop = False
        '
        'frmRunAdvancedAnalysis
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(1012, 903)
        Me.Controls.Add(Me.TabBarriers)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdRun)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "frmRunAdvancedAnalysis"
        Me.Text = "FIPEX - Advanced Analysis"
        Me.TabBarriers.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.frmLayersInclude.ResumeLayout(False)
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox8.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.Farme3.ResumeLayout(False)
        Me.Farme3.PerformLayout()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.GroupBox9.ResumeLayout(False)
        Me.GroupBox9.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox10.ResumeLayout(False)
        Me.GroupBox10.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage6.ResumeLayout(False)
        Me.TabPage6.PerformLayout()
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).EndInit()
        Me.frmDirection.ResumeLayout(False)
        Me.frmDirection.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage5.ResumeLayout(False)
        Me.TabPage5.PerformLayout()
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
        Me.GroupBox13.ResumeLayout(False)
        Me.GroupBox13.PerformLayout()
        Me.GroupBox11.ResumeLayout(False)
        Me.GroupBox11.PerformLayout()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox12.ResumeLayout(False)
        Me.GroupBox12.PerformLayout()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdRun As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents TabBarriers As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents chkTotalPathDownHab As System.Windows.Forms.CheckBox
    Friend WithEvents chkTotalDownHab As System.Windows.Forms.CheckBox
    Friend WithEvents chkTotalUpHab As System.Windows.Forms.CheckBox
    Friend WithEvents chkPathDownHab As System.Windows.Forms.CheckBox
    Friend WithEvents chkDownHab As System.Windows.Forms.CheckBox
    Friend WithEvents chkUpHab As System.Windows.Forms.CheckBox
    Friend WithEvents frmLayersInclude As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents lstLineLayers As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lstLineHabCls As System.Windows.Forms.ListBox
    Friend WithEvents lstLineHabUnits As System.Windows.Forms.ListBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents lstLineHabQuan As System.Windows.Forms.ListBox
    Friend WithEvents cmdChngLineCls As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lstPolyLayers As System.Windows.Forms.CheckedListBox
    Friend WithEvents lstPolyUnit As System.Windows.Forms.ListBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents cmdChngPolyCls As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lstPolyHabCls As System.Windows.Forms.ListBox
    Friend WithEvents lstPolyHabQuan As System.Windows.Forms.ListBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdRemove2 As System.Windows.Forms.Button
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents lstVlsExcld As System.Windows.Forms.ListBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents lstFtrsExcld As System.Windows.Forms.ListBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents lstLyrsExcld As System.Windows.Forms.ListBox
    Friend WithEvents Farme3 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdAddExcld As System.Windows.Forms.Button
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents lstFields As System.Windows.Forms.ListBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents lstValues As System.Windows.Forms.ListBox
    Friend WithEvents lstLayers As System.Windows.Forms.ListBox
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox9 As System.Windows.Forms.GroupBox
    Friend WithEvents lstBarrierField As System.Windows.Forms.ListBox
    Friend WithEvents lstNaturalTFField As System.Windows.Forms.ListBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cmdSelectNaturalTF As System.Windows.Forms.Button
    Friend WithEvents lstPermField As System.Windows.Forms.ListBox
    Friend WithEvents cmdBarrierID As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cmdSelectBarrierPerm As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents chkLstBarriersLayers As System.Windows.Forms.CheckedListBox
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents chkDistanceLimit As System.Windows.Forms.CheckBox
    Friend WithEvents chkDCISectional As System.Windows.Forms.CheckBox
    Friend WithEvents cmdDCIModelDir As System.Windows.Forms.Button
    Friend WithEvents txtDCIModelDir As System.Windows.Forms.TextBox
    Friend WithEvents txtRInstallDir As System.Windows.Forms.TextBox
    Friend WithEvents cmdRInstallDir As System.Windows.Forms.Button
    Friend WithEvents chkDCI As System.Windows.Forms.CheckBox
    Friend WithEvents PictureBox6 As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox10 As System.Windows.Forms.GroupBox
    Friend WithEvents rdoAdvancedNet As System.Windows.Forms.RadioButton
    Friend WithEvents rdoBasicConnect As System.Windows.Forms.RadioButton
    Friend WithEvents chkAdvConnect As System.Windows.Forms.CheckBox
    Friend WithEvents chkNaturalTF As System.Windows.Forms.CheckBox
    Friend WithEvents chkBarrierPerm As System.Windows.Forms.CheckBox
    Friend WithEvents chkConnect As System.Windows.Forms.CheckBox
    Friend WithEvents cmdAddGDB As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ChkDBFOutput As System.Windows.Forms.CheckBox
    Friend WithEvents txtTablesPrefix As System.Windows.Forms.TextBox
    Friend WithEvents txtGDB As System.Windows.Forms.TextBox
    Friend WithEvents TabPage6 As System.Windows.Forms.TabPage
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Friend WithEvents frmDirection As System.Windows.Forms.GroupBox
    Friend WithEvents OptDown As System.Windows.Forms.RadioButton
    Friend WithEvents OptUp As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtOrder As System.Windows.Forms.TextBox
    Friend WithEvents ChkMaxOrd As System.Windows.Forms.CheckBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents rdoNatExp1 As System.Windows.Forms.RadioButton
    Friend WithEvents rdoSigmoid As System.Windows.Forms.RadioButton
    Friend WithEvents rdoLinear As System.Windows.Forms.RadioButton
    Friend WithEvents txtMaxDistance As System.Windows.Forms.TextBox
    Friend WithEvents chkDistanceDecay As System.Windows.Forms.CheckBox
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents GroupBox11 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox12 As System.Windows.Forms.GroupBox
    Friend WithEvents rdoCircle As System.Windows.Forms.RadioButton
    Friend WithEvents PictureBox7 As System.Windows.Forms.PictureBox
    Friend WithEvents lblGeneralDesc As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents lstLineLengthUnits As System.Windows.Forms.ListBox
    Friend WithEvents lstLineLength As System.Windows.Forms.ListBox
    Friend WithEvents GroupBox13 As System.Windows.Forms.GroupBox
    Friend WithEvents rdoHabArea As System.Windows.Forms.RadioButton
    Friend WithEvents rdoHabLength As System.Windows.Forms.RadioButton
    Friend WithEvents PictureBox8 As System.Windows.Forms.PictureBox
End Class
