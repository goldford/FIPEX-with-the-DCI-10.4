<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmVisualizeDecisionsAndNet
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVisualizeDecisionsAndNet))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmdBrowseResultsTab = New System.Windows.Forms.Button()
        Me.cmdHighlight = New System.Windows.Forms.Button()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.txtResultsTab = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmdBrowseDecisionsTab = New System.Windows.Forms.Button()
        Me.txtDecisionsTab = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lblStep1Status = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cmdBrowseConnectivityTab = New System.Windows.Forms.Button()
        Me.lblStep2Status = New System.Windows.Forms.Label()
        Me.txtConnectivityTab = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblStep3Status = New System.Windows.Forms.Label()
        Me.lstLineLayers = New System.Windows.Forms.CheckedListBox()
        Me.grpLines = New System.Windows.Forms.GroupBox()
        Me.lblLineFieldSet = New System.Windows.Forms.Label()
        Me.cmdSetLineField = New System.Windows.Forms.Button()
        Me.label23 = New System.Windows.Forms.Label()
        Me.cmdAddFieldLine = New System.Windows.Forms.Button()
        Me.txtNewLinefield = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.lstLineFieldsDouble = New System.Windows.Forms.ListBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.grpPolygons = New System.Windows.Forms.GroupBox()
        Me.cmdSetPolyField = New System.Windows.Forms.Button()
        Me.lblPolyFieldSet = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cmdAddFieldPoly = New System.Windows.Forms.Button()
        Me.txtNewPolyField = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.lstPolyFieldsDouble = New System.Windows.Forms.ListBox()
        Me.lstPolyLayers = New System.Windows.Forms.CheckedListBox()
        Me.lstBudgets = New System.Windows.Forms.ListBox()
        Me.grpStep6 = New System.Windows.Forms.GroupBox()
        Me.lblStep6bStatus = New System.Windows.Forms.Label()
        Me.lblStep6aStatus = New System.Windows.Forms.Label()
        Me.cmdCalcCumPerm = New System.Windows.Forms.Button()
        Me.lblStep6 = New System.Windows.Forms.Label()
        Me.lblStep4Status = New System.Windows.Forms.Label()
        Me.lblStep5Status = New System.Windows.Forms.Label()
        Me.lblStep1Warning = New System.Windows.Forms.Label()
        Me.lblStep2warning = New System.Windows.Forms.Label()
        Me.lblStep3Warning = New System.Windows.Forms.Label()
        Me.lblStep4Warning = New System.Windows.Forms.Label()
        Me.lblStep5Warning = New System.Windows.Forms.Label()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpLines.SuspendLayout()
        Me.grpPolygons.SuspendLayout()
        Me.grpStep6.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.My.Resources.Resources.FiPEx_LOGOv3b_90x90
        Me.PictureBox1.Location = New System.Drawing.Point(22, 12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(90, 90)
        Me.PictureBox1.TabIndex = 16
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(148, 37)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(363, 104)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = resources.GetString("Label1.Text")
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(177, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(326, 16)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Visualize Decisions and Network Accessibility"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(18, 160)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(219, 13)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "1. Select Optimisation 'Results' Table"
        '
        'cmdBrowseResultsTab
        '
        Me.cmdBrowseResultsTab.Location = New System.Drawing.Point(31, 176)
        Me.cmdBrowseResultsTab.Name = "cmdBrowseResultsTab"
        Me.cmdBrowseResultsTab.Size = New System.Drawing.Size(117, 23)
        Me.cmdBrowseResultsTab.TabIndex = 21
        Me.cmdBrowseResultsTab.Text = "Browse for Table"
        Me.cmdBrowseResultsTab.UseVisualStyleBackColor = True
        '
        'cmdHighlight
        '
        Me.cmdHighlight.Location = New System.Drawing.Point(51, 49)
        Me.cmdHighlight.Name = "cmdHighlight"
        Me.cmdHighlight.Size = New System.Drawing.Size(118, 23)
        Me.cmdHighlight.TabIndex = 22
        Me.cmdHighlight.Text = "Highlight Decisions"
        Me.cmdHighlight.UseVisualStyleBackColor = True
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(290, 753)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 23
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'txtResultsTab
        '
        Me.txtResultsTab.Location = New System.Drawing.Point(154, 178)
        Me.txtResultsTab.Name = "txtResultsTab"
        Me.txtResultsTab.Size = New System.Drawing.Size(366, 20)
        Me.txtResultsTab.TabIndex = 24
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(19, 214)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(232, 13)
        Me.Label4.TabIndex = 25
        Me.Label4.Text = "2. Select Optimisation 'Decisions' Table"
        '
        'cmdBrowseDecisionsTab
        '
        Me.cmdBrowseDecisionsTab.Location = New System.Drawing.Point(32, 230)
        Me.cmdBrowseDecisionsTab.Name = "cmdBrowseDecisionsTab"
        Me.cmdBrowseDecisionsTab.Size = New System.Drawing.Size(117, 23)
        Me.cmdBrowseDecisionsTab.TabIndex = 26
        Me.cmdBrowseDecisionsTab.Text = "Browse for Table"
        Me.cmdBrowseDecisionsTab.UseVisualStyleBackColor = True
        '
        'txtDecisionsTab
        '
        Me.txtDecisionsTab.Location = New System.Drawing.Point(155, 230)
        Me.txtDecisionsTab.Name = "txtDecisionsTab"
        Me.txtDecisionsTab.Size = New System.Drawing.Size(365, 20)
        Me.txtDecisionsTab.TabIndex = 27
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(19, 619)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(148, 13)
        Me.Label5.TabIndex = 29
        Me.Label5.Text = "5. Select Budget Amount"
        '
        'lblStep1Status
        '
        Me.lblStep1Status.AutoSize = True
        Me.lblStep1Status.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStep1Status.Location = New System.Drawing.Point(540, 160)
        Me.lblStep1Status.Name = "lblStep1Status"
        Me.lblStep1Status.Size = New System.Drawing.Size(72, 13)
        Me.lblStep1Status.TabIndex = 30
        Me.lblStep1Status.Text = "step1status"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(19, 267)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(174, 13)
        Me.Label6.TabIndex = 31
        Me.Label6.Text = "3. Select 'Connectivity' Table"
        '
        'cmdBrowseConnectivityTab
        '
        Me.cmdBrowseConnectivityTab.Location = New System.Drawing.Point(32, 283)
        Me.cmdBrowseConnectivityTab.Name = "cmdBrowseConnectivityTab"
        Me.cmdBrowseConnectivityTab.Size = New System.Drawing.Size(117, 23)
        Me.cmdBrowseConnectivityTab.TabIndex = 32
        Me.cmdBrowseConnectivityTab.Text = "Browse for Table"
        Me.cmdBrowseConnectivityTab.UseVisualStyleBackColor = True
        '
        'lblStep2Status
        '
        Me.lblStep2Status.AutoSize = True
        Me.lblStep2Status.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStep2Status.Location = New System.Drawing.Point(540, 214)
        Me.lblStep2Status.Name = "lblStep2Status"
        Me.lblStep2Status.Size = New System.Drawing.Size(72, 13)
        Me.lblStep2Status.TabIndex = 33
        Me.lblStep2Status.Text = "step2status"
        '
        'txtConnectivityTab
        '
        Me.txtConnectivityTab.Location = New System.Drawing.Point(154, 286)
        Me.txtConnectivityTab.Name = "txtConnectivityTab"
        Me.txtConnectivityTab.Size = New System.Drawing.Size(365, 20)
        Me.txtConnectivityTab.TabIndex = 34
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(18, 318)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(361, 13)
        Me.Label7.TabIndex = 35
        Me.Label7.Text = "4. Select Network fields to update (layers from FIPEX settings)"
        '
        'lblStep3Status
        '
        Me.lblStep3Status.AutoSize = True
        Me.lblStep3Status.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStep3Status.Location = New System.Drawing.Point(540, 267)
        Me.lblStep3Status.Name = "lblStep3Status"
        Me.lblStep3Status.Size = New System.Drawing.Size(72, 13)
        Me.lblStep3Status.TabIndex = 36
        Me.lblStep3Status.Text = "step3status"
        '
        'lstLineLayers
        '
        Me.lstLineLayers.FormattingEnabled = True
        Me.lstLineLayers.Location = New System.Drawing.Point(21, 19)
        Me.lstLineLayers.Name = "lstLineLayers"
        Me.lstLineLayers.Size = New System.Drawing.Size(306, 94)
        Me.lstLineLayers.TabIndex = 37
        '
        'grpLines
        '
        Me.grpLines.Controls.Add(Me.lblLineFieldSet)
        Me.grpLines.Controls.Add(Me.cmdSetLineField)
        Me.grpLines.Controls.Add(Me.label23)
        Me.grpLines.Controls.Add(Me.cmdAddFieldLine)
        Me.grpLines.Controls.Add(Me.txtNewLinefield)
        Me.grpLines.Controls.Add(Me.Label9)
        Me.grpLines.Controls.Add(Me.lstLineFieldsDouble)
        Me.grpLines.Controls.Add(Me.Label8)
        Me.grpLines.Controls.Add(Me.lstLineLayers)
        Me.grpLines.Location = New System.Drawing.Point(22, 343)
        Me.grpLines.Name = "grpLines"
        Me.grpLines.Size = New System.Drawing.Size(652, 127)
        Me.grpLines.TabIndex = 38
        Me.grpLines.TabStop = False
        Me.grpLines.Text = "Lines"
        '
        'lblLineFieldSet
        '
        Me.lblLineFieldSet.AutoSize = True
        Me.lblLineFieldSet.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLineFieldSet.Location = New System.Drawing.Point(333, 40)
        Me.lblLineFieldSet.Name = "lblLineFieldSet"
        Me.lblLineFieldSet.Size = New System.Drawing.Size(37, 13)
        Me.lblLineFieldSet.TabIndex = 45
        Me.lblLineFieldSet.Text = "None"
        '
        'cmdSetLineField
        '
        Me.cmdSetLineField.Location = New System.Drawing.Point(451, 89)
        Me.cmdSetLineField.Name = "cmdSetLineField"
        Me.cmdSetLineField.Size = New System.Drawing.Size(47, 23)
        Me.cmdSetLineField.TabIndex = 44
        Me.cmdSetLineField.Text = "Set"
        Me.cmdSetLineField.UseVisualStyleBackColor = True
        '
        'label23
        '
        Me.label23.AutoSize = True
        Me.label23.Location = New System.Drawing.Point(343, 19)
        Me.label23.Name = "label23"
        Me.label23.Size = New System.Drawing.Size(51, 13)
        Me.label23.TabIndex = 43
        Me.label23.Text = "Field Set:"
        '
        'cmdAddFieldLine
        '
        Me.cmdAddFieldLine.Location = New System.Drawing.Point(569, 66)
        Me.cmdAddFieldLine.Name = "cmdAddFieldLine"
        Me.cmdAddFieldLine.Size = New System.Drawing.Size(59, 23)
        Me.cmdAddFieldLine.TabIndex = 42
        Me.cmdAddFieldLine.Text = "Add"
        Me.cmdAddFieldLine.UseVisualStyleBackColor = True
        '
        'txtNewLinefield
        '
        Me.txtNewLinefield.Location = New System.Drawing.Point(546, 40)
        Me.txtNewLinefield.Name = "txtNewLinefield"
        Me.txtNewLinefield.Size = New System.Drawing.Size(100, 20)
        Me.txtNewLinefield.TabIndex = 41
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(554, 19)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(74, 13)
        Me.Label9.TabIndex = 40
        Me.Label9.Text = "Add new field:"
        '
        'lstLineFieldsDouble
        '
        Me.lstLineFieldsDouble.FormattingEnabled = True
        Me.lstLineFieldsDouble.Location = New System.Drawing.Point(420, 40)
        Me.lstLineFieldsDouble.Name = "lstLineFieldsDouble"
        Me.lstLineFieldsDouble.Size = New System.Drawing.Size(120, 43)
        Me.lstLineFieldsDouble.TabIndex = 39
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(426, 19)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(103, 13)
        Me.Label8.TabIndex = 38
        Me.Label8.Text = "Fields (type: Double)"
        '
        'grpPolygons
        '
        Me.grpPolygons.Controls.Add(Me.cmdSetPolyField)
        Me.grpPolygons.Controls.Add(Me.lblPolyFieldSet)
        Me.grpPolygons.Controls.Add(Me.Label12)
        Me.grpPolygons.Controls.Add(Me.cmdAddFieldPoly)
        Me.grpPolygons.Controls.Add(Me.txtNewPolyField)
        Me.grpPolygons.Controls.Add(Me.Label11)
        Me.grpPolygons.Controls.Add(Me.Label10)
        Me.grpPolygons.Controls.Add(Me.lstPolyFieldsDouble)
        Me.grpPolygons.Controls.Add(Me.lstPolyLayers)
        Me.grpPolygons.Location = New System.Drawing.Point(22, 476)
        Me.grpPolygons.Name = "grpPolygons"
        Me.grpPolygons.Size = New System.Drawing.Size(646, 129)
        Me.grpPolygons.TabIndex = 39
        Me.grpPolygons.TabStop = False
        Me.grpPolygons.Text = "Polygons"
        '
        'cmdSetPolyField
        '
        Me.cmdSetPolyField.Location = New System.Drawing.Point(451, 92)
        Me.cmdSetPolyField.Name = "cmdSetPolyField"
        Me.cmdSetPolyField.Size = New System.Drawing.Size(47, 23)
        Me.cmdSetPolyField.TabIndex = 46
        Me.cmdSetPolyField.Text = "Set"
        Me.cmdSetPolyField.UseVisualStyleBackColor = True
        '
        'lblPolyFieldSet
        '
        Me.lblPolyFieldSet.AutoSize = True
        Me.lblPolyFieldSet.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPolyFieldSet.Location = New System.Drawing.Point(333, 49)
        Me.lblPolyFieldSet.Name = "lblPolyFieldSet"
        Me.lblPolyFieldSet.Size = New System.Drawing.Size(37, 13)
        Me.lblPolyFieldSet.TabIndex = 46
        Me.lblPolyFieldSet.Text = "None"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(343, 26)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(51, 13)
        Me.Label12.TabIndex = 46
        Me.Label12.Text = "Field Set:"
        '
        'cmdAddFieldPoly
        '
        Me.cmdAddFieldPoly.Location = New System.Drawing.Point(569, 68)
        Me.cmdAddFieldPoly.Name = "cmdAddFieldPoly"
        Me.cmdAddFieldPoly.Size = New System.Drawing.Size(59, 23)
        Me.cmdAddFieldPoly.TabIndex = 43
        Me.cmdAddFieldPoly.Text = "Add"
        Me.cmdAddFieldPoly.UseVisualStyleBackColor = True
        '
        'txtNewPolyField
        '
        Me.txtNewPolyField.Location = New System.Drawing.Point(540, 42)
        Me.txtNewPolyField.Name = "txtNewPolyField"
        Me.txtNewPolyField.Size = New System.Drawing.Size(100, 20)
        Me.txtNewPolyField.TabIndex = 43
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(554, 26)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(74, 13)
        Me.Label11.TabIndex = 43
        Me.Label11.Text = "Add new field:"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(418, 26)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(103, 13)
        Me.Label10.TabIndex = 43
        Me.Label10.Text = "Fields (type: Double)"
        '
        'lstPolyFieldsDouble
        '
        Me.lstPolyFieldsDouble.FormattingEnabled = True
        Me.lstPolyFieldsDouble.Location = New System.Drawing.Point(414, 42)
        Me.lstPolyFieldsDouble.Name = "lstPolyFieldsDouble"
        Me.lstPolyFieldsDouble.Size = New System.Drawing.Size(120, 43)
        Me.lstPolyFieldsDouble.TabIndex = 43
        '
        'lstPolyLayers
        '
        Me.lstPolyLayers.FormattingEnabled = True
        Me.lstPolyLayers.Location = New System.Drawing.Point(21, 21)
        Me.lstPolyLayers.Name = "lstPolyLayers"
        Me.lstPolyLayers.Size = New System.Drawing.Size(306, 94)
        Me.lstPolyLayers.TabIndex = 19
        '
        'lstBudgets
        '
        Me.lstBudgets.FormattingEnabled = True
        Me.lstBudgets.Location = New System.Drawing.Point(43, 648)
        Me.lstBudgets.Name = "lstBudgets"
        Me.lstBudgets.Size = New System.Drawing.Size(69, 108)
        Me.lstBudgets.TabIndex = 40
        '
        'grpStep6
        '
        Me.grpStep6.Controls.Add(Me.lblStep6bStatus)
        Me.grpStep6.Controls.Add(Me.lblStep6aStatus)
        Me.grpStep6.Controls.Add(Me.cmdCalcCumPerm)
        Me.grpStep6.Controls.Add(Me.lblStep6)
        Me.grpStep6.Controls.Add(Me.cmdHighlight)
        Me.grpStep6.Location = New System.Drawing.Point(239, 619)
        Me.grpStep6.Name = "grpStep6"
        Me.grpStep6.Size = New System.Drawing.Size(429, 113)
        Me.grpStep6.TabIndex = 41
        Me.grpStep6.TabStop = False
        '
        'lblStep6bStatus
        '
        Me.lblStep6bStatus.AutoSize = True
        Me.lblStep6bStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStep6bStatus.Location = New System.Drawing.Point(244, 85)
        Me.lblStep6bStatus.Name = "lblStep6bStatus"
        Me.lblStep6bStatus.Size = New System.Drawing.Size(79, 13)
        Me.lblStep6bStatus.TabIndex = 50
        Me.lblStep6bStatus.Text = "step6bstatus"
        '
        'lblStep6aStatus
        '
        Me.lblStep6aStatus.AutoSize = True
        Me.lblStep6aStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStep6aStatus.Location = New System.Drawing.Point(68, 85)
        Me.lblStep6aStatus.Name = "lblStep6aStatus"
        Me.lblStep6aStatus.Size = New System.Drawing.Size(79, 13)
        Me.lblStep6aStatus.TabIndex = 49
        Me.lblStep6aStatus.Text = "step6astatus"
        '
        'cmdCalcCumPerm
        '
        Me.cmdCalcCumPerm.Location = New System.Drawing.Point(212, 49)
        Me.cmdCalcCumPerm.Name = "cmdCalcCumPerm"
        Me.cmdCalcCumPerm.Size = New System.Drawing.Size(152, 23)
        Me.cmdCalcCumPerm.TabIndex = 25
        Me.cmdCalcCumPerm.Text = "Calculate Cumulative Perm."
        Me.cmdCalcCumPerm.UseVisualStyleBackColor = True
        '
        'lblStep6
        '
        Me.lblStep6.AutoSize = True
        Me.lblStep6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStep6.Location = New System.Drawing.Point(10, 20)
        Me.lblStep6.Name = "lblStep6"
        Me.lblStep6.Size = New System.Drawing.Size(335, 13)
        Me.lblStep6.TabIndex = 24
        Me.lblStep6.Text = "6. Highlight and Calculate Cumulative Permeability Scores"
        '
        'lblStep4Status
        '
        Me.lblStep4Status.AutoSize = True
        Me.lblStep4Status.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStep4Status.Location = New System.Drawing.Point(540, 318)
        Me.lblStep4Status.Name = "lblStep4Status"
        Me.lblStep4Status.Size = New System.Drawing.Size(72, 13)
        Me.lblStep4Status.TabIndex = 42
        Me.lblStep4Status.Text = "step4status"
        '
        'lblStep5Status
        '
        Me.lblStep5Status.AutoSize = True
        Me.lblStep5Status.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStep5Status.Location = New System.Drawing.Point(130, 648)
        Me.lblStep5Status.Name = "lblStep5Status"
        Me.lblStep5Status.Size = New System.Drawing.Size(72, 13)
        Me.lblStep5Status.TabIndex = 43
        Me.lblStep5Status.Text = "step5status"
        '
        'lblStep1Warning
        '
        Me.lblStep1Warning.AutoSize = True
        Me.lblStep1Warning.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStep1Warning.ForeColor = System.Drawing.Color.DarkRed
        Me.lblStep1Warning.Location = New System.Drawing.Point(540, 181)
        Me.lblStep1Warning.Name = "lblStep1Warning"
        Me.lblStep1Warning.Size = New System.Drawing.Size(82, 13)
        Me.lblStep1Warning.TabIndex = 44
        Me.lblStep1Warning.Text = "step1warning"
        '
        'lblStep2warning
        '
        Me.lblStep2warning.AutoSize = True
        Me.lblStep2warning.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStep2warning.ForeColor = System.Drawing.Color.DarkRed
        Me.lblStep2warning.Location = New System.Drawing.Point(540, 233)
        Me.lblStep2warning.Name = "lblStep2warning"
        Me.lblStep2warning.Size = New System.Drawing.Size(82, 13)
        Me.lblStep2warning.TabIndex = 45
        Me.lblStep2warning.Text = "step2warning"
        '
        'lblStep3Warning
        '
        Me.lblStep3Warning.AutoSize = True
        Me.lblStep3Warning.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStep3Warning.ForeColor = System.Drawing.Color.DarkRed
        Me.lblStep3Warning.Location = New System.Drawing.Point(540, 289)
        Me.lblStep3Warning.Name = "lblStep3Warning"
        Me.lblStep3Warning.Size = New System.Drawing.Size(82, 13)
        Me.lblStep3Warning.TabIndex = 46
        Me.lblStep3Warning.Text = "step3warning"
        '
        'lblStep4Warning
        '
        Me.lblStep4Warning.AutoSize = True
        Me.lblStep4Warning.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStep4Warning.ForeColor = System.Drawing.Color.DarkRed
        Me.lblStep4Warning.Location = New System.Drawing.Point(448, 318)
        Me.lblStep4Warning.Name = "lblStep4Warning"
        Me.lblStep4Warning.Size = New System.Drawing.Size(82, 13)
        Me.lblStep4Warning.TabIndex = 47
        Me.lblStep4Warning.Text = "step4warning"
        '
        'lblStep5Warning
        '
        Me.lblStep5Warning.AutoSize = True
        Me.lblStep5Warning.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStep5Warning.ForeColor = System.Drawing.Color.DarkRed
        Me.lblStep5Warning.Location = New System.Drawing.Point(130, 668)
        Me.lblStep5Warning.Name = "lblStep5Warning"
        Me.lblStep5Warning.Size = New System.Drawing.Size(82, 13)
        Me.lblStep5Warning.TabIndex = 48
        Me.lblStep5Warning.Text = "step5warning"
        '
        'frmVisualizeDecisionsAndNet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(686, 802)
        Me.Controls.Add(Me.lblStep5Warning)
        Me.Controls.Add(Me.lblStep4Warning)
        Me.Controls.Add(Me.lblStep3Warning)
        Me.Controls.Add(Me.lblStep2warning)
        Me.Controls.Add(Me.lblStep1Warning)
        Me.Controls.Add(Me.lblStep5Status)
        Me.Controls.Add(Me.lblStep4Status)
        Me.Controls.Add(Me.grpStep6)
        Me.Controls.Add(Me.lstBudgets)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.grpPolygons)
        Me.Controls.Add(Me.grpLines)
        Me.Controls.Add(Me.lblStep3Status)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtConnectivityTab)
        Me.Controls.Add(Me.lblStep2Status)
        Me.Controls.Add(Me.cmdBrowseConnectivityTab)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.lblStep1Status)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtDecisionsTab)
        Me.Controls.Add(Me.cmdBrowseDecisionsTab)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtResultsTab)
        Me.Controls.Add(Me.cmdBrowseResultsTab)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Name = "frmVisualizeDecisionsAndNet"
        Me.Text = "Visualize Decisions And Network Accessibility"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpLines.ResumeLayout(False)
        Me.grpLines.PerformLayout()
        Me.grpPolygons.ResumeLayout(False)
        Me.grpPolygons.PerformLayout()
        Me.grpStep6.ResumeLayout(False)
        Me.grpStep6.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmdBrowseResultsTab As System.Windows.Forms.Button
    Friend WithEvents cmdHighlight As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents txtResultsTab As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdBrowseDecisionsTab As System.Windows.Forms.Button
    Friend WithEvents txtDecisionsTab As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lblStep1Status As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmdBrowseConnectivityTab As System.Windows.Forms.Button
    Friend WithEvents lblStep2Status As System.Windows.Forms.Label
    Friend WithEvents txtConnectivityTab As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lblStep3Status As System.Windows.Forms.Label
    Friend WithEvents lstLineLayers As System.Windows.Forms.CheckedListBox
    Friend WithEvents grpLines As System.Windows.Forms.GroupBox
    Friend WithEvents cmdAddFieldLine As System.Windows.Forms.Button
    Friend WithEvents txtNewLinefield As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents lstLineFieldsDouble As System.Windows.Forms.ListBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents grpPolygons As System.Windows.Forms.GroupBox
    Friend WithEvents cmdAddFieldPoly As System.Windows.Forms.Button
    Friend WithEvents txtNewPolyField As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents lstPolyFieldsDouble As System.Windows.Forms.ListBox
    Friend WithEvents lstPolyLayers As System.Windows.Forms.CheckedListBox
    Friend WithEvents lblLineFieldSet As System.Windows.Forms.Label
    Friend WithEvents cmdSetLineField As System.Windows.Forms.Button
    Friend WithEvents label23 As System.Windows.Forms.Label
    Friend WithEvents cmdSetPolyField As System.Windows.Forms.Button
    Friend WithEvents lblPolyFieldSet As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents lstBudgets As System.Windows.Forms.ListBox
    Friend WithEvents grpStep6 As System.Windows.Forms.GroupBox
    Friend WithEvents lblStep6 As System.Windows.Forms.Label
    Friend WithEvents lblStep4Status As System.Windows.Forms.Label
    Friend WithEvents cmdCalcCumPerm As System.Windows.Forms.Button
    Friend WithEvents lblStep5Status As System.Windows.Forms.Label
    Friend WithEvents lblStep6bStatus As System.Windows.Forms.Label
    Friend WithEvents lblStep6aStatus As System.Windows.Forms.Label
    Friend WithEvents lblStep1Warning As System.Windows.Forms.Label
    Friend WithEvents lblStep2warning As System.Windows.Forms.Label
    Friend WithEvents lblStep3Warning As System.Windows.Forms.Label
    Friend WithEvents lblStep4Warning As System.Windows.Forms.Label
    Friend WithEvents lblStep5Warning As System.Windows.Forms.Label
End Class
