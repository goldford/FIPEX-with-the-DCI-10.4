
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.NetworkAnalysis
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Catalog
Imports System.Windows.Forms
Imports System.IO ' For use with file name checking in DCI section
'Imports Microsoft.VisualStudio.Shell.Interop

' the following statements were added for the 
' checkperm class
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Security.Principal
Imports System                     ' Added Jan 16, 2012 - following instructions here:
' http://msdn.microsoft.com/en-us/library/6ka1wd3w(v=VS.85).aspx
'Imports System.Web

Public Class frmRunAdvancedAnalysis

    Private m_app As ESRI.ArcGIS.Framework.IApplication
    Private pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Private m_FiPEx As FishPassageExtension
    Private m_UtilityNetworkAnalysisExt As ESRI.ArcGIS.EditorExt.IUtilityNetworkAnalysisExt

    Private helpprovider1 As Help
    ' global variable for feature class selected (used in exclusions)
    Private m_pFeatureClass As IFeatureClass

    ' used to show / hide run button ('advancedanalysis' or 'optionsmenu')
    Public m_FIPEXOptionsContext = "advancedanalysis"

    ' ALL THESE GLOBAL VARIABLES SHOULD BE ELIMINATED
    ' AND MOVED TO LOCAL VARIABLES
    Private m_iPolysCount As Integer = 0      ' number of polygon layers currently using
    Private m_iLinesCount As Integer = 0      ' number of lines layers currently using
    Private m_iBarrierCount As Integer = 0    ' number of barriers currently using
    Private m_iExclusions As Integer = 0      ' number of exclusions
    Private m_sPolyLayer, m_sLineLayer, m_sBarrierIDLayer As String
    Private m_sLayerType As String

    'Private _HabParam As String
    Private m_LLayersFields As List(Of LineLayerToAdd) = New List(Of LineLayerToAdd)
    Private m_PLayersFields As List(Of PolyLayerToAdd) = New List(Of PolyLayerToAdd)
    Private m_lExclusions As List(Of LayerToExclude) = New List(Of LayerToExclude)
    ' variable to hold list of barrier layers and ID fields (if set)
    Private m_lBarrierIDs As List(Of BarrierIDObj) = New List(Of BarrierIDObj)

    ' These single object persistent (module level) variables are used
    ' to pass and retrieve data from the pop-out forms
    Private m_LLayerToAdd As LineLayerToAdd
    Private m_PLayerToAdd As PolyLayerToAdd
    Private m_BarrierIDObj As BarrierIDObj

    Public m_bRun As Boolean = False
    ' these two variables monitor where user clicks/selects - used to open help docs
    'Private selectionContext As IVsUserContext
    'Private trackSelection As ITrackSelection


    Public Function Form_Initialize(ByVal m_application As ESRI.ArcGIS.Framework.IApplication) As Boolean

        'Dim pMap As ESRI.ArcGIS.Carto.IMap
        'Dim pLayer As ESRI.ArcGIS.Carto.ILayer
        'Dim pFLayer As ESRI.ArcGIS.Carto.IFeatureLayer2
        'Dim pGeoDataset As ESRI.ArcGIS.Geodatabase.IGeoDataset

        Try
            m_app = m_application
            pMxDoc = CType(m_app.Document, ESRI.ArcGIS.ArcMapUI.IMxDocument)
            m_FiPEx = FishPassageExtension.GetExtension

            ' Obtain a reference to Utility Network Analysis Ext in the current Doc
            'Dim pUID As New ESRI.ArcGIS.esriSystem.UID
            'pUID.Value = "{98528F9B-B971-11D2-BABD-00C04FA33C20}"

            'Dim pExtension As ESRI.ArcGIS.esriSystem.IExtension = m_application.FindExtensionByCLSID(pUID)
            'm_UtilityNetworkAnalysisExt = CType(pExtension, IUtilityNetworkAnalysisExt)
            m_UtilityNetworkAnalysisExt = FishPassageExtension.GetUNAExt


            ' Obtain a reference to the DFOBarriersAnalysis Extension
            'Dim pUID2 As New ESRI.ArcGIS.esriSystem.UID
            'pUID2.Value = "FiPEx.DFOBarriersAnalysisExtension"

            'pExtension = m_application.FindExtensionByCLSID(pUID2)

            'm_DFOExt = CType(pExtension, FiPEx.DFOBarriersAnalysisExtension)

            Return True

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.ToString, "Initialize Form")
            Return False
        End Try
    End Function
    Private Sub saveoptions(ByVal canceleverything As Boolean)
        Dim sGDB As String
        Dim bDBF As Boolean
        Dim bUpHab, bTotalUpHab, bDownHab, bTotalDownHab, bPathDownHab, bTotalPathDownHab As Boolean
        Dim sPrefix As String
        Dim sDirection As String
        Dim bConnectTab, bBarrierPerm, bNaturalYN As Boolean
        Dim bDCISectional, bDCI, bUseHabLength, bUseHabArea As Boolean

        '2020
        Dim bAdvConnectTab As Boolean = False
        Dim bDistanceLim As Boolean = False
        Dim dMaxDist As Double = 0.0
        Dim bDistanceDecay As Boolean = False
        Dim sDDFunction As String = "none"

        Dim bMaximum As Boolean
        Dim iOrderNum As Integer
        Dim bGLPKTables As Boolean
        Dim sGLPKModelDir As String
        Dim sGnuWinDir As String

        Dim i As Integer = 0

        ' Set the analysis direction
        If OptUp.Checked = True Then
            sDirection = "up"
        Else
            sDirection = "down"
        End If

        ' Set the Order

        Dim sOrderNum As String = TxtOrder.Text
        If Not Integer.TryParse(sOrderNum, iOrderNum) Then
            MessageBox.Show("Number in 'Order' Textbox must be an integer. Setting OrderNum to maximum (entire network).")
            iOrderNum = 10000
        Else
            iOrderNum = Convert.ToInt32(TxtOrder.Text)
        End If

        ' Set the Maximum
        If ChkMaxOrd.Checked = True Then
            bMaximum = True
        Else
            bMaximum = False
        End If

        ' Set Connectivity Table output
        If chkConnect.Checked = True Then
            bConnectTab = True
        Else
            bConnectTab = False
        End If


        If chkBarrierPerm.Checked = True Then
            bBarrierPerm = True
        Else
            bBarrierPerm = False
        End If

        If chkNaturalTF.Checked = True Then
            bNaturalYN = True
        Else
            bNaturalYN = False
        End If

        ' Set DBF output parameters
        If ChkDBFOutput.Checked = True And m_lBarrierIDs.Count > 0 Then
            bDBF = True
            sGDB = txtGDB.Text
            sPrefix = txtTablesPrefix.Text

            ' halt the save if there is no GDB or prefix set
            If sGDB = "" Or sGDB = "n/a" Then
                MessageBox.Show("Please select a Geodatabase for table output", "Missing Parameter:", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                Exit Sub
            ElseIf sPrefix = "" Or sPrefix = "n/a" Then
                MessageBox.Show("Please select a prefix for table(s) output", "Missing Parameter:", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                Exit Sub
            End If
        Else
            bDBF = False
            sGDB = "n/a"
            sPrefix = "n/a"
        End If

        Dim sDCIModelDir As String = txtDCIModelDir.Text
        Dim sRInstallDir As String = txtRInstallDir.Text

        ' If DCI Output box is checked and enabled, and bDBF is true 
        ' then confirm DCI output
        If chkDCI.Checked = True And bDBF = True And chkDCI.Enabled = True Then
            If sDCIModelDir Is Nothing Or sDCIModelDir = "" Or sDCIModelDir = "n/a" Then
                MessageBox.Show("Please browse to the directory the DCI Model files are installed in ('Advanced' Tab) or uncheck the 'calculate DCI' box.", "Missing Parameter:", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                Exit Sub

            Else ' check if there are permissions to the directory
                'Dim pc As New CheckPerm
                'pc.Permission = "Modify"
                'If Not pc.CheckPerm(sDCIModelDir) Then
                '    MsgBox("You do not have the necessary permission to create, delete, or modify files in the DCI Model installation directory." + _
                '    " Please change directory, attain necessary permissions, or uncheck DCI Output checkbox in 'Advanced Tab.'")
                '    Exit Sub
                'End If

                ' Check that the user currently has file permissions to write to 
                ' this directory
                Dim bPermissionCheck
                bPermissionCheck = FileWriteDeleteCheck(sDCIModelDir)
                If bPermissionCheck = False Then
                    MsgBox("File / folder permission check: " & Str(bPermissionCheck))
                    MsgBox("It appears you do not have write permission to the DCI Model Directory.  Write permission to this directory is needed in order to run DCI Analysis.")
                    txtDCIModelDir.Text = ""
                    Exit Sub
                End If


                ' check if the DCI model files are present in the selected directory
                Dim fFile As New FileInfo(sDCIModelDir + "/run.r")
                If Not fFile.Exists Then
                    MessageBox.Show("DCI Directory does not contain required 'run.r' file", "DCI Model File Missing", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                    Exit Sub
                End If

                bDCI = True
                If chkDCISectional.Checked = True Then
                    bDCISectional = True
                Else
                    bDCISectional = False
                End If

            End If
        Else
            bDCI = False
            bDCISectional = False
            sDCIModelDir = "n/a"
        End If

        ' If DCI Output box is checked and enabled, and bDBF is true 
        ' then confirm 'R' program directory is correct
        If chkDCI.Checked = True And bDBF = True And chkDCI.Enabled = True Then
            If sRInstallDir Is Nothing Or sRInstallDir = "" Or sRInstallDir = "n/a" Then
                MessageBox.Show("Please browse to the 'R' stats program installation directory ('Advanced' Tab)", "Missing Parameter:", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
                Exit Sub

            Else
                ' check if the R program files are present in the selected directory
                Dim fFile As New FileInfo(sRInstallDir + "/bin/rterm.exe")
                If Not fFile.Exists Then
                    Dim fFile2 As New FileInfo(sRInstallDir + "/bin/i386/rterm.exe")
                    If Not fFile2.Exists Then
                        MessageBox.Show("'R' Program files not found in R installation directory provided ('Advanced Tab').  Please browse to the correct directory or install R.", "'R' Program Files Missing", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                        Exit Sub
                    End If
                End If

                bDCI = True
            End If
        Else
            bDCI = False
            sRInstallDir = "n/a"
        End If

        '2020 
        If rdoHabLength.Checked = True Then
            bUseHabLength = True
            bUseHabArea = False
        ElseIf rdoHabArea.Checked = True Then
            bUseHabLength = False
            bUseHabArea = True
        Else
            bUseHabLength = False
            bUseHabArea = False
        End If

        '2020
        'If chkAdvConnect.Checked = True Then
        'bAdvConnectTab = True
        'Else
        'bAdvConnectTab = False
        'End If
        ' 2020 - user no longer selects this checkbox
        '      - instead 'advanced' analysis turned on if 'apply distance limit' is checked

        If chkDistanceLimit.Checked = True Then
            bDistanceLim = True
            bAdvConnectTab = True
        Else
            bDistanceLim = False
            bAdvConnectTab = False
        End If

        ' iOrderNum = Convert.ToInt32(TxtOrder.Text)

        '2020
        Dim sMaxDist = txtMaxDistance.Text
        If Double.TryParse(sMaxDist, dMaxDist) = True Then
            dMaxDist = Convert.ToDouble(sMaxDist)
        Else
            dMaxDist = 0.0
        End If

        If bDistanceLim = True And dMaxDist = 0 Then
            MsgBox("The maximum / cut-off distance must be set > 0 to continue if the distance limit is on. Please check and re-save / re-run.")
            canceleverything = True
            Exit Sub
        End If
        'MsgBox("dMaxDist after conversionto double: " & dMaxDist)

        'Dim myInt As Integer = 0
        'If Not Integer.TryParse(MyString, myInt) Then
        'MessageBox.Show("This is not an integer")
        'End If

        If chkDistanceDecay.Checked = True Then
            bDistanceDecay = True
            If rdoLinear.Checked = True Then
                sDDFunction = "linear"
            ElseIf rdoNatExp1.Checked = True Then
                sDDFunction = "natexp1"
            ElseIf rdoCircle.Checked = True Then
                sDDFunction = "circle"
            ElseIf rdoSigmoid.Checked = True Then
                sDDFunction = "sigmoid"
            Else
                sDDFunction = "none"
            End If
        Else
            bDistanceDecay = False
            sDDFunction = "none"
        End If

        ' Set habitat statistics parameters
        If chkUpHab.Checked = True Then
            bUpHab = True
        Else
            bUpHab = False
        End If

        If chkTotalUpHab.Checked = True Then
            bTotalUpHab = True
        Else
            bTotalUpHab = False
        End If

        If chkDownHab.Checked = True Then
            bDownHab = True
        Else
            bDownHab = False
        End If

        If chkTotalDownHab.Checked = True Then
            bTotalDownHab = True
        Else
            bTotalDownHab = False
        End If

        If chkPathDownHab.Checked = True Then
            bPathDownHab = True
        Else
            bPathDownHab = False
        End If

        If chkTotalPathDownHab.Checked = True Then
            bTotalPathDownHab = True
        Else
            bTotalPathDownHab = False
        End If

        ' BETTER IN FORM_LOAD?
        ' 1.0 For each of the layers in the global list
        '   2.0 If it matches a layer in the listbox
        '     3.0a Leave it be
        '   2.1 Else if it doesn't match a layer in the listbox
        '     3.0b Remove that object from the list
        Dim m As Integer
        Dim bMatch As Boolean
        For i = m_LLayersFields.Count - 1 To 0 Step -1
            bMatch = False
            For m = 0 To lstLineLayers.Items.Count - 1
                If m_LLayersFields.Item(i).Layer = lstLineLayers.Items.Item(m).ToString Then
                    bMatch = True
                End If
            Next
            If bMatch = False Then
                m_LLayersFields.RemoveAt(i)
            End If
        Next
        For i = m_PLayersFields.Count - 1 To 0 Step -1
            bMatch = False
            For m = 0 To lstPolyLayers.Items.Count - 1
                If m_PLayersFields.Item(i).Layer = lstPolyLayers.Items.Item(m).ToString Then
                    bMatch = True
                End If
            Next
            If bMatch = False Then
                m_PLayersFields.RemoveAt(i)
            End If
        Next
        For i = m_lBarrierIDs.Count - 1 To 0 Step -1
            bMatch = False
            For m = 0 To chkLstBarriersLayers.Items.Count - 1
                If m_lBarrierIDs.Item(i).Layer = chkLstBarriersLayers.Items.Item(m).ToString Then
                    bMatch = True
                End If
            Next
            If bMatch = False Then
                m_lBarrierIDs.RemoveAt(i)
            End If
        Next

        bGLPKTables = False
        sGLPKModelDir = "not set"
        sGnuWinDir = "not set"

        m_FiPEx.pPropset.SetProperty("direction", sDirection)
        m_FiPEx.pPropset.SetProperty("ordernum", iOrderNum)
        m_FiPEx.pPropset.SetProperty("maximum", bMaximum)
        m_FiPEx.pPropset.SetProperty("connecttab", bConnectTab)
        m_FiPEx.pPropset.SetProperty("advconnecttab", bAdvConnectTab)
        m_FiPEx.pPropset.SetProperty("barrierperm", bBarrierPerm)
        m_FiPEx.pPropset.SetProperty("naturalyn", bNaturalYN)
        m_FiPEx.pPropset.SetProperty("dciyn", bDCI)
        m_FiPEx.pPropset.SetProperty("dcisectionalyn", bDCISectional)
        m_FiPEx.pPropset.SetProperty("sRInstallDir", sRInstallDir)
        m_FiPEx.pPropset.SetProperty("sDCIModelDir", sDCIModelDir)

        '2020
        m_FiPEx.pPropset.SetProperty("bUseHabLength", bUseHabLength)
        m_FiPEx.pPropset.SetProperty("bUseHabArea", bUseHabArea)
        m_FiPEx.pPropset.SetProperty("bDistanceLim", bDistanceLim)
        m_FiPEx.pPropset.SetProperty("dMaxDist", dMaxDist)
        m_FiPEx.pPropset.SetProperty("bDistanceDecay", bDistanceDecay)
        m_FiPEx.pPropset.SetProperty("sDDFunction", sDDFunction)

        m_FiPEx.pPropset.SetProperty("bDBF", bDBF)
        m_FiPEx.pPropset.SetProperty("sGDB", sGDB)
        m_FiPEx.pPropset.SetProperty("TabPrefix", sPrefix)

        m_FiPEx.pPropset.SetProperty("UpHab", bUpHab)
        m_FiPEx.pPropset.SetProperty("TotalUpHab", bTotalUpHab)
        m_FiPEx.pPropset.SetProperty("DownHab", bDownHab)
        m_FiPEx.pPropset.SetProperty("TotalDownHab", bTotalDownHab)
        m_FiPEx.pPropset.SetProperty("PathDownHab", bPathDownHab)
        m_FiPEx.pPropset.SetProperty("TotalPathDownHab", bTotalPathDownHab)

        m_FiPEx.pPropset.SetProperty("bGLPKTables", bGLPKTables)
        m_FiPEx.pPropset.SetProperty("sGLPKModelDir", sGLPKModelDir)
        m_FiPEx.pPropset.SetProperty("sGnuWinDir", sGnuWinDir)

        m_iPolysCount = m_PLayersFields.Count
        m_iLinesCount = m_LLayersFields.Count
        m_iExclusions = m_lExclusions.Count
        m_iBarrierCount = m_lBarrierIDs.Count

        ' Putting a check here to be sure that the counts correspond to the 
        ' number of checked items in the list... in debugging this happened
        ' and it could not be gotten rid of.  
        If lstLineLayers.CheckedItems.Count <> m_iLinesCount Then
            System.Windows.Forms.MessageBox.Show("There is a problem matching the number of items checked in the lines listbox to the number of line layers in the global variable.  Please try again.", "Global Variable Mismatch")
            m_LLayersFields.Clear()
            m_iLinesCount = m_LLayersFields.Count
        End If
        If lstPolyLayers.CheckedItems.Count <> m_iPolysCount Then
            System.Windows.Forms.MessageBox.Show("There is a problem matching the number of items checked in the polys listbox to the number of polys layers in the global variable.  Please try again.", "Global Variable Mismatch")
            m_PLayersFields.Clear()
            m_iPolysCount = m_PLayersFields.Count
        End If

        m_FiPEx.pPropset.SetProperty("numPolys", m_iPolysCount)

        ' variables to hold property names
        Dim sIncLayer, sLengthField, sLengthUnits, sHabClassField, sHabQuanField, sHabUnits As String

        If m_iPolysCount > 0 Then
            For i = 0 To m_iPolysCount - 1
                sIncLayer = "IncPoly" + i.ToString
                sHabClassField = "PolyClassField" + i.ToString
                sHabQuanField = "PolyQuanField" + i.ToString
                sHabUnits = "PolyUnitField" + i.ToString
                m_FiPEx.pPropset.SetProperty(sIncLayer, m_PLayersFields.Item(i).Layer)
                m_FiPEx.pPropset.SetProperty(sHabClassField, m_PLayersFields.Item(i).HabClsField)
                m_FiPEx.pPropset.SetProperty(sHabQuanField, m_PLayersFields.Item(i).HabQuanField)
                m_FiPEx.pPropset.SetProperty(sHabUnits, m_PLayersFields.Item(i).HabUnitField)
            Next
        End If

        m_FiPEx.pPropset.SetProperty("numLines", m_iLinesCount)

        If m_iLinesCount > 0 Then
            For i = 0 To m_iLinesCount - 1
                sIncLayer = "IncLine" + i.ToString
                sLengthField = "LineLengthField" + i.ToString
                sLengthUnits = "LineLengthUnits" + i.ToString
                sHabQuanField = "LineHabQuanField" + i.ToString
                sHabUnits = "LineHabUnits" + i.ToString
                m_FiPEx.pPropset.SetProperty(sIncLayer, m_LLayersFields.Item(i).Layer)

                ' double check the new property names for lines, fix issues
                m_FiPEx.pPropset.SetProperty(sLengthField, m_LLayersFields.Item(i).LengthField)
                m_FiPEx.pPropset.SetProperty(sLengthUnits, m_LLayersFields.Item(i).LengthUnits)
                m_FiPEx.pPropset.SetProperty(sHabQuanField, m_LLayersFields.Item(i).HabQuanField)
                m_FiPEx.pPropset.SetProperty("LineHabClassField" + i.ToString, m_LLayersFields.Item(i).HabClsField)
                m_FiPEx.pPropset.SetProperty(sHabUnits, m_LLayersFields.Item(i).HabUnits)
            Next
        End If

        m_FiPEx.pPropset.SetProperty("numExclusions", m_iExclusions)

        Dim sExcldLayer, sExcldFeature, sExcldValue As String

        If m_iExclusions > 0 Then
            For i = 0 To m_iExclusions - 1
                sExcldLayer = "ExcldLayer" + i.ToString
                sExcldFeature = "ExcldFeature" + i.ToString
                sExcldValue = "ExcldValue" + i.ToString
                m_FiPEx.pPropset.SetProperty(sExcldLayer, m_lExclusions.Item(i).Layer)
                m_FiPEx.pPropset.SetProperty(sExcldFeature, m_lExclusions.Item(i).Feature)
                m_FiPEx.pPropset.SetProperty(sExcldValue, m_lExclusions.Item(i).Value)
            Next
        End If

        m_FiPEx.pPropset.SetProperty("numBarrierIDs", m_iBarrierCount)

        ' This tool always requires a barriers layer to be selected if tables are to be output so 
        ' exit the save process if one isn't selected and if the user has selected DBF output (and the option is active).
        If m_iBarrierCount = 0 And bDBF = True And ChkDBFOutput.Enabled = True Then
            MessageBox.Show("FIPEX requires a barriers layer to output DBF Tables.  Please select one in the listbox in the 'Barriers' Tab.", "Missing Parameter:", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)
            Exit Sub
        End If


        Dim sBarrierIDLayer, sBarrierIDField, sBarrierPermField, sBarrierNaturalYNField As String
        If m_iBarrierCount > 0 Then
            For i = 0 To m_iBarrierCount - 1
                sBarrierIDLayer = "BarrierIDLayer" + i.ToString
                sBarrierIDField = "BarrierIDField" + i.ToString
                sBarrierPermField = "BarrierPermField" + i.ToString
                sBarrierNaturalYNField = "BarrierNaturalYNField" + i.ToString
                m_FiPEx.pPropset.SetProperty(sBarrierIDLayer, m_lBarrierIDs.Item(i).Layer)
                m_FiPEx.pPropset.SetProperty(sBarrierIDField, m_lBarrierIDs.Item(i).Field)


                ' When a barriers layer is checked in the lstbox a barrier object is created and
                ' all perm and naturalyn fields set to 'not set' by default.  Also, there is a <none>
                ' option in the popup form.  
                ' Users are allowed to set to <none> - this setting allows room for a default/assumption to 
                ' be made - i.e., if <none> is set for barrier permeability then assume that all barriers are
                ' impermeable.  If <none> is set for barrier natural TF then assume all barriers are manmade.  
                If bDCI = True Then
                    If m_lBarrierIDs.Item(i).PermField = "Not set" Then
                        MessageBox.Show("You have not selected a barrier permeability field to use in DCI analysis in the 'Barriers Tab'." _
                        + "FIPEX will assume all barriers are 100% impassable (permeability value = 0).", "Missing Parameter:", MessageBoxButtons.OK, _
                        MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)

                        m_lBarrierIDs.Item(i).PermField = "<None>"
                    End If
                    m_FiPEx.pPropset.SetProperty(sBarrierPermField, m_lBarrierIDs.Item(i).PermField)

                    If m_lBarrierIDs.Item(i).NaturalYNField = "Not set" Then
                        MessageBox.Show("You have not selected a barrier Natural TF field to use in DCI analysis in the 'Barriers Tab'." _
                        + "FIPEX will assume all barriers are manmade ('Natural TF' value = F).", "Missing Parameter:", MessageBoxButtons.OK, _
                        MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, False)

                        m_lBarrierIDs.Item(i).NaturalYNField = "<None>"
                    End If
                    m_FiPEx.pPropset.SetProperty(sBarrierNaturalYNField, m_lBarrierIDs.Item(i).NaturalYNField)
                Else
                    m_FiPEx.pPropset.SetProperty(sBarrierPermField, m_lBarrierIDs.Item(i).PermField)
                    m_FiPEx.pPropset.SetProperty(sBarrierNaturalYNField, m_lBarrierIDs.Item(i).NaturalYNField)
                End If

            Next
        End If

        m_FiPEx.m_bLoaded = True ' Set the bLoaded variable true so next open of form will load the saved values

        ' save settings closes form if user opens from options menu
        If m_FIPEXOptionsContext = "optionsmenu" Then
            Me.Close()
        End If


    End Sub
    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim canceleverything As Boolean = False
        saveoptions(canceleverything)

    End Sub
    Private Sub Options_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' Loads settings from extension
        ' populates form
        '
        ' NOTE: YOU MUST HAVE TAB SELECTED TO SET AN ITEM 
        ' IN A LISTBOX AS SELECTED (PROPERLY)
        '

        ' 2020 if user opens from Options Menu then hide the 'run' button
        If m_FIPEXOptionsContext = "optionsmenu" Then
            cmdRun.Enabled = False
            cmdRun.Visible = False
        End If

        Dim k As Integer = 0                ' loop counter
        Dim j As Integer = 0                ' loop counter
        Dim m As Integer = 0                ' another loop counter
        Dim sPolyLayer As String            ' a polygon layer saved in stream
        Dim sLineLayer As String            ' a line layer saved in stream
        Dim sExcldLayer As String

        ' Variables to hold extension settings
        Dim sGDB As String = "n/a"
        Dim bDBF As Boolean = False

        Dim bUpHab, bTotalUpHab, bDownHab, bTotalDownHab, bPathDownHab, bTotalPathDownHab As Boolean

        Dim sPrefix As String = "n/a"
        Dim sDirection As String = "up"
        Dim bConnectTab As Boolean = False
        Dim bAdvConnectTab As Boolean = False
        Dim bBarrierPerm As Boolean = False
        Dim bNaturalTF As Boolean = False
        Dim bDCI As Boolean = False
        Dim bDCISectional As Boolean = False
        Dim sRInstallDir As String = "n/a"
        Dim sDCIModelDir As String = "n/a"

        '2020
        Dim bUseHabLength As Boolean = True
        Dim bUseHabArea As Boolean = False
        Dim bDistanceLim As Boolean = False
        Dim bDistanceDecay As Boolean = False
        Dim dMaxDist As Double = 0.0
        Dim sDDFunction As String = "none"

        Dim bMaximum As Boolean = False
        Dim iOrderNum As Integer = 1

        Dim bGLPKTables As Boolean
        Dim sGLPKModelDir As String
        Dim sGnuWinDir As String

        Dim pMxDoc As IMxDocument
        Dim pMap As IMap
        Dim pDoc As IDocument = m_app.Document

        ' hook into ArcMap
        pMxDoc = CType(pDoc, IMxDocument)
        pMap = pMxDoc.FocusMap

        Dim pFeatureLayer As IFeatureLayer
        Dim i As Integer

        ' Reset the global variables (they don't get cleared after form is closed)
        m_LLayersFields = New List(Of LineLayerToAdd)
        m_PLayersFields = New List(Of PolyLayerToAdd)
        m_lExclusions = New List(Of LayerToExclude)
        m_lBarrierIDs = New List(Of BarrierIDObj)

        ' Default to the first tab
        TabBarriers.SelectTab(0)

        ' default 'change...' buttons to not selectable
        cmdChngPolyCls.Enabled = False
        cmdChngLineCls.Enabled = False
        cmdBarrierID.Enabled = False
        cmdSelectBarrierPerm.Enabled = False
        cmdSelectNaturalTF.Enabled = False

        ' default select DCI directory buttons to disabled
        cmdDCIModelDir.Enabled = False
        txtDCIModelDir.Enabled = False
        cmdRInstallDir.Enabled = False
        txtRInstallDir.Enabled = False
        chkDCISectional.Enabled = False

        '2020
        chkConnect.Enabled = False
        chkAdvConnect.Enabled = False
        chkDistanceLimit.Enabled = False
        chkDistanceDecay.Enabled = False
        txtMaxDistance.Enabled = False
        rdoHabArea.Enabled = False
        rdoHabLength.Enabled = False
        rdoLinear.Enabled = False
        rdoNatExp1.Enabled = False
        rdoCircle.Enabled = False
        rdoSigmoid.Enabled = False
        ' lstLineHabQuan.Enabled = False
        ' lstLineHabUnits.Enabled = False


        ' obtain reference to current geometric network
        ' and get all participating point feature classes so only
        ' those are included in the barriers list. 
        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        pNetworkAnalysisExt = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExt)
        Dim pGeometricNetwork As IGeometricNetwork
        pGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        Dim pEnumSimJuncFCs As ESRI.ArcGIS.Geodatabase.IEnumFeatureClass = pGeometricNetwork.ClassesByType(esriFeatureType.esriFTSimpleJunction)
        Dim pEnumComJuncFCs As ESRI.ArcGIS.Geodatabase.IEnumFeatureClass = pGeometricNetwork.ClassesByType(esriFeatureType.esriFTComplexJunction)
        Dim pFeatureClass As IFeatureClass
        Dim bJuncMatch As Boolean

        ' load all layers in TOC to the lstbox
        For i = 0 To pMap.LayerCount - 1
            If pMap.Layer(i).Valid = True Then
                If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                    Dim pLayer As ILayer = pMap.Layer(i)
                    pFeatureLayer = CType(pLayer, IFeatureLayer)
                    If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
                    Or pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Then
                        lstLineLayers.Items.Add(pMap.Layer(i).Name)
                    ElseIf pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                        lstPolyLayers.Items.Add(pMap.Layer(i).Name)
                    ElseIf pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint _
                    Or pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint Then

                        ' Can't use the code below because if there are multiple networks loaded this 
                        ' would cause conflict
                        'If pFeatureLayer.FeatureClass.FeatureType = esriFeatureType.esriFTComplexJunction _
                        'Or pFeatureLayer.FeatureClass.FeatureType = esriFeatureType.esriFTSimpleJunction Then
                        'End If

                        ' This point should be participating in the network
                        bJuncMatch = False
                        'Initial set of feature class
                        pEnumSimJuncFCs.Reset()
                        pFeatureClass = pEnumSimJuncFCs.Next
                        Do Until pFeatureClass Is Nothing
                            ' If there's a match then flag it to add to the listbox
                            If pFeatureLayer.FeatureClass.FeatureClassID = pFeatureClass.FeatureClassID Then
                                bJuncMatch = True
                            End If
                            pFeatureClass = pEnumSimJuncFCs.Next
                        Loop

                        pEnumComJuncFCs.Reset()
                        pFeatureClass = pEnumComJuncFCs.Next
                        Do Until pFeatureClass Is Nothing
                            If pFeatureLayer.FeatureClass.FeatureClassID = pFeatureClass.FeatureClassID Then
                                bJuncMatch = True
                            End If
                            pFeatureClass = pEnumComJuncFCs.Next
                        Loop

                        ' If there was a match between the FC and the junction FCs in the network
                        ' then add it to the listbox in the options menu
                        If bJuncMatch = True Then
                            chkLstBarriersLayers.Items.Add(pMap.Layer(i).Name)
                        End If
                    End If
                End If ' LayerType Check
            End If
        Next

        ' Read the properties saved in the extension, if they are saved
        Try
            If m_FiPEx.m_bLoaded = True Then

                sDirection = Convert.ToString(m_FiPEx.pPropset.GetProperty("direction"))
                iOrderNum = Convert.ToInt32((m_FiPEx.pPropset.GetProperty("ordernum")))
                bMaximum = Convert.ToBoolean((m_FiPEx.pPropset.GetProperty("maximum")))
                bConnectTab = Convert.ToBoolean(m_FiPEx.pPropset.GetProperty("connecttab"))
                bAdvConnectTab = Convert.ToBoolean(m_FiPEx.pPropset.GetProperty("advconnecttab"))
                bBarrierPerm = Convert.ToBoolean(m_FiPEx.pPropset.GetProperty("barrierperm"))
                bNaturalTF = Convert.ToBoolean(m_FiPEx.pPropset.GetProperty("naturalyn"))
                bDCI = Convert.ToBoolean(m_FiPEx.pPropset.GetProperty("dciyn"))
                bDCISectional = Convert.ToBoolean(m_FiPEx.pPropset.GetProperty("dcisectionalyn"))
                sDCIModelDir = Convert.ToString(m_FiPEx.pPropset.GetProperty("sDCIModelDir"))
                sRInstallDir = Convert.ToString(m_FiPEx.pPropset.GetProperty("sRInstallDir"))

                '2020

                Try
                    bUseHabLength = Convert.ToBoolean(m_FiPEx.pPropset.GetProperty("bUseHabLength"))
                Catch ex As Exception
                    bUseHabLength = False
                End Try


                Try
                    bUseHabArea = Convert.ToBoolean(m_FiPEx.pPropset.GetProperty("bUseHabArea"))
                Catch ex As Exception
                    bUseHabArea = False
                End Try

                Try
                    bDistanceDecay = Convert.ToBoolean(m_FiPEx.pPropset.GetProperty("bDistanceDecay"))
                Catch ex As Exception
                    'MessageBox.Show("Trouble loading FIPEX property bDistanceDecay. Setting it to False.")
                    bDistanceDecay = False
                End Try

                Try
                    bDistanceLim = Convert.ToBoolean(m_FiPEx.pPropset.GetProperty("bDistanceLim"))
                Catch ex As Exception
                    'MessageBox.Show("Trouble loading FIPEX property bDistanceLim. Setting it to False.")
                    bDistanceLim = False
                End Try

                Try
                    dMaxDist = Convert.ToDouble(m_FiPEx.pPropset.GetProperty("dMaxDist"))
                Catch ex As Exception
                    'MessageBox.Show("Trouble loading FIPEX property dMaxDist. Setting it to 0.")
                    dMaxDist = 0.0
                End Try

                Try
                    sDDFunction = Convert.ToString(m_FiPEx.pPropset.GetProperty("sDDFunction"))
                Catch ex As Exception
                    'MessageBox.Show("Trouble loading FIPEX property sDDFunction. Setting it to 'none'.")
                    sDDFunction = "none"
                End Try

                bDBF = Convert.ToBoolean(m_FiPEx.pPropset.GetProperty("bDBF"))
                sGDB = Convert.ToString(m_FiPEx.pPropset.GetProperty("sGDB"))
                sPrefix = Convert.ToString(m_FiPEx.pPropset.GetProperty("TabPrefix"))

                bUpHab = Convert.ToBoolean(m_FiPEx.pPropset.GetProperty("UpHab"))
                bTotalUpHab = Convert.ToBoolean(m_FiPEx.pPropset.GetProperty("TotalUpHab"))
                bDownHab = Convert.ToBoolean(m_FiPEx.pPropset.GetProperty("DownHab"))
                bTotalDownHab = Convert.ToBoolean(m_FiPEx.pPropset.GetProperty("TotalDownHab"))
                bPathDownHab = Convert.ToBoolean(m_FiPEx.pPropset.GetProperty("PathDownHab"))
                bTotalPathDownHab = Convert.ToBoolean(m_FiPEx.pPropset.GetProperty("TotalPathDownHab"))

                m_iPolysCount = Convert.ToInt32(m_FiPEx.pPropset.GetProperty("numPolys"))

                Dim PolyHabLayerObj As New PolyLayerToAdd(Nothing, Nothing, Nothing, Nothing)

                ' match any of the polygon layers saved in stream to those in listboxes and select
                If m_iPolysCount > 0 Then
                    For k = 0 To m_iPolysCount - 1
                        sPolyLayer = Convert.ToString(m_FiPEx.pPropset.GetProperty("IncPoly" + k.ToString)) ' get poly layer

                        PolyHabLayerObj = New PolyLayerToAdd(Nothing, Nothing, Nothing, Nothing)
                        With PolyHabLayerObj
                            .Layer = sPolyLayer
                            .HabClsField = Convert.ToString(m_FiPEx.pPropset.GetProperty("PolyClassField" + k.ToString))
                            .HabQuanField = Convert.ToString(m_FiPEx.pPropset.GetProperty("PolyQuanField" + k.ToString))
                            .HabUnitField = Convert.ToString(m_FiPEx.pPropset.GetProperty("PolyUnitField" + k.ToString))
                        End With

                        ' Load that object into the list
                        m_PLayersFields.Add(PolyHabLayerObj)
                        m = 0 ' loop to set currently selected
                        For m = 0 To lstPolyLayers.Items.Count - 1                ' for each item in list
                            If sPolyLayer = lstPolyLayers.Items.Item(m).ToString Then      ' if the two match
                                lstPolyLayers.SetSelected(m, True)                ' set it selected
                                lstPolyLayers.SetItemChecked(m, True)             ' set it checked
                            End If
                        Next
                    Next
                End If

                m_iLinesCount = Convert.ToInt32(m_FiPEx.pPropset.GetProperty("numLines"))
                ' match any of the line layers saved in stream to those in listboxes and select

                Dim LineHabLayerObj As New LineLayerToAdd(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing) ' layer to hold parameters to send to property

                If m_iLinesCount > 0 Then
                    For j = 0 To m_iLinesCount - 1
                        sLineLayer = Convert.ToString(m_FiPEx.pPropset.GetProperty("IncLine" + j.ToString)) ' get line layer
                        LineHabLayerObj = New LineLayerToAdd(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                        With LineHabLayerObj
                            .Layer = sLineLayer
                            .LengthField = Convert.ToString(m_FiPEx.pPropset.GetProperty("LineLengthField" + j.ToString))
                            .LengthUnits = Convert.ToString(m_FiPEx.pPropset.GetProperty("LineLengthUnits" + j.ToString))
                            .HabClsField = Convert.ToString(m_FiPEx.pPropset.GetProperty("LineHabClassField" + j.ToString))
                            .HabQuanField = Convert.ToString(m_FiPEx.pPropset.GetProperty("LineHabQuanField" + j.ToString))
                            .HabUnits = Convert.ToString(m_FiPEx.pPropset.GetProperty("LineHabUnits" + j.ToString))
                        End With

                        ' add to the module level list
                        m_LLayersFields.Add(LineHabLayerObj)

                        m = 0
                        For m = 0 To lstLineLayers.Items.Count - 1                ' for each item in list
                            If sLineLayer = lstLineLayers.Items.Item(m).ToString Then
                                ' if the two match
                                'lstLineLayers.SetItemChecked(m, True)

                                lstLineLayers.SetSelected(m, True)

                                lstLineLayers.SetItemCheckState(m, Windows.Forms.CheckState.Checked)
                                'lstLineLayers.s()
                            End If
                        Next
                    Next
                End If

                m_iExclusions = Convert.ToInt32(m_FiPEx.pPropset.GetProperty("numExclusions"))
                Dim ExcldLayerObj As New LayerToExclude(Nothing, Nothing, Nothing)
                ' match any of the line layers saved in stream to those in listboxes and select
                If m_iExclusions > 0 Then
                    For j = 0 To m_iExclusions - 1

                        sExcldLayer = Convert.ToString(m_FiPEx.pPropset.GetProperty("ExcldLayer" + j.ToString)) ' get line layer
                        ExcldLayerObj = New LayerToExclude(Nothing, Nothing, Nothing)
                        With ExcldLayerObj
                            .Layer = sExcldLayer
                            .Feature = Convert.ToString(m_FiPEx.pPropset.GetProperty("ExcldFeature" + j.ToString))
                            .Value = Convert.ToString(m_FiPEx.pPropset.GetProperty("ExcldValue" + j.ToString))
                        End With

                        ' add to the module level list
                        m_lExclusions.Add(ExcldLayerObj)

                        ' Add exclusion to the exclude list
                        lstLyrsExcld.Items.Add(ExcldLayerObj.Layer)
                        lstFtrsExcld.Items.Add(ExcldLayerObj.Feature)
                        lstVlsExcld.Items.Add(ExcldLayerObj.Value)
                    Next
                End If

                Dim sBarrierIDLayer As String
                Dim sBarrierIDField As String
                Dim sBarrierPermField As String
                Dim sBarrierNaturalYNField As String

                Dim pBarrierIDObj As New BarrierIDObj(Nothing, Nothing, Nothing, Nothing, Nothing)
                m_iBarrierCount = Convert.ToInt32(m_FiPEx.pPropset.GetProperty("numBarrierIDs"))
                If m_iBarrierCount > 0 Then
                    For j = 0 To m_iBarrierCount - 1
                        sBarrierIDLayer = Convert.ToString(m_FiPEx.pPropset.GetProperty("BarrierIDLayer" + j.ToString))
                        sBarrierIDField = Convert.ToString(m_FiPEx.pPropset.GetProperty("BarrierIDField" + j.ToString))
                        sBarrierPermField = Convert.ToString(m_FiPEx.pPropset.GetProperty("BarrierPermField" + j.ToString))
                        sBarrierNaturalYNField = Convert.ToString(m_FiPEx.pPropset.GetProperty("BarrierNaturalYNField" + j.ToString))

                        pBarrierIDObj = New BarrierIDObj(sBarrierIDLayer, sBarrierIDField, sBarrierPermField, sBarrierNaturalYNField, Nothing)

                        m_lBarrierIDs.Add(pBarrierIDObj)
                        m = 0 ' loop to set currently selected
                        ' YOU MUST HAVE TAB SELECTED TO SET AN ITEM 
                        ' IN A LISTBOX AS SELECTED (PROPERlY)
                        TabBarriers.SelectTab(2)
                        For m = 0 To chkLstBarriersLayers.Items.Count - 1                ' for each item in list
                            If sBarrierIDLayer = chkLstBarriersLayers.Items.Item(m).ToString Then      ' if the two match

                                chkLstBarriersLayers.SetSelected(m, True)         ' set it selected
                                chkLstBarriersLayers.SetItemChecked(m, True)      ' set it checked
                            End If
                        Next
                        TabBarriers.SelectTab(0)
                    Next

                    ' If there are no barriers then disable DBF section in 'Advanced Tab'
                Else
                    ChkDBFOutput.Enabled = False
                End If

                Try
                    bGLPKTables = Convert.ToBoolean(m_FiPEx.pPropset.GetProperty("bGLPKTables"))
                    sGLPKModelDir = Convert.ToString(m_FiPEx.pPropset.GetProperty("sGLPKModelDir"))
                    sGnuWinDir = Convert.ToString(m_FiPEx.pPropset.GetProperty("sGnuWinDir"))
                Catch ex As Exception

                End Try



                ' If no settings are being loaded from the extension stream then 
                ' no barriers layer will be selected so disable the DBF output by default 
            Else
                ' load defaults

                sDirection = "up"
                iOrderNum = 999
                bMaximum = True
                bConnectTab = False
                bAdvConnectTab = False
                bBarrierPerm = False
                bNaturalTF = False
                bDCI = False
                bDCISectional = False
                sDCIModelDir = "not set"
                sRInstallDir = "not set"

                '2020
                bUseHabLength = True
                bUseHabArea = False
                bDistanceLim = False
                dMaxDist = 0.0
                bDistanceDecay = False
                sDDFunction = "none"

                bDBF = False
                sGDB = "not set"
                sPrefix = "not set"

                bUpHab = True
                bTotalUpHab = True
                bDownHab = False
                bTotalDownHab = False
                bPathDownHab = False
                bTotalPathDownHab = False
                m_iPolysCount = 0
                m_iLinesCount = 0
                m_iExclusions = 0
                m_iBarrierCount = 0

                bGLPKTables = False
                sGLPKModelDir = "not set"
                sGnuWinDir = "not set"

            End If

        Catch ex As Exception
            MsgBox("Error loading stored FIPEX Properties. This is normal if it's the first time loading FIPEX Options or if FIPEX has been updated. " & ex.Message)

            sDirection = "up"
            iOrderNum = 999
            bMaximum = True
            bConnectTab = False
            bAdvConnectTab = False
            bBarrierPerm = False
            bNaturalTF = False
            bDCI = False
            bDCISectional = False
            sDCIModelDir = "not set"
            sRInstallDir = "not set"

            '2020
            bUseHabLength = True
            bUseHabArea = False
            bDistanceLim = False
            dMaxDist = 0.0
            bDistanceDecay = False
            sDDFunction = "none"

            bDBF = False
            sGDB = "not set"
            sPrefix = "not set"

            bUpHab = True
            bTotalUpHab = True
            bDownHab = False
            bTotalDownHab = False
            bPathDownHab = False
            bTotalPathDownHab = False
            m_iPolysCount = 0
            m_iLinesCount = 0
            m_iExclusions = 0
            m_iBarrierCount = 0

            bGLPKTables = False
            sGLPKModelDir = "not set"
            sGnuWinDir = "not set"
        End Try

        ' Set Analysis Direction
        If sDirection = "down" Then
            OptUp.Checked = False
            OptDown.Checked = True
        Else
            OptDown.Checked = False
            OptUp.Checked = True
        End If

        'Set the Maximum
        If bMaximum = True Then
            ChkMaxOrd.Checked = True
            TxtOrder.Enabled = False
        Else
            ChkMaxOrd.Checked = False
        End If

        ' Set the Order
        TxtOrder.Text = Convert.ToString(iOrderNum)

        If bConnectTab = True Then
            chkConnect.Checked = True
        Else
            chkConnect.Checked = False
        End If

        ' 2020
        If bAdvConnectTab = True Then
            chkAdvConnect.Checked = True
        Else
            chkAdvConnect.Checked = False
        End If


        ' default the things below dbf checkbox to disabled...
        cmdAddGDB.Enabled = False
        txtGDB.Enabled = False
        txtTablesPrefix.Enabled = False
        chkConnect.Enabled = False
        chkAdvConnect.Enabled = False
        chkBarrierPerm.Enabled = False
        chkNaturalTF.Enabled = False


        '...including things in DCI group
        chkDCI.Enabled = False
        chkDCISectional.Enabled = False

        ' Set DBF output parameters
        If bDBF = True Then
            ChkDBFOutput.Checked = True

            cmdAddGDB.Enabled = True  ' enable the controls
            'txtGDB.Enabled = True
            txtTablesPrefix.Enabled = True
            chkConnect.Enabled = True
            chkAdvConnect.Enabled = True
            chkNaturalTF.Enabled = True
            chkBarrierPerm.Enabled = True

            ' enable DCI controls only if perm and natural yn are checked
            If bBarrierPerm = True And bNaturalTF = True And (bConnectTab = True Or bAdvConnectTab = True) Then
                chkDCI.Enabled = True
            End If

            txtGDB.Text = sGDB
            txtTablesPrefix.Text = sPrefix
        Else
            ChkDBFOutput.Checked = False
            txtGDB.Text = "n/a"
            txtTablesPrefix.Text = "n/a"
        End If

        ' Set the habitat stats checkboxes
        If bUpHab = True Then
            chkUpHab.Checked = True
        Else
            chkUpHab.Checked = False
        End If

        If bTotalUpHab = True Then
            chkTotalUpHab.Checked = True
        Else
            chkTotalUpHab.Checked = False
        End If

        If bDownHab = True Then
            chkDownHab.Checked = True
        Else
            chkDownHab.Checked = False
        End If

        If bTotalDownHab = True Then
            chkTotalDownHab.Checked = True
        Else
            chkTotalDownHab.Checked = False
        End If

        If bPathDownHab = True Then
            chkPathDownHab.Checked = True
        Else
            chkPathDownHab.Checked = False
        End If

        If bTotalPathDownHab = True Then
            chkTotalPathDownHab.Checked = True
        Else
            chkTotalPathDownHab.Checked = False
        End If

        ' make list boxes uneditable
        'lstLineHabCls.Enabled = False
        'lstPolyHabCls.Enabled = False

        ' Set the Barrier Permeability Field Option on/off
        If bBarrierPerm = True Then
            chkBarrierPerm.Checked = True
        Else
            chkBarrierPerm.Checked = False
        End If

        ' Set the Natural Barrier YN option 
        If bNaturalTF = True Then
            chkNaturalTF.Checked = True
        Else
            chkNaturalTF.Checked = False
        End If

        ' Set the DCI output y/n 
        If bDCI = True Then

            chkDCI.Checked = True
            chkDCISectional.Checked = bDCISectional
            txtDCIModelDir.Text = sDCIModelDir
            txtRInstallDir.Text = sRInstallDir

            '2020
            chkDistanceLimit.Enabled = True
            chkDistanceLimit.Checked = bDistanceLim
            rdoHabArea.Checked = bUseHabArea
            rdoHabLength.Checked = bUseHabLength

            txtMaxDistance.Text = dMaxDist
            chkDistanceDecay.Checked = bDistanceDecay
            If sDDFunction = "linear" Then
                rdoLinear.Checked = True
            ElseIf sDDFunction = "natexp1" Then
                rdoNatExp1.Checked = True
            ElseIf sDDFunction = "circle" Then
                rdoCircle.Checked = True
            ElseIf sDDFunction = "sigmoid" Then
                rdoSigmoid.Checked = True
            End If

        Else
            chkDCI.Checked = False
            chkDCISectional.Checked = False
            txtDCIModelDir.Text = "n/a"
            txtRInstallDir.Text = "n/a"

            '2020
            rdoHabArea.Checked = False
            rdoHabLength.Checked = True
            chkDistanceLimit.Checked = False
            txtMaxDistance.Text = 0.0
            chkDistanceDecay.Checked = False
            bDistanceDecay = False
            rdoLinear.Checked = False
            rdoNatExp1.Checked = False
            rdoCircle.Checked = False
            rdoSigmoid.Checked = False
        End If

    End Sub
    Private Sub ChkDBFOutput_Click()

        ' If the checkbox is checked then enable
        ' some checkboxes
        If ChkDBFOutput.Checked = True Then
            txtTablesPrefix.Enabled = True
            txtGDB.Enabled = True
            chkConnect.Enabled = True
            chkAdvConnect.Enabled = True
            cmdAddGDB.Enabled = True

            ' enable DCI controls only if perm and natural yn are checked
            'If chkBarrierPerm.Checked = True And chkNaturalYN.Checked = True Then
            '    chkDCI.Enabled = True
            'End If

        Else
            txtTablesPrefix.Enabled = False
            txtGDB.Enabled = False
            chkConnect.Enabled = False
            chkAdvConnect.Enabled = False
            cmdAddGDB.Enabled = False

        End If
    End Sub
    Private Sub lstLineLayers_ItemCheck(ByVal sender As System.Object, ByVal e As ItemCheckEventArgs) Handles lstLineLayers.ItemCheck
        ' -----------------------------------------------------------------------
        ' This sub does the following:
        ' 1. Updates the global list of included habitat layers and corresponding
        '    quantity and quality fields
        ' 2. Updates the boxes containing the current fields on the form
        ' 3. Update the exclusions listbox containing layers to choose from 

        ' Because of the way the this finds matches, if there are two layers with 
        ' the same name in the global list, then the *last* one encountered
        ' will be used.  Layers with the same name should be avoided!!
        '
        ' 1.0 If the item in the list is being checked
        '   2.0a Clear the listboxes for habitat quantity, class, and unit
        '   2.1a If there are line layers in the list
        '       (loaded from extension settings to global variable)
        '     3.0a For each of the layers in this global list 
        '     3.1a If the layer in the global list matches the selected item
        '     3.2a Populate a variable object to pass to the ChooseHabParam form
        '     3.3a Populate the listboxes with units, quantity, and habitat field
        '   2.2a Else if there are no line layers listed
        '     3.0a Add "not set" or "n/a" to unit, class, and quantity listboxes
        '     3.0a Add an object to the global list of line layers using
        '   2.3a Add the layer to the exclusions list
        ' 1.1 Else if the item in the list is being unchecked
        '   2.0b Remove it from the global list 
        '   2.1b Remove it from the exclusions lists
        ' -----------------------------------------------------------------------

        Dim sLayer As String = lstLineLayers.SelectedItem.ToString
        Dim pLayerToAdd As New LineLayerToAdd(sLayer, Nothing, Nothing, Nothing, Nothing, Nothing)

        ' If it is checked now check if it is part of the properties variable
        If e.NewValue = CheckState.Checked Then
            ' if the form is loading here, the selectedindex event will not have
            ' turned on the changepoly command button so we need to do it here

            cmdChngLineCls.Enabled = True
            lstLineLength.Items.Clear()
            lstLineLengthUnits.Items.Clear()
            lstLineHabCls.Items.Clear()
            lstLineHabQuan.Items.Clear()
            lstLineHabUnits.Items.Clear()
            Dim bMatch As Boolean = False
            Dim sLineLengthUnits, sLineHabUnits As String

            ' If there are any layers in the list then loop through the list to see if there
            ' is a match.  If there is a match then use the class and quantity fields associated
            ' to populate the boxes.  If there is no match then add "not set" to the boxes and
            ' an object to the global list of layers. 
            If m_LLayersFields.Count <> 0 Then
                For i As Integer = 0 To m_LLayersFields.Count - 1
                    ' it really shouldn't match, it should have been removed from the list if 
                    ' it was unchecked - unless this is called when the form is loading
                    ' If the form is loading here then need to add the cls fields to the textboxes
                    ' - normally this would be done in the selectindexchange event but this was already
                    ' called to help us find the item that is being checked in this routine.

                    If m_LLayersFields.Item(i).Layer = sLayer Then

                        ' (if the form is loading it needs to populate this object 
                        '  in case the user clicks directly on the 'change' button
                        '  without highlighting a layer in the listbox)
                        If m_LLayerToAdd Is Nothing Then
                            m_LLayerToAdd = New LineLayerToAdd(m_LLayersFields.Item(i).Layer, m_LLayersFields.Item(i).LengthField, m_LLayersFields.Item(i).LengthUnits, _
m_LLayersFields.Item(i).HabQuanField, m_LLayersFields.Item(i).HabClsField, m_LLayersFields.Item(i).HabUnits)
                        Else
                            m_LLayerToAdd.Layer = m_LLayersFields.Item(i).Layer
                            m_LLayerToAdd.LengthField = m_LLayersFields.Item(i).LengthField
                            m_LLayerToAdd.LengthUnits = m_LLayersFields.Item(i).LengthUnits

                            m_LLayerToAdd.HabClsField = m_LLayersFields.Item(i).HabClsField
                            m_LLayerToAdd.HabQuanField = m_LLayersFields.Item(i).HabQuanField
                            m_LLayerToAdd.HabUnits = m_LLayersFields.Item(i).HabUnits
                        End If

                        lstLineLength.Items.Add(m_LLayersFields.Item(i).LengthField) ' load class field to box
                        sLineLengthUnits = m_LLayersFields.Item(i).LengthUnits
                        lstLineLengthUnits.Items.Add(sLineLengthUnits) ' load quantity field to box

                        lstLineHabCls.Items.Add(m_LLayersFields.Item(i).HabClsField) ' load class field to box
                        lstLineHabQuan.Items.Add(m_LLayersFields.Item(i).HabQuanField) ' load quantity field to box

                        sLineHabUnits = m_LLayersFields.Item(i).HabUnits
                        lstLineHabUnits.Items.Add(sLineHabUnits)

                        'If sLineUnit = "Metres" Then
                        '    lstLineUnit.Items.Add("m")
                        'ElseIf sLineUnit = "Kilometres" Then
                        '    lstLineUnit.Items.Add("km")
                        'ElseIf sLineUnit = "Square Metres" Then
                        '    lstLineUnit.Items.Add("m^2")
                        'ElseIf sLineUnit = "Feet" Then
                        '    lstLineUnit.Items.Add("ft")
                        'ElseIf sLineUnit = "Miles" Then
                        '    lstLineUnit.Items.Add("mi")
                        'ElseIf sLineUnit = "Square Miles" Then
                        '    lstLineUnit.Items.Add("mi^2")
                        'ElseIf sLineUnit = "Hectares" Then
                        '    lstLineUnit.Items.Add("ha")
                        'ElseIf sLineUnit = "Acres" Then
                        '    lstLineUnit.Items.Add("ac")
                        'Else
                        '    lstLineUnit.Items.Add("n/a")
                        'End If

                        bMatch = True
                        'Exit Sub                        
                    End If
                Next
            End If

            ' If no existing layer was found in the global list then add it to the list
            If bMatch = False Then

                lstLineLength.Items.Add("Not set")
                lstLineLength.Items.Add("n/a")
                lstLineHabCls.Items.Add("Not set")
                lstLineHabQuan.Items.Add("Not set")
                lstLineHabUnits.Items.Add("n/a")

                pLayerToAdd.LengthField = "Not set"
                pLayerToAdd.LengthUnits = "Not set"

                pLayerToAdd.HabClsField = "Not set"
                pLayerToAdd.HabQuanField = "Not set"
                pLayerToAdd.HabUnits = "Not set"

                m_LLayersFields.Add(pLayerToAdd)
            End If
            ' Add the layer to the exclusions list
            lstLayers.Items.Add(sLayer)

        Else ' If it has been unchecked
            If m_LLayersFields.Count <> 0 Then
                For i As Integer = m_LLayersFields.Count - 1 To 0 Step -1
                    If m_LLayersFields.Item(i).Layer = sLayer Then
                        m_LLayersFields.RemoveAt(i)
                    End If
                Next
            End If

            'Remove the layer from the exclusions list
            ' (this removes the first one encountered only and then exits)
            ' avoid layers with the same name!
            For i As Integer = 0 To lstLayers.Items.Count - 1
                If lstLayers.Items.Item(i).ToString = sLayer Then
                    lstLayers.Items.RemoveAt(i)
                    Exit For
                End If
            Next

        End If
    End Sub
    Private Sub lstPolyLayers_ItemCheck(ByVal sender As System.Object, ByVal e As ItemCheckEventArgs) Handles lstPolyLayers.ItemCheck

        ' -----------------------------------------------------------------------
        ' This sub does the following:
        ' 1. Updates the global list of included habitat layers and corresponding
        '    quantity and quality fields
        ' 2. Updates the boxes containing the current fields on the form
        ' 3. Update the exclusions listbox containing layers to choose from 

        ' Because of the way the this finds matches, if there are two layers with 
        ' the same name in the global list, then the *last* one encountered
        ' will be used.  Layers with the same name should be avoided!!
        '
        ' 1.0 If the item in the list is being checked
        '   2.0a Clear the listboxes for habitat quantity, class, and unit
        '   2.1a If there are polygon layers in the list
        '       (loaded from extension settings to global variable)
        '     3.0a For each of the layers in this global list 
        '     3.1a If the layer in the global list matches the selected item
        '     3.2a Populate a variable object to pass to the ChooseHabParam form
        '     3.3a Populate the listboxes with units, quantity, and habitat field
        '   2.2a Else if there are no polygon layers listed
        '     3.0a Add "not set" or "n/a" to unit, class, and quantity listboxes
        '     3.0a Add an object to the global list of polygon layers using
        '   2.3a Add the layer to the exclusions list
        ' 1.2 Else if the item in the list is being unchecked
        '   2.0b Remove it from the global list 
        '   2.1b Remove it from the exclusions lists
        ' -----------------------------------------------------------------------

        Dim sLayer As String = lstPolyLayers.SelectedItem.ToString
        Dim pLayerToAdd As New PolyLayerToAdd(sLayer, Nothing, Nothing, Nothing)

        ' If it is checked now check if it is part of the properties variable
        If e.NewValue = CheckState.Checked Then
            ' if the form is loading here, the selectedindex event will not have
            ' turned on the changepoly command button so we need to do it here
            cmdChngPolyCls.Enabled = True

            lstPolyHabCls.Items.Clear()
            lstPolyHabQuan.Items.Clear()
            lstPolyUnit.Items.Clear()
            Dim bMatch As Boolean = False
            Dim sPolyUnit As String

            ' If there are any layers in the list then loop through the list to see if there
            ' is a match.  If there is a match then use the class and quantity fields associated
            ' to populate the boxes.  If there is no match then add "not set" to the boxes and
            ' an object to the global list of layers. 
            If m_PLayersFields.Count <> 0 Then

                For i As Integer = 0 To m_PLayersFields.Count - 1
                    ' it really shouldn't match; it should have been removed from the list if 
                    ' it was unchecked. It could match, though, if this is the first time the
                    ' form has been loaded.
                    ' If the form is loading here then need to add the cls fields to the textboxes
                    ' - normally this would be done in the selectindexchange event but this was already
                    ' called to help us find the item that is being checked in this routine. 

                    ' If there is a match then update the textboxes corresponding to class and quantity fields
                    If m_PLayersFields.Item(i).Layer = sLayer Then

                        ' (if the form is loading it needs to populate this object 
                        '  in case the user clicks directly on the 'change' button
                        '  without highlighting a layer in the listbox)
                        If m_PLayerToAdd Is Nothing Then
                            m_PLayerToAdd = New PolyLayerToAdd(m_PLayersFields.Item(i).Layer, _
m_PLayersFields.Item(i).HabQuanField, m_PLayersFields.Item(i).HabClsField, m_PLayersFields.Item(i).HabUnitField)
                        Else
                            ' set the module level variable using these items
                            ' which can then be passed to the choosehabparam form
                            m_PLayerToAdd.Layer = m_PLayersFields.Item(i).Layer
                            m_PLayerToAdd.HabClsField = m_PLayersFields.Item(i).HabClsField
                            m_PLayerToAdd.HabQuanField = m_PLayersFields.Item(i).HabQuanField
                            m_PLayerToAdd.HabUnitField = m_PLayersFields.Item(i).HabUnitField
                        End If

                        lstPolyHabCls.Items.Add(m_PLayersFields.Item(i).HabClsField) ' load class field to box
                        lstPolyHabQuan.Items.Add(m_PLayersFields.Item(i).HabQuanField) ' load quantity field to box
                        sPolyUnit = m_PLayersFields.Item(i).HabUnitField
                        lstPolyUnit.Items.Add(sPolyUnit)
                        'If sPolyUnit = "Metres" Then
                        '    lstPolyUnit.Items.Add("m")
                        'ElseIf sPolyUnit = "Kilometres" Then
                        '    lstPolyUnit.Items.Add("km")
                        'ElseIf sPolyUnit = "Square Metres" Then
                        '    lstPolyUnit.Items.Add("m^2")
                        'ElseIf sPolyUnit = "Feet" Then
                        '    lstPolyUnit.Items.Add("ft")
                        'ElseIf sPolyUnit = "Miles" Then
                        '    lstPolyUnit.Items.Add("mi")
                        'ElseIf sPolyUnit = "Square Miles" Then
                        '    lstPolyUnit.Items.Add("mi^2")
                        'ElseIf sPolyUnit = "Hectares" Then
                        '    lstPolyUnit.Items.Add("ha")
                        'ElseIf sPolyUnit = "Acres" Then
                        '    lstPolyUnit.Items.Add("ac")
                        'Else
                        '    lstPolyUnit.Items.Add("n/a")
                        'End If
                        bMatch = True
                        'Exit Sub
                    End If
                Next
            End If
            ' If no existing layer was found in the global list then add it to the list
            If bMatch = False Then
                lstPolyHabCls.Items.Add("Not set") ' load class field to box
                lstPolyHabQuan.Items.Add("Not set") ' load quantity field to box
                lstPolyUnit.Items.Add("n/a")
                pLayerToAdd.HabClsField = "Not set" 'need to put something in these fields
                pLayerToAdd.HabQuanField = "Not set" ' or next .add will crash script
                pLayerToAdd.HabUnitField = "Not set"
                m_PLayersFields.Add(pLayerToAdd)
            End If

            ' Add the layer to the exclusions list
            lstLayers.Items.Add(sLayer)

        Else ' If it has been unchecked then remove the layer from the global list
            If m_PLayersFields.Count <> 0 Then
                For i As Integer = m_PLayersFields.Count - 1 To 0 Step -1
                    If m_PLayersFields.Item(i).Layer = sLayer Then
                        m_PLayersFields.RemoveAt(i)
                    End If
                Next
            End If

            'Remove the layer from the exclusions list
            ' (this removes the first one encountered only and then exits)
            ' avoid layers with the same name!
            For i As Integer = 0 To lstLayers.Items.Count - 1
                If lstLayers.Items.Item(i).ToString = sLayer Then
                    lstLayers.Items.RemoveAt(i)
                    Exit For
                End If
            Next
        End If ' checked state

    End Sub
    Private Sub chkLstBarriersLayers_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles chkLstBarriersLayers.ItemCheck
        ' -----------------------------------------------------------------------
        ' This sub does the following:
        ' 1. Updates the global list of barrier ID Layers and Fields
        ' 2. Updates the boxes containing the current field on the form

        ' Because of the way the this finds matches, if there are two layers with 
        ' the same name in the global list, then the *last* one encountered
        ' will be used.  Layers with the same name should be avoided!!
        '
        ' 1.0 If the item in the list is being checked
        '   2.0a Clear the listboxes for barrier ID field
        '   2.1a If there are layers in the list
        '       (loaded from extension settings to global variable)
        '     3.0a For each of the layers in this global list 
        '     3.1a If the layer in the global list matches the selected item
        '     3.2a Populate a variable object to pass to the ChooseHabParam form
        '     3.3a Populate the listboxes with units, quantity, and habitat field
        '   2.2a Else if there are no polygon layers listed
        '     3.0a Add "not set" or "n/a" to unit, class, and quantity listboxes
        '     3.0a Add an object to the global list of polygon layers using
        '   2.3a Add the layer to the exclusions list
        ' 1.2 Else if the item in the list is being unchecked
        '   2.0b Remove it from the global list 
        '   2.1b Remove it from the exclusions lists
        ' -----------------------------------------------------------------------

        Dim sLayer As String = chkLstBarriersLayers.SelectedItem.ToString
        Dim pBarrierIDObj As New BarrierIDObj(sLayer, Nothing, Nothing, Nothing, Nothing)

        ' If it is checked now check if it is part of the properties variable
        If e.NewValue = CheckState.Checked Then
            ' if the form is loading here, the selectedindex event will not have
            ' turned on the changebarrier command button so we need to do it here
            cmdBarrierID.Enabled = True

            lstBarrierField.Items.Clear()
            lstPermField.Items.Clear()
            lstNaturalTFField.Items.Clear()

            cmdSelectBarrierPerm.Enabled = True
            cmdSelectNaturalTF.Enabled = True

            Dim bMatch As Boolean = False

            ' If there are any layers in the list then loop through the list to see if there
            ' is a match.  If there is a match then use the class and quantity fields associated
            ' to populate the boxes.  If there is no match then add "not set" to the boxes and
            ' an object to the global list of layers. 
            If m_lBarrierIDs.Count <> 0 Then

                For i As Integer = 0 To m_lBarrierIDs.Count - 1
                    ' it really shouldn't match; it should have been removed from the list if 
                    ' it was unchecked. It could match, though, if this is the first time the
                    ' form has been loaded.
                    ' If the form is loading here then need to add the cls fields to the textboxes
                    ' - normally this would be done in the selectindexchange event but this was already
                    ' called to help us find the item that is being checked in this routine. 

                    ' If there is a match then update the textboxes corresponding to class and quantity fields
                    If m_lBarrierIDs.Item(i).Layer = sLayer Then

                        ' (if the form is loading it needs to populate this object 
                        '  in case the user clicks directly on the 'change' button
                        '  without highlighting a layer in the listbox)
                        If m_BarrierIDObj Is Nothing Then
                            m_BarrierIDObj = New BarrierIDObj(m_lBarrierIDs.Item(i).Layer, _
                                                              m_lBarrierIDs.Item(i).Field, _
                                                              m_lBarrierIDs.Item(i).PermField, _
                                                              m_lBarrierIDs.Item(i).NaturalYNField, _
                                                              Nothing)
                        Else
                            ' set the module level variable using these items
                            ' which can then be passed to the choosehabparam 
                            ' or the other forms for choosing perm or natural y/n
                            m_BarrierIDObj.Layer = m_lBarrierIDs.Item(i).Layer
                            m_BarrierIDObj.Field = m_lBarrierIDs.Item(i).Field
                            m_BarrierIDObj.PermField = m_lBarrierIDs.Item(i).PermField
                            m_BarrierIDObj.NaturalYNField = m_lBarrierIDs.Item(i).NaturalYNField

                        End If

                        lstBarrierField.Items.Add(m_lBarrierIDs.Item(i).Field)
                        lstPermField.Items.Add(m_lBarrierIDs.Item(i).PermField)
                        lstNaturalTFField.Items.Add(m_lBarrierIDs.Item(i).NaturalYNField)
                        bMatch = True
                        'Exit Sub
                    End If
                Next
            End If
            ' If no existing layer was found in the global list then add it to the list
            If bMatch = False Then
                lstBarrierField.Items.Add("Not set") ' load class field to box
                lstPermField.Items.Add("Not set")
                lstNaturalTFField.Items.Add("Not set")
                pBarrierIDObj.Field = "Not set" 'need to put something in these fields
                pBarrierIDObj.PermField = "Not set"
                pBarrierIDObj.NaturalYNField = "Not set"
                m_lBarrierIDs.Add(pBarrierIDObj)
            End If

        Else ' If it has been unchecked then remove the layer from the global list
            cmdSelectBarrierPerm.Enabled = False
            cmdSelectNaturalTF.Enabled = False
            If m_lBarrierIDs.Count <> 0 Then
                For i As Integer = m_lBarrierIDs.Count - 1 To 0 Step -1
                    If m_lBarrierIDs.Item(i).Layer = sLayer Then
                        m_lBarrierIDs.RemoveAt(i)
                    End If
                Next
            End If
        End If ' checked state

        ' If the count of barriers is zero then disable the dbf output checkbox in
        ' 'advanced' tab
        If m_lBarrierIDs.Count = 0 Then
            ChkDBFOutput.Enabled = False
        Else
            ChkDBFOutput.Enabled = True
        End If

    End Sub
    Private Sub lstLineLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstLineLayers.SelectedIndexChanged
        ' Populates two listboxes if there are habitat class and quantity fields set for it in property set
        ' MAY CAUSE ISSUE IF LAYERS SHARE A COMMON NAME - temp sol'n: will use first layer encountered in propset

        ' ----------------------------------------------------------------------------------------
        ' 1.0 Get the selected layer in the listbox as a variable
        ' 1.1 Clear the listboxes containing habitat quantity, class and units
        '   2.0 If the list layer item is being checked then
        '   2.1a enable the 'change button'
        '     3.0a If settings from this extension have been loaded then
        '       4.0a If there are any polygon layers in the global variable
        '            that holds layers included
        '         5.0a For each of those polygon layers
        '           6.0a If the layer in the global variable list matches the selected layer
        '           6.1a Set a module level variable containing all parameters for that field
        '                (that can be used to be passed to the ChooseHabPAram form)
        '           6.2a Load the class, quantity, and unit fields into the appropriate listboxes
        '           6.3a Check a variable saying there has been a match found
        '           6.4a Exit this for loop to avoid multiple adds (see issue outlined above)
        '  7.0 If nothing has been added to the class, quantity, and unit listboxes
        '  7.1 Set a module level variable containing all parameters for that field
        '     (that can be used to be passed to the ChooseHabPAram form)
        '  7.2 Update the listboxes with 'not set'
        ' -----------------------------------------------------------------------------------------
        ' clear the listboxes

        lstLineLength.Items.Clear()
        lstLineLengthUnits.Items.Clear()

        lstLineHabCls.Items.Clear()
        lstLineHabQuan.Items.Clear()
        lstLineHabUnits.Items.Clear()

        Dim indexes As ListBox.SelectedIndexCollection = Me.lstLineLayers.SelectedIndices

        ' check what item was selected
        ' sometimes things weren't selected, the box was just clicked. 
        'Dim pSelectedIndexCollection As System.Windows.Forms.CheckedListBox.SelectedIndexCollection
        If indexes.Count > 0 Then


            m_sLineLayer = lstLineLayers.SelectedItem.ToString
            Dim bSet As Boolean = False

            ' Default the 'change button' to disabled
            cmdChngLineCls.Enabled = False

            Dim iSelectedIndex As Integer = lstLineLayers.SelectedIndex
            Dim sLineHabUnits, sLineLengthUnits As String

            ' if checked there should be some parameters for it
            If lstLineLayers.GetItemCheckState(iSelectedIndex) = CheckState.Checked Then
                cmdChngLineCls.Enabled = True ' Enable the change button
                If m_FiPEx.m_bLoaded = True Then ' if any properties have been loaded
                    If m_LLayersFields.Count <> 0 Then
                        ' then the list in this class should have been loaded too
                        ' For each object in the list
                        For i As Integer = 0 To m_LLayersFields.Count - 1
                            ' if there is a match between layers
                            If m_LLayersFields.Item(i).Layer = m_sLineLayer Then
                                ' set the module level variable using these items
                                ' which can then be passed to the choosehabparam form
                                If m_LLayerToAdd Is Nothing Then
                                    m_LLayerToAdd = New LineLayerToAdd(m_LLayersFields.Item(i).Layer, m_LLayersFields.Item(i).LengthField, m_LLayersFields.Item(i).LengthUnits, _
    m_LLayersFields.Item(i).HabQuanField, m_LLayersFields.Item(i).HabClsField, m_LLayersFields.Item(i).HabUnits)
                                Else
                                    m_LLayerToAdd.Layer = m_LLayersFields.Item(i).Layer
                                    m_LLayerToAdd.LengthField = m_LLayersFields.Item(i).LengthField
                                    m_LLayerToAdd.LengthUnits = m_LLayersFields.Item(i).LengthUnits
                                    m_LLayerToAdd.HabClsField = m_LLayersFields.Item(i).HabClsField
                                    m_LLayerToAdd.HabQuanField = m_LLayersFields.Item(i).HabQuanField
                                    m_LLayerToAdd.HabUnits = m_LLayersFields.Item(i).HabUnits
                                End If

                                lstLineLength.Items.Add(m_LLayersFields.Item(i).LengthField)
                                sLineLengthUnits = m_LLayersFields.Item(i).LengthUnits
                                lstLineLengthUnits.Items.Add(sLineLengthUnits)

                                lstLineHabCls.Items.Add(m_LLayersFields.Item(i).HabClsField)
                                lstLineHabQuan.Items.Add(m_LLayersFields.Item(i).HabQuanField)

                                ' Add a abbreviated unit to the Lines Unit Listbox
                                sLineHabUnits = m_LLayersFields.Item(i).HabUnits
                                lstLineHabUnits.Items.Add(sLineHabUnits)
                                'If sLineUnit = "Metres" Then
                                '    lstLineUnit.Items.Add("m")
                                'ElseIf sLineUnit = "Kilometres" Then
                                '    lstLineUnit.Items.Add("km")
                                'ElseIf sLineUnit = "Square Metres" Then
                                '    lstLineUnit.Items.Add("m^2")
                                'ElseIf sLineUnit = "Feet" Then
                                '    lstLineUnit.Items.Add("ft")
                                'ElseIf sLineUnit = "Miles" Then
                                '    lstLineUnit.Items.Add("mi")
                                'ElseIf sLineUnit = "Square Miles" Then
                                '    lstLineUnit.Items.Add("mi^2")
                                'ElseIf sLineUnit = "Hectares" Then
                                '    lstLineUnit.Items.Add("ha")
                                'ElseIf sLineUnit = "Acres" Then
                                '    lstLineUnit.Items.Add("ac")
                                'Else
                                '    lstLineUnit.Items.Add("n/a")
                                'End If

                                bSet = True
                                Exit For ' exit after one item is added, to avoid multiple adds - see possible issue above
                            End If
                        Next
                    End If
                End If


            End If

            If bSet = False Then
                ' set the module level variable using these items
                ' which can then be passed to the choosehabparam form
                If m_LLayerToAdd Is Nothing Then ' check if it has been created already
                    ' else it will crash
                    m_LLayerToAdd = New LineLayerToAdd(m_sLineLayer, "Not set", "Not set", "Not set", "Not set", "Not set")
                Else
                    m_LLayerToAdd.Layer = m_sLineLayer

                    m_LLayerToAdd.LengthField = "Not set"
                    m_LLayerToAdd.LengthUnits = "Not set"

                    m_LLayerToAdd.HabClsField = "Not set"
                    m_LLayerToAdd.HabQuanField = "Not set"
                    m_LLayerToAdd.HabUnits = "Not set"
                End If

                lstLineLength.Items.Add("Not set")
                lstLineLengthUnits.Items.Add("n/a")

                lstLineHabCls.Items.Add("Not set")
                lstLineHabQuan.Items.Add("Not set")
                lstLineHabUnits.Items.Add("n/a")
            End If

        End If
    End Sub
    Private Sub lstPolyLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstPolyLayers.SelectedIndexChanged
        ' Populates two listboxes if there are habitat class and quantity fields set for it in property set
        ' MAY CAUSE ISSUE IF LAYERS SHARE A COMMON NAME - temp sol'n: will use first layer encountered in propset
        ' ----------------------------------------------------------------------------------------
        ' 1.0 Get the selected layer in the listbox as a variable
        ' 1.1 Clear the listboxes containing habitat quantity, class and units
        '   2.0 If the list layer item is being checked then
        '   2.1a enable the 'change button'
        '     3.0a If settings from this extension have been loaded then
        '       4.0a If there are any polygon layers in the global variable
        '            that holds layers included
        '         5.0a For each of those polygon layers
        '           6.0a If the layer in the global variable list matches the selected layer
        '           6.1a Set a module level variable containing all parameters for that field
        '                (that can be used to be passed to the ChooseHabPAram form)
        '           6.2a Load the class, quantity, and unit fields into the appropriate listboxes
        '           6.3a Check a variable saying there has been a match found
        '           6.4a Exit this for loop to avoid multiple adds (see issue outlined above)
        '  7.0 If nothing has been added to the class, quantity, and unit listboxes
        '  7.1 Set a module level variable containing all parameters for that field
        '     (that can be used to be passed to the ChooseHabPAram form)
        '  7.2 Update the listboxes with 'not set'
        ' -----------------------------------------------------------------------------------------

        ' clear the listboxes
        lstPolyHabCls.Items.Clear()
        lstPolyHabQuan.Items.Clear()
        lstPolyUnit.Items.Clear()

        Dim indexes As ListBox.SelectedIndexCollection = Me.lstPolyLayers.SelectedIndices
        If indexes.Count > 0 Then

            ' check what item was selected
            m_sPolyLayer = lstPolyLayers.SelectedItem.ToString
            Dim bSet As Boolean = False ' to check if lstItems have been set

            ' Default the 'change button' to disabled
            cmdChngPolyCls.Enabled = False

            Dim iSelectedIndex As Integer = lstPolyLayers.SelectedIndex
            Dim sPolyUnit As String

            ' check if the selected item is checked
            If lstPolyLayers.GetItemCheckState(iSelectedIndex) = CheckState.Checked Then
                cmdChngPolyCls.Enabled = True ' Enable the change button
                If m_FiPEx.m_bLoaded = True Then
                    ' if any properties have been loaded
                    If m_PLayersFields.Count <> 0 Then
                        For i As Integer = 0 To m_PLayersFields.Count - 1
                            ' if there is a match between layers
                            If m_PLayersFields.Item(i).Layer = m_sPolyLayer Then
                                If m_PLayerToAdd Is Nothing Then
                                    m_PLayerToAdd = New PolyLayerToAdd(m_PLayersFields.Item(i).Layer, _
    m_PLayersFields.Item(i).HabQuanField, m_PLayersFields.Item(i).HabClsField, m_PLayersFields.Item(i).HabUnitField)
                                Else
                                    ' set the module level variable using these items
                                    ' which can then be passed to the choosehabparam form
                                    m_PLayerToAdd.Layer = m_PLayersFields.Item(i).Layer
                                    m_PLayerToAdd.HabClsField = m_PLayersFields.Item(i).HabClsField
                                    m_PLayerToAdd.HabQuanField = m_PLayersFields.Item(i).HabQuanField
                                    m_PLayerToAdd.HabUnitField = m_PLayersFields.Item(i).HabUnitField
                                End If
                                lstPolyHabCls.Items.Add(m_PLayersFields.Item(i).HabClsField) ' load class field to box
                                lstPolyHabQuan.Items.Add(m_PLayersFields.Item(i).HabQuanField) ' load quantity field to box

                                sPolyUnit = m_PLayersFields.Item(i).HabUnitField
                                lstPolyUnit.Items.Add(sPolyUnit)
                                'If sPolyUnit = "Metres" Then
                                '    lstPolyUnit.Items.Add("m")
                                'ElseIf sPolyUnit = "Kilometres" Then
                                '    lstPolyUnit.Items.Add("km")
                                'ElseIf sPolyUnit = "Square Metres" Then
                                '    lstPolyUnit.Items.Add("m^2")
                                'ElseIf sPolyUnit = "Feet" Then
                                '    lstPolyUnit.Items.Add("ft")
                                'ElseIf sPolyUnit = "Miles" Then
                                '    lstPolyUnit.Items.Add("mi")
                                'ElseIf sPolyUnit = "Square Miles" Then
                                '    lstPolyUnit.Items.Add("mi^2")
                                'ElseIf sPolyUnit = "Hectares" Then
                                '    lstPolyUnit.Items.Add("ha")
                                'ElseIf sPolyUnit = "Acres" Then
                                '    lstPolyUnit.Items.Add("ac")
                                'Else
                                '    lstPolyUnit.Items.Add("n/a")
                                'End If
                                bSet = True
                                Exit For ' exit after one item is added, to avoid multiple adds - see possible issue above
                            End If
                        Next
                    End If
                End If
            End If

            If bSet = False Then

                ' set the module level variable using these items
                ' which can then be passed to the choosehabparam form
                If m_PLayerToAdd Is Nothing Then ' check if it has been created already
                    ' else it will crash
                    m_PLayerToAdd = New PolyLayerToAdd(m_sPolyLayer, "Not set", "Not set", "Not set")
                Else
                    m_PLayerToAdd.Layer = m_sPolyLayer
                    m_PLayerToAdd.HabClsField = "Not set"
                    m_PLayerToAdd.HabQuanField = "Not set"
                    m_PLayerToAdd.HabUnitField = "Not set"
                End If
                lstPolyHabCls.Items.Add("Not set")
                lstPolyHabQuan.Items.Add("Not set")
                lstPolyUnit.Items.Add("n/a")
            End If
        End If
    End Sub
    Private Sub chkLstBarriersLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLstBarriersLayers.SelectedIndexChanged

        ' Populates a listbox if there are barrier fields for the layer in the property set
        ' MAY CAUSE ISSUE IF LAYERS SHARE A COMMON NAME - temp sol'n: will use first layer encountered in propset

        ' ----------------------------------------------------------------------------------------
        ' 1.0 Get the selected layer in the listbox as a variable
        ' 1.1 Clear the listbox containing barrier ID Field
        '   2.0 If the list layer item is being checked then
        '   2.1a enable the 'change button'
        '     3.0a If settings from this extension have been loaded then
        '       4.0a If there are any layers in the global variable
        '         5.0a For each of those layers
        '           6.0a If the layer in the global variable list matches the selected layer
        '           6.1a Set a module level variable containing all parameters for that field
        '                (that can be used to be passed to the ChooseHabPAram form)
        '           6.2a Load the class, quantity, and unit fields into the appropriate listboxes
        '           6.3a Check a variable saying there has been a match found
        '           6.4a Exit this for loop to avoid multiple adds (see issue outlined above)
        '  7.0 If nothing has been added to the class, quantity, and unit listboxes
        '  7.1 Set a module level variable containing all parameters for that field
        '     (that can be used to be passed to the ChooseHabPAram form)
        '  7.2 Update the listboxes with 'not set'
        ' -----------------------------------------------------------------------------------------
        ' clear the listboxes
        lstBarrierField.Items.Clear()
        lstPermField.Items.Clear()
        lstNaturalTFField.Items.Clear()

        ' check what item was selected
        m_sBarrierIDLayer = chkLstBarriersLayers.SelectedItem.ToString
        Dim bSet As Boolean = False

        ' Default the 'change ID...' button to disabled
        cmdBarrierID.Enabled = False
        cmdSelectBarrierPerm.Enabled = False
        cmdSelectNaturalTF.Enabled = False

        Dim iSelectedIndex As Integer = chkLstBarriersLayers.SelectedIndex

        ' if checked there should be some parameters for it
        If chkLstBarriersLayers.GetItemCheckState(iSelectedIndex) = CheckState.Checked Then

            cmdBarrierID.Enabled = True ' Enable the change button
            cmdSelectNaturalTF.Enabled = True
            cmdSelectBarrierPerm.Enabled = True


            If m_FiPEx.m_bLoaded = True Then ' if any properties have been loaded
                If m_lBarrierIDs.Count <> 0 Then
                    ' then the list in this class should have been loaded too
                    ' For each object in the list
                    For i As Integer = 0 To m_lBarrierIDs.Count - 1
                        ' if there is a match between layers
                        If m_lBarrierIDs.Item(i).Layer = m_sBarrierIDLayer Then
                            ' set the module level variable using these items
                            ' which can then be passed to the choosehabparam form
                            If m_BarrierIDObj Is Nothing Then
                                m_BarrierIDObj = New BarrierIDObj(m_lBarrierIDs.Item(i).Layer, _
                                                                  m_lBarrierIDs.Item(i).Field, _
                                                                  m_lBarrierIDs.Item(i).PermField, _
                                                                  m_lBarrierIDs.Item(i).NaturalYNField, _
                                                                  Nothing)
                            Else
                                m_BarrierIDObj.Layer = m_lBarrierIDs.Item(i).Layer
                                m_BarrierIDObj.Field = m_lBarrierIDs.Item(i).Field
                                m_BarrierIDObj.PermField = m_lBarrierIDs.Item(i).PermField
                                m_BarrierIDObj.NaturalYNField = m_lBarrierIDs.Item(i).NaturalYNField
                            End If
                            lstBarrierField.Items.Add(m_lBarrierIDs.Item(i).Field)
                            lstPermField.Items.Add(m_lBarrierIDs.Item(i).PermField)
                            lstNaturalTFField.Items.Add(m_lBarrierIDs.Item(i).NaturalYNField)
                            bSet = True
                            Exit For ' exit after one item is added, to avoid multiple adds - see possible issue above
                        End If
                    Next
                End If
            End If
        End If

        If bSet = False Then
            ' set the module level variable using these items
            ' which can then be passed to the changeBarrierField form
            If m_BarrierIDObj Is Nothing Then ' check if it has been created already
                ' else it will crash
                m_BarrierIDObj = New BarrierIDObj(m_sBarrierIDLayer, _
                                                  "Not set", _
                                                  "Not set", _
                                                  "Not set", _
                                                  Nothing)
            Else
                m_BarrierIDObj.Layer = m_sBarrierIDLayer
                m_BarrierIDObj.Field = "Not set"
                m_BarrierIDObj.PermField = "Not set"
                m_BarrierIDObj.NaturalYNField = "Not set"
            End If
            lstBarrierField.Items.Add("Not set")
            lstPermField.Items.Add("Not set")
            lstNaturalTFField.Items.Add("Not set")
        End If



    End Sub
    Private Sub ChkDBFOutput_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkDBFOutput.CheckedChanged
        ' If the checkbox is checked then enable
        ' the table prefix textbox
        If ChkDBFOutput.Checked = True Then
            txtTablesPrefix.Enabled = True
            txtGDB.Enabled = True
            chkConnect.Enabled = True
            chkAdvConnect.Enabled = True
            cmdAddGDB.Enabled = True
            chkBarrierPerm.Enabled = True
            chkNaturalTF.Enabled = True


            ' enable DCI controls only if perm and natural yn are checked
            If chkBarrierPerm.Checked = True And chkNaturalTF.Checked = True And chkConnect.Checked = True Then
                chkDCI.Enabled = True
            End If

        Else
            txtTablesPrefix.Enabled = False
            txtGDB.Enabled = False
            chkConnect.Enabled = False
            chkAdvConnect.Enabled = False
            cmdAddGDB.Enabled = False
            chkBarrierPerm.Enabled = False
            chkNaturalTF.Enabled = False
            chkDCI.Enabled = False
            chkDistanceLimit.Enabled = False
            rdoCircle.Enabled = False
            rdoLinear.Enabled = False
            rdoSigmoid.Enabled = False
            rdoNatExp1.Enabled = False
        End If
    End Sub
    Private Sub cmdAddGDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddGDB.Click
        Dim pGxDialog As IGxDialog
        Dim pDstFilter As IGxObjectFilter
        Dim pEnumGx As IEnumGxObject

        pDstFilter = New GxFilterFileGeodatabases
        pGxDialog = New GxDialog

        Dim pFilterCol As IGxObjectFilterCollection
        pFilterCol = CType(pGxDialog, IGxObjectFilterCollection)

        pFilterCol.AddFilter(pDstFilter, True)

        pGxDialog.Title = "Browse for Geodatabase"
        pGxDialog.AllowMultiSelect = False

        pGxDialog.DoModalOpen(0, pEnumGx)

        If pEnumGx Is Nothing Then Exit Sub

        'Me.ZOrder(0) 'sendtoback?
        pEnumGx.Reset()

        Dim pGxDatabase As IGxDatabase
        Dim pGxObject As IGxObject = pEnumGx.Next
        pGxDatabase = CType(pGxObject, IGxDatabase)

        If pGxDatabase Is Nothing Then
            Exit Sub
        Else
            txtGDB.Text = pGxDatabase.Workspace.PathName
        End If
        txtGDB.Text = pGxDatabase.Workspace.PathName
    End Sub
    Private Sub cmdChngLineCls_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChngLineCls.Click

        m_sLayerType = "line"
        Dim lLayerToAdd As New List(Of LineLayerToAdd)

        If m_LLayerToAdd IsNot Nothing Then
            Using MyForm As New FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.ChooseLineHabParam(m_sLineLayer, m_LLayerToAdd)
                If MyForm.Form_Initialize(m_app) Then
                    MyForm.ShowDialog()
                    m_LLayerToAdd = MyForm.m_LayerToAdd
                End If
            End Using
            ' update the lstboxes now
            'Me.LHabParamList = m_LLayerToAdd
            lstLineLength.Items.Clear()
            lstLineLength.Items.Add(m_LLayerToAdd.LengthField)

            lstLineHabCls.Items.Clear()
            lstLineHabCls.Items.Add(m_LLayerToAdd.HabClsField)

            lstLineHabQuan.Items.Clear()
            lstLineHabQuan.Items.Add(m_LLayerToAdd.HabQuanField)

            lstLineLengthUnits.Items.Clear()
            lstLineHabUnits.Items.Clear()

            Dim sLineHabUnits As String = m_LLayerToAdd.HabUnits
            Dim sLineLengthUnits As String = m_LLayerToAdd.LengthUnits

            lstLineHabUnits.Items.Add(sLineHabUnits)
            lstLineLengthUnits.Items.Add(sLineLengthUnits)

            ' also need to update the list variable
            ' add it to the property, which will add it to the list
            lLayerToAdd.Add(m_LLayerToAdd)
            LHabParamList = lLayerToAdd
        End If
    End Sub
    Private Sub cmdChngPolyCls_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChngPolyCls.Click
        m_sLayerType = "poly"
        Dim lLayerToAdd As New List(Of PolyLayerToAdd)

        If m_PLayerToAdd IsNot Nothing Then
            Using MyForm As New FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.ChoosePolyHabParam(m_sPolyLayer, m_PLayerToAdd)
                If MyForm.Form_Initialize(m_app) Then
                    MyForm.ShowDialog()
                    m_PLayerToAdd = MyForm.m_LayerToAdd
                End If
            End Using
            ' update the lstboxes now
            'Me.PHabParamList = m_PLayerToAdd
            lstPolyHabCls.Items.Clear()
            lstPolyHabCls.Items.Add(m_PLayerToAdd.HabClsField)
            lstPolyHabQuan.Items.Clear()
            lstPolyHabQuan.Items.Add(m_PLayerToAdd.HabQuanField)

            lstPolyUnit.Items.Clear()
            Dim sPolyUnit As String = m_PLayerToAdd.HabUnitField
            lstPolyUnit.Items.Add(sPolyUnit)

            lLayerToAdd.Add(m_PLayerToAdd)
            PHabParamList = lLayerToAdd
        End If
    End Sub
    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub
    Public Property BarrierParamList() As List(Of BarrierIDObj)
        Get
            Return m_lBarrierIDs
        End Get
        Set(ByVal value As List(Of BarrierIDObj))

            Dim pBarrierIDObj As BarrierIDObj
            pBarrierIDObj = value(0)
            Dim sLayer As String = pBarrierIDObj.Layer
            Dim comparer As New CompareLayerNamePred2(sLayer)

            ' Need an object to hold returned object, if it exists...
            Dim CheckedBarrierIDObj As BarrierIDObj
            Dim iIndex As Integer

            ' This find function will pass in the m_LLayersFields object to the comparerobject function
            ' which has its variable -_name- set above
            CheckedBarrierIDObj = m_lBarrierIDs.Find(AddressOf comparer.CompareNames)
            iIndex = m_lBarrierIDs.FindIndex(AddressOf comparer.CompareNames)

            ' If there is a match found then change the other parameters of the object
            If CheckedBarrierIDObj IsNot Nothing Then
                With CheckedBarrierIDObj
                    If .Layer IsNot Nothing Then

                        ' Change the fields to equal incoming object's
                        .Field = pBarrierIDObj.Field
                        .PermField = pBarrierIDObj.PermField
                        .NaturalYNField = pBarrierIDObj.NaturalYNField
                        ' Now need to delete the old object and put it the modified one back into the list.
                        m_lBarrierIDs.Item(iIndex).Field = pBarrierIDObj.Field
                        m_lBarrierIDs.Item(iIndex).PermField = pBarrierIDObj.PermField
                        m_lBarrierIDs.Item(iIndex).NaturalYNField = pBarrierIDObj.NaturalYNField

                    End If
                End With
            Else 'Otherwise just add the object to the list
                m_lBarrierIDs.Add(pBarrierIDObj)
            End If

        End Set
    End Property
    Public Property PHabParamList() As List(Of PolyLayerToAdd)
        Get
            Return m_PLayersFields
        End Get
        Set(ByVal value As List(Of PolyLayerToAdd))
            Dim pLayerToAdd As PolyLayerToAdd

            pLayerToAdd = value(0)
            Dim sLayer As String = pLayerToAdd.Layer
            Dim comparer As New ComparePolyLayerNamePred(sLayer)

            ' Need an object to hold returned object, if it exists...
            Dim CheckedLayerToAddObj As PolyLayerToAdd
            Dim iIndex As Integer

            ' This find function will pass in the m_LLayersFields object to the comparerobject function
            ' which has its variable -_name- set above
            CheckedLayerToAddObj = m_PLayersFields.Find(AddressOf comparer.CompareNames)
            iIndex = m_PLayersFields.FindIndex(AddressOf comparer.CompareNames)

            ' If there is a match found then change the other parameters of the object
            If CheckedLayerToAddObj IsNot Nothing Then
                With CheckedLayerToAddObj
                    If .Layer IsNot Nothing Then
                        ' Change the fields to equal incoming object's
                        .HabClsField = pLayerToAdd.HabClsField
                        .HabQuanField = pLayerToAdd.HabQuanField
                        .HabUnitField = pLayerToAdd.HabUnitField

                        ' Now need to delete the old object and put it the modified one back into the list.
                        m_PLayersFields.Item(iIndex).HabClsField = pLayerToAdd.HabClsField
                        m_PLayersFields.Item(iIndex).HabQuanField = pLayerToAdd.HabQuanField
                        m_PLayersFields.Item(iIndex).HabUnitField = pLayerToAdd.HabUnitField
                    End If
                End With
            Else 'Otherwise just add the object to the list
                m_PLayersFields.Add(pLayerToAdd)
            End If

        End Set
    End Property
    Public Property LHabParamList() As List(Of LineLayerToAdd)
        Get
            Return m_LLayersFields
        End Get
        Set(ByVal value As List(Of LineLayerToAdd))

            Dim pLayerToAdd As LineLayerToAdd

            pLayerToAdd = value(0)
            Dim sLayer As String = pLayerToAdd.Layer
            Dim comparer As New CompareLineLayerNamePred(sLayer)

            ' Need an object to hold returned object, if it exists...
            Dim CheckedLayerToAddObj As LineLayerToAdd
            Dim iIndex As Integer

            ' This find function will pass in the m_LLayersFields object to the comparerobject function
            ' which has its variable -_name- set above
            CheckedLayerToAddObj = m_LLayersFields.Find(AddressOf comparer.CompareNames)
            iIndex = m_LLayersFields.FindIndex(AddressOf comparer.CompareNames)

            ' If there is a match found then change the other parameters of the object
            If CheckedLayerToAddObj IsNot Nothing Then
                With CheckedLayerToAddObj
                    If .Layer IsNot Nothing Then
                        ' Change the fields to equal incoming object's
                        .LengthField = pLayerToAdd.LengthField
                        .LengthUnits = pLayerToAdd.LengthUnits
                        .HabClsField = pLayerToAdd.HabClsField
                        .HabQuanField = pLayerToAdd.HabQuanField
                        .HabUnits = pLayerToAdd.HabUnits

                        ' Now need to delete the old object and put it the modified one back into the list.
                        m_LLayersFields.Item(iIndex).LengthField = pLayerToAdd.LengthField
                        m_LLayersFields.Item(iIndex).LengthUnits = pLayerToAdd.LengthUnits
                        m_LLayersFields.Item(iIndex).HabClsField = pLayerToAdd.HabClsField
                        m_LLayersFields.Item(iIndex).HabQuanField = pLayerToAdd.HabQuanField
                        m_LLayersFields.Item(iIndex).HabUnits = pLayerToAdd.HabUnits

                    End If
                End With
            Else 'Otherwise just add the object to the list
                m_LLayersFields.Add(pLayerToAdd)
            End If
        End Set
    End Property

    ' Use this Predicate object to define search terms and return comparison result
    Private Class CompareLineLayerNamePred
        Private _name As String
        Public Sub New(ByVal name As String)
            _name = name
        End Sub
        Public Function CompareNames(ByVal obj As LineLayerToAdd) As Boolean
            Return (_name = obj.Layer)
        End Function
    End Class
    Private Class ComparePolyLayerNamePred
        Private _name As String
        Public Sub New(ByVal name As String)
            _name = name
        End Sub
        Public Function CompareNames(ByVal obj As PolyLayerToAdd) As Boolean
            Return (_name = obj.Layer)
        End Function
    End Class

    ' Use this Predicate object to define search terms and return comparison result
    Private Class CompareLayerNamePred2
        Private _name As String
        Public Sub New(ByVal name As String)
            _name = name
        End Sub
        Public Function CompareNames(ByVal obj As BarrierIDObj) As Boolean
            Return (_name = obj.Layer)
        End Function
    End Class
    Private Sub ChkMaxOrd_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If ChkMaxOrd.Checked = True Then
            TxtOrder.Enabled = False
        Else
            TxtOrder.Enabled = True
        End If
    End Sub
    Private Sub cmdRemove2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemove2.Click
        ' -------------------------------------------------------
        ' 1. Removes the selected exclusion from the list
        ' 2. Updates global exclusions list
        ' -------------------------------------------------------

        Dim i, iselectedindex As Integer

        ' Get the selected layer, feature, and value
        Dim sExcldLayer, sExcldFeature, sExcldValue As String
        Try

            sExcldLayer = lstLyrsExcld.SelectedItem.ToString
            iselectedindex = lstLyrsExcld.SelectedIndex
            ' there is a bug in the setselected vb.net code or my code preventing index of 0 
            ' from being selected. need to get index if there is no selection
            If lstFtrsExcld.SelectedItems.Count = 0 Then
                sExcldFeature = lstFtrsExcld.Items.Item(iselectedindex).ToString
            Else
                sExcldFeature = lstFtrsExcld.SelectedItem.ToString
            End If
            If lstVlsExcld.SelectedItems.Count = 0 Then
                sExcldValue = lstVlsExcld.Items.Item(iselectedindex).ToString
            Else
                sExcldValue = lstVlsExcld.SelectedItem.ToString
            End If
            ' Remove Selected from the listboxes
            ' have to have the layers removed last or else it 
            ' triggers an index changed
            lstFtrsExcld.Items.RemoveAt(iselectedindex)
            lstVlsExcld.Items.RemoveAt(iselectedindex)
            lstLyrsExcld.Items.RemoveAt(iselectedindex)

            ' Check the global exclusions list for a match
            ' If there is a match then remove the exclusion from the list
            For i = m_lExclusions.Count - 1 To 0 Step -1
                If m_lExclusions(i).Layer.ToString = sExcldLayer Then
                    If m_lExclusions(i).Feature.ToString = sExcldFeature Then
                        If m_lExclusions(i).Value.ToString = sExcldValue Then
                            m_lExclusions.RemoveAt(i)
                        End If
                    End If
                End If
            Next

        Catch ex As Exception
            MsgBox("Error trying to remove exclusion from list" + ex.Message)
        End Try
    End Sub
    Private Sub lstLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstLayers.SelectedIndexChanged
        ' --------------------------------------------------------
        ' 1. Matches layer name in list with layer name in TOC
        ' 2. Loads field names for that layer into the listbox
        ' --------------------------------------------------------

        Dim pMxDoc As IMxDocument
        Dim pMap As IMap
        Dim pFeatureLayer As IFeatureLayer
        Dim i As Integer
        Dim sLayer As String
        Dim pFields As IFields
        Dim pField As IField

        ' hook into ArcMap
        Dim pDocument As IDocument = m_app.Document
        pMxDoc = CType(pDocument, IMxDocument)
        pMap = pMxDoc.FocusMap

        ' clear the fields and values listboxes
        lstFields.Items.Clear()
        lstValues.Items.Clear()

        ' See which layer was selected
        sLayer = lstLayers.SelectedItem.ToString
        'i = 0
        'For i = 0 To lstLayers.ListCount - 1
        '    If lstLayers.Selected(i) = True Then
        '        sLayer = lstLayers.List(i)
        '    End If
        'Next

        i = 0
        For i = 0 To pMap.LayerCount - 1
            If pMap.Layer(i).Valid = True Then
                If pMap.Layer(i).Name = sLayer Then Exit For
            End If
        Next

        Dim pLayer As ILayer = pMap.Layer(i)
        pFeatureLayer = CType(pLayer, IFeatureLayer)

        ' Get the feature class corresponding to feature layer
        m_pFeatureClass = pFeatureLayer.FeatureClass
        pFields = m_pFeatureClass.Fields ' get the fields from the featureclass

        i = 0
        ' Put the fields in the next listbox
        For i = 0 To pFields.FieldCount - 1
            pField = pFields.Field(i)
            lstFields.Items.Add(pField.Name)
        Next

    End Sub
    Private Sub lstFields_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstFields.SelectedIndexChanged
        ' -----------------------------------------------------
        ' Populates the lstValues box with unique values of
        ' selected attribute (field)
        ' -----------------------------------------------------

        Dim sFieldname As String
        'Dim i As Integer
        Dim pFeatureCursor As IFeatureCursor
        Dim pDataStats As IDataStatistics
        'Dim pEnumVar As ESRI.ArcGIS.esriSystem.IEnumVariantSimple
        Dim vVar As Object

        ' clear the values listbox
        lstValues.Items.Clear()

        ' Get selected field
        sFieldname = lstFields.SelectedItem.ToString

        'i = 0
        'For i = 0 To lstFields.ListCount - 1
        '    If lstFields.Selected(i) = True Then
        '        sFieldname = lstFields.List(i)
        '        Exit For
        '    End If
        'Next

        pFeatureCursor = m_pFeatureClass.Search(Nothing, False)

        ' Setup the datastatistics and get the unique values of the "Id" field
        pDataStats = New DataStatistics
        Dim pCursor As ICursor = CType(pFeatureCursor, ICursor)
        Dim pEnumerator As IEnumerator

        With pDataStats
            .Cursor = pCursor
            .Field = sFieldname
            pEnumerator = .UniqueValues
        End With

        'pEnumVar = CType(pEnumerator, ESRI.ArcGIS.esriSystem.IEnumVariantSimple)
        ' Display the no. of unique values and the actual values
        'Debug.Print "No. of unique values is:"; pDataStats.UniqueValueCount
        'Debug.Print "The values are:"

        pEnumerator.Reset()
        pEnumerator.MoveNext()
        vVar = pEnumerator.Current

        While Not IsNothing(vVar)
            lstValues.Items.Add(vVar.ToString)
            pEnumerator.MoveNext()
            vVar = pEnumerator.Current
        End While
    End Sub
    Private Sub cmdAddExcld_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddExcld.Click
        ' --------------------------------------------------------------
        ' 1. Adds the selected layer, feature, and value to the excludes
        '    listbox below
        ' 2. Updates the global exclusions list
        ' --------------------------------------------------------------

        Dim i As Integer
        Dim sExcldLyr As String
        Dim sExcldFtr As String
        Dim sExcldVal As String
        Dim duplicate As Boolean

        duplicate = False  ' initialize duplicate entry check variable

        ' Get selected layer
        sExcldLyr = lstLayers.SelectedItem.ToString
        'i = 0
        'For i = 0 To lstLayers.ListCount - 1
        '    If lstLayers.Selected(i) = True Then
        '        sExcldLyr = lstLayers.List(i)
        '        Exit For
        '    End If
        'Next

        ' Get selected field
        sExcldFtr = lstFields.SelectedItem.ToString

        'i = 0
        'For i = 0 To lstFields.ListCount - 1
        '    If lstFields.Selected(i) = True Then
        '        sExcldFld = lstFields.List(i)
        '        Exit For
        '    End If
        'Next

        ' Get selected value
        sExcldVal = lstValues.SelectedItem.ToString
        'i = 0
        'For i = 0 To lstValues.ListCount - 1
        '    If lstValues.Selected(i) = True Then
        '        sExcldVal = lstValues.List(i)
        '        Exit For
        '    End If
        'Next

        ' If any of the items are not selected show a message and prompt
        ' use to make sure something is selected in each listbox. 
        If sExcldLyr Is Nothing Or sExcldFtr Is Nothing Or sExcldVal Is Nothing Then
            System.Windows.Forms.MessageBox.Show("Please make sure that a layer, feature, and value are selected to add to exclusions", "Incomplete Exclusion List")
            Exit Sub
        End If

        ' Check that excludes list doesn't contain duplicate values
        i = 0
        For i = 0 To lstLyrsExcld.Items.Count - 1
            If lstLyrsExcld.Items.Item(i).ToString = sExcldLyr Then
                If lstFtrsExcld.Items.Item(i).ToString = sExcldFtr Then
                    If lstVlsExcld.Items.Item(i).ToString = sExcldVal Then
                        duplicate = True
                    End If
                End If
            End If
        Next

        ' Check and make sure there is a unique value for the field
        If sExcldVal <> "" Then
            If duplicate <> True Then
                ' add to excludes listboxes
                lstLyrsExcld.Items.Add(sExcldLyr)
                lstFtrsExcld.Items.Add(sExcldFtr)
                lstVlsExcld.Items.Add(sExcldVal)

                ' Add it to the global exclusions list
                Dim LayerToExcldObj As New LayerToExclude(sExcldLyr, sExcldFtr, sExcldVal)
                Dim bMatch As Boolean = False
                ' Make sure exclusion does not already exist in the global list
                ' If they do then do not add this one to the global list
                If m_lExclusions.Count <> 0 Then
                    For i = 0 To m_lExclusions.Count - 1
                        If m_lExclusions(i).Layer = sExcldLyr Then
                            If m_lExclusions(i).Feature = sExcldFtr Then
                                If m_lExclusions(i).Value = sExcldVal Then
                                    bMatch = True
                                End If
                            End If
                        End If
                    Next
                End If
                If bMatch = False Then
                    m_lExclusions.Add(LayerToExcldObj)
                End If

            Else
                System.Windows.Forms.MessageBox.Show("Exclusion already exists in list", "Duplicate Exclusion")
                Exit Sub
            End If
            ' if no value in unique value box
        Else
            System.Windows.Forms.MessageBox.Show("No value to add, please select one or choose another field.", "No Unique Value Found")
            Exit Sub
        End If

        ' save whatever is in the excludes listboxes for use later
        ' If there are layers in the excludes listbox

        'If lstLyrsExcld.ListCount <> 0 Then

        '    ' Make it into a comma delimited string for config file
        '    i = 0
        '    For i = 0 To lstLyrsExcld.ListCount - 1
        '        If i = 0 Then
        '            m_sLyrs = lstLyrsExcld.List(i)
        '        Else
        '            m_sLyrs = m_sLyrs + "," + lstLyrsExcld.List(i)
        '        End If
        '    Next

        '    i = 0
        '    For i = 0 To lstFldsExcld.ListCount - 1
        '        If i = 0 Then
        '            m_sFlds = lstFldsExcld.List(i)
        '        Else
        '            m_sFlds = m_sFlds + "," + lstFldsExcld.List(i)
        '        End If
        '    Next

        '    i = 0
        '    For i = 0 To lstVlsExcld.ListCount - 1
        '        If i = 0 Then
        '            m_sVls = lstVlsExcld.List(i)
        '        Else
        '            m_sVls = m_sVls + "," + lstVlsExcld.List(i)
        '        End If
        '    Next

        '    ' Update the arrays
        '    ' split the string by comma
        '    m_aLyrs = Split(m_sLyrs, ",")
        '    ReDim Preserve m_aLyrs(UBound(m_aLyrs))

        '    ' split the string by comma
        '    m_aFlds = Split(m_sFlds, ",")
        '    ReDim Preserve m_aFlds(UBound(m_aFlds))

        '    ' split the string by comma
        '    m_aVls = Split(m_sVls, ",")
        '    ReDim Preserve m_aVls(UBound(m_aVls))

        '    ' otherwise set the strings as empty
        'Else

        '    m_sLyrs = "none,"
        '    m_sFlds = "none,"
        '    m_sVls = "none,"

        '    ReDim m_aLyrs(0)
        '    m_aLyrs(0) = "none"
        '    ReDim m_aFlds(0)
        '    m_aFlds(0) = "none"
        '    ReDim m_aVls(0)
        '    m_aVls(0) = "none"

        'End If
    End Sub

    Private Sub lstLyrsExcld_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstLyrsExcld.SelectedIndexChanged
        ' -----------------------------------------------------------
        ' 1. Match Features and values listbox Selection to the Layer
        ' -----------------------------------------------------------

        Dim i As Integer

        ' Selects corresponding value and field
        i = 0
        For i = 0 To lstLyrsExcld.Items.Count - 1
            If lstLyrsExcld.SelectedIndex = i Then
                lstFtrsExcld.SetSelected(i, True)
                lstVlsExcld.SetSelected(i, True)
            Else
                lstFtrsExcld.SetSelected(i, False)
                lstVlsExcld.SetSelected(i, False)
            End If
        Next
    End Sub

    Private Sub lstFtrsExcld_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstFtrsExcld.SelectedIndexChanged
        ' ---------------------------------------------------------------------
        ' 1. Match Layers and values listbox Selection to the Feature Selection
        ' ---------------------------------------------------------------------
        'Dim i As Integer

        '' Selects corresponding value and field
        'i = 0
        'For i = 0 To lstFtrsExcld.Items.Count - 1
        '    If lstFtrsExcld.SelectedIndex = i Then
        '        lstLyrsExcld.SetSelected(i, True)
        '        lstVlsExcld.SetSelected(i, True)
        '    Else
        '        lstLyrsExcld.SetSelected(i, False)
        '        lstVlsExcld.SetSelected(i, False)
        '    End If
        'Next
    End Sub

    Private Sub lstVlsExcld_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstVlsExcld.SelectedIndexChanged
        ' ---------------------------------------------------------------------
        ' 1. Match Layers and Features listbox Selection to the Values Selection
        ' ---------------------------------------------------------------------
        'Dim i As Integer

        '' Selects corresponding value and field
        'i = 0
        'For i = 0 To lstVlsExcld.Items.Count - 1
        '    If lstVlsExcld.SelectedIndex = i Then
        '        lstLyrsExcld.SetSelected(i, True)
        '        lstFtrsExcld.SetSelected(i, True)
        '    Else
        '        lstLyrsExcld.SetSelected(i, False)
        '        lstFtrsExcld.SetSelected(i, False)
        '    End If
        'Next
    End Sub

    Private Sub cmdBarrierID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBarrierID.Click

        ' This command... 
        ' creates a new list of barrier layers and associated IDS (BarrierIDObj)
        ' Passes that to a form where the associated ID can be changed.  The form
        ' then passes back the updated object. 

        Dim lBarrierIDs As New List(Of BarrierIDObj)

        If m_BarrierIDObj IsNot Nothing Then
            Using MyForm As New FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.frmChooseBarrID(m_sBarrierIDLayer, m_BarrierIDObj) 'm_sPOlyLayer was here!
                If MyForm.Form_Initialize(m_app) Then
                    MyForm.ShowDialog()
                    m_BarrierIDObj = MyForm.m_BarrierIDObj2
                End If
            End Using
            ' update the lstboxes now
            'Me.PHabParamList = m_BarrierIDObj
            lstBarrierField.Items.Clear()
            lstBarrierField.Items.Add(m_BarrierIDObj.Field)

            ' also need to update the list variable
            ' add it to the property, which will add it to the list
            lBarrierIDs.Add(m_BarrierIDObj)
            BarrierParamList = lBarrierIDs
        End If
    End Sub

    Private Sub cmdSelectBarrierPerm_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelectBarrierPerm.Click
        ' This command... 
        ' creates a new list of barrier layers and associated IDS (BarrierIDObj)
        ' Passes that to a form where the associated ID can be changed.  The form
        ' then passes back the updated object. 

        Dim lBarrierIDs As New List(Of BarrierIDObj)

        If m_BarrierIDObj IsNot Nothing Then
            Using MyForm As New FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.frmChooseBarrierPerm(m_sBarrierIDLayer, m_BarrierIDObj)
                If MyForm.Form_Initialize(m_app) Then
                    MyForm.ShowDialog()
                    m_BarrierIDObj = MyForm.m_BarrierIDObj3
                End If
            End Using
            ' update the lstboxes now
            'Me.PHabParamList = m_BarrierIDObj
            lstPermField.Items.Clear()
            lstPermField.Items.Add(m_BarrierIDObj.PermField)

            ' also need to update the list variable
            ' add it to the property, which will add it to the list
            lBarrierIDs.Add(m_BarrierIDObj)
            BarrierParamList = lBarrierIDs
        End If
    End Sub

    Private Sub chkBarrierPerm_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBarrierPerm.CheckedChanged

        ' Get reference to the list items in the lstBarriers listbox 
        ' If one is selected and if that selected layer is checked
        ' then set the permeability button to 'active' 
        ' Also, enable/disable the DCI checkbox based on whether this check is active

        Dim i As Integer = chkLstBarriersLayers.SelectedIndex
        'Dim indexChecked As Integer

        If chkBarrierPerm.CheckState = CheckState.Checked Then
            'If i <> -1 Then
            '    For Each indexChecked In chkLstBarriersLayers.CheckedIndices
            '        If indexChecked = i Then
            '            cmdSelectBarrierPerm.Enabled = True
            '        End If
            '    Next
            'End If
            If chkNaturalTF.Checked = True And ChkDBFOutput.Checked = True And (chkConnect.Checked = True Or chkAdvConnect.Checked = True) Then
                chkDCI.Enabled = True
            End If
        Else
            'cmdSelectBarrierPerm.Enabled = False
            chkDCI.Enabled = False
        End If

    End Sub

    Private Sub chkNaturalYN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNaturalTF.CheckedChanged
        ' Get reference to the list items in the lstBarriers listbox 
        ' If one is selected and if that selected layer is checked
        ' then set the permeability button to 'active' 
        ' if it is being unchecked then disable DCI output

        Dim i As Integer = chkLstBarriersLayers.SelectedIndex
        'Dim indexChecked As Integer

        If chkNaturalTF.CheckState = CheckState.Checked Then
            'If i <> -1 Then
            '    For Each indexChecked In chkLstBarriersLayers.CheckedIndices
            '        If indexChecked = i Then
            '            cmdSelectNaturalYN.Enabled = True

            '        End If
            '    Next
            'End If
            If chkBarrierPerm.Checked = True And ChkDBFOutput.Checked = True And (chkConnect.Checked = True Or chkAdvConnect.Checked = True) Then
                chkDCI.Enabled = True
            End If
        Else
            'cmdSelectNaturalYN.Enabled = False
            chkDCI.Enabled = False
        End If
    End Sub

    Private Sub cmdSelectNaturalYN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelectNaturalTF.Click
        ' This command... 
        ' creates a new list of barrier layers and associated IDS (BarrierIDObj)
        ' Passes that to a form where the associated ID can be changed.  The form
        ' then passes back the updated object. 

        Dim lBarrierIDs As New List(Of BarrierIDObj)

        If m_BarrierIDObj IsNot Nothing Then
            Using MyForm As New FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.frmChooseNaturalYN(m_sBarrierIDLayer, m_BarrierIDObj)
                If MyForm.Form_Initialize(m_app) Then
                    MyForm.ShowDialog()
                    m_BarrierIDObj = MyForm.m_BarrierIDObj4
                End If
            End Using
            ' update the lstboxes now
            'Me.PHabParamList = m_BarrierIDObj
            lstNaturalTFField.Items.Clear()
            lstNaturalTFField.Items.Add(m_BarrierIDObj.NaturalYNField)

            ' also need to update the list variable
            ' add it to the property, which will add it to the list
            lBarrierIDs.Add(m_BarrierIDObj)
            BarrierParamList = lBarrierIDs
        End If
    End Sub

    Private Sub ChkDBFOutput_EnabledChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkDBFOutput.EnabledChanged
        ' If the DBF Output Checkbox is disabled then we want all 
        ' other options in this group disabled. 
        ' If it is enables then there are additional checks for a few
        ' before enabling them.  
        If ChkDBFOutput.Enabled = False Then
            cmdAddGDB.Enabled = False
            txtTablesPrefix.Enabled = False
            chkBarrierPerm.Enabled = False
            chkNaturalTF.Enabled = False
            chkConnect.Enabled = False
            chkAdvConnect.Enabled = False
            chkDCI.Enabled = False
        Else
            If ChkDBFOutput.Checked = True Then
                cmdAddGDB.Enabled = True
                txtTablesPrefix.Enabled = True
                chkConnect.Enabled = True
                chkAdvConnect.Enabled = False
                chkBarrierPerm.Enabled = True
                chkNaturalTF.Enabled = True
                chkDCI.Enabled = True
            End If

        End If
    End Sub

    Private Sub chkConnect_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkConnect.CheckedChanged

        If chkConnect.CheckState = CheckState.Checked Then
            If chkBarrierPerm.Checked = True And ChkDBFOutput.Checked = True And chkNaturalTF.Checked = True Then
                chkDCI.Enabled = True
            End If
        Else
            chkDCI.Enabled = False
        End If

    End Sub

    'Private Sub chkDCI_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    'If chkDCI.CheckState = CheckState.Checked Then
    'cmdDCIModelDir.Enabled = True
    'txtDCIModelDir.Enabled = True
    ' cmdRInstallDir.Enabled = True
    '  txtRInstallDir.Enabled = True
    '   chkDCISectional.Enabled = True
    'Else
    '   cmdDCIModelDir.Enabled = False
    '    txtDCIModelDir.Enabled = False
    '     cmdRInstallDir.Enabled = False
    '      txtRInstallDir.Enabled = False
    '       chkDCISectional.Enabled = False
    '    End If

    ' End Sub

    ' Private Sub chkDCI_EnabledChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    MsgBox("chkDCI triggered")
    '    If chkDCI.Enabled = True And chkDCI.Checked = True Then
    '        cmdDCIModelDir.Enabled = True
    '        txtDCIModelDir.Enabled = True
    '        cmdRInstallDir.Enabled = True
    '        txtRInstallDir.Enabled = True
    '        chkDCISectional.Enabled = True
    '    Else
    '        cmdDCIModelDir.Enabled = False
    '        txtDCIModelDir.Enabled = False
    '        cmdRInstallDir.Enabled = False
    '        txtRInstallDir.Enabled = False
    '        chkDCISectional.Enabled = False
    '    End If
    'End Sub

    Public Function FileWriteDeleteCheck(ByVal sDCIOutputDir As String) As Boolean

        Dim FILE_NAME As String = "FiPExPermTEST1.txt"
        If File.Exists(sDCIOutputDir + "\" + FILE_NAME) Then
            MsgBox("tempmsg: this is the file name tested: " + sDCIOutputDir + FILE_NAME)
            MsgBox("test file already exists in DCI output directory")
        End If

        Try
            Dim path As String = sDCIOutputDir + "\" + FILE_NAME
            Dim sw As StreamWriter = File.CreateText(path)
            sw.Close()

            ' Ensure that the target does not exist.
            File.Delete(path)

            Return True

        Catch e As Exception
            MsgBox("The following exception was found: " & e.Message)
            Return False
        End Try

    End Function
    Private Sub cmdRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRun.Click
        Dim canceleverything As Boolean = False
        saveoptions(canceleverything)
        If canceleverything = False Then
            m_bRun = True
            Me.Close()
        Else
            Exit Sub
        End If
        

    End Sub

    Private Sub chkAdvConnect_CheckedChanged(sender As Object, e As EventArgs) Handles chkAdvConnect.CheckedChanged
        ' to enable DCI, users must click either this Advanced Connectivity box or Basic Connectivity table box
        ' they also must have checked barrierperm, dbf out tables, natural tf fields
        If chkAdvConnect.CheckState = CheckState.Checked Then
            If chkBarrierPerm.Checked = True And ChkDBFOutput.Checked = True And chkNaturalTF.Checked = True Then
                If chkDCI.Enabled = False Then
                    chkDCI.Enabled = True
                End If
            End If
        Else
            If chkConnect.CheckState = False Then
                chkDCI.Enabled = False
            End If
        End If
    End Sub

    Private Sub chkDCI_CheckedChanged_1(sender As Object, e As EventArgs) Handles chkDCI.CheckedChanged
        'MsgBox("chkDCI checked changed 1 triggered")
        If chkDCI.CheckState = CheckState.Checked Then
            cmdDCIModelDir.Enabled = True
            txtDCIModelDir.Enabled = True
            cmdRInstallDir.Enabled = True
            txtRInstallDir.Enabled = True
            chkDCISectional.Enabled = True
            chkDistanceLimit.Enabled = True

            rdoHabLength.Enabled = True
            rdoHabArea.Enabled = True

            If rdoNatExp1.Checked = True Then
                PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.NatExp1DD
            End If
            If rdoCircle.Checked = True Then
                PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.CircleDD
            End If
            If rdoSigmoid.Checked = True Then
                PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.SigmoidDD
            End If
            If rdoLinear.Checked = True Then
                PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.LinearDD
            End If
        Else
            cmdDCIModelDir.Enabled = False
            txtDCIModelDir.Enabled = False
            cmdRInstallDir.Enabled = False
            txtRInstallDir.Enabled = False
            chkDCISectional.Enabled = False
            chkDistanceLimit.Enabled = False
            chkDistanceDecay.Enabled = False

            rdoHabLength.Enabled = False
            rdoHabArea.Enabled = False

            rdoLinear.Enabled = False
            rdoNatExp1.Enabled = False
            rdoCircle.Enabled = False
            rdoSigmoid.Enabled = False
            PictureBox7.Image = Nothing
        End If
    End Sub

    Private Sub chkDCI_EnabledChanged1(sender As Object, e As EventArgs) Handles chkDCI.EnabledChanged
        'MsgBox("chkDCI enabled changed triggered")
        If chkDCI.Enabled = True And chkDCI.Checked = True Then
            cmdDCIModelDir.Enabled = True
            txtDCIModelDir.Enabled = True
            cmdRInstallDir.Enabled = True
            txtRInstallDir.Enabled = True
            chkDCISectional.Enabled = True
            chkDistanceLimit.Enabled = True

            rdoHabLength.Enabled = True
            rdoHabArea.Enabled = True
        Else
            cmdDCIModelDir.Enabled = False
            txtDCIModelDir.Enabled = False
            cmdRInstallDir.Enabled = False
            txtRInstallDir.Enabled = False
            chkDCISectional.Enabled = False
            chkDistanceLimit.Enabled = False
            chkDistanceDecay.Enabled = False

            rdoHabLength.Enabled = False
            rdoHabArea.Enabled = False

            rdoLinear.Enabled = False
            rdoCircle.Enabled = False
            rdoSigmoid.Enabled = False
            rdoNatExp1.Enabled = False

        End If
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub cmdRInstallDir_Click_1(sender As Object, e As EventArgs) Handles cmdRInstallDir.Click
        ' This button will ask the user to browse to the directory where the R 
        ' Program files are installed.  It will check for the /bin/rterm.exe
        ' program.  

        Dim fdlg As FolderBrowserDialog = New FolderBrowserDialog
        fdlg.Description = "Browse to 'R' Statistical Software Installation Dir"
        fdlg.RootFolder = System.Environment.SpecialFolder.MyComputer

        If fdlg.ShowDialog = Windows.Forms.DialogResult.OK Then

            ' determine if requisite DCI model file(s) exist in this directory
            Dim sFileName As String
            sFileName = fdlg.SelectedPath + "/bin/rterm.exe"
            Dim fFile As New FileInfo(sFileName)

            If Not fFile.Exists Then
                If Not fFile.Exists Then
                    Dim fFile2 As New FileInfo(fdlg.SelectedPath + "/bin/i386/rterm.exe")
                    If Not fFile2.Exists Then
                        MessageBox.Show("'R' Program files not found in R installation directory provided ('Advanced Tab').  Please browse to the correct directory or install R.", "'R' Program Files Missing", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                        Exit Sub
                    End If
                End If
            End If

            txtRInstallDir.Text = fdlg.SelectedPath
        End If
    End Sub

    Private Sub cmdDCIModelDir_Click_1(sender As Object, e As EventArgs) Handles cmdDCIModelDir.Click
        ' This button will ask the user to browse to the directory where the DCI
        ' model functions are, check the directory for the presence of the model
        ' functions, and ensure the user has proper permissions to create files
        ' in the directory. 

        Dim fdlg As FolderBrowserDialog = New FolderBrowserDialog
        fdlg.Description = "Browse to DCI Model Installation Directory"
        fdlg.RootFolder = System.Environment.SpecialFolder.MyComputer

        If fdlg.ShowDialog = Windows.Forms.DialogResult.OK Then

            ' determine if requisite DCI model file(s) exist in this directory
            Dim sFileName As String
            sFileName = fdlg.SelectedPath + "/FIPEX.run.DCI.r"
            Dim fFile As New FileInfo(sFileName)

            If Not fFile.Exists Then
                MessageBox.Show("Directory does not contain required 'FIPEX.run.DCI.r' file")
                Exit Sub
            End If


            ' Check that the user currently has file permissions to write to 
            ' this directory
            Dim bPermissionCheck
            bPermissionCheck = FileWriteDeleteCheck(fdlg.SelectedPath)
            If bPermissionCheck = False Then
                MsgBox("File / folder permission check: " & Str(bPermissionCheck))
                MsgBox("It appears you do not have write permission to the DCI Model Directory.  Write permission to this directory is needed in order to run DCI Analysis.")
                txtDCIModelDir.Text = ""
                Exit Sub
            End If

            ' Check that the user currently has file permissions to write to 
            ' this directory
            'Dim pc As New CheckPerm
            'pc.Permission = "Modify"
            'If Not pc.CheckPerm(fdlg.SelectedPath) Then
            '    MsgBox("You don't currently have permission to write files to that directory.")
            '    Exit Sub
            'End If

            ' Set the path in the text dialogue to save to extension stream
            txtDCIModelDir.Text = fdlg.SelectedPath
        End If
    End Sub

    Private Sub chkDistanceLimit_CheckedChanged(sender As Object, e As EventArgs) Handles chkDistanceLimit.CheckedChanged
        'MsgBox("chkDCI checked changed 1 triggered")
        If chkDistanceLimit.CheckState = CheckState.Checked Then
            txtMaxDistance.Enabled = True


            Dim dMaxdistance As Double = 0.0
            Try
                dMaxdistance = Convert.ToDouble(txtMaxDistance.Text)
            Catch ex As Exception
                dMaxdistance = 0.0
            End Try

            If dMaxdistance > 0 And chkDistanceLimit.Checked = True Then
                chkDistanceDecay.Enabled = True
            Else
                chkDistanceDecay.Enabled = False
            End If

            If dMaxdistance > 0 And chkDistanceDecay.Checked = True And chkDistanceDecay.Enabled = True Then
                rdoLinear.Enabled = True
                rdoNatExp1.Enabled = True
                rdoCircle.Enabled = True
                rdoSigmoid.Enabled = True
                If rdoNatExp1.Checked = True Then
                    PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.NatExp1DD
                End If
                If rdoCircle.Checked = True Then
                    PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.CircleDD
                End If
                If rdoSigmoid.Checked = True Then
                    PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.SigmoidDD
                End If
                If rdoLinear.Checked = True Then
                    PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.LinearDD
                End If
            Else
                rdoLinear.Enabled = False
                rdoNatExp1.Enabled = False
                rdoCircle.Enabled = False
                rdoSigmoid.Enabled = False
                PictureBox7.Image = Nothing

            End If




        Else
            txtMaxDistance.Enabled = False
            chkDistanceDecay.Enabled = False
            rdoLinear.Enabled = False
            rdoNatExp1.Enabled = False
            rdoCircle.Enabled = False
            rdoSigmoid.Enabled = False
            PictureBox7.Image = Nothing

        End If
    End Sub

    Private Sub chkDistanceLimit_EnabledChanged(sender As Object, e As EventArgs) Handles chkDistanceLimit.EnabledChanged
        If chkDistanceLimit.Enabled = True And chkDistanceLimit.Checked = True Then
            txtMaxDistance.Enabled = True
            chkDistanceDecay.Enabled = True

        Else
            txtMaxDistance.Enabled = False
            chkDistanceDecay.Enabled = False
            rdoLinear.Enabled = False
            rdoCircle.Enabled = False
            rdoNatExp1.Enabled = False
            rdoSigmoid.Enabled = False
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles rdoLinear.CheckedChanged
        If rdoLinear.Checked = True Then
            PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.LinearDD
        End If
    End Sub

    Private Sub chkDistanceDecay_CheckedChanged_1(sender As Object, e As EventArgs) Handles chkDistanceDecay.CheckedChanged
        If chkDistanceDecay.Checked = True Then
            rdoCircle.Enabled = True
            rdoLinear.Enabled = True
            rdoSigmoid.Enabled = True
            rdoNatExp1.Enabled = True

            ' ensure at least one radio button is checked (default to linear)
            If rdoLinear.Checked = False And
               rdoCircle.Checked = False And
               rdoNatExp1.Checked = False And
               rdoSigmoid.Checked = False Then
                rdoLinear.Checked = True
            End If

            If rdoNatExp1.Checked = True Then
                PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.NatExp1DD
            End If
            If rdoCircle.Checked = True Then
                PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.CircleDD
            End If
            If rdoSigmoid.Checked = True Then
                PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.SigmoidDD
            End If
            If rdoLinear.Checked = True Then
                PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.LinearDD
            End If
        Else
            rdoCircle.Enabled = False
            rdoLinear.Enabled = False
            rdoSigmoid.Enabled = False
            rdoNatExp1.Enabled = False
            PictureBox7.Image = Nothing
        End If
    End Sub

    Private Sub chkDistanceDecay_EnabledChanged(sender As Object, e As EventArgs) Handles chkDistanceDecay.EnabledChanged
        ' max sure at least one radio box is checked by default if the distance decay function is enabled
        If chkDistanceDecay.Enabled = True Then
            If rdoLinear.Checked = False And
                rdoCircle.Checked = False And
                rdoNatExp1.Checked = False And
                rdoSigmoid.Checked = False Then
                rdoLinear.Checked = True
            End If
        Else
            rdoLinear.Checked = False
            rdoCircle.Checked = False
            rdoNatExp1.Checked = False
            rdoSigmoid.Checked = False
            rdoLinear.Checked = False
        End If
    End Sub

    Private Sub rdoNatExp1_CheckedChanged(sender As Object, e As EventArgs) Handles rdoNatExp1.CheckedChanged
        If rdoNatExp1.Checked = True Then
            PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.NatExp1DD
        End If
    End Sub

    Private Sub rdoCircle_CheckedChanged(sender As Object, e As EventArgs) Handles rdoCircle.CheckedChanged
        If rdoCircle.Checked = True Then
            PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.CircleDD
        End If
    End Sub

    Private Sub rdoSigmoid_CheckedChanged(sender As Object, e As EventArgs) Handles rdoSigmoid.CheckedChanged
        If rdoSigmoid.Checked = True Then
            PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.SigmoidDD
        End If
    End Sub

    Private Sub txtMaxDistance_TextChanged(sender As Object, e As EventArgs) Handles txtMaxDistance.TextChanged
        Dim dMaxdistance As Double = 0.0
        Try
            dMaxdistance = Convert.ToDouble(txtMaxDistance.Text)
        Catch ex As Exception
            dMaxdistance = 0.0
        End Try


        If dMaxdistance > 0 And chkDistanceLimit.Checked = True Then
            chkDistanceDecay.Enabled = True
        Else
            chkDistanceDecay.Enabled = False
        End If


        If dMaxdistance > 0 And chkDistanceDecay.Checked = True And chkDistanceDecay.Enabled = True Then
            rdoLinear.Enabled = True
            rdoNatExp1.Enabled = True
            rdoCircle.Enabled = True
            rdoSigmoid.Enabled = True
            If rdoNatExp1.Checked = True Then
                PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.NatExp1DD
            End If
            If rdoCircle.Checked = True Then
                PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.CircleDD
            End If
            If rdoSigmoid.Checked = True Then
                PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.SigmoidDD
            End If
            If rdoLinear.Checked = True Then
                PictureBox7.Image = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.My.Resources.Resources.LinearDD
            End If
        Else
            rdoLinear.Enabled = False
            rdoNatExp1.Enabled = False
            rdoCircle.Enabled = False
            rdoSigmoid.Enabled = False
            PictureBox7.Image = Nothing

        End If

    End Sub

 
End Class