'Created: Apr, 2021 by G Oldford
'Purpose: add and populate attributes containing segment and sub-segment labels
'To do: 
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.ArcMap
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.NetworkAnalysis
Imports ESRI.ArcGIS.esriSystem


Public Class frmLabelHypergraph
    Private m_app As ESRI.ArcGIS.Framework.IApplication
    Private pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Private m_FiPEx As FishPassageExtension
    Private m_UtilityNetworkAnalysisExt As ESRI.ArcGIS.EditorExt.IUtilityNetworkAnalysisExt
    Private m_pFWorkspace As IFeatureWorkspace

    ' below needed? -GLO 2021
    Private m_lEIDs As List(Of Integer)
    Private m_lLineLayersFields As List(Of LayersAndFCIDAndCumulativePassField) = New List(Of LayersAndFCIDAndCumulativePassField)
    Private m_lPolyLayersFields As List(Of LayersAndFCIDAndCumulativePassField) = New List(Of LayersAndFCIDAndCumulativePassField)
    Private m_lSinks As List(Of Integer) = New List(Of Integer)

    ' for saving / restoring original network
    Private m_pOriginalBarriersListSaved As IEnumNetEID
    Private m_pOriginalEdgeFlagsList As IEnumNetEID
    Private m_pOriginaljuncFlagsList As IEnumNetEID

    Private Property bFoundEID As Boolean


    Public Function Form_Initialize(ByVal m_application As ESRI.ArcGIS.Framework.IApplication) As Boolean
        Try
            m_app = m_application
            pMxDoc = CType(m_app.Document, ESRI.ArcGIS.ArcMapUI.IMxDocument)
            m_FiPEx = FishPassageExtension.GetExtension

            m_UtilityNetworkAnalysisExt = FishPassageExtension.GetUNAExt

            Return True

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.ToString, "Initialize Form")
            Return False
        End Try
    End Function
    Private Sub cmdRunLabel_Click(sender As Object, e As EventArgs) Handles cmdRunLabel.Click
        ' 1/ expose EIDs for nodes as attribute
        ' 2/ label segmetns and subsegments

        ' ================================== Pre-check network =====================================
        Dim bContinue As Boolean = True

        ' removed check for extension state - should already by set and checked 
        ' via FishPassageExtension.vb - GO Apr 2021
        ' check for FIPEX extension issues
        Check1_Network(bContinue)
        'If bContinue = False Then Exit Sub

        ' check for issues with networks loaded in ArcMap
        Check2_Network(bContinue)
        If bContinue = False Then Exit Sub

        ' check for issues with flags or barriers
        Check3_Network(bContinue)
        If bContinue = False Then Exit Sub

        ' check editing sessions already
        Check4_Network(bContinue)
        If bContinue = False Then Exit Sub

        ' check for issues and save original flags & barriers
        Prep1_SaveNetworkBarriersFlags(bContinue)
        If bContinue = False Then Exit Sub

        ' ensure all layers in TOC are selectable
        Prep2_CheckMapTOC(bContinue)
        If bContinue = False Then Exit Sub

        ' Left off here Apr 28. TO DO 
        ' (- add selection set save for editing att. tables to expose node EIDs)
        Prep3_FindBranchAndSourceJunctions(bContinue)
        If bContinue = False Then Exit Sub

       
        ' Label Segments - use SelectAndUpdateFeaturesObject (see frmVisualizeDecisionsAndNet.vb)
        ' upstream trace from all barriers / sinks one at a time (ignore branch nodes)

        ' Label Sub-Segments
        ' downstream trace from all branches, source, barrier nodes, one at a time

    End Sub
    Private Sub Check1_Network(ByRef bContinue As Boolean)
        ' removed check for extension state - should already by set and checked 
        ' via FishPassageExtension.vb - GO Apr 2021
        ' run check 1a: is FIPEX extension active? If no, exit
        ' run check 1b: is FIPEX extension loaded and user settings accesisble? If no, exit.
        lblCheck1.Text = "Checking for FIPEX... success."

    End Sub
    Private Sub Check2_Network(ByRef bContinue As Boolean)
        ' run check 2a: is there a network? If no exit.
        ' run check 2b: is there more than one network? If yes, exit.
        lblCheck2.Text = "Searching for networks..."

        ' =============== Check for network =====================
        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        If m_UtilityNetworkAnalysisExt Is Nothing Then
            MsgBox("A Network must be loaded in ArcMap to use this tool.  Exiting.")
            bContinue = False
            lblCheck2.Text = "Searching for networks... no UNA Extension found."
            Exit Sub
        Else
            pNetworkAnalysisExt = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExt)
            If pNetworkAnalysisExt.NetworkCount = 0 Then
                MsgBox("A Network must be loaded in ArcMap to use this tool.  Exiting.")
                bContinue = False
                lblCheck2.Text = "Searching for networks... no geometric networks."
                Exit Sub
            End If
        End If
        Dim pGeometricNetwork As IGeometricNetwork
        Try
            pGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        Catch ex As Exception
            MsgBox("Trouble getting current geometric network.  Must have an active geometric network in ArcMap to " &
                   "use this toolset. Exiting.")
            bContinue = False
            lblCheck2.Text = "Searching for networks... cannot access network object."
            Exit Sub
        End Try

        lblCheck2.Text = "Searching for networks... success."


    End Sub
    Private Sub Check3_Network(ByRef bContinue As Boolean)

        lblCheck3.Text = "Scanning network..."

        ' check flags
        Dim pNetworkAnalysisExtFlags As INetworkAnalysisExtFlags
        Dim iFlagNumber As Integer ' for junction flags
        'Dim eFlagNumber As Integer ' for edge flags

        pNetworkAnalysisExtFlags = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtFlags)
        iFlagNumber = pNetworkAnalysisExtFlags.JunctionFlagCount
        If iFlagNumber = 0 Then
            ' exit
            Dim result = Windows.Forms.MessageBox.Show("At least on flag must be set. Exiting.")
            bContinue = False
            lblCheck3.Text = "Scanning network... no flags."
            Exit Sub
        ElseIf iFlagNumber > 1 Then
            ' warning
            Dim result = Windows.Forms.MessageBox.Show("FIPEX detects " & iFlagNumber & " flags are currently set. Running analysis " &
                                                       "with multiple flags may be slow and if networks upstream of flags overlap, " &
                                                       "labels will be overwritten. Continue?", "continue", _
                                                       Windows.Forms.MessageBoxButtons.YesNo)
            If result = Windows.Forms.DialogResult.No Then
                bContinue = False
                lblCheck3.Text = "Scanning network... > 1 flag found."
                Exit Sub
            End If
        End If

        ' check barriers
        Dim pNetworkAnalysisExtBarriers As INetworkAnalysisExtBarriers
        pNetworkAnalysisExtBarriers = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtBarriers)
        Dim iJunctionBarrierCount As Integer = pNetworkAnalysisExtBarriers.JunctionBarrierCount
        Dim iEdgeBarrierCount As Integer = pNetworkAnalysisExtBarriers.EdgeBarrierCount

        If iEdgeBarrierCount > 0 Then
            Dim result = Windows.Forms.MessageBox.Show("FIPEX detects " & iEdgeBarrierCount & " barriers are currently set on a " &
                                                       "network EDGE (versus junction). FIPEX can only use junction barriers and " &
                                                       "will ignore edge barriers. Continue?", "continue", _
                                                       Windows.Forms.MessageBoxButtons.YesNo)
            If result = Windows.Forms.DialogResult.No Then
                bContinue = False
                lblCheck3.Text = "Scanning network... edge barriers found."
                Exit Sub
            End If
        End If

        If iJunctionBarrierCount = 0 Then
            Dim result = Windows.Forms.MessageBox.Show("FIPEX detects " & iJunctionBarrierCount & " barriers are currently set. " &
                                                       "Run analysis with no barriers set?", "continue", _
                                                       Windows.Forms.MessageBoxButtons.YesNo)
            If result = Windows.Forms.DialogResult.No Then
                bContinue = False
                lblCheck3.Text = "Scanning network... no barriers."
                Exit Sub
            End If
        End If

        lblCheck3.Text = "Scanning network... success."

    End Sub
    Private Sub Check4_Network(ByRef bContinue As Boolean)

        lblCheck4.Text = "Checking for edit sessions..."

        Dim pEditor As ESRI.ArcGIS.Editor.IEditor
        Dim pID As New UID

        ' get reference to editor extension
        Try
            pID.Value = "{F8842F20-BB23-11D0-802B-0000F8037368}"
            pEditor = m_app.FindExtensionByCLSID(pID)
            If pEditor Is Nothing Then
                MsgBox("Error getting reference to the Editor extension. Exiting. ")
                bContinue = False
                lblCheck4.Text = "Checking for edit sessions...could not find Editor extension"
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Error getting reference to the Editor extension. Exiting. " + ex.Message)
            bContinue = False
            lblCheck4.Text = "Checking for edit sessions...could not find Editor extension"
            Exit Sub
        End Try

        ' start one edit session for this refined list
        If pEditor.EditState = ESRI.ArcGIS.Editor.esriEditState.esriStateEditing Then
            MsgBox("FIPEX detects an open editing session. Please save work, close the editing session, and try again.")
            lblCheck4.Text = "Checking for edit sessions... found open editing session."
            bContinue = False
            Exit Sub
        End If

        lblCheck4.Text = "Checking for edit sessions...no active sessions"

    End Sub
    Private Sub Prep1_SaveNetworkBarriersFlags(ByRef bContinue As Boolean)
        ' copied from Runanalysis Apr 28, 2021 - GO
        lblPrep1.Text = "Checking flags and barriers..."

        ' 1/ save original barriers and flags to global objects
        ' 2/ check for edge flags and barriers (not supported)
        ' 3/ check flag / barrier consistency - all on junctions

        ' ============================ SETUP =====================================
        Dim pNetworkAnalysisExtFlags As INetworkAnalysisExtFlags
        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        Dim pNetwork As INetwork
        Dim pNetElements As INetElements
        Dim pTraceTasks As ITraceTasks
        Dim pNetworkAnalysisExtResults As INetworkAnalysisExtResults
        Dim pNetworkAnalysisExtBarriers As INetworkAnalysisExtBarriers
        Dim pOriginalBarriersListGEN As IEnumNetEIDBuilderGEN
        Dim pOriginalEdgeFlagsListGEN As IEnumNetEIDBuilderGEN
        Dim pOriginaljuncFlagsListGEN As IEnumNetEIDBuilderGEN
        Dim pOriginalBarriersList As IEnumNetEID

        Dim pOriginalBarriersListSavedGEN As IEnumNetEIDBuilderGEN
        Dim iOriginalJunctionBarrierCount As Integer
        Dim bFlagDisplay As IFlagDisplay
        Dim bEID, i As Integer
        pNetworkAnalysisExt = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExt)
        Dim pGeometricNetwork As IGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        pGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        pNetwork = pGeometricNetwork.Network
        pNetElements = CType(pNetwork, INetElements)
        pTraceTasks = CType(m_UtilityNetworkAnalysisExt, ITraceTasks)
        pNetworkAnalysisExtResults = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtResults)

        ' clear any leftover results from previous calls to the cmd
        pNetworkAnalysisExtResults.ClearResults()

        ' QI the Flags and barriers
        pNetworkAnalysisExtFlags = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtFlags)
        pNetworkAnalysisExtBarriers = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtBarriers)
        pOriginalBarriersListGEN = New EnumNetEIDArray
        pOriginalEdgeFlagsListGEN = New EnumNetEIDArray
        pOriginaljuncFlagsListGEN = New EnumNetEIDArray

        ' 2020 note - these saved' objects are for the 'advanced connectivity' 
        '             'saved' is so if user selects advanced connectivity table 
        '              this object will not be modified and can be user to restor

        pOriginalBarriersListSavedGEN = New EnumNetEIDArray
        iOriginalJunctionBarrierCount = pNetworkAnalysisExtBarriers.JunctionBarrierCount

        ' ============================ SAVE ORIGINAL BARRIERS =================================
        For i = 0 To pNetworkAnalysisExtBarriers.JunctionBarrierCount - 1
            ' Use bFlagDisplay to retrieve EIDs of the barriers for later
            bFlagDisplay = CType(pNetworkAnalysisExtBarriers.JunctionBarrier(i), IFlagDisplay)
            bEID = pNetElements.GetEID(bFlagDisplay.FeatureClassID, bFlagDisplay.FID, bFlagDisplay.SubID, esriElementType.esriETJunction)
            pOriginalBarriersListGEN.Add(bEID)
            pOriginalBarriersListSavedGEN.Add(bEID)
        Next

        ' QI to and get an array object that has 'count' and 'next' methods
        pOriginalBarriersList = CType(pOriginalBarriersListGEN, IEnumNetEID)
        ' 2020: for branch connectivity feature - GO
        ' save the original list so it can be restored later
        m_pOriginalBarriersListSaved = CType(pOriginalBarriersListSavedGEN, IEnumNetEID)

        ' ============================ SAVE ORIGINAL FLAGS =================================
        i = 0
        For i = 0 To pNetworkAnalysisExtFlags.JunctionFlagCount - 1
            ' Use the bFlagDisplay to retrieve the EIDs of the junction flags
            bFlagDisplay = CType(pNetworkAnalysisExtFlags.JunctionFlag(i), IFlagDisplay)
            bEID = pNetElements.GetEID(bFlagDisplay.FeatureClassID, bFlagDisplay.FID, bFlagDisplay.SubID, esriElementType.esriETJunction)
            pOriginaljuncFlagsListGEN.Add(bEID)
        Next

        ' QI to and get an array interface that has 'count' and 'next' methods
        m_pOriginaljuncFlagsList = CType(pOriginaljuncFlagsListGEN, IEnumNetEID)

        ' ============================ CHECK FOR EDGE FLAGS =================================
        i = 0
        For i = 0 To pNetworkAnalysisExtFlags.EdgeFlagCount - 1
            ' Use the bFlagDisplay to retrieve EIDs of the Edge flags for later
            bFlagDisplay = CType(pNetworkAnalysisExtFlags.EdgeFlag(i), IFlagDisplay)
            bEID = pNetElements.GetEID(bFlagDisplay.FeatureClassID, bFlagDisplay.FID, bFlagDisplay.SubID, esriElementType.esriETEdge)
            pOriginalEdgeFlagsListGEN.Add(bEID)
        Next

        ' QI to and get an array interface that has 'count' and 'next' methods
        m_pOriginalEdgeFlagsList = CType(pOriginalEdgeFlagsListGEN, IEnumNetEID)

        ' If there are no flags set exit sub
        If pNetworkAnalysisExtFlags.JunctionFlagCount = 0 Then
            MsgBox("There are no flags set on junctions.  Please Set flags only on network junctions.")
            lblPrep1.Text = "Checking flags and barriers... no junction flags found."
            bContinue = False
            Exit Sub
        End If

        ' ============================ FLAG CONSISTENCY CHECK ==============================
        ' Check for consistency that flags are all on barriers or all on non-barriers.
        Dim sFlagCheck As String
        sFlagCheck = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.SharedSubs.flagcheck2021(m_pOriginalBarriersListSaved, m_pOriginalEdgeFlagsList, m_pOriginaljuncFlagsList)

        If sFlagCheck = "error" Then
            MsgBox("Inconsistent flags - some are on edges.  Please Set flags only on network junctions.")
            lblPrep1.Text = "Checking flags and barriers... flags inconsistent (set on both edges and junctions)."
            bContinue = False
            Exit Sub
        End If
        ' ============================ END FLAG CONSISTENCY CHECK ============================

        lblPrep1.Text = "Checking flags and barriers... success."

    End Sub
    Private Sub Prep2_CheckMapTOC(ByRef bContinue As Boolean)

        lblPrep2.Text = "Checking TOC..."

        Dim pMxDocument As IMxDocument = CType(My.ArcMap.Application.Document, IMxDocument)
        Dim pMap As IMap = pMxDocument.FocusMap
        Dim pFLyrSlct As IFeatureLayer
        Dim i As Integer = 0

        ' Make sure all the layers in the TOC are selectable
        Try
            For i = 0 To pMap.LayerCount - 1
                If pMap.Layer(i).Valid = True Then
                    If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                        pFLyrSlct = CType(pMap.Layer(i), IFeatureLayer)
                        pFLyrSlct.Selectable = True
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox("Problem encountered setting all layers as selectable. Exiting. " & ex.Message)
            bContinue = False
            lblPrep2.Text = "Checking TOC... could not set all layers as selectable."
            Exit Sub
        End Try
      
        lblPrep2.Text = "Checking TOC... success."

    End Sub
    Private Sub Prep3_FindBranchAndSourceJunctions(ByRef bContinue As Boolean)
        lblPrep3.Text = "Searching for branch junctions and source nodes..."

        lblPrep3.Text = "Searching for branch junctions and source nodes... success"

    End Sub
End Class